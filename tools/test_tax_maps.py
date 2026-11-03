"""HCAD tax-map export, offline UI regression, and optional live print checks."""
from __future__ import annotations
import argparse
from copy import deepcopy
from dataclasses import asdict
import importlib
import json
from pathlib import Path
import tempfile
import time
import unittest
from unittest.mock import patch

import fitz
from shapely.geometry import Polygon, box
from shapely.affinity import translate

from core import tax_map as TM
from core.boundary_map import MapCase, MapProperty, application_boundary, map_case, parcel_polygon, read_map_case
from core.boundary_map_render import export_map
from core.paths import sample_input_dir

OUT = DATA = CASES = None
LIVE = UI = False
SAMPLES = {
    'mls': 'MLS_Calculation_Sheet_Case_Info_v13_FILLED_SAMPLE.xlsx',
    'mbl': 'MBL_Calculation_Sheet_Case_Info_v3_FILLED_SAMPLE.xlsx',
    'mls_renewal': 'MLS_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx',
    'mbl_renewal': 'MBL_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx',
}


def fixture():
    properties = [MapProperty(2, '0000000000001', '100 TEST ST', '100', 'TEST ST', 'SFR', '100'),
                  MapProperty(3, '0000000000002', '0 TEST ST', '', 'TEST ST', 'EXC', '')]
    case = MapCase('mls', 'MLS TEST', 'Special Minimum Lot Size', '100 block of Test Street',
                   'north side', 'Test Street', ['First Street', 'Second Street'], '5,000 Square Feet',
                   'E4', properties, 'test-workbook-hash')
    matches = []
    for i, prop in enumerate(properties):
        x, y = 3100000 + i*50, 13800000
        feature = {'attributes': {'HCAD_NUM': prop.account}, 'geometry': {'rings': [
            [[x,y], [x,y+100], [x+50,y+100], [x+50,y], [x,y]]]}}
        matches.append({'property': asdict(prop), 'feature': feature, 'match_method': 'account'})
    return {'case': asdict(case), 'matched': matches, 'context': [], 'roads': [], 'sources': [], 'warnings': ['Test warning: review retained parcel.']}


def fixture_pdf():
    with fitz.open() as doc:
        p = doc.new_page(width=792, height=612)
        p.insert_text((560, 55), 'Account Detail\nHCAD_Map\nTEST FIXTURE')
        p.draw_rect(fitz.Rect(100, 170, 460, 280), color=(1,1,0), width=3)
        return doc.tobytes()


class TaxMapTests(unittest.TestCase):
    def setUp(self):
        from unittest.mock import patch
        from tools.test_designation_lookup import no_lookup
        external=patch('core.designation_lookup.lookup',side_effect=no_lookup);external.start();self.addCleanup(external.stop)
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory()
        cls.folder = Path(OUT or cls.temp.name).resolve()
        cls.folder.mkdir(parents=True, exist_ok=True)
        cls.data = fixture()
        cls.pdf = fixture_pdf()
        cls.results = []

    @classmethod
    def tearDownClass(cls):
        (cls.folder/'results.json').write_text(json.dumps(cls.results, indent=2), encoding='utf-8')
        cls.temp.cleanup()

    def rendered(self, settings=None):
        client = TM.HCADPrintClient(self.folder/'fixture_cache')
        with patch.object(client, 'print_map', return_value=self.pdf):
            return TM.render_tax_map(self.data, settings or TM.default_settings(), client)

    def test_geometry_includes_excluded_and_unaddressed_properties(self):
        web = TM.build_web_map(self.data, TM.default_settings())
        layer = web['operationalLayers'][-1]['featureCollection']['layers'][0]
        geometry = layer['featureSet']['features'][0]['geometry']
        selected = parcel_polygon({'geometry': geometry})
        for match in self.data['matched']:
            self.assertTrue(selected.covers(parcel_polygon(match['feature'])))
        self.assertEqual(selected.area, 10000)
        symbol = layer['layerDefinition']['drawingInfo']['renderer']['symbol']
        self.assertEqual(symbol['outline']['color'], [255,255,0,255])
        self.assertEqual(symbol['style'], 'esriSFSNull')
        self.assertEqual(symbol['color'], [0,0,0,0])
        self.assertEqual(len(geometry['rings']), 1)
        encoded = json.dumps(web)
        self.assertNotIn('100 TEST ST', encoded)
        self.assertNotIn('test-workbook-hash', encoded)
        self.assertEqual(web['mapOptions']['rotation'], 0)

    def test_contiguous_boundary_bridges_streets_and_has_no_holes(self):
        data = deepcopy(self.data)
        # Opposite parcel rows separated by a street and a second block gap.
        data['matched'] = []
        for index, (dx, dy) in enumerate(((0, 0), (200, 0), (0, 160), (200, 160))):
            match = deepcopy(self.data['matched'][index % 2])
            poly = translate(parcel_polygon(self.data['matched'][0]['feature']), dx, dy)
            match['feature']['geometry'] = TM.boundary_geometry(poly)
            data['matched'].append(match)
        gap = box(3100080, 13800020, 3100120, 13800080)
        data['context'] = [{'attributes': {'HCAD_NUM': 'unlisted'},
                            'geometry': TM.boundary_geometry(gap)}]
        original = deepcopy(data)
        # Legacy settings must not reintroduce disconnected boundaries.
        for include_street in (False, True):
            settings = dict(TM.default_settings(), include_street=include_street)
            boundary, _ = TM.map_extent(data, settings)
            self.assertEqual(boundary.geom_type, 'Polygon')
            self.assertTrue(boundary.is_valid)
            self.assertEqual(len(boundary.interiors), 0)
            self.assertTrue(boundary.covers(gap))
            for match in data['matched']:
                self.assertTrue(boundary.covers(parcel_polygon(match['feature'])))
            web = TM.build_web_map(data, settings)
            geometry = web['operationalLayers'][-1]['featureCollection']['layers'][0]['featureSet']['features'][0]['geometry']
            self.assertEqual(len(geometry['rings']), 1)
        self.assertEqual(data, original, 'HCAD parcel shapes and case data must stay intact')

    def test_stepped_edges_are_preserved_and_interior_holes_removed(self):
        data = deepcopy(self.data)
        # Two street-separated rows; one block has deeper lots than the other.
        data['matched'] = []
        for index, bounds in enumerate(((0, 0, 100, 100), (150, -60, 250, 100),
                                         (0, 210, 100, 310), (150, 210, 250, 310))):
            match = deepcopy(self.data['matched'][index % 2])
            shape = translate(box(*bounds), 3100000, 13800000)
            match['feature']['geometry'] = TM.boundary_geometry(shape)
            data['matched'].append(match)
        boundary = TM.tax_map_boundary(data)
        self.assertEqual(boundary.geom_type, 'Polygon')
        self.assertEqual(len(boundary.interiors), 0)
        self.assertLess(boundary.area, application_boundary(data).convex_hull.area)
        for match in data['matched']:
            self.assertTrue(boundary.covers(parcel_polygon(match['feature'])))
        donut = box(0, 0, 300, 300).difference(box(100, 100, 200, 200))
        data['matched'] = [deepcopy(self.data['matched'][0])]
        data['matched'][0]['feature']['geometry'] = TM.boundary_geometry(donut)
        boundary = TM.tax_map_boundary(data)
        self.assertEqual(len(boundary.interiors), 0)
        self.assertEqual(boundary.area, 90000)

    def test_zoom_pan_and_invalid_framing(self):
        _, original = TM.map_extent(self.data, TM.default_settings())
        settings = dict(TM.default_settings(), zoom=2, pan_x=40, pan_y=60)
        _, changed = TM.map_extent(self.data, settings)
        self.assertAlmostEqual(changed[2]-changed[0], (original[2]-original[0])/2)
        self.assertAlmostEqual((changed[0]+changed[2])/2-(original[0]+original[2])/2, 40)
        self.assertAlmostEqual((changed[1]+changed[3])/2-(original[1]+original[3])/2, 60)
        for value in (0, 5, float('nan')):
            with self.assertRaises(TM.MapError):
                TM.map_extent(self.data, dict(settings, zoom=value))

    def test_warnings_do_not_block_save_or_download(self):
        rendered = self.rendered(dict(TM.default_settings(), pan_x=900))
        self.assertTrue(any('outside' in s for s in rendered['issues']))
        p1 = TM.save_tax_map(self.data, TM.default_settings(), rendered, self.folder/'saved')
        p2 = TM.save_tax_map(self.data, TM.default_settings(), rendered, self.folder/'saved')
        self.assertNotEqual(p1, p2)
        self.assertEqual(p1.read_bytes(), p2.read_bytes())
        downloaded = export_map(rendered, self.folder/'download.pdf')
        self.assertEqual(downloaded.read_bytes(), self.pdf)
        png = export_map(rendered, self.folder/'download.png')
        self.assertTrue(png.read_bytes().startswith(b'\x89PNG'))

    def test_print_cache_reuse_refresh_and_errors(self):
        web = TM.build_web_map(self.data, TM.default_settings())
        response = json.dumps({'results': [{'paramName':'Output_File', 'value':{'url':'https://arcweb.hcad.org/server/rest/directories/arcgisoutput/test.pdf'}}]}).encode()
        client = TM.HCADPrintClient(self.folder/'cache_check'/str(time.monotonic_ns()))
        with patch.object(client, 'request', side_effect=[response, self.pdf]) as request:
            self.assertEqual(client.print_map(web), self.pdf)
            self.assertEqual(request.call_count, 2)
            self.assertEqual(client.print_map(web), self.pdf)
            self.assertEqual(request.call_count, 2)
            self.assertTrue(client.cached)
        client.use_cache = False
        with patch.object(client, 'request', side_effect=[response, self.pdf]) as request:
            client.print_map(web)
            self.assertEqual(request.call_count, 2)
        for response in (b'not json', b'{"error":{"message":"Service unavailable"}}',
                         b'{"results":[{"paramName":"Output_File","value":{"url":"https://example.com/file.pdf"}}]}'):
            with patch.object(client, 'request', return_value=response), self.assertRaises(TM.MapError):
                client.print_map(web)

    def test_all_four_calculation_sheet_types_are_read_only(self):
        if not CASES:
            self.skipTest('Use --cases with four completed case workbooks; instructional placeholders are not map inputs')
        cases = json.loads(Path(CASES).read_text(encoding='utf-8'))
        self.assertTrue(set(SAMPLES).issubset({c['wizard'] for c in cases}))
        for entry in cases:
            key, source = entry['wizard'], Path(entry['workbook'])
            before = source.read_bytes()
            case = read_map_case(source, key)
            self.assertTrue(case.properties)
            self.assertEqual(case.wizard, key)
            self.assertEqual(source.read_bytes(), before)

    def test_live_hcad_prints(self):
        if not LIVE or not DATA:
            self.skipTest('Use --live --data <saved parcel data> to test HCAD')
        source = Path(DATA)
        files = [source] if source.is_file() else sorted(source.glob('*/parcel_data.json'))
        self.assertTrue(files)
        for source in files:
            with self.subTest(case=source.parent.name):
                data = json.loads(source.read_text(encoding='utf-8'))
                boundary = TM.tax_map_boundary(data)
                self.assertEqual(boundary.geom_type, 'Polygon')
                self.assertEqual(len(boundary.interiors), 0)
                self.assertEqual(len(TM.boundary_geometry(boundary)['rings']), 1)
                for match in data['matched']:
                    self.assertTrue(boundary.buffer(.001).covers(parcel_polygon(match['feature'])))
                folder = self.folder/'live'/source.parent.name
                folder.mkdir(parents=True, exist_ok=True)
                client = TM.HCADPrintClient(folder/'cache', progress=lambda s: print(s, flush=True))
                result = TM.render_tax_map(data, TM.default_settings(), client)
                path = TM.save_tax_map(data, TM.default_settings(), result, folder)
                with fitz.open(path) as doc:
                    self.assertEqual(len(doc), 1)
                    self.assertIn('HCAD_Map', doc[0].get_text())
                    self.assertTrue(any(d.get('color') and d['color'][0]>.9 and d['color'][1]>.9 and d['color'][2]<.1
                                        for d in doc[0].get_drawings()), 'No yellow boundary was printed')
                    doc[0].get_pixmap(dpi=115, alpha=False).save(folder/'page-1.png')
                self.results.append({'live_case':data['case']['reference'], 'pdf':str(path), 'properties':len(data['matched']), 'pages':1})

    def test_all_four_wizard_buttons_download_and_attachment_selection(self):
        if not UI:
            self.skipTest('Use --ui for the local Tk wizard checks')
        import tkinter as tk
        from core import session, tax_map_ui as UI_MAP
        names = {'mls':'MailoutWizard', 'mbl':'MBLWizard', 'mls_renewal':'MLSRenewalWizard', 'mbl_renewal':'MBLRenewalWizard'}
        root = tk.Tk()
        root.withdraw()
        def descendants(widget):
            for child in widget.winfo_children():
                yield child
                yield from descendants(child)
        def fail_dialog(title, message, **kwargs):
            raise AssertionError(f'{title}: {message}')
        def wait(preview):
            deadline = time.monotonic()+25
            root.update()
            while (preview.rendered is None or preview._loading) and time.monotonic()<deadline:
                root.update()
                time.sleep(.02)
            self.assertIsNotNone(preview.rendered, preview.review.get('1.0','end'))
            self.assertFalse(preview._loading)
        try:
            with patch.object(session,'SETTINGS_PATH',self.folder/'test_session.json'), patch.object(session,'load',return_value={}), \
                 patch('tkinter.messagebox.showerror',fail_dialog), \
                 patch.object(UI_MAP,'read_map_case',side_effect=lambda p,k,**kw:map_case(dict(self.data,case=dict(self.data['case'],wizard=k)))), \
                 patch.object(UI_MAP,'fetch_map_data',side_effect=lambda c,client:dict(deepcopy(self.data),case=asdict(c))), \
                 patch.object(TM.HCADPrintClient,'print_map',return_value=self.pdf):
                for key, name in names.items():
                    owner = None
                    with self.subTest(wizard=key):
                        try:
                            module = importlib.import_module(f'wizards.{key}.wizard')
                            owner = getattr(module,name)(root)
                            output = self.folder/'ui'/key
                            output.mkdir(parents=True,exist_ok=True)
                            workbook = sample_input_dir(key)/SAMPLES[key]
                            if hasattr(owner,'vars'):
                                owner.vars['excel'].set(str(workbook));owner.vars['output_folder'].set(str(output));owner.vars['map'].set('ORIGINAL_BOUNDARY.pdf')
                            else:
                                owner.files.update(workbook=str(workbook),out=str(output),map='ORIGINAL_BOUNDARY.pdf')
                            buttons = [w for w in descendants(owner) if 'text' in w.keys() and w.cget('text')=='Create Tax Map']
                            self.assertEqual(len(buttons),1)
                            buttons[0].invoke()
                            preview = owner._tax_map_preview
                            wait(preview)
                            self.assertEqual(str(preview.use['state']),'normal')
                            preview.pan(1,0);preview.zoom(.15)
                            self.assertEqual(str(preview.use['state']),'disabled')
                            preview.start_load();wait(preview)
                            preview.reset();wait(preview)
                            self.assertEqual(preview.rendered_settings,TM.default_settings())
                            self.assertFalse(hasattr(preview, 'include_street'))
                            geometry = preview.rendered['web_map']['operationalLayers'][-1]['featureCollection']['layers'][0]['featureSet']['features'][0]['geometry']
                            self.assertEqual(len(geometry['rings']), 1)
                            with patch.object(TM.HCADPrintClient,'print_map',side_effect=TM.MapError('Test outage')):
                                preview.start_load(refresh=True);wait(preview)
                            self.assertEqual(str(preview.use['state']),'normal')
                            self.assertIn('previous tax map',preview.status.get())
                            dest = output/'Downloaded_Tax_Map.pdf'
                            with patch('tkinter.filedialog.asksaveasfilename',return_value=str(dest)):
                                preview.download.invoke()
                            self.assertEqual(dest.read_bytes(),self.pdf)
                            preview.use.invoke()
                            selected = owner.vars['tax_map'].get() if hasattr(owner,'vars') else owner.files['tax_map']
                            original = owner.vars['map'].get() if hasattr(owner,'vars') else owner.files['map']
                            self.assertEqual(Path(selected).read_bytes(),self.pdf)
                            self.assertEqual(original,'ORIGINAL_BOUNDARY.pdf')
                            saved = json.loads((self.folder/'test_session.json').read_text())
                            self.assertEqual(saved[key]['tax_map'],selected)
                            self.results.append({'wizard_ui':key,'download_and_use':True,'boundary_map_preserved':True})
                        finally:
                            if owner is not None:
                                preview=getattr(owner,'_tax_map_preview',None)
                                if preview is not None and preview.winfo_exists():preview.close()
                                owner.destroy()
                            root.update()
        finally:
            root.destroy()


def main():
    global OUT, DATA, CASES, LIVE, UI
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--out');parser.add_argument('--data');parser.add_argument('--cases');parser.add_argument('--live',action='store_true');parser.add_argument('--ui',action='store_true')
    args=parser.parse_args()
    OUT, DATA, CASES, LIVE, UI = args.out,args.data,args.cases,args.live,args.ui
    result=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(TaxMapTests))
    return 0 if result.wasSuccessful() else 1


if __name__=='__main__':
    raise SystemExit(main())
