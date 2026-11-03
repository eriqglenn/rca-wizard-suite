"""Boundary-map regression checks. Synthetic workbooks; no user files modified.

Run: python tools/test_boundary_map.py
Optional --real-data DIRECTORY uses saved Crystalwood extracts for visual QA.
"""
from __future__ import annotations

from dataclasses import asdict
from contextlib import contextmanager
from pathlib import Path
import copy
import gc
import importlib
import io
import json
import os
import sys
import tempfile
import time
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from pypdf import PdfReader
from shapely.geometry import Polygon, box
from shapely.affinity import affine_transform, rotate

WF = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(WF))
from core import session
from core.boundary_map import (MAP_TYPES, MapError, PublicMapClient, account_number,
                               application_boundary, assert_map_current, fetch_map_data,
                               fingerprint, parcel_polygon, read_map_case, format_block_range, road_key,
                               ROAD_SERVICE, COUNTY_ROAD_SERVICE, load_case_roads)
from core.boundary_map_render import default_settings, footer_lines, render_map, save_map, pdf_png


def workbook(path, wizard='mls', props=None):
    props = props or [(12103, 'CRYSTALWOOD', '1047810000018', 'SFR'),
                      (12107, 'CRYSTALWOOD', '1047810000017', 'VAC')]
    w = Workbook()
    s1 = w.active
    s1.title = 'Step 1 - Input Properties Info'
    s1.append(['Str_num', 'Str_name', 'Str_sfx', 'Site_addr_1', 'ALT_TAX_ID', 'Property Description', 'Land Use'])
    for num, street, account, lu in props:
        s1.append([num, street, 'DR', f'{num} {street} DR', account, f'{num} (TEST LOT)', lu])
    s2 = w.create_sheet('Step 2 - Main Calculation Sheet')
    s2['A1'] = MAP_TYPES[wizard][0].upper()
    s2['B2'] = 999
    s2['B4'] = 'Crystalwood Drive'
    s2['D4'] = '12100 and 12200 blocks of Crystalwood Drive'
    s2['B5'], s2['D5'] = 'Cool Mist Drive', 'Woodcliff Drive'
    s2['B6'] = 'North'
    if wizard.endswith('_renewal'):
        s2['D2'] = 5000 if wizard == 'mls_renewal' else 15
    else:
        s2['E4'] = '5000-Square Foot' if wizard == 'mls' else '15-Foot'
        if wizard == 'mbl':
            s2['L16'] = 15
    row = 14 if wizard.endswith('_renewal') else 16
    s2.cell(row, 1, 'Property Description')
    s2.cell(row, 2, 'Land Use')
    for i, (num, street, account, lu) in enumerate(props, row + 1):
        s2.cell(i, 1, f'{num} (TEST LOT)')
        s2.cell(i, 2, lu)
    w.save(path)
    w.close()
    return path


def feature(polygon, account, number, street='CRYSTALWOOD'):
    rings = [list(polygon.exterior.coords)] + [list(r.coords) for r in polygon.interiors]
    return {'attributes': {'OBJECTID': number, 'HCAD_NUM': account, 'site_str_num': number,
                           'site_str_name': street, 'site_str_sfx': 'DR', 'tax_year': 2025},
            'geometry': {'rings': rings}}


def fixture(case):
    x, y = 3166000, 13851000
    matches = []
    for i, p in enumerate(case.properties):
        f = feature(box(x + 70*i, y, x + 70*(i+1), y+115), p.account, int(p.number))
        matches.append({'property': asdict(p), 'feature': f})
    length = 70*len(matches)
    roads = []
    for i, (name, coords) in enumerate([
        ('CRYSTALWOOD DR', [(x-300,y-30),(x+length+300,y-30)]),
        ('COOL MIST DR', [(x-30,y-300),(x-30,y+600)]),
        ('WOODCLIFF DR', [(x+length+30,y-300),(x+length+30,y+600)]),
    ]):
        roads.append({'attributes': {'OBJECTID': i+1, 'FULL_NAME': name}, 'geometry': {'paths': [coords]}})
    return {'case': asdict(case), 'matched': matches, 'context': [m['feature'] for m in matches],
            'roads': roads, 'extent': [x-400,y-400,x+length+400,y+700],
            'sources': [{'service': 'synthetic test fixture', 'retrieved_utc': '2026-09-08T00:00:00Z'}],
            'warnings': []}


class BoundaryMapTests(unittest.TestCase):
    def setUp(self):
        from tools.test_designation_lookup import no_lookup
        external=patch('core.designation_lookup.lookup',side_effect=no_lookup);external.start();self.addCleanup(external.stop)
        # Dispose of prior tests' Tk cycles on the main thread before a new
        # preview starts its background workbook reader.
        gc.collect()
        self.temp = tempfile.TemporaryDirectory(prefix='rca_boundary_test_')
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    @contextmanager
    def windows_save_handle(self, path, exclusive=False):
        """Hold the replacement access seen on the user's Excel workbook."""
        import ctypes
        from ctypes import wintypes
        kernel=ctypes.WinDLL('kernel32',use_last_error=True)
        kernel.CreateFileW.argtypes=[wintypes.LPCWSTR,wintypes.DWORD,wintypes.DWORD,ctypes.c_void_p,wintypes.DWORD,wintypes.DWORD,wintypes.HANDLE]
        kernel.CreateFileW.restype=wintypes.HANDLE
        kernel.CloseHandle.argtypes=[wintypes.HANDLE]
        handle=kernel.CreateFileW(str(path),0x10000,0 if exclusive else 7,None,3,0x80,None)
        if handle==ctypes.c_void_p(-1).value:raise ctypes.WinError(ctypes.get_last_error())
        try:yield
        finally:kernel.CloseHandle(handle)

    @unittest.skipUnless(os.name=='nt','Windows file sharing')
    def test_workbook_becoming_open_in_excel_invalidates_original_path_cache(self):
        from core.workbooks import WorkbookReadCache
        source=workbook(self.folder/'open.xlsx');before=source.read_bytes()
        cache=WorkbookReadCache();self.assertEqual(cache.resolve(source),str(source))
        with self.windows_save_handle(source):
            with self.assertRaises(PermissionError):source.read_bytes()
            resolved=Path(cache.resolve(source))
            self.assertNotEqual(resolved,source);self.assertEqual(resolved.read_bytes(),before)
            self.assertEqual(read_map_case(source,'mls').workbook_hash,fingerprint(resolved))
            self.assertEqual(cache.resolve(source),str(resolved))
        self.assertEqual(source.read_bytes(),before)

    @unittest.skipUnless(os.name=='nt','Windows file sharing')
    def test_exclusive_lock_cannot_reuse_an_older_cached_workbook(self):
        from core.workbooks import WorkbookReadCache
        source=workbook(self.folder/'exclusive.xlsx');cache=WorkbookReadCache();cache.resolve(source)
        with self.windows_save_handle(source,exclusive=True),patch('core.workbooks.time.sleep'):
            with self.assertRaisesRegex(PermissionError,'Finish saving in Excel'):
                cache.resolve(source)
            self.assertIsNone(cache._key)
        self.assertEqual(cache.resolve(source),str(source))

    @unittest.skipUnless(os.name=='nt','Windows file sharing')
    def test_all_four_maps_load_refresh_save_and_validate_with_excel_open(self):
        import tkinter as tk
        from core.boundary_map_ui import MapPreview
        names={'mls':'MailoutWizard','mbl':'MBLWizard','mls_renewal':'MLSRenewalWizard','mbl_renewal':'MBLRenewalWizard'}
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        with patch.object(session,'SETTINGS_PATH',self.folder/'settings.json'),patch.object(session,'load',return_value={}), \
                patch('core.boundary_map.fetch_map_data',side_effect=lambda case,client:fixture(case)), \
                patch('core.council_district.resolve_districts',side_effect=OSError('Council service offline')), \
                patch('tkinter.messagebox.showerror',side_effect=AssertionError):
            for key,name in names.items():
                owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),name)(root)
                source=workbook(self.folder/(key+'_held.xlsx'),key);before=source.read_bytes()
                if hasattr(owner,'vars'):
                    owner.vars['excel'].set(str(source));owner.vars['output_folder'].set(str(self.folder/key))
                else:owner.files.update(workbook=str(source),out=str(self.folder/key))
                with self.windows_save_handle(source):
                    preview=MapPreview(owner)
                    for refresh in (False,True):
                        if refresh:preview.refresh_data()
                        deadline=time.monotonic()+15
                        while preview.rendered is None and time.monotonic()<deadline:
                            root.update();time.sleep(.02)
                        self.assertIsNotNone(preview.rendered,preview.review.get('1.0','end'))
                        self.assertFalse(preview.rendered['blocking'])
                    preview.accept()
                    selected=owner.vars['map'].get() if hasattr(owner,'vars') else owner.files['map']
                    self.assertTrue(Path(selected).is_file());owner._check_boundary_map()
                self.assertEqual(source.read_bytes(),before)
                owner.destroy();root.update()

    def assert_addresses_inside_own_lots(self, data, result, overrides=None):
        addressed={f'parcel:{m["property"]["row"]}':m for m in data['matched']}
        labels=[l for l in result['labels'] if l['key'].startswith('parcel:')]
        self.assertEqual({l['key'] for l in labels},set(addressed))
        for label in labels:
            matched=addressed[label['key']]
            self.assertEqual(label['text'],(overrides or {}).get(label['key'], matched['property']['label'] or matched['property']['number'] or '0'))
            polygon=rotate(parcel_polygon(matched['feature']),result['rotation'],origin=result['rotation_origin'])
            polygon=affine_transform(polygon,result['page_transform'])
            self.assertTrue(polygon.covers(Polygon(label['corners'])),label)
            self.assertNotIn('leader',label)
        words=set(PdfReader(io.BytesIO(result['pdf'])).pages[0].extract_text().split())
        self.assertFalse(words.intersection({'SFR','MF','COM','VAC','EXC','Unaddressed'}))

    def test_four_workbook_types_titles_units_and_embedded_provenance(self):
        for key in MAP_TYPES:
            with self.subTest(wizard=key):
                path = workbook(self.folder / (key+'.xlsx'), key)
                case = read_map_case(path, key)
                self.assertEqual(case.title, MAP_TYPES[key][0])
                self.assertEqual(case.minimum, '5,000 Square Feet' if key.startswith('mls') else '15 Feet')
                self.assertEqual(case.reference, MAP_TYPES[key][1]+' 999'+('REN' if key.endswith('_renewal') else ''))
                data = fixture(case)
                settings = default_settings(data)
                result = render_map(data, settings)
                self.assertFalse(result['blocking'], result['blocking'])
                self.assert_addresses_inside_own_lots(data,result)
                pdf = save_map(data, settings, result, self.folder / key)
                assert_map_current(pdf, path, key)
                page = PdfReader(pdf).pages[0]
                text = page.extract_text()
                self.assertIn(case.title, text)
                self.assertIn(case.minimum, text)
                self.assertIn(case.boundary_legend, text)
                self.assertIn('12107', text)
                self.assertIn('12100-12200 blocks',text)
                self.assertNotIn('12100 and 12200',text)
                self.assertNotIn('All properties within',text)
                fonts = []
                page.extract_text(visitor_text=lambda t, cm, tm, f, size: fonts.append((t.strip(),size)) if 'DRIVE' in t else None)
                self.assertEqual(len(fonts), 3)
                self.assertTrue(all(s==18 for t,s in fonts))
                self.assertEqual(pdf.read_bytes(), result['pdf'])
                again = save_map(data, settings, result, self.folder / key)
                self.assertNotEqual(pdf, again)

    def test_block_ranges_use_hyphens_without_changing_other_and_words(self):
        for original,expected in [
            ('100 and 200 block of Main Street','100-200 block of Main Street'),
            ('100 block and 200 block of Main Street','100-200 block of Main Street'),
            ('100 & 200 blocks of Main Street','100-200 blocks of Main Street'),
            ('100- 200 blocks of Main Street','100-200 blocks of Main Street'),
            ('100 to 200 blocks of Main Street','100-200 blocks of Main Street'),
            ('12,100 and 12,200 blocks of Main Street','12,100-12,200 blocks of Main Street'),
            ('100 block of Rock and Roll Street','100 block of Rock and Roll Street'),
            ('100 block of Main and 200 block of Oak','100 block of Main and 200 block of Oak')]:
            self.assertEqual(format_block_range(original),expected)
        data=fixture(read_map_case(workbook(self.folder/'range.xlsx'),'mls'))
        data['case']['sides']='north and south sides'
        text=' '.join(PdfReader(io.BytesIO(render_map(data,default_settings(data))['pdf'])).pages[0].extract_text().split())
        self.assertIn('north and south sides',text)
        self.assertIn('between Cool Mist Drive and',text)

    def test_old_workbook_without_dimension_metadata_can_be_read(self):
        import re
        from zipfile import ZipFile,ZIP_DEFLATED
        source=workbook(self.folder/'dimensions.xlsx')
        with ZipFile(source) as z: entries={n:z.read(n) for n in z.namelist()}
        for name in list(entries):
            if name.startswith('xl/worksheets/sheet') and name.endswith('.xml'):
                entries[name]=re.sub(rb'<dimension\b[^>]*/>',b'',entries[name])
        with ZipFile(source,'w',ZIP_DEFLATED) as z:
            for name,content in entries.items():z.writestr(name,content)
        case=read_map_case(source,'mls')
        self.assertEqual(len(case.properties),2)
        self.assertEqual(case.minimum,'5,000 Square Feet')

    def test_numbered_streets_match_city_ordinals_and_use_public_label(self):
        self.assertEqual(road_key('E 22 Street'),road_key('E 22ND ST'))
        self.assertNotEqual(road_key('E 22 Street'),road_key('E 23RD ST'))
        data=fixture(read_map_case(workbook(self.folder/'numbered.xlsx'),'mls'))
        data['case']['cross_streets'][0]='E 22 Street'
        data['roads'][1]['attributes']['FULL_NAME']='E 22ND ST'
        result=render_map(data,default_settings(data))
        self.assertFalse(result['blocking'],result['blocking'])
        self.assertIn('EAST 22ND STREET',[l['text'] for l in result['labels']])

        for key in MAP_TYPES:
            with self.subTest(wizard=key):
                case = read_map_case(workbook(self.folder/(key+'_ordinals.xlsx'),key),key)
                case.block = '100 block of E 18TH Street'
                case.street = 'E 18TH ST'
                case.cross_streets = ['E 22ND Street', 'E 21ST Street']
                data = fixture(case)
                original_case = copy.deepcopy(data['case'])
                result = render_map(data,default_settings(data))
                footer = ' '.join(result['footer_lines'])
                self.assertIn('E 18th Street',footer)
                self.assertIn('E 22nd Street',footer)
                self.assertIn('E 21st Street',footer)
                self.assertEqual(data['case'],original_case)
                self.assertIn('EAST 18TH STREET',[l['text'] for l in result['labels']])
                self.assertIn('EAST 22ND STREET',[l['text'] for l in result['labels']])
                page_text = PdfReader(io.BytesIO(result['pdf'])).pages[0].extract_text()
                self.assertIn('E 18th Street',page_text)
                self.assertIn('EAST 18TH STREET',page_text)
                case.cross_streets = ['E 23Rd Street', 'E 21st Street']
                self.assertIn('E 23rd Street',' '.join(footer_lines(case)))

    def test_preview_orientation_framing_drag_reset_and_refresh(self):
        import tkinter as tk
        from core.boundary_map_ui import MapPreview
        from core.boundary_map import ROAD_SERVICE
        from wizards.mls.wizard import MailoutWizard
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        with patch.object(session,'SETTINGS_PATH',self.folder/'controls-session.json'),patch.object(session,'load',return_value={}):
            owner=MailoutWizard(root);self.addCleanup(owner.destroy)
            source=workbook(self.folder/'controls.xlsx')
            owner.vars['excel'].set(str(source));owner.vars['output_folder'].set(str(self.folder/'controls'))
            with patch.object(MapPreview,'start_load'):
                preview=MapPreview(owner);self.addCleanup(preview.close)
                preview.data=fixture(read_map_case(source,'mls'));preview.reset();root.update();preview.redraw()
                preview.refresh_button.configure(state='normal')
            def descendants(w):
                for child in w.winfo_children():
                    yield child
                    yield from descendants(child)
            def click(caption):
                next(w for w in descendants(preview) if 'text' in w.keys() and w.cget('text')==caption).invoke()
                preview.redraw()
            original=default_settings(preview.data)
            self.assertTrue(preview.interior_lot_lines.get())
            click('Show tax-parcel boundaries')
            self.assertFalse(preview.settings['show_interior_lot_lines'])
            self.assertEqual(str(preview.use.cget('state')), 'normal')
            self.assertEqual(str(preview.download.cget('state')), 'normal')
            click('Reset layout')
            self.assertTrue(preview.interior_lot_lines.get())
            click('North up');self.assertEqual(preview.rendered['rotation'],0)
            scale=preview.rendered['scale']
            preview.zoom.set(.8);preview.changed();preview.redraw()
            self.assertAlmostEqual(preview.rendered['scale'],scale*.8)
            self.assert_addresses_inside_own_lots(preview.data,preview.rendered)
            transform=preview.rendered['page_transform']
            click('Right');self.assertAlmostEqual(preview.rendered['page_transform'][4],transform[4]-20)
            click('Left');self.assertAlmostEqual(preview.rendered['page_transform'][4],transform[4])
            click('Up');self.assertAlmostEqual(preview.rendered['page_transform'][5],transform[5]-20)
            click('Down');self.assertAlmostEqual(preview.rendered['page_transform'][5],transform[5])
            click('Reset layout');self.assertEqual(preview.settings,original)
            click('Rotate address labels 90 degrees')
            self.assertTrue(preview.vertical_addresses.get())
            self.assertTrue(all(abs(l['angle'])==90 for l in preview.rendered['labels'] if l['key'].startswith('parcel:')))
            click('Enclose the street between parcel rows');self.assertTrue(preview.settings['include_street'])
            self.assert_addresses_inside_own_lots(preview.data,preview.rendered)
            click('Reset layout')
            preview.page_scale=1
            for prefix in ('parcel:','street:'):
                item=next(l for l in preview.rendered['labels'] if l['key'].startswith(prefix))
                with patch.object(preview,'page_point',side_effect=lambda e:(e.x,e.y)):
                    preview.drag_start(SimpleNamespace(x=item['x'],y=item['y']))
                    preview.drag_end(SimpleNamespace(x=item['x']+2,y=item['y']))
                preview.redraw()
                self.assertEqual(preview.settings['label_offsets'][item['key']],[2,0])
            self.assert_addresses_inside_own_lots(preview.data,preview.rendered)
            click('Reset layout');self.assertEqual(preview.settings,original)
            data=fixture(read_map_case(source,'mls'));cache_modes=[]
            def query(client,service,*args,**kwargs):
                cache_modes.append(client.use_cache)
                return data['roads'] if service==ROAD_SERVICE else data['context']
            with patch.object(PublicMapClient,'query',query):
                preview.refresh_button.invoke()
                deadline=time.monotonic()+10
                while preview.rendered is None and time.monotonic()<deadline:
                    root.update();time.sleep(.02)
                self.assertIsNotNone(preview.rendered)
                self.assertTrue(cache_modes)
                self.assertFalse(any(cache_modes))
                self.assertFalse(preview.rendered['blocking'],preview.rendered['blocking'])
            preview.close();self.assertTrue(preview.cancel.is_set())

    def test_network_failure_can_be_retried_and_closing_cancels(self):
        import tkinter as tk
        from core.boundary_map_ui import MapPreview
        from core.boundary_map import ROAD_SERVICE
        from wizards.mls.wizard import MailoutWizard
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        with patch.object(session,'SETTINGS_PATH',self.folder/'retry-session.json'),patch.object(session,'load',return_value={}):
            owner=MailoutWizard(root);self.addCleanup(owner.destroy)
            source=workbook(self.folder/'retry.xlsx')
            owner.vars['excel'].set(str(source));owner.vars['output_folder'].set(str(self.folder/'retry'))
            with patch.object(PublicMapClient,'query',side_effect=MapError('Service temporarily unavailable')):
                preview=MapPreview(owner);self.addCleanup(preview.close)
                deadline=time.monotonic()+10
                while not preview.status.get().startswith('Map not created') and time.monotonic()<deadline:
                    root.update();time.sleep(.02)
                self.assertIn('Service temporarily unavailable',preview.review.get('1.0','end'))
                self.assertTrue(preview.use.instate(['disabled']))
            data=fixture(read_map_case(source,'mls'))
            def query(client,service,*args,**kwargs):return data['roads'] if service==ROAD_SERVICE else data['context']
            with patch.object(PublicMapClient,'query',query):
                preview.refresh_button.invoke();deadline=time.monotonic()+10
                while preview.rendered is None and time.monotonic()<deadline:
                    root.update();time.sleep(.02)
                self.assertIsNotNone(preview.rendered)
                self.assertFalse(preview.use.instate(['disabled']))
            preview.close();self.assertTrue(preview.cancel.is_set())

    def test_minimum_never_guessed_from_property_areas(self):
        path = workbook(self.folder/'missing.xlsx')
        w=load_workbook(path)
        w.worksheets[1]['E4']=None
        w.save(path); w.close()
        with self.assertRaisesRegex(MapError, 'minimum'):
            read_map_case(path,'mls')

    def test_account_leading_zero_and_duplicate_conflicts(self):
        for value in ('0071050000002',71050000002,'71050000002.0','7.1050000002E10'):
            self.assertEqual(account_number(value),'0071050000002')
        self.assertEqual(account_number('0001050000002'),'0001050000002')
        with self.assertRaises(MapError):account_number('0000000000000')
        self.assertEqual(account_number(130550000022), '0130550000022')
        self.assertEqual(account_number('130550000022.0'), '0130550000022')
        with self.assertRaises(MapError): account_number('1234')
        path=workbook(self.folder/'duplicates.xlsx', props=[(12103,'CRYSTALWOOD','1047810000018','SFR'),
                                                          (12107,'CRYSTALWOOD','1047810000018','SFR')])
        with self.assertRaisesRegex(MapError,'same HCAD account'):
            read_map_case(path,'mls')

    def test_wrong_wizard_and_street_are_rejected(self):
        path=workbook(self.folder/'wrong.xlsx')
        with self.assertRaisesRegex(MapError,'different wizard'):
            read_map_case(path,'mbl')
        w=load_workbook(path); w.worksheets[1]['B4']='Unrelated Street'; w.save(path); w.close()
        with self.assertRaisesRegex(MapError,'none of the Step 1'):
            read_map_case(path,'mls')

    def test_workbook_land_use_override_and_unknown_classification(self):
        path=workbook(self.folder/'use.xlsx')
        w=load_workbook(path); w.worksheets[1]['B17']='MF'; w.save(path); w.close()
        self.assertEqual(read_map_case(path,'mls').properties[0].land_use,'MF')
        w=load_workbook(path); w.worksheets[1]['B17']='UNKNOWN'; w.save(path); w.close()
        with self.assertRaisesRegex(MapError,'Land Use'):
            read_map_case(path,'mls')

    def test_only_outer_boundary_is_emphasized_with_interior_gap(self):
        import fitz
        for key in MAP_TYPES:
            data = fixture(read_map_case(workbook(self.folder/(key+'_outline.xlsx'),key),key))
            first = data['matched'][0]
            x,y,_,_ = parcel_polygon(first['feature']).bounds
            outer = box(x,y,x+250,y+250)
            hole = box(x+90,y+90,x+160,y+160)
            first['feature'] = feature(outer.difference(hole), first['property']['account'],12103)
            data['matched'] = [first]
            data['context'] = [first['feature'],feature(hole,'1047810000099',12105)]
            before = copy.deepcopy(data)
            for enclose_street in (False,True):
                settings = default_settings(data)
                settings['include_street'] = enclose_street
                rendered = render_map(data,settings)
                self.assert_addresses_inside_own_lots(data,rendered)
                with fitz.open(stream=rendered['pdf'],filetype='pdf') as doc:
                    drawings = doc[0].get_drawings()
                    emphasized = [d for d in drawings if abs((d['width'] or 0)-1.65)<.001]
                    self.assertEqual(len(emphasized),1)
                    self.assertEqual(len(emphasized[0]['items']),1,'The interior gap must not get a bold outline')
                    self.assertIsNone(emphasized[0]['fill'])
                    parcels = [d for d in drawings if d.get('color') and abs(d['color'][0]-177/255)<.001]
                    self.assertTrue(parcels)
                    self.assertTrue(all(abs(d['width']-.5)<.001 for d in parcels))
                    self.assertTrue(all(d['fill'] is None or min(d['fill'])>.99 for d in parcels))
                self.assertEqual(data,before,'Drawing rules must not change source parcel geometry')

    def test_interior_lot_lines_toggle_exports_without_changing_boundary_or_labels(self):
        import fitz
        from core.boundary_map_render import export_map
        for key in MAP_TYPES:
            data = fixture(read_map_case(workbook(self.folder/(key+'_interior.xlsx'),key),key))
            before = copy.deepcopy(data)
            for angle in (0, 27):
                settings = default_settings(data)
                self.assertTrue(settings['show_interior_lot_lines'])
                settings.update(rotation=angle, include_street=True)
                results = {}
                for show in (True, False):
                    settings['show_interior_lot_lines'] = show
                    rendered = render_map(data,settings)
                    results[show] = rendered
                    self.assert_addresses_inside_own_lots(data,rendered)
                    with fitz.open(stream=rendered['pdf'],filetype='pdf') as doc:
                        # Sample the shared parcel edge away from either address.
                        x,y,x1,y1 = parcel_polygon(data['matched'][0]['feature']).bounds
                        first = parcel_polygon(data['matched'][0]['feature'])
                        second = parcel_polygon(data['matched'][1]['feature'])
                        center = first.union(second).centroid
                        from shapely.geometry import Point
                        point = rotate(Point(x1,y+20), angle, origin=(center.x,center.y))
                        point = affine_transform(point,rendered['page_transform'])
                        pix = doc[0].get_pixmap(matrix=fitz.Matrix(4,4),alpha=False)
                        px,py = round(point.x*4),round((792-point.y)*4)
                        darkest = min(min(pix.pixel(i,j)[:3]) for i in range(px-2,px+3) for j in range(py-2,py+3))
                        self.assertLess(darkest,220) if show else self.assertGreater(darkest,250)
                        bold = [d for d in doc[0].get_drawings() if abs((d['width'] or 0)-1.65)<.001]
                        self.assertTrue(bold)
                        self.assertTrue(all(d['fill'] is None for d in bold))
                    if angle == 0:
                        saved = save_map(data,settings,rendered,self.folder/f'{key}_{show}')
                        self.assertEqual(saved.read_bytes(),rendered['pdf'])
                        meta = json.loads(saved.with_suffix('.boundary.json').read_text())
                        self.assertEqual(meta['settings']['show_interior_lot_lines'],show)
                        downloaded = self.folder/f'{key}_{show}_download.pdf'
                        export_map(rendered,downloaded)
                        self.assertEqual(downloaded.read_bytes(),rendered['pdf'])
                self.assertEqual(results[True]['labels'],results[False]['labels'])
                self.assertEqual(results[True]['page_transform'],results[False]['page_transform'])
            self.assertEqual(data,before)

    def test_map_text_edits_persist_in_all_four_wizards(self):
        from core.map_text import saved_layout
        from core.boundary_map_render import export_map, STREET_FONT_SIZE
        for key in MAP_TYPES:
            data=fixture(read_map_case(workbook(self.folder/(key+'_text.xlsx'),key),key))
            before=copy.deepcopy(data)
            settings=default_settings(data)
            original=render_map(data,settings)
            street=next(l for l in original['labels'] if l['key'].startswith('street:'))
            settings['text_overrides']={'parcel:2':'12103A',street['key']:'CRYSTAL WOOD DRIVE',
                                        'legend:1':'12100 block of E. 22ND Street,'}
            edited=render_map(data,settings)
            self.assert_addresses_inside_own_lots(data,edited,settings['text_overrides'])
            self.assertEqual(next(l['text'] for l in edited['labels'] if l['key']=='parcel:2'),'12103A')
            changed_street=next(l for l in edited['labels'] if l['key']==street['key'])
            self.assertEqual(changed_street['text'],'CRYSTAL WOOD DRIVE')
            self.assertEqual(changed_street['size'],STREET_FONT_SIZE)
            self.assertEqual(edited['footer_lines'][1],'12100 block of E. 22nd Street,')
            self.assertEqual(edited['page_transform'],original['page_transform'])
            self.assertEqual(data,before,'Text edits must not alter workbook/property matching data')
            saved=save_map(data,settings,edited,self.folder/(key+'_saved'))
            self.assertEqual(saved.read_bytes(),edited['pdf'])
            restored=saved_layout(saved,data,default_settings(data))
            self.assertEqual(restored,settings)
            self.assertEqual(render_map(data,restored)['labels'],edited['labels'])
            download=self.folder/(key+'_edited_download.pdf')
            export_map(edited,download)
            self.assertEqual(download.read_bytes(),edited['pdf'])
            with PdfReader(io.BytesIO(edited['pdf'])) as doc:
                text=doc.pages[0].extract_text()
                self.assertIn('12103A',text)
                self.assertIn('CRYSTAL WOOD DRIVE',text)
                self.assertIn('22nd Street',text)
            other=copy.deepcopy(data)
            other['case']['workbook_hash']='another workbook'
            self.assertEqual(saved_layout(saved,other,default_settings(other)),default_settings(other))
            other=copy.deepcopy(data)
            other['case']['wizard']='different wizard'
            self.assertEqual(saved_layout(saved,other,default_settings(data)),default_settings(data))
            for edit_key,value in [('legend:1','W'*190),('parcel:2',''),(street['key'],'W'*190)]:
                bad=copy.deepcopy(settings)
                bad['text_overrides'][edit_key]=value
                with self.assertRaises(MapError):render_map(data,bad)

    def test_click_edits_text_cancel_invalid_and_reset_preserve_map(self):
        import tkinter as tk
        from core.boundary_map_ui import MapPreview
        from wizards.mls.wizard import MailoutWizard
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        with patch.object(session,'SETTINGS_PATH',self.folder/'edit-session.json'),patch.object(session,'load',return_value={}):
            owner=MailoutWizard(root);self.addCleanup(owner.destroy)
            source=workbook(self.folder/'click_edit.xlsx')
            owner.vars['excel'].set(str(source));owner.vars['output_folder'].set(str(self.folder/'out'))
            with patch.object(MapPreview,'start_load'):
                preview=MapPreview(owner);self.addCleanup(preview.close)
                preview.data=fixture(read_map_case(source,'mls'));preview.reset();preview.redraw()
            preview.page_scale=1
            for key,value in [('parcel:2','12103A'),('street:CRYSTALWOOD DR','CRYSTAL WOOD DRIVE'),('legend:1','12100 block of Crystalwood Drive,')]:
                if key.startswith('street:'):
                    key=next(l['key'] for l in preview.rendered['labels'] if l['key'].startswith('street:'))
                item=next(l for l in preview.rendered['editable_labels'] if l['key']==key)
                x=item.get('x',(item['bounds'][0]+item['bounds'][2])/2)
                y=item.get('y',(item['bounds'][1]+item['bounds'][3])/2)
                with patch.object(preview,'page_point',return_value=(x,y)),patch('tkinter.simpledialog.askstring',return_value=value) as dialog:
                    preview.drag_start(SimpleNamespace());preview.drag_end(SimpleNamespace())
                    dialog.assert_called_once()
                self.assertEqual(preview.settings['text_overrides'][key],value)
                self.assertTrue(preview.use.instate(['!disabled']))
                self.assertTrue(preview.download.instate(['!disabled']))
            expected=copy.deepcopy(preview.settings['text_overrides'])
            preview.reset();preview.redraw()
            self.assertEqual(preview.settings['text_overrides'],expected)
            prior=preview.rendered['pdf']
            with patch('tkinter.simpledialog.askstring',return_value=None):preview.edit_text('parcel:2')
            self.assertEqual(prior,preview.rendered['pdf'])
            with patch('tkinter.simpledialog.askstring',return_value=''),patch('tkinter.messagebox.showwarning') as warning:
                preview.edit_text('parcel:2');warning.assert_called_once()
            self.assertEqual(prior,preview.rendered['pdf'])
            self.assertEqual(preview.settings['text_overrides'],expected)
            saved=save_map(preview.data,preview.settings,preview.rendered,self.folder/'restore')
            owner.vars['map'].set(str(saved))
            with patch.object(MapPreview,'start_load'):
                reopened=MapPreview(owner);self.addCleanup(reopened.close)
                reopened.events.put(('ready',preview.data));reopened.poll();reopened.redraw()
                self.assertEqual(reopened.settings['text_overrides'],expected)
                self.assertEqual(reopened.rendered['labels'],preview.rendered['labels'])
                # A parcel-data refresh retains unsaved wording for this case.
                reopened.settings['text_overrides']['parcel:2']='12103B'
                reopened.events.put(('ready',preview.data));reopened.poll();reopened.redraw()
                self.assertEqual(reopened.settings['text_overrides']['parcel:2'],'12103B')
                reopened.close()
            preview.reset_text();preview.redraw()
            self.assertEqual(preview.settings['text_overrides'],{})
            self.assertEqual(next(l['text'] for l in preview.rendered['labels'] if l['key']=='parcel:2'),'12103')

    def test_individual_street_size_rotation_save_and_reset_all_four(self):
        from core.map_text import saved_layout
        from core.boundary_map_render import export_map, STREET_FONT_SIZE
        for key in MAP_TYPES:
            data=fixture(read_map_case(workbook(self.folder/(key+'_styles.xlsx'),key),key))
            settings=default_settings(data)
            original=render_map(data,settings)
            streets=[i for i in original['labels'] if i['key'].startswith('street:')]
            selected=streets[0]['key']
            settings['label_offsets'][selected]=[7,9]
            positioned=render_map(data,settings)
            settings['street_styles'][selected]={'size':28.5,'angle':-35}
            edited=render_map(data,settings)
            self.assertEqual(edited['footer_lines'],original['footer_lines'])
            self.assertEqual(edited['page_transform'],original['page_transform'])
            for prior,after in zip(positioned['labels'],edited['labels']):
                self.assertEqual((prior['key'],prior['x'],prior['y']),(after['key'],after['x'],after['y']))
                if after['key']==selected:
                    self.assertEqual(after['size'],28.5)
                    self.assertEqual(after['angle'],-35)
                    self.assertNotEqual(after['corners'],prior['corners'])
                else:self.assertEqual(prior,after)
            self.assert_addresses_inside_own_lots(data,edited)
            saved=save_map(data,settings,edited,self.folder/(key+'_style_save'))
            restored=saved_layout(saved,data,default_settings(data))
            self.assertEqual(restored['street_styles'],settings['street_styles'])
            self.assertEqual(render_map(data,restored)['labels'],edited['labels'])
            downloaded=self.folder/(key+'_styled_download.pdf')
            export_map(edited,downloaded)
            self.assertEqual(downloaded.read_bytes(),edited['pdf'])
            for size,angle in [(0,0),(80,0),(18,181),(float('nan'),0),(18,float('inf'))]:
                bad=copy.deepcopy(settings);bad['street_styles'][selected]={'size':size,'angle':angle}
                with self.assertRaises(MapError):render_map(data,bad)
            settings['street_styles'][selected]={'size':12,'angle':None}
            automatic=render_map(data,settings)
            item=next(l for l in automatic['labels'] if l['key']==selected)
            self.assertEqual(item['size'],12)
            self.assertEqual(item['angle'],streets[0]['angle'])
            settings['street_styles'].clear()
            self.assertEqual(render_map(data,settings)['labels'],positioned['labels'])

    def test_street_appearance_dialog_apply_drag_restore_and_validation(self):
        import tkinter as tk
        from core.boundary_map_ui import MapPreview
        from wizards.mls.wizard import MailoutWizard
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        with patch.object(session,'SETTINGS_PATH',self.folder/'style_session.json'),patch.object(session,'load',return_value={}):
            owner=MailoutWizard(root);self.addCleanup(owner.destroy)
            source=workbook(self.folder/'style_ui.xlsx')
            owner.vars['excel'].set(str(source));owner.vars['output_folder'].set(str(self.folder/'out'))
            with patch.object(MapPreview,'start_load'):
                preview=MapPreview(owner);self.addCleanup(preview.close)
                preview.data=fixture(read_map_case(source,'mls'));preview.reset();preview.redraw()
            preview.street_appearance_button.invoke()
            editor=preview._street_editor;editor.withdraw()
            self.addCleanup(editor.destroy)
            key=editor.selected_key()
            editor.size.set(25);editor.automatic.set(False);editor.toggle_automatic();editor.angle.set(47)
            editor.apply()
            item=next(l for l in preview.rendered['labels'] if l['key']==key)
            self.assertEqual((item['size'],item['angle']),(25,47))
            preview.page_scale=1
            self.assertEqual(preview.label_at(item['x'],item['y'])['key'],key)
            with patch.object(preview,'page_point',side_effect=[(item['x'],item['y']),(item['x']+8,item['y']+10)]):
                preview.drag_start(SimpleNamespace());preview.drag_end(SimpleNamespace())
            preview.redraw()
            moved=next(l for l in preview.rendered['labels'] if l['key']==key)
            self.assertAlmostEqual(moved['x'],item['x']+8);self.assertAlmostEqual(moved['y'],item['y']+10)
            with patch.object(preview,'page_point',return_value=(moved['x'],moved['y'])):
                preview.street_context(SimpleNamespace())
            self.assertEqual(editor.selected_key(),key)
            prior=preview.rendered['pdf']
            editor.size.set(100)
            with patch('tkinter.messagebox.showwarning') as warning:
                editor.apply();warning.assert_called_once()
            self.assertEqual(prior,preview.rendered['pdf'])
            editor.reset_selected()
            self.assertNotIn(key,preview.settings['street_styles'])
            self.assertEqual(next(l['size'] for l in preview.rendered['labels'] if l['key']==key),18)
            editor.choice.current(1);editor.load_values()
            other=editor.selected_key()
            editor.size.set(14);editor.automatic.set(False);editor.angle.set(90);editor.apply()
            self.assertEqual(preview.settings['street_styles'][other],{'size':14,'angle':90})
            preview.settings['text_overrides'][key]='CRYSTALWOOD'
            preview.reset();preview.redraw()
            self.assertEqual(preview.settings['street_styles'],{})
            self.assertEqual(preview.settings['text_overrides'][key],'CRYSTALWOOD')

    def test_holes_and_multipart_geometry_are_preserved(self):
        p=Polygon([(0,0),(100,0),(100,100),(0,100)], holes=[[(20,20),(80,20),(80,80),(20,80)]])
        f=feature(p,'1047810000018',12103)
        f['geometry']['rings'].append(list(box(150,0,180,30).exterior.coords))
        shape=parcel_polygon(f)
        self.assertAlmostEqual(shape.area,7300)
        self.assertFalse(shape.covers(box(30,30,40,40)))

    def test_exc_stays_inside_boundary_with_address_only_in_all_four_wizards(self):
        for key in MAP_TYPES:
            with self.subTest(wizard=key):
                path=workbook(self.folder/(key+'_exc.xlsx'),key,props=[(12103,'CRYSTALWOOD','1047810000018','SFR'),
                                                                    (12107,'CRYSTALWOOD','1047810000017','EXC')])
                data=fixture(read_map_case(path,key))
                exc=parcel_polygon(data['matched'][1]['feature'])
                for include_street in (False,True):
                    self.assertTrue(application_boundary(data,include_street).buffer(.001).covers(exc))
                    settings=default_settings(data); settings['include_street']=include_street
                    rendered=render_map(data,settings)
                    self.assertFalse(rendered['blocking'],rendered['blocking'])
                    self.assert_addresses_inside_own_lots(data,rendered)

    def test_exc_only_property_list_can_be_mapped(self):
        for key in MAP_TYPES:
            path=workbook(self.folder/(key+'_all_exc.xlsx'),key,props=[(12103,'CRYSTALWOOD','1047810000018','EXC')])
            data=fixture(read_map_case(path,key))
            self.assertTrue(application_boundary(data,True).buffer(.001).covers(parcel_polygon(data['matched'][0]['feature'])))

    def test_unlisted_parcel_stays_outside_when_exc_is_inside(self):
        data=fixture(read_map_case(workbook(self.folder/'gap.xlsx'),'mls'))
        data['matched'][1]['property']['land_use']='EXC'
        x,y,_,top=parcel_polygon(data['matched'][0]['feature']).bounds
        data['matched'][1]['feature']['geometry']['rings']=[list(box(x+100,y,x+170,top).exterior.coords)]
        unlisted=feature(box(x+70,y,x+100,top),'1047810000099',12105)
        data['context']=[m['feature'] for m in data['matched']]+[unlisted]
        boundary=application_boundary(data,True)
        self.assertLess(boundary.intersection(parcel_polygon(unlisted)).area,.001)
        self.assertTrue(boundary.buffer(.001).covers(parcel_polygon(data['matched'][1]['feature'])))

    def test_older_generated_maps_warn_without_requiring_regeneration(self):
        from pypdf import PdfWriter
        source = workbook(self.folder / 'legacy.xlsx')
        data = fixture(read_map_case(source, 'mls'))
        rendered = render_map(data, default_settings(data))
        writer = PdfWriter(clone_from=io.BytesIO(rendered['pdf']))
        old = self.folder / 'old_map.pdf'
        for version, detail in [('V1', 'EXC boundary correction'), ('V2', 'land-use labels'),
                                ('V3', 'block-range formatting'), ('V4', 'City of Houston seal')]:
            with self.subTest(version=version):
                writer.add_metadata({'/Keywords': f'RCA_BOUNDARY_MAP_{version}|mls|{data["case"]["workbook_hash"]}'})
                with old.open('wb') as stream:
                    writer.write(stream)
                before = old.read_bytes()
                warning = assert_map_current(old, source, 'mls')
                self.assertIn(detail, warning)
                self.assertIn('You can continue generating documents with the selected map.', warning)
                self.assertEqual(before, old.read_bytes())

    def test_stale_map_warns_for_same_timestamp_change_and_wrong_wizard(self):
        path = workbook(self.folder / 'source.xlsx')
        data = fixture(read_map_case(path, 'mls'))
        result = render_map(data, default_settings(data))
        pdf = self.folder / 'copied_without_sidecar.pdf'
        pdf.write_bytes(result['pdf'])
        self.assertIsNone(assert_map_current(pdf, path, 'mls'))
        self.assertIn('changed', assert_map_current(pdf, path, 'mbl'))
        before = path.stat()
        content = bytearray(path.read_bytes())
        content[-1] ^= 1
        path.write_bytes(content)
        os.utime(path, ns=(before.st_atime_ns, before.st_mtime_ns))
        self.assertIn('changed', assert_map_current(pdf, path, 'mls'))
        self.assertEqual(result['pdf'], pdf.read_bytes())

    def test_unavailable_comparison_warns_and_external_map_needs_no_provenance(self):
        from pypdf import PdfWriter
        source = workbook(self.folder / 'comparison.xlsx')
        data = fixture(read_map_case(source, 'mls'))
        result = render_map(data, default_settings(data))
        pdf = self.folder / 'map.pdf'
        pdf.write_bytes(result['pdf'])
        with patch('core.boundary_map.fingerprint', side_effect=PermissionError('Excel is saving')):
            self.assertIn('could not be compared', assert_map_current(pdf, source, 'mls'))
        writer = PdfWriter(clone_from=io.BytesIO(result['pdf']))
        writer.add_metadata({'/Keywords': 'Externally prepared boundary map'})
        with pdf.open('wb') as stream:
            writer.write(stream)
        self.assertIsNone(assert_map_current(pdf, source, 'mls'))
        pdf.write_bytes(b'not a readable PDF')
        with self.assertRaisesRegex(MapError, 'Could not read'):
            assert_map_current(pdf, source, 'mls')

    def test_all_four_wizards_generate_letters_with_old_or_mismatched_maps(self):
        import tkinter as tk
        from zipfile import ZipFile
        from pypdf import PdfWriter
        from core.paths import sample_input_dir
        from core.boundary_map_ui import MapControls, wizard_paths
        names = {'mls': 'MailoutWizard', 'mbl': 'MBLWizard',
                 'mls_renewal': 'MLSRenewalWizard', 'mbl_renewal': 'MBLRenewalWizard'}
        samples = {
            'mls': 'MLS_Calculation_Sheet_Case_Info_v13_FILLED_SAMPLE.xlsx',
            'mbl': 'MBL_Calculation_Sheet_Case_Info_v3_FILLED_SAMPLE.xlsx',
            'mls_renewal': 'MLS_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx',
            'mbl_renewal': 'MBL_Renewal_Calculation_Sheet_Case_Info_v2_FILLED_SAMPLE.xlsx',
        }
        root = tk.Tk()
        root.withdraw()
        self.addCleanup(root.destroy)
        case = read_map_case(workbook(self.folder / 'different_case.xlsx'), 'mls')
        data = fixture(case)
        rendered = render_map(data, default_settings(data))
        with patch.object(session, 'SETTINGS_PATH', self.folder / 'settings.json'), \
                patch.object(session, 'load', return_value={}), \
                patch('tkinter.messagebox.showerror', side_effect=AssertionError('Unexpected blocking error')), \
                patch('tkinter.messagebox.askyesno', side_effect=AssertionError('Unexpected confirmation')), \
                patch('tkinter.messagebox.askokcancel', side_effect=AssertionError('Unexpected confirmation')):
            for key, name in names.items():
                owner = getattr(importlib.import_module('wizards.' + key + '.wizard'), name)(root)
                source = self.folder / (key + '.xlsx')
                wb = load_workbook(sample_input_dir(key) / samples[key], data_only=True)
                wb.save(source)
                wb.close()
                source_before = source.read_bytes()
                controls_parent = tk.Frame(root)
                for version in ('V1', 'V5'):
                    with self.subTest(wizard=key, map_version=version):
                        pdf = self.folder / (key + '_' + version + '.pdf')
                        writer = PdfWriter(clone_from=io.BytesIO(rendered['pdf']))
                        writer.add_metadata({'/Keywords': f'RCA_BOUNDARY_MAP_{version}|mls|{case.workbook_hash}'})
                        with pdf.open('wb') as stream:
                            writer.write(stream)
                        map_before = pdf.read_bytes()
                        output = self.folder / (key + '_' + version)
                        output.mkdir()
                        if key == 'mls':
                            owner.vars['excel'].set(str(source))
                            owner.vars['map'].set(str(pdf))
                            owner.vars['output_folder'].set(str(output))
                        else:
                            owner.files.update(workbook=str(source), map=str(pdf), out=str(output))
                        controls = MapControls(controls_parent, owner, 0)
                        self.assertIn('Warning:', controls.status.cget('text'))
                        self.assertIn('You can continue', controls.status.cget('text'))
                        self.assertEqual('#B45309', controls.status.cget('foreground'))
                        self.assertIn('Warning:', owner._check_boundary_map())
                        # Execute the actual notification actions and rendering.
                        # MLS's worker runs inline here to avoid a Tk mainloop in the test.
                        if key == 'mls':
                            def inline_worker(*, target, **kwargs):
                                return SimpleNamespace(start=target)
                            with patch('wizards.mls.wizard.threading.Thread', side_effect=inline_worker), \
                                    patch.object(owner, 'after', side_effect=lambda delay, callback: callback()), \
                                    patch.object(owner, '_done'), \
                                    patch.object(owner, '_error', side_effect=AssertionError):
                                owner._run(preview=True)
                        else:
                            with patch.object(owner, '_notify_done'):
                                if key == 'mls_renewal':
                                    owner._preview_notice()
                                else:
                                    role = 'owner_template' if key == 'mbl_renewal' else 'template'
                                    owner._letters(role, 'Notification_Preview.docx', 'preview')
                        docs = list(output.glob('*.docx'))
                        self.assertEqual(1, len(docs), str(output))
                        with ZipFile(docs[0]) as archive:
                            self.assertIn(b'<w:drawing>', archive.read('word/document.xml'))
                        self.assertEqual(str(pdf), wizard_paths(owner)[1])
                        self.assertEqual(map_before, pdf.read_bytes())
                        self.assertEqual(source_before, source.read_bytes())
                        controls.destroy()
                controls_parent.destroy()
                owner.destroy()
                root.update()


    def test_public_query_paginates_and_rejects_truncation(self):
        class Fake(PublicMapClient):
            def request(self, service, params):
                if params.get('returnIdsOnly'):
                    return {'objectIdFieldName':'OBJECTID','objectIds':list(range(301))}
                ids=[int(s) for s in params['objectIds'].split(',')]
                return {'spatialReference':{'wkid':2278},
                        'features':[{'attributes':{'OBJECTID':i}} for i in ids]}
        client=Fake(self.folder/'cache')
        self.assertEqual(len(client.query('https://example.invalid')),301)
        with patch.object(client,'request',side_effect=AssertionError('network')):
            self.assertEqual(len(client.query('https://example.invalid')),301)
        class Truncated(Fake):
            def request(self, service, params):
                d=super().request(service,params)
                if 'features' in d: d['features']=d['features'][:-1]
                return d
        with self.assertRaisesRegex(MapError,'incomplete'):
            Truncated(self.folder/'badcache').query('https://example.invalid')

    def test_missing_parcel_or_address_mismatch_cannot_silently_drop_rows(self):
        case=read_map_case(workbook(self.folder/'match.xlsx'),'mls')
        data=fixture(case)
        class Client(PublicMapClient):
            def query(self,*args,**kwargs): return [data['matched'][0]['feature']]
        with self.assertRaisesRegex(MapError,'Row 3'):
            fetch_map_data(case,Client(self.folder/'cache'))
        bad=copy.deepcopy(data['matched'][0]['feature']); bad['attributes']['site_str_num']=99999
        with patch.object(Client,'query',return_value=[bad,data['matched'][1]['feature']]):
            with self.assertRaisesRegex(MapError,'(?i)address/account mismatch'):
                fetch_map_data(case,Client(self.folder/'cache'))

    def test_county_street_fallback_uses_exact_names_and_recorded_aliases(self):
        case=read_map_case(workbook(self.folder/'county.xlsx'),'mls')
        data=fixture(case)
        county=copy.deepcopy(data['roads'][1])
        county['attributes']={'OBJECTID':123,'STREET':'SERVICE RD','ALIAS_1':'COOL MIST DR'}
        unrelated=copy.deepcopy(county)
        unrelated['attributes']={'OBJECTID':124,'STREET':'COOL MIST ST'}
        def query(client,service,*args,**kwargs):
            if service==ROAD_SERVICE: return [data['roads'][0],data['roads'][2]]
            if service==COUNTY_ROAD_SERVICE: return [county,unrelated]
            return data['context']
        with patch.object(PublicMapClient,'query',query):
            result=fetch_map_data(case,PublicMapClient(self.folder/'cache'))
        recovered=[r for r in result['roads'] if r.get('_road_service')==COUNTY_ROAD_SERVICE]
        self.assertEqual(len(recovered),1)
        self.assertEqual(recovered[0]['geometry'],county['geometry'])
        self.assertEqual(road_key(recovered[0]['attributes']['FULL_NAME']),road_key(case.cross_streets[0]))
        self.assertTrue(any('Harris County street geometry' in w for w in result['warnings']))
        rendered=render_map(result,default_settings(result))
        self.assertFalse(rendered['blocking'],rendered['blocking'])
        self.assert_addresses_inside_own_lots(result,rendered)

    def test_missing_cross_street_warns_and_saves_in_all_four_wizards(self):
        import tkinter as tk
        from core.boundary_map_ui import MapPreview
        names={'mls':'MailoutWizard','mbl':'MBLWizard','mls_renewal':'MLSRenewalWizard','mbl_renewal':'MBLRenewalWizard'}
        root=tk.Tk(); root.withdraw(); self.addCleanup(root.destroy)
        with patch.object(session,'SETTINGS_PATH',self.folder/'settings.json'), patch.object(session,'load',return_value={}):
            for key,name in names.items():
                with self.subTest(wizard=key):
                    source=workbook(self.folder/(key+'.xlsx'),key)
                    case=read_map_case(source,key); data=fixture(case)
                    def query(client,service,*args,**kwargs):
                        if service==ROAD_SERVICE: return [data['roads'][0],data['roads'][2]]
                        if service==COUNTY_ROAD_SERVICE: return []
                        return data['context']
                    with patch.object(PublicMapClient,'query',query):
                        fetched=fetch_map_data(case,PublicMapClient(self.folder/'cache'))
                    owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),name)(root)
                    if hasattr(owner,'vars'):
                        owner.vars['excel'].set(str(source)); owner.vars['output_folder'].set(str(self.folder/key))
                    else: owner.files.update(workbook=str(source),out=str(self.folder/key))
                    with patch.object(MapPreview,'start_load'):
                        preview=MapPreview(owner); preview.withdraw()
                        preview.data=fetched; preview.reset(); preview.redraw()
                        self.assertFalse(preview.rendered['blocking'],preview.rendered['blocking'])
                        self.assertEqual(str(preview.use.cget('state')),'normal')
                        self.assertIn('Map created with warnings',preview.status.get())
                        self.assertIn('Cool Mist Drive from Step 2',preview.review.get('1.0','end'))
                        self.assert_addresses_inside_own_lots(fetched,preview.rendered)
                        with patch('tkinter.messagebox.showerror',side_effect=AssertionError): preview.use.invoke()
                        selected=owner.vars['map'].get() if hasattr(owner,'vars') else owner.files['map']
                        owner._check_boundary_map()
                        metadata=json.loads(Path(selected).with_suffix('.boundary.json').read_text())
                        self.assertTrue(any('could not be located' in w for w in metadata['warnings']))
                        text=' '.join(PdfReader(selected).pages[0].extract_text().split())
                        self.assertIn('between Cool Mist Drive and Woodcliff Drive',text)
                    owner.destroy(); root.update()

    def test_street_service_failure_still_renders_north_up_and_cancellation_stops(self):
        case=read_map_case(workbook(self.folder/'noroads.xlsx'),'mls')
        data=fixture(case)
        def query(client,service,*args,**kwargs):
            if service in (ROAD_SERVICE,COUNTY_ROAD_SERVICE): raise MapError('Street service offline')
            return data['context']
        with patch.object(PublicMapClient,'query',query):
            fetched=fetch_map_data(case,PublicMapClient(self.folder/'cache'))
        settings=default_settings(fetched)
        self.assertEqual(settings['rotation'],0)
        rendered=render_map(fetched,settings)
        self.assertFalse(rendered['blocking'],rendered['blocking'])
        self.assertEqual(sum(l['key'].startswith('street:') for l in rendered['labels']),3)
        self.assert_addresses_inside_own_lots(fetched,rendered)
        self.assertTrue(save_map(fetched,settings,rendered,self.folder/'out').is_file())
        client=PublicMapClient(self.folder/'cancel');client.cancel.set()
        with self.assertRaisesRegex(MapError,'cancelled'):
            load_case_roads(case,client,data['extent'],[])

    def test_street_label_placement_warnings_allow_save(self):
        case=read_map_case(workbook(self.folder/'streetmove.xlsx'),'mls')
        data=fixture(case); settings=default_settings(data)
        settings['label_offsets']['street:'+road_key(case.cross_streets[0])]=[10000,0]
        result=render_map(data,settings)
        self.assertFalse(result['blocking'],result['blocking'])
        self.assertTrue(any('label is outside the map frame' in w for w in result['issues']))
        self.assertTrue(save_map(data,settings,result,self.folder/'moved').is_file())

    def test_stale_account_uses_address_and_legal_to_resolve_two_parcels(self):
        from core.boundary_map import ROAD_SERVICE
        case=read_map_case(workbook(self.folder/'stale.xlsx'),'mls')
        case.properties[0].legal_description='TRS 15 & 16A BLK 15'
        data=fixture(case)
        correct=copy.deepcopy(data['matched'][0]['feature'])
        correct['attributes'].update(HCAD_NUM='0130200000034',legal_dscr_1=case.properties[0].legal_description)
        other=copy.deepcopy(correct)
        other['attributes'].update(HCAD_NUM='0130200000016',legal_dscr_1='TR 16B BLK 15')
        stale=copy.deepcopy(correct)
        stale['attributes'].update(HCAD_NUM=case.properties[0].account,site_str_name='ROSEDALE')
        second=data['matched'][1]['feature']
        class Client(PublicMapClient):
            def query(self,service,where='1=1',extent=None,**kw):
                if service==ROAD_SERVICE: return data['roads']
                if extent is not None: return [correct,second]
                if where.startswith('HCAD_NUM'): return [stale,second]
                return [correct,other]
        result=fetch_map_data(case,Client(self.folder/'cache'))
        self.assertEqual(result['matched'][0]['feature']['attributes']['HCAD_NUM'],'0130200000034')
        self.assertEqual(result['matched'][0]['match_method'],'address and legal description')
        self.assertTrue(any('stale' in w for w in result['warnings']))
        case.properties[0].legal_description=''
        with self.assertRaisesRegex(MapError,'2 matches'):
            fetch_map_data(case,Client(self.folder/'cache'))

    def test_unaddressed_parcel_preserves_account_and_legal_description(self):
        path=workbook(self.folder/'no_address.xlsx',props=[(12103,'CRYSTALWOOD','1047810000018','SFR'),(0,'','1047810000017','EXC')])
        w=load_workbook(path); s=w.worksheets[0]; s['D3']=0; s['H1']='Legal_Dscr_1'; s['H3']='LT 1 BLK 8'; w.save(path); w.close()
        p=read_map_case(path,'mls').properties[1]
        self.assertEqual(p.number,''); self.assertEqual(p.legal_description,'LT 1 BLK 8')
        self.assertEqual(p.account,'1047810000017')

    def test_city_fallback_keeps_one_consistent_parcel_edition(self):
        from core.boundary_map import CITY_PARCEL_SERVICE,ROAD_SERVICE
        case=read_map_case(workbook(self.folder/'retained.xlsx'),'mls')
        data=fixture(case)
        first,second=[m['feature'] for m in data['matched']]
        newer=copy.deepcopy(first)
        newer['geometry']['rings']=[list(parcel_polygon(first).union(parcel_polygon(second)).exterior.coords)]
        def city(f):
            a=f['attributes']
            return {**f,'attributes':{'OBJECTID':a['OBJECTID'],'HCAD_NUM':a['HCAD_NUM'],
                                      'Str_num':a['site_str_num'],'Str_name':a['site_str_name'],
                                      'Str_sfx':a['site_str_sfx'],'Tax_Year':'2024'}}
        class Client(PublicMapClient):
            def query(self,service,where='1=1',extent=None,**kw):
                if service==ROAD_SERVICE: return data['roads']
                if service==CITY_PARCEL_SERVICE:
                    return [city(f) for f in (first,second) if extent is not None or f['attributes']['HCAD_NUM'] in where]
                return [newer]
        result=fetch_map_data(case,Client(self.folder/'cache'))
        self.assertEqual(len(result['matched']),2)
        self.assertTrue(all(m['feature'].get('_parcel_service')==CITY_PARCEL_SERVICE for m in result['matched']))
        self.assertAlmostEqual(application_boundary(result).area,16100)

    def test_missing_street_type_is_allowed_but_conflicting_types_are_not(self):
        from core.boundary_map import address_agrees
        case=read_map_case(workbook(self.folder/'suffix.xlsx'),'mls')
        f=fixture(case)['matched'][0]['feature']
        f['attributes']['site_str_sfx']=''
        self.assertTrue(address_agrees(case.properties[0],f))
        f['attributes']['site_str_sfx']='ST'
        self.assertFalse(address_agrees(case.properties[0],f))

    def test_default_mbl_is_an_explicit_geographic_example(self):
        from core.workbooks import calculation_template_path
        case=read_map_case(calculation_template_path('mbl'),'mbl')
        self.assertEqual(case.reference,'MBL EXAMPLE')
        self.assertEqual(case.sides,'north side')
        self.assertEqual(case.street,'Haver Street')
        self.assertEqual(len(case.properties),5)
        self.assertIn('INSTRUCTIONAL EXAMPLE',case.warnings[0])

    def test_narrow_lot_address_fits_inside_without_a_callout(self):
        data=fixture(read_map_case(workbook(self.folder/'strip.xlsx'),'mls'))
        x,y,_,top=parcel_polygon(data['matched'][1]['feature']).bounds
        data['matched'][1]['feature']['geometry']['rings']=[list(box(x,y,x+3,top).exterior.coords)]
        result=render_map(data,default_settings(data))
        self.assertFalse(result['blocking'],result['blocking'])
        self.assert_addresses_inside_own_lots(data,result)

    def test_unaddressed_lots_show_zero_in_all_four_wizards(self):
        for key in MAP_TYPES:
            data=fixture(read_map_case(workbook(self.folder/(key+'_blank.xlsx'),key),key))
            data['matched'][1]['property'].update(label='',number='',land_use='EXC')
            result=render_map(data,default_settings(data))
            self.assertFalse(result['blocking'],result['blocking'])
            self.assertEqual(result['selected_count'],2)
            self.assertEqual(next(l for l in result['labels'] if l['key']=='parcel:3')['text'],'0')
            self.assertEqual(data['matched'][1]['property']['number'],'')
            self.assert_addresses_inside_own_lots(data,result)

    def test_repeated_address_stays_on_each_distinct_parcel(self):
        data=fixture(read_map_case(workbook(self.folder/'repeated.xlsx'),'mls'))
        data['matched'][1]['property']['label']=data['matched'][0]['property']['label']
        result=render_map(data,default_settings(data))
        self.assertFalse(result['blocking'],result['blocking'])
        self.assert_addresses_inside_own_lots(data,result)

    def test_address_fits_actual_shape_without_covering_parcel_hole(self):
        from core.boundary_map_render import fit_address, label_rectangle, register_fonts
        register_fonts()
        parcel=box(0,0,100,100).difference(box(20,20,80,80)).union(box(150,0,180,20))
        placement=fit_address('12103',parcel,0,[])
        self.assertIsNotNone(placement)
        x,y,angle,size=placement
        self.assertTrue(parcel.covers(label_rectangle('12103',x,y,angle,size)))

    def test_address_drag_cannot_move_into_another_lot(self):
        data=fixture(read_map_case(workbook(self.folder/'drag.xlsx'),'mls'))
        settings=default_settings(data)
        first=render_map(data,settings)
        a,b=[l for l in first['labels'] if l['key'].startswith('parcel:')]
        settings['label_offsets'][a['key']]=[b['x']-a['x'],b['y']-a['y']]
        moved=render_map(data,settings)
        self.assertFalse(moved['blocking'],moved['blocking'])
        self.assertTrue(any('kept inside its own lot' in s for s in moved['issues']))
        self.assert_addresses_inside_own_lots(data,moved)
        settings['label_offsets'][a['key']]=[2,0]
        moved=render_map(data,settings)
        self.assertEqual(next(l for l in moved['labels'] if l['key']==a['key'])['x'],a['x']+2)
        self.assert_addresses_inside_own_lots(data,moved)

    def test_unreadably_small_parcel_warning_allows_saving(self):
        data=fixture(read_map_case(workbook(self.folder/'tiny.xlsx'),'mls'))
        x,y,_,top=parcel_polygon(data['matched'][1]['feature']).bounds
        data['matched'][1]['feature']['geometry']['rings']=[list(box(x,y,x+.1,top).exterior.coords)]
        settings=default_settings(data); result=render_map(data,settings)
        self.assertFalse(result['blocking'])
        self.assertTrue(any('12107: its lot is too small' in s for s in result['issues']))
        path=save_map(data,settings,result,self.folder/'tiny')
        self.assertEqual(path.read_bytes(),result['pdf'])
        self.assertTrue(json.loads(path.with_suffix('.boundary.json').read_text())['warnings'])

    def test_clipped_parcels_and_overlapping_labels_warn_but_allow_saving(self):
        data=fixture(read_map_case(workbook(self.folder/'layout.xlsx'),'mls'))
        settings=default_settings(data); settings['zoom']=5
        result=render_map(data,settings)
        self.assertFalse(result['blocking'])
        self.assertTrue(any('outside the map frame' in note for note in result['issues']))
        self.assertTrue(save_map(data,settings,result,self.folder/'bad').is_file())

    def test_export_preserves_preview_bytes_and_existing_files_on_failure(self):
        from core.boundary_map_render import export_map
        data=fixture(read_map_case(workbook(self.folder/'export.xlsx'),'mls'))
        result=render_map(data,default_settings(data))
        target=self.folder/'chosen.pdf';target.write_bytes(b'prior map')
        with patch('core.boundary_map_render.os.replace',side_effect=PermissionError('file locked')):
            with self.assertRaises(PermissionError): export_map(result,target)
        self.assertEqual(target.read_bytes(),b'prior map')
        self.assertFalse(list(self.folder.glob('.boundary_map_*.tmp')))
        export_map(result,target)
        self.assertEqual(target.read_bytes(),result['pdf'])
        with self.assertRaises(MapError):export_map(result,self.folder/'bad.txt')
        with self.assertRaises(MapError):export_map(None,self.folder/'missing.pdf')

    def test_all_four_download_and_use_with_parcel_layout_warnings(self):
        import tkinter as tk
        from core.boundary_map_ui import MapPreview
        from PIL import Image
        names={'mls':'MailoutWizard','mbl':'MBLWizard','mls_renewal':'MLSRenewalWizard','mbl_renewal':'MBLRenewalWizard'}
        root=tk.Tk();root.withdraw();self.addCleanup(root.destroy)
        with patch.object(session,'SETTINGS_PATH',self.folder/'settings.json'),patch.object(session,'load',return_value={}):
            for key,name in names.items():
                with self.subTest(wizard=key):
                    owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),name)(root)
                    source=workbook(self.folder/(key+'_warnings.xlsx'),key)
                    out=self.folder/key
                    if hasattr(owner,'vars'):
                        owner.vars['excel'].set(str(source));owner.vars['output_folder'].set(str(out))
                    else:owner.files.update(workbook=str(source),out=str(out))
                    with patch.object(MapPreview,'start_load'):
                        preview=MapPreview(owner);preview.withdraw()
                        self.addCleanup(preview.close)
                        self.assertTrue(preview.download.instate(['disabled']))
                        preview.data=fixture(read_map_case(source,key))
                        preview.data['matched'][0]['property'].update(label='',number='',land_use='EXC')
                        f=preview.data['matched'][1]['feature'];x,y,_,top=parcel_polygon(f).bounds
                        f['geometry']['rings']=[list(box(x,y,x+.1,top).exterior.coords)]
                        preview.reset();preview.redraw()
                        self.assertTrue(preview.rendered['issues'])
                        self.assertFalse(preview.use.instate(['disabled']))
                        self.assertFalse(preview.download.instate(['disabled']))
                        zero=next(l for l in preview.rendered['labels'] if l['key']=='parcel:2')
                        self.assertEqual(zero['text'],'0')
                        reviewed=preview.rendered['pdf']
                        original_selection=owner.vars['map'].get() if hasattr(owner,'vars') else owner.files.get('map','')
                        with patch('tkinter.filedialog.asksaveasfilename',return_value=''):
                            preview.download.invoke()
                        for suffix in ('pdf','png'):
                            destination=self.folder/(key+'_download.'+suffix)
                            with patch('tkinter.filedialog.asksaveasfilename',return_value=str(destination)),patch('tkinter.messagebox.showerror',side_effect=AssertionError):
                                preview.download.invoke()
                            self.assertTrue(destination.is_file())
                            if suffix=='pdf':self.assertEqual(destination.read_bytes(),reviewed)
                            else:
                                with Image.open(destination) as image:self.assertEqual(image.format,'PNG')
                            self.assertFalse(preview._closed)
                            selected=owner.vars['map'].get() if hasattr(owner,'vars') else owner.files.get('map','')
                            self.assertEqual(selected,original_selection)
                        with patch('tkinter.filedialog.asksaveasfilename',return_value=str(self.folder/'locked.pdf')),patch('core.boundary_map_ui.export_map',side_effect=PermissionError('file locked')),patch('tkinter.messagebox.showerror') as error:
                            preview.download.invoke()
                            self.assertIn('locked',error.call_args.args[1])
                        self.assertFalse(preview.use.instate(['disabled']))
                        with patch('tkinter.messagebox.showerror',side_effect=AssertionError):preview.use.invoke()
                        selected=Path(owner.vars['map'].get() if hasattr(owner,'vars') else owner.files['map'])
                        self.assertEqual(selected.read_bytes(),reviewed)
                        meta=json.loads(selected.with_suffix('.boundary.json').read_text())
                        self.assertTrue(any('too small' in note for note in meta['warnings']))
                        owner._check_boundary_map()
                    owner.destroy();root.update()

    def test_all_four_buttons_attachment_updates_and_preview_save(self):
        import tkinter as tk
        from core.boundary_map_ui import MapControls, MapPreview
        names={'mls':'MailoutWizard','mbl':'MBLWizard','mls_renewal':'MLSRenewalWizard','mbl_renewal':'MBLRenewalWizard'}
        root=tk.Tk(); root.withdraw()
        self.addCleanup(root.destroy)
        with patch.object(session,'SETTINGS_PATH',self.folder/'settings.json'), patch.object(session,'load',return_value={}):
            for key,name in names.items():
                owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),name)(root)
                owner.pack(fill='both',expand=True); root.update()
                def descendants(w):
                    for ch in w.winfo_children():
                        yield ch
                        yield from descendants(ch)
                self.assertTrue(any(isinstance(w,MapControls) for w in descendants(owner)),key)
                source=workbook(self.folder/(key+'.xlsx'),key)
                out=self.folder/key
                if hasattr(owner,'vars'):
                    owner.vars['excel'].set(str(source)); owner.vars['output_folder'].set(str(out))
                else:
                    owner.files.update(workbook=str(source),out=str(out))
                with patch.object(MapPreview,'start_load'):
                    preview=MapPreview(owner)
                    preview.data=fixture(read_map_case(source,key))
                    preview.reset(); root.update(); preview.redraw()
                    self.assertFalse(preview.rendered['blocking'])
                    address=next(l for l in preview.rendered['labels'] if l['key'].startswith('parcel:'))
                    preview.settings['label_offsets'][address['key']]=[10000,0]
                    preview.redraw()
                    self.assertNotIn(address['key'],preview.settings['label_offsets'])
                    self.assert_addresses_inside_own_lots(preview.data,preview.rendered)
                    with patch('tkinter.messagebox.showerror',side_effect=AssertionError):
                        preview.accept()
                    selected=owner.vars['map'].get() if hasattr(owner,'vars') else owner.files['map']
                    self.assertTrue(Path(selected).is_file())
                    owner._check_boundary_map()
                owner.destroy(); root.update()

    def test_async_preview_load_and_workbook_change_before_accept(self):
        import tkinter as tk
        from core.boundary_map_ui import MapPreview
        from wizards.mls.wizard import MailoutWizard
        root=tk.Tk(); root.withdraw()
        self.addCleanup(root.destroy)
        with patch.object(session,'SETTINGS_PATH',self.folder/'settings.json'), patch.object(session,'load',return_value={}):
            owner=MailoutWizard(root)
            self.addCleanup(owner.destroy)
            source=workbook(self.folder/'async.xlsx')
            owner.vars['excel'].set(str(source)); owner.vars['output_folder'].set(str(self.folder/'out'))
            data=fixture(read_map_case(source,'mls'))
            def query(client,service,*args,**kwargs):
                return data['roads'] if 'LandbaseandRoads' in service else data['context']
            with patch.object(PublicMapClient,'query',query):
                preview=MapPreview(owner)
                self.addCleanup(preview.close)
                deadline=time.monotonic()+8
                while preview.rendered is None and time.monotonic()<deadline:
                    root.update(); time.sleep(.02)
                self.assertIsNotNone(preview.rendered,preview.review.get('1.0','end'))
                self.assertFalse(preview.rendered['blocking'])
                content=bytearray(source.read_bytes()); content[-1]^=1; source.write_bytes(content)
                with patch('tkinter.messagebox.showerror',side_effect=AssertionError):
                    preview.accept()
                saved=Path(owner.vars['map'].get())
                self.assertTrue(saved.is_file())
                meta=json.loads(saved.with_suffix('.boundary.json').read_text())
                self.assertTrue(any('changed after this preview' in note for note in meta['warnings']))


def real_visuals(data_folder, output):
    """Exercise all four adapters with real Crystalwood shapes and explicit test values."""
    output.mkdir(parents=True,exist_ok=True)
    raw=json.loads((data_folder/'crystal_wood_full_street.json').read_text())
    context=json.loads((data_folder/'crystal_wood_full_context.json').read_text())
    roads=json.loads((data_folder/'crystal_wood_roads.json').read_text())
    numbers={12103,12107,12111,12115,12119,12203,12207,12211,12215}
    features=[f for f in raw['features'] if f['attributes']['site_str_num'] in numbers]
    props=[(f['attributes']['site_str_num'],'CRYSTALWOOD',f['attributes']['HCAD_NUM'],'SFR') for f in features]
    for key in MAP_TYPES:
        source=workbook(output/(key+'_test.xlsx'),key,props)
        case=read_map_case(source,key)
        data={'case':asdict(case),'matched':[{'property':asdict(p),'feature':next(f for f in features if f['attributes']['HCAD_NUM']==p.account)} for p in case.properties],
              'context':context['features'],'roads':roads['features'],'sources':[raw['_source'],roads['_source']],
              'extent':[0,0,0,0],'warnings':['Visual QA fixture: example minimum and case number; not an official case.']}
        # Record the dates in the fixture, retaining the unmodified public sources.
        for s in data['sources']:
            s.setdefault('retrieved_utc','2026-09-08T00:00:00Z')
        settings=default_settings(data); result=render_map(data,settings)
        if result['blocking']: raise AssertionError(result['blocking'])
        (output/(key+'.pdf')).write_bytes(result['pdf'])
        (output/(key+'.png')).write_bytes(pdf_png(result['pdf'],144))
        (output/(key+'_data.json')).write_text(json.dumps(data),encoding='utf-8')
        print(key,case.title,case.minimum,len(case.properties),'parcels; no layout issues')


if __name__=='__main__':
    if '--real-data' in sys.argv:
        real_visuals(Path(sys.argv[sys.argv.index('--real-data')+1]),WF.parent/'build'/'map_qa'/'crystalwood')
    else:
        unittest.main(verbosity=2)
