"""PC report and map regression tests use synthetic cases, never user workbooks."""
from pathlib import Path
import copy, io, json, tempfile, time, unittest
from unittest.mock import patch
from openpyxl import load_workbook
from pypdf import PdfReader
from shapely.affinity import affine_transform, rotate
from shapely.geometry import Polygon
from tools.test_boundary_map import workbook, fixture
from core.boundary_map import read_map_case, parcel_polygon, fingerprint
from core.boundary_map_render import render_map, default_settings
from core.map_text import saved_layout
from wizards.mls.pc_data import read_pc_facts, attach_pc_rows, save_preferences, load_preferences
from wizards.mls.pc_map import render_pc_map, pc_defaults, save_pc_map
from wizards.mls.pc_landuse import COLORS, NAMES, LEGEND_USES, attach_land_use, selected_use, context_use
from wizards.mls.pc_report import report_context, fill_template, generate_pc_report

def case_fixture(folder):
    path=workbook(Path(folder)/'pc_test.xlsx',props=[(12103,'CRYSTALWOOD','1047810000018','SFR'),(12107,'CRYSTALWOOD','1047810000017','SFR'),(12111,'CRYSTALWOOD','1047810000016','COM'),(12115,'CRYSTALWOOD','1047810000015','SFR'),(12119,'CRYSTALWOOD','1047810000014','EXC')])
    wb=load_workbook(path);s1=wb.worksheets[0];s1.cell(1,8,'Legal_Dscr_1')
    for i in range(2,7):s1.cell(i,8,'TEST LOT')
    s=wb.worksheets[1];s.title='Step 2 - Calculation Sheet'
    for r,area,signed in [(17,2000,'Y'),(18,5000,'N'),(19,5000,'Y'),(20,8000,'N'),(21,10000,'Y')]:s.cell(r,4,area);s.cell(r,3,signed)
    s['A15']=4;s['F8']=.35;s['H11']=.75
    ci=wb.create_sheet('Case Info')
    for pair in [('PC_Hearing_Date','September 24, 2026'),('Protest_Count',1),('Subdivision_Plat_Year',1938),('Structure_Age_Statement','The houses originate from the 1930s.'),('Subdivision_Name','Test subdivision')]:ci.append(pair)
    wb.save(path);wb.close()
    facts=read_pc_facts(path);data=attach_pc_rows(fixture(read_map_case(path,'mls')),facts,[18])
    return path,facts,data

class PCReportTests(unittest.TestCase):
    def setUp(self):
        from tools.test_designation_lookup import no_lookup
        external=patch('core.designation_lookup.lookup',side_effect=no_lookup);external.start();self.addCleanup(external.stop)
        self.temp=tempfile.TemporaryDirectory(prefix='pc_report_tests_');self.addCleanup(self.temp.cleanup)
        self.folder=Path(self.temp.name);self.path,self.facts,self.data=case_fixture(self.folder)

    def test_metric_denominators_and_all_ties(self):
        f=self.facts
        self.assertEqual((f['all_count'],f['eligible_count'],f['excluded_count']),(5,4,1))
        self.assertEqual(f['sfr_lot_percent'],'75');self.assertEqual(f['support_area_percent'],'35')
        self.assertEqual(f['qualifying_count'],3);self.assertEqual(f['qualifying_area_percent'],'90')
        self.assertEqual(f['signed_count'],2);self.assertFalse(f['warnings'])

    def test_join_uses_description_not_row_order(self):
        facts=copy.deepcopy(self.facts);facts['rows'].reverse()
        result=attach_pc_rows(self.data,facts,[18])
        self.assertEqual([m['pc']['signed'] for m in result['matched']],[True,False,True,False,True])
        self.assertTrue(result['matched'][1]['pc']['protested'])
        facts['rows'].append(copy.deepcopy(facts['rows'][0]))
        result=attach_pc_rows(self.data,facts)
        self.assertFalse(result['matched'][-1]['pc']['known']);self.assertTrue(result['warnings'])

    def test_missing_area_does_not_publish_partial_statistics(self):
        wb=load_workbook(self.path);wb.worksheets[1]['D18']=None;wb.save(self.path);wb.close()
        f=read_pc_facts(self.path)
        self.assertIsNone(f['qualifying_count']);self.assertEqual(f['support_area_percent'],'[review]');self.assertTrue(f['warnings'])

    def test_full_reference_legend_and_exact_palette(self):
        expected={'SFR':'#FFFAD5','MF':'#FED362','COM':'#D9797A','OFF':'#FD8AFF','IND':'#BAA9D9',
                  'PUB':'#ADDFF8','TRN':'#FFFFFF','PRK':'#4ADA49','VAC':'#A5A5A5','AGR':'#5D8944'}
        self.assertEqual({k:COLORS[k] for k in LEGEND_USES},expected)
        rendered=render_pc_map(self.data,pc_defaults(self.data))
        pdf_text=PdfReader(io.BytesIO(rendered['pdf'])).pages[0].extract_text()
        self.assertEqual(len(rendered['legend_entries']),14)
        for caption in [*NAMES.values()][:-1]+['Property owner signed to support application',
             'Property owner protesting application','Properties that do not meet the minimum lot size','Proposed Application Area']:
            self.assertIn(caption,' '.join(pdf_text.split()))
        self.assertNotIn('Excluded property',pdf_text)
        self.assertGreater(min(e['bounds'][1] for e in rendered['legend_entries'] if 'bounds' in e),155)

    def test_step1_detailed_use_does_not_change_eligibility(self):
        wb=load_workbook(self.path);s=wb.worksheets[0]
        s.cell(1,9,'group_dscr');s.cell(1,10,'Property Description')
        for n,row in enumerate(self.facts['rows'],2):
            s.cell(n,10,row['description']);s.cell(n,9,'Office' if row['land_use']=='COM' else 'Single-Family Residential')
        wb.save(self.path);wb.close()
        facts=read_pc_facts(self.path);result=attach_pc_rows(self.data,facts)
        self.assertEqual(result['matched'][2]['pc']['land_use'],'COM')
        self.assertEqual(selected_use(result['matched'][2]),'OFF')
        self.assertEqual(selected_use(result['matched'][4]),'SFR')  # EXC is not a color.
        self.assertEqual((facts['sfr_lot_percent'],facts['support_area_percent']),('75','35'))

    def test_city_use_join_requires_account_and_matching_geometry(self):
        from unittest.mock import Mock
        features=[]
        for m,lu in zip(self.data['matched'],['Office','Industrial','Park & Open Spaces','Transportation & Utility','Agriculture Production']):
            f=copy.deepcopy(m['feature']);f['attributes'].update(group_dscr=lu,Tax_Year='2024');features.append(f)
        client=Mock();client.query.return_value=features;client.sources=[{'service':'test'}]
        result=attach_land_use(self.data,client)
        self.assertEqual([context_use(m['feature']) for m in result['matched']],['OFF','IND','PRK','TRN','AGR'])
        self.assertNotIn('pc_land_use',self.data['matched'][0]['feature'])
        # An account on a different footprint must not color an unrelated lot.
        features[0]['geometry']=copy.deepcopy(features[1]['geometry'])
        result=attach_land_use(self.data,client)
        self.assertEqual(context_use(result['matched'][0]['feature']),'UNK')
        client.query.side_effect=RuntimeError('offline')
        result=attach_land_use(self.data,client)
        self.assertTrue(any('could not be loaded' in w for w in result['warnings']))
        self.assertTrue(render_pc_map(result,pc_defaults(result))['pdf'])

    def test_patterns_and_colors_are_the_same_in_lots_and_legend(self):
        import fitz
        data=copy.deepcopy(self.data)
        categories=['COM','OFF','TRN','VAC','AGR']
        for m,code in zip(data['matched'],categories):m['pc'].update(map_land_use=NAMES[code],signed=False,protested=False,below=False)
        result=render_pc_map(data,pc_defaults(data))
        self.assertEqual([r['category'] for r in result['pc_land_uses']],categories)
        pdf=fitz.open(stream=result['pdf'],filetype='pdf');page=pdf[0]
        drawings=page.get_drawings()
        for code in categories:
            if code=='TRN':continue
            rgb=tuple(int(COLORS[code][i:i+2],16)/255 for i in (1,3,5))
            fills=[d for d in drawings if d['fill'] and max(abs(a-b) for a,b in zip(d['fill'],rgb))<.001]
            self.assertTrue(any(d['rect'].width>20 for d in fills),code+' parcel fill')
            self.assertTrue(any(10<d['rect'].width<12 for d in fills),code+' legend fill')
        # Both hatch patterns must be real vector strokes, not just captions.
        data['matched'][0]['pc']['below']=True
        result=render_pc_map(data,pc_defaults(data));pdf=fitz.open(stream=result['pdf'],filetype='pdf')
        colors=[d['color'] for d in pdf[0].get_drawings() if d['color']]
        for expected in [(85/255,200/255,185/255),(119/255,)*3]:
            self.assertTrue(any(max(abs(a-b) for a,b in zip(color,expected))<.001 for color in colors))

    def test_support_arrows_face_street_on_both_sides_and_after_rotation(self):
        from shapely.geometry import Point, LineString, box
        from shapely.affinity import translate
        from shapely.ops import unary_union
        from core.boundary_map_render import _road_lines
        from wizards.mls.pc_map import support_arrow
        data=copy.deepcopy(self.data)
        south=data['matched'][1]
        polygon=translate(parcel_polygon(south['feature']),yoff=-175)
        south['feature']['geometry']['rings']=[list(polygon.exterior.coords)]
        south['pc'].update(signed=True,protested=False)
        for rotation in (0,37,90):
            settings=pc_defaults(data);settings['rotation']=rotation
            result=render_pc_map(data,settings)
            roads=unary_union([affine_transform(rotate(line,rotation,origin=result['rotation_origin']),result['page_transform']) for line in _road_lines(data,data['case']['street'])])
            symbols=[s for s in result['pc_symbols'] if s['kind']=='support']
            self.assertEqual(len(symbols),4)
            for symbol in symbols:
                x,y=symbol['center'];arrow=support_arrow(x,y,symbol['size'],symbol['angle'])
                tip=Point(list(arrow.exterior.coords)[1])
                self.assertLess(tip.distance(roads),Point(x,y).distance(roads))
                match=next(m for m in data['matched'] if m['property']['row']==symbol['row'])
                parcel=affine_transform(rotate(parcel_polygon(match['feature']),rotation,origin=result['rotation_origin']),result['page_transform'])
                self.assertTrue(parcel.covers(arrow))
                self.assertTrue(box(*symbol['bounds']).covers(arrow))
            if rotation==0:
                self.assertAlmostEqual(symbols[0]['angle'],-90)
                self.assertAlmostEqual(next(s for s in symbols if s['row']==south['property']['row'])['angle'],90)
        data['roads']=[]
        result=render_pc_map(data,pc_defaults(data))
        self.assertTrue(any('arrow direction' in w for w in result['issues']))
        self.assertFalse(result['blocking'])

    def test_map_styles_symbols_save_and_regular_map_unchanged(self):
        normal=render_map(self.data,default_settings(self.data))['pdf']
        settings=pc_defaults(self.data);rendered=render_pc_map(self.data,settings)
        self.assertEqual(len(PdfReader(io.BytesIO(rendered['pdf'])).pages),1)
        self.assertEqual(len(rendered['labels']),8)
        self.assertEqual({s['kind'] for s in rendered['pc_symbols']},{'support','protest'})
        for item in rendered['labels']:
            if item['key'].startswith('parcel:'):
                match=next(m for m in self.data['matched'] if str(m['property']['row'])==item['key'].split(':')[1])
                p=affine_transform(rotate(parcel_polygon(match['feature']),rendered['rotation'],origin=rendered['rotation_origin']),rendered['page_transform'])
                self.assertTrue(p.covers(Polygon(item['corners'])))
        street=next(i for i in rendered['labels'] if i['key'].startswith('street:'))
        settings['street_styles']={street['key']:{'size':13,'angle':30}}
        result=render_pc_map(self.data,settings);changed=next(i for i in result['labels'] if i['key']==street['key'])
        self.assertEqual((changed['size'],changed['angle']),(13,30))
        saved=save_pc_map(self.data,settings,result,self.folder)
        restored=saved_layout(saved,self.data,pc_defaults(self.data))
        self.assertEqual(restored['street_styles'],settings['street_styles'])
        before=PdfReader(io.BytesIO(normal)).pages[0].get_contents().get_data()
        after=PdfReader(io.BytesIO(render_map(self.data,default_settings(self.data))['pdf'])).pages[0].get_contents().get_data()
        self.assertEqual(before,after)  # PDF creation timestamps may differ.

    def test_word_context_and_source_facts_not_carried_over(self):
        ctx,notes=report_context(self.facts,{'agenda':'VII','notice_sent':True,'protested_rows':[18]})
        ctx['attachment_list']='1. PC Map'
        self.assertIn('75%',ctx['sfr_sentence']);self.assertIn('35%',ctx['support_sentence']);self.assertEqual(ctx['qualifying_area_percent'],'90')
        self.assertIn('One (1) written protest',ctx['protest_sentence'])
        self.assertEqual(ctx['qualifying_lots'],'three (3)')
        doc=fill_template(ctx);text='\n'.join(p.text for p in doc.paragraphs)
        self.assertNotIn('{{',text);self.assertNotIn('Bartlett',text);self.assertNotIn('915',text)
        self.assertIn('five (5) lots',text);self.assertIn('eligible application area',text)
        empty,_=report_context(self.facts,{'protest_count':''})
        self.assertIn('Review',empty['protest_sentence']);self.assertIn('[Confirm',empty['notice_sentence'])

    def test_preferences_do_not_leak_to_changed_workbook(self):
        before=self.path.read_bytes();opts={'agenda':'VII','protested_rows':[18]}
        save_preferences(self.path,self.folder,self.facts['workbook_hash'],opts)
        self.assertEqual(load_preferences(self.path,self.folder,self.facts['workbook_hash']),opts)
        self.assertEqual(load_preferences(self.path,self.folder,'changed'),{})
        self.assertEqual(before,self.path.read_bytes())

    def test_generate_area_only_and_pc_map_selection(self):
        import tkinter as tk
        from core import session
        from wizards.mls.wizard import MailoutWizard
        from wizards.mls.pc_ui import PCWorkspace
        root=tk.Tk();root.withdraw()
        council={'value':'H','districts':['H'],'member_value':'Test Member H','warnings':[]}
        with patch.object(session,'load',return_value={}),patch('core.boundary_map.fetch_map_data',side_effect=lambda c,client:fixture(c)),patch('core.council_district.resolve_districts',return_value=council),patch('core.council_district.add_members',side_effect=lambda result,cache:result),patch('wizards.mls.pc_ui.attach_land_use',side_effect=lambda d,c:d),patch('tkinter.messagebox.showerror',side_effect=AssertionError):
            owner=MailoutWizard(root);owner.vars['excel'].set(str(self.path));owner.vars['output_folder'].set(str(self.folder));owner.vars['map'].set('original.pdf')
            def texts(widget):
                result=[]
                for child in widget.winfo_children():
                    try:result.append(child.cget('text'))
                    except tk.TclError:pass
                    result+=texts(child)
                return result
            self.assertNotIn('Generate PC Map',texts(owner));owner._show_step(3)
            self.assertIn('Generate PC Map',texts(owner));self.assertIn('Generate PC Draft Report',texts(owner))
            window=PCWorkspace(owner)
            end=time.monotonic()+20
            while window.facts is None and time.monotonic()<end:root.update();time.sleep(.02)
            self.assertIsNotNone(window.facts)
            window.make_map();end=time.monotonic()+20
            while window.preview.rendered is None and time.monotonic()<end:root.update();time.sleep(.02)
            self.assertIsNotNone(window.preview.rendered,window.preview.status.get());window.preview.accept();root.update()
            self.assertIn('_PC_Map_',window.pc_map.get());self.assertEqual(owner.vars['map'].get(),'original.pdf')
            self.assertEqual(fingerprint(window.source),window.facts['workbook_hash'])
            self.assertEqual('H',window.facts['case_info']['Council_District'])
            self.assertEqual('Test Member H',window.facts['case_info']['Council_Member'])
            window.close();owner.destroy()
        root.destroy()

def main():
    import sys
    if '--office' in sys.argv:
        out=Path(sys.argv[sys.argv.index('--office')+1]);out.mkdir(parents=True,exist_ok=True)
        source,facts,data=case_fixture(out)
        render=render_pc_map(data,pc_defaults(data));pc=save_pc_map(data,pc_defaults(data),render,out)
        from core.boundary_map_render import save_map
        boundary=save_map(data,default_settings(data),render_map(data,default_settings(data)),out)
        options={'agenda':'VII','notice_sent':True,'protested_rows':[18],'pc_map':str(pc),'boundary_map':str(boundary),
                 'include_calculation':True,'application_pdf':str(boundary)}
        before=source.read_bytes();result=generate_pc_report(source,facts,options,out)
        assert result['pdf'],result['notes'];assert source.read_bytes()==before
        reader=PdfReader(result['pdf']);print('OFFICE_REPORT',json.dumps(result));print('PAGES',len(reader.pages))
        assert not result['notes'],result['notes']
        (out/'result.json').write_text(json.dumps(result,indent=2))
        return 0
    result=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(PCReportTests))
    return 0 if result.wasSuccessful() else 1

if __name__=='__main__':raise SystemExit(main())
