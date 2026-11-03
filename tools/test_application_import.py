"""Behavioral tests for application boundaries, field provenance and City xlsx."""
import copy
from datetime import datetime
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch
import zipfile

from lxml import etree as ET
from openpyxl import load_workbook

from core.application_ocr import parse_boundary, extract_cover, table_hints, detected_wizard, designation_suggestions
from core.application_import import (numeric_range, parcel_row, save_calculation,
    record_pdf_facts, suggested_land_use, sheet_paths, block_description, find_parcels, enrich_city_records, N,
    case_number, optional_measurement)
from core.workbooks import calculation_template_path


def fields():
    return {'application_number':'910','street':'Cortlandt Street','block_range':'2000',
      'cross_1':'E 20th Street','cross_2':'E 22nd Street','sides':'east',
      'date_received':'01/20/2026','applicant_name':'Applicant Test','applicant_address':'2014 Cortlandt St',
      'applicant_city':'Houston','applicant_state':'TX','applicant_zip':'77008',
      'planner_name':'Planner Test','council_district':'C'}


def parcel(number=2014,account='0200650000017',area=6599,state='A1'):
    feature={'attributes':{'OBJECTID':number,'HCAD_NUM':account,'site_str_num':number,
      'site_str_name':'CORTLANDT','site_str_sfx':'ST','owner_name_1':'Owner Test',
      'mail_addr_1':'PO BOX 10','mail_city':'HOUSTON','mail_state':'TX','mail_zip':'07708',
      'tax_year':'2025','state_class':state,'land_sqft':area,'legal_dscr_1':'LT 17 BLK 69',
      'legal_dscr_2':'HOUSTON HEIGHTS','activeAccount_flag':'Y'},
      'geometry':{'rings':[[[0,0],[100,0],[100,60],[0,60],[0,0]]]}}
    row=parcel_row(feature);row['include']=True
    return row


class ApplicationImportTests(unittest.TestCase):
    def setUp(self):
        from unittest.mock import patch
        from tools.test_designation_lookup import no_lookup
        external=patch('core.designation_lookup.lookup',side_effect=no_lookup);external.start();self.addCleanup(external.stop)
    def _same_name_regions(self):
        rows=[]
        for i,(x,city) in enumerate(((0,'HOUSTON'),(100000,'BAYTOWN'),(200000,'LA PORTE'))):
            f=parcel(number=1701,account=f'{520000000000+i:013d}')['feature']
            f['attributes'].update(OBJECTID=i+1,site_str_name='MAIN',site_str_pfx='W',site_city=city)
            f['geometry']['rings']=[[[px+x,py] for px,py in ring] for ring in f['geometry']['rings']]
            rows.append(f)
        fields_={**fields(),'street':'W. Main Street','block_range':'1700',
                 'cross_1':'Woodhead Street','cross_2':'Dunlavy Street','sides':'north and south'}
        def roads_at(x):
            paths=[('W MAIN ST',[[x,-10],[x+200,-10]]),
                   ('WOODHEAD ST',[[x,-100],[x,100]]),
                   ('DUNLAVY ST',[[x+200,-100],[x+200,100]])]
            return [{'attributes':{'FULL_NAME':name},'geometry':{'paths':[path]}} for name,path in paths]
        return rows,fields_,roads_at

    def test_same_name_towns_resolved_by_both_cross_streets(self):
        from unittest.mock import MagicMock
        rows,f,roads_at=self._same_name_regions()
        client=MagicMock();client.sources=[]
        client.query.side_effect=[rows,[rows[0]]]
        def local_roads(case,client,extent,warnings):
            self.assertLess(extent[2]-extent[0],12000,'Never query the entire countywide span.')
            return roads_at(0) if extent[0]<50000 else []
        with tempfile.TemporaryDirectory() as tmp,patch('core.application_import.PublicMapClient',return_value=client), \
             patch('core.application_import.load_case_roads',side_effect=local_roads), \
             patch('core.application_import.enrich_city_records'):
            result=find_parcels(f,tmp)
        self.assertEqual([rows[0]['attributes']['HCAD_NUM']],[p['account'] for p in result['parcels'] if p['include']])
        self.assertTrue(any('BAYTOWN' in w and 'LA PORTE' in w and '2 same-name' in w for w in result['warnings']))

    def test_same_name_areas_without_unique_cross_streets_are_not_guessed(self):
        from unittest.mock import MagicMock
        rows,f,roads_at=self._same_name_regions()
        for ambiguous in (False,True):
            with self.subTest(ambiguous=ambiguous):
                client=MagicMock();client.sources=[];client.query.return_value=rows
                def local_roads(case,client,extent,warnings):
                    if not ambiguous:return roads_at(0)[:2] if extent[0]<50000 else []
                    return roads_at(round((extent[0]+160)/100000)*100000)
                with tempfile.TemporaryDirectory() as tmp,patch('core.application_import.PublicMapClient',return_value=client), \
                     patch('core.application_import.load_case_roads',side_effect=local_roads):
                    with self.assertRaisesRegex(ValueError,'one unique area'):
                        find_parcels(f,tmp)
                self.assertEqual(client.query.call_count,1,'Do not import an arbitrary town or load a huge extent.')

    def test_reviewed_save_keeps_selection_with_incomplete_data_all_wizards(self):
        for wizard in ('mls','mbl','mls_renewal','mbl_renewal'):
            with self.subTest(wizard=wizard), tempfile.TemporaryDirectory() as tmp:
                f={**fields(), 'application_number':'', 'cross_2':'Corrected cross street',
                   'date_received':'unreadable', 'original_effective_date':'unreadable',
                   'renewal_minimum':'unreadable'}
                rows=[parcel(area=None,state='UNKNOWN'),parcel(2018,'0200650000018')]
                rows[0].update(building_line='unreadable', original_lot_area='unreadable',
                               original_building_line='unreadable',support='unclear')
                rows[1]['land_use']='EXC'
                data={'fields':fields(),'parcels':rows}
                before=copy.deepcopy(data)
                dest=save_calculation(Path(tmp)/'case.xlsx',f,data,{},wizard=wizard,reviewed=True)
                self.assertEqual(data,before,'Saving must not change the reviewed selection or source facts.')
                w=load_workbook(dest,read_only=True,data_only=False)
                try:
                    s1,s2,ci=w.worksheets
                    headers={c.value:c.column for c in next(s1.iter_rows(max_row=1)) if c.value}
                    self.assertEqual(s1.cell(2,headers['TAX_ID']).value,'0200650000017')
                    self.assertEqual(s1.cell(3,headers['TAX_ID']).value,'0200650000018')
                    self.assertIsNone(s1.cell(2,headers['Total_Land_Area']).value)
                    self.assertIn(s1.cell(2,headers['group_dscr']).value,(None,''))
                    self.assertEqual(s1.cell(3,headers['group_dscr']).value,'Excluded')
                    self.assertIsNone(s2['B2'].value)
                    self.assertIsNone(s2['B3'].value)
                    self.assertEqual(s2['D5'].value,'Corrected cross street')
                finally:w.close()
                audit=json.loads(dest.with_name(dest.stem+'_sources').joinpath('Import review.json').read_text())
                self.assertTrue(audit['planner_accepted_selection'])
                self.assertEqual(audit['submitted_fields']['date_received'],'unreadable')
                self.assertEqual(len([p for p in audit['parcels'] if p['include']]),2)
                self.assertTrue(any('wording changed' in n for n in audit['warnings']))
                self.assertTrue(dest.with_name(dest.stem+'_sources').joinpath('Review notes.txt').is_file())

    def test_save_button_accepts_changed_boundary_and_blank_case_all_wizards(self):
        from core import application_import_ui as ui
        for wizard in ('mls','mbl','mls_renewal','mbl_renewal'):
            with self.subTest(wizard=wizard),tempfile.TemporaryDirectory() as tmp:
                f={**fields(),'application_number':'','cross_2':''}
                dest=Path(tmp)/'from_button.xlsx'
                dialog=SimpleNamespace(busy=False,data={'fields':fields(),'parcels':[parcel(area=None,state='UNKNOWN')]},
                    case_reviewed=SimpleNamespace(get=lambda:False),parcels_reviewed=SimpleNamespace(get=lambda:True),
                    values=lambda:f,wizard=wizard,wizard_name=wizard,application={},
                    owner=SimpleNamespace(files={'out':tmp}),close=lambda:None,
                    progress=lambda message:None,cancel=None,start=lambda work,done:done(work()))
                with (patch.object(ui.filedialog,'asksaveasfilename',return_value=str(dest)), 
                     patch.object(ui.messagebox,'showinfo') as info, 
                     patch.object(ui.messagebox,'showerror',side_effect=AssertionError('Save was blocked')), 
                     patch.object(ui,'select_imported_calculation') as select):
                    ui.ApplicationImport.save(dialog)
                self.assertTrue(dest.is_file())
                select.assert_called_once_with(dialog.owner,dest)
                self.assertEqual(info.call_args.args[0],'Calculation Sheet Created')

    def test_current_form_selects_main_applicant_and_ignores_printed_boundary(self):
        words=[]
        def add(y,text,x=.065,w=.13):words.append({'text':text,'x':x,'y':y-.006,'w':w,'h':.012})
        add(.178,'Describe the location of the application area.',w=.5)
        add(.196,'Example: 1600-1700 block of Planning St, north and south sides, between Renew St and Welcome Ave.',w=.85)
        add(.220,'600-700 blocks of W Cottage St, north and south sides,',w=.75)
        add(.245,'between Watson St and Moss St',w=.45)
        add(.271,'Date and name of Planner at your Preliminary Meeting:',w=.7)
        add(.386,'MainApplicant (Required)',x=.38,w=.3)
        add(.416,"Applicant's Name:",w=.142);add(.416,'Main Person',x=.235,w=.2)
        add(.5,'Address:',w=.068);add(.5,'631 W Cottage St, Houston TX 77009',x=.155,w=.55)
        add(.578,'Co - Applicant (Optional)',x=.38,w=.3)
        add(.606,"Applicant's Name:",w=.142);add(.606,'Other Person',x=.235,w=.2)
        add(.795,'File Number:',w=.094);add(.795,'93REN',x=.2,w=.08)
        got=extract_cover(words)
        self.assertEqual(got['applicant_name'],'Main Person')
        self.assertEqual(got['applicant_city'],'Houston')
        self.assertEqual(got['cross_1'],'Watson St');self.assertEqual(got['cross_2'],'Moss St')
        self.assertEqual(got['application_number'],'93REN')
        blank=[w for w in words if not .21<w['y']<.25]
        self.assertNotIn('street',extract_cover(blank))

    def test_recalculation_prefers_excel_and_keeps_original_unchanged(self):
        from core.recalc import recalculated_copy
        with tempfile.TemporaryDirectory() as tmp:
            src=save_calculation(Path(tmp)/'case.xlsx',fields(),{'parcels':[parcel()]},{})
            original=src.read_bytes()
            def excel(source,dest):
                self.assertNotEqual(Path(source).parent,Path(dest).parent)
                Path(dest).write_bytes(b'calculated test result');return Path(dest)
            with patch('core.recalc.find_soffice',return_value='soffice'), \
                 patch('core.office_convert.recalc_with_excel',side_effect=excel), \
                 patch('core.recalc.subprocess.run',side_effect=AssertionError('Excel already succeeded')):
                got=recalculated_copy(src)
                self.assertEqual(got.read_bytes(),b'calculated test result')
            self.assertEqual(src.read_bytes(),original)

    def test_libreoffice_output_must_not_be_its_staged_input(self):
        from core.recalc import recalculated_copy
        with tempfile.TemporaryDirectory() as tmp:
            src=save_calculation(Path(tmp)/'case.xlsx',fields(),{'parcels':[parcel()]},{})
            def convert(command,**kwargs):
                source=Path(command[-1]);out=Path(command[command.index('--outdir')+1])/source.name
                self.assertNotEqual(source,out)
                out.write_bytes(b'converted test result')
            with patch('core.recalc.find_soffice',return_value='soffice'), \
                 patch('core.office_convert.recalc_with_excel',return_value=None), \
                 patch('core.recalc.subprocess.run',side_effect=convert):
                got=recalculated_copy(src)
                self.assertEqual(got.read_bytes(),b'converted test result')

    def test_each_wizard_uses_its_template_and_named_case_fields(self):
        for wizard in ('mls','mbl','mls_renewal','mbl_renewal'):
            with self.subTest(wizard=wizard),tempfile.TemporaryDirectory() as tmp:
                f={**fields(),'renewal_minimum':'6,000' if wizard.startswith('mls') else '25',
                   'original_ordinance_no':'2006-0123','original_effective_date':'02/01/2006'}
                row=parcel(account='0071050000002');row.update(support='Y',building_line=22,
                    original_building_line=20,original_lot_area=6200)
                source=calculation_template_path(wizard);original=source.read_bytes()
                dest=save_calculation(Path(tmp)/'case.xlsx',f,{'parcels':[row]}, {},wizard=wizard)
                from core.recalc import needs_recalc
                self.assertTrue(needs_recalc(dest),'New workbooks with shared formulas must trigger recalculation.')
                self.assertEqual(source.read_bytes(),original)
                with zipfile.ZipFile(source) as a,zipfile.ZipFile(dest) as b:
                    paths=sheet_paths(a)
                    for name in a.namelist():
                        if name not in paths.values() and name!='xl/workbook.xml':self.assertEqual(a.read(name),b.read(name),name)
                    for name in paths.values():
                        x,y=ET.fromstring(a.read(name)),ET.fromstring(b.read(name))
                        for c in x.findall('.//s:c[s:f]',N):
                            ref=c.get('r')
                            if 'sheet2' in name and not wizard.endswith('_renewal') and ref in ('B8','D8'):continue
                            self.assertEqual(c.find('s:f',N).text,y.find(f'.//s:c[@r="{ref}"]/s:f',N).text,(wizard,name,ref))
                w=load_workbook(dest,read_only=True,data_only=False)
                try:
                    s1,s2,ci=w.worksheets
                    self.assertEqual(s1['AX2'].value,'0071050000002');self.assertIsNone(s1['AX3'].value)
                    self.assertEqual(s1['A2'].value,2014)
                    values={row[0].value:row[1].value for row in ci.iter_rows(max_col=2)}
                    self.assertEqual(values['Applicant_Name'],'Applicant Test')
                    self.assertEqual(values['Planner_Name'],'Planner Test')
                    self.assertEqual(values['Council_District'],'C')
                    self.assertEqual(values['Applicant_Address_1'],'2014 Cortlandt St')
                    self.assertIsNone(values['Planner_Phone']);self.assertIsNone(values['Sign_Count_Number'])
                    self.assertEqual(s2['B2'].value,'910REN' if wizard.endswith('_renewal') else '910')
                    self.assertEqual(s2['D4'].value,'2000 block of Cortlandt Street')
                    if wizard.endswith('_renewal'):
                        self.assertIn('RENEWAL',s2['A1'].value)
                        self.assertEqual(s2['D2'].value,6000 if wizard.startswith('mls') else 25)
                        self.assertEqual(values['Original_Ordinance_No'],'2006-0123')
                        self.assertEqual(values['Original_Effective_Date'],datetime(2006,2,1))
                    if wizard=='mbl':
                        self.assertEqual(s2['E17'].value,22);self.assertIsNone(s2['E18'].value)
                        self.assertEqual(s2['C17'].value,'Y');self.assertIn('$Z99',s2['D8'].value)
                    elif wizard=='mls_renewal':
                        self.assertEqual(s2['D15'].value,6200);self.assertIsNone(s2['D16'].value)
                        self.assertEqual(s2['E15'].data_type,'f');self.assertIsNone(s2['C15'].value)
                    elif wizard=='mbl_renewal':
                        self.assertEqual(s2['C15'].value,20);self.assertEqual(s2['D15'].value,22)
                        self.assertIsNone(s2['D16'].value);self.assertEqual(s2['E15'].data_type,'f')
                finally:w.close()
                audit=json.loads(dest.with_name('case_sources').joinpath('Import review.json').read_text())
                self.assertEqual(audit['wizard'],wizard)

    def test_missing_historical_values_never_inherit_examples_or_current_hcad(self):
        for wizard in ('mbl','mls_renewal','mbl_renewal'):
            with self.subTest(wizard=wizard),tempfile.TemporaryDirectory() as tmp:
                dest=save_calculation(Path(tmp)/'case.xlsx',fields(),{'parcels':[parcel()]},{},wizard=wizard)
                w=load_workbook(dest,read_only=True)
                try:
                    s=w.worksheets[1]
                    if wizard.endswith('_renewal'):self.assertIsNone(s['D2'].value)
                    for ref in ({'mbl':['E17','E18'],'mls_renewal':['D15','D16'],
                                 'mbl_renewal':['C15','D15','D16']}[wizard]):self.assertIsNone(s[ref].value)
                finally:w.close()
                audit=json.loads(dest.with_name('case_sources').joinpath('Import review.json').read_text())
                self.assertTrue(audit['warnings'])

    def test_renewal_numbers_units_and_invalid_measurements(self):
        self.assertEqual(case_number('MLS 247 ren','mls_renewal'),'247REN')
        self.assertEqual(case_number('93','mbl_renewal'),'93REN')
        for val,key in [('MBL 12','mls'),('247REN','mls'),('not a number','mbl')]:
            with self.assertRaises(ValueError):case_number(val,key)
        self.assertIsNone(optional_measurement('','measurement'))
        self.assertEqual(optional_measurement('0','measurement'),0)
        for val in ['-1','nan','inf','twenty']:
            with self.assertRaises(ValueError):optional_measurement(val,'measurement')

    def test_other_templates_cannot_silently_omit_properties_after_row_83(self):
        rows=[parcel(account=f'{71050000001+i:013d}') for i in range(84)]
        for key in ('mbl','mls_renewal','mbl_renewal'):
            with self.subTest(wizard=key),tempfile.TemporaryDirectory() as tmp:
                with self.assertRaisesRegex(ValueError,'83 properties'):
                    save_calculation(Path(tmp)/'case.xlsx',fields(),{'parcels':rows},{},wizard=key)

    def test_form_type_and_minimum_suggestions_do_not_use_lot_areas(self):
        self.assertEqual(detected_wizard('Special Minimum Lot Size Block PLANNING Renew RECEIVED Application to DEVELOPMENT'),'mls_renewal')
        for family,phrase in [('mls','Lot Size'),('mbl','Building Line')]:
            self.assertEqual(detected_wizard('Application for Special Minimum '+phrase+' Block'),family)
            self.assertEqual(detected_wizard('Application for Special Minimum '+phrase+' Block Renewal'),family+'_renewal')
        p={'page':2,'text':'HCAD\nLand Area: 5,000 Square Feet\nTax Account 0621310000001'}
        self.assertEqual(designation_suggestions([p])[0],{})
        p['text']='Special Minimum Lot Size Renewal\n5,950 Square Feet\nReference: MLS 247REN\nSpecial Minimum Lot Size Boundary'
        self.assertEqual(designation_suggestions([p])[0],{'minimum_lot_size':'5950'})
        f,s,w=designation_suggestions([p,{'page':3,'text':p['text'].replace('5,950','6,000')}])
        self.assertEqual(f,{});self.assertTrue(w)
        self.assertEqual(parse_boundary('Example: 1600-1700 block of Planning St, north and south sides, between Renew St and Welcome Ave.'),{})

    def test_large_review_area_uses_bounded_land_area_requests(self):
        from unittest.mock import MagicMock
        from urllib.parse import urlencode
        rows=[parcel(number=2000+i,account=f'{621310000001+i:013d}') for i in range(111)]
        client=MagicMock()
        def query(service,where,fields):
            self.assertLess(len(urlencode({'where':where,'returnIdsOnly':'true','f':'json'})),1900)
            accounts=where.partition('(')[2].rstrip(')').replace("'",'').split(',')
            return [{'attributes':{'HCAD_NUM':a,'Total_Land_Area':5616,'Tax_Year':'2024'}} for a in accounts]
        client.query.side_effect=query
        enrich_city_records(rows,client)
        self.assertEqual(client.query.call_count,3)
        self.assertTrue(all(r['land_area']==5616 and r['secondary_attributes'] for r in rows))

    def test_public_accounts_with_multiple_leading_zeros_reach_step1_unchanged(self):
        row=parcel(account='0071050000002')
        self.assertEqual(row['account'],'0071050000002')
        with tempfile.TemporaryDirectory() as tmp:
            dest=save_calculation(Path(tmp)/'zeros.xlsx',fields(),{'parcels':[row]}, {})
            w=load_workbook(dest,read_only=True)
            try:
                self.assertEqual(w.worksheets[0]['AX2'].value,'0071050000002')
                self.assertEqual(w.worksheets[0]['AX2'].data_type,'s')
            finally:w.close()

    def test_bad_public_records_warn_without_aborting_valid_parcels(self):
        selected=parcel()['feature'];bad=parcel(number=2048)['feature'];bad['attributes']['HCAD_NUM']='bad data'
        nearby=parcel(number=109,account='0071050000003')['feature']
        missing=parcel(number=113)['feature'];missing['attributes']['HCAD_NUM']=''
        from unittest.mock import MagicMock
        client=MagicMock();client.sources=[]
        client.query.side_effect=[[selected,bad],[selected,bad,nearby,missing]]
        with tempfile.TemporaryDirectory() as tmp,patch('core.application_import.PublicMapClient',return_value=client), \
                patch('core.application_import.load_case_roads',return_value=[]),patch('core.application_import.enrich_city_records'):
            result=find_parcels({**fields(),'sides':'north and south'},tmp)
            self.assertEqual({r['account'] for r in result['parcels']},{'0200650000017','0071050000003'})
            self.assertEqual(sum(r['include'] for r in result['parcels']),1)
            self.assertEqual(len(result['unresolved_records']),2)
            self.assertTrue(any(r['in_requested_street_range'] for r in result['unresolved_records']))
            self.assertTrue(any('2048 CORTLANDT' in s and 'needs review' in s for s in result['warnings']))
            dest=save_calculation(Path(tmp)/'review.xlsx',fields(),result,{})
            audit=json.loads(dest.with_name('review_sources').joinpath('Import review.json').read_text())
            self.assertEqual(len(audit['unresolved_records']),2)

    def test_hcad_record_requires_labelled_account_and_keeps_leading_zero(self):
        import fitz
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'record.pdf'
            with fitz.open() as doc:
                page=doc.new_page();page.insert_text((40,50),'Account: 0852840000003\nLand 6,490 SF')
                doc.save(path)
            got=record_pdf_facts(path)
            self.assertEqual(got['account'],'0852840000003');self.assertEqual(got['land_area'],6490)

    def test_no_results_pdf_cannot_be_matched_by_filename(self):
        import fitz
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'HCAD_0852840000003.pdf'
            with fitz.open() as doc:
                page=doc.new_page();page.insert_text((40,50),'HCAD Property Search\nNo Results Found')
                doc.save(path)
            with patch('core.application_ocr.read_application',side_effect=AssertionError('No OCR needed')):
                got=record_pdf_facts(path)
            self.assertEqual(got['account'],'');self.assertIsNone(got['land_area'])

    def test_small_print_account_uses_ocr_region_not_filename(self):
        import fitz
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'HCAD_wrong_0852840000028.pdf'
            with fitz.open() as doc:doc.new_page();doc.save(path)
            with patch('core.application_ocr.read_application',return_value={'pages':[{'text':'Account: Land 6,490 SF'}]}), \
                 patch('core.application_ocr.read_small_account_text',return_value='Account: 0852840000003'):
                got=record_pdf_facts(path)
            self.assertEqual(got['account'],'0852840000003');self.assertEqual(got['land_area'],6490)

    def test_boundary_range_not_side_conjunction(self):
        got=parse_boundary('2200 and 2300 blocks of Rosedale Street, north and south sides, between Hutchins Street and Emancipation Avenue')
        self.assertEqual(got['block_range'],'2200-2300')
        self.assertEqual(got['sides'],'north and south')
        self.assertEqual(got['cross_2'],'Emancipation Avenue')
        self.assertEqual(numeric_range('2200-2300'),(2200,2399))
        self.assertEqual(numeric_range('4806-4834'),(4806,4834))
        with self.assertRaises(ValueError):numeric_range('2300-2200')
        self.assertEqual(block_description({'block_range':'4806-4834','street':'Edfield Street'}),'4800 block of Edfield Street')

    def test_form_examples_are_not_application_boundary(self):
        self.assertEqual(parse_boundary('North side of Golden Retriever Drive between Boxer and Schnauzer Streets'),{})

    def test_compass_cross_streets_are_not_application_sides(self):
        got=parse_boundary('2000 block of Cortlandt Street east side between E 20TH Street and E 22ND Street')
        self.assertEqual(got['sides'],'east');self.assertEqual(got['street'],'Cortlandt Street')

    def test_application_discrepancy_does_not_overwrite_hcad(self):
        row=parcel()
        got=parcel_row(row['feature'],[{'account':row['account'],'area':6600,'number':'2014'}])
        self.assertEqual(got['land_area'],6599)
        self.assertTrue(any('6,600' in s and '6,599' in s for s in got['notes']))

    def test_unknown_use_and_zero_area_are_not_invented(self):
        row=parcel(area=0,state='B2')
        self.assertIsNone(row['land_area']);self.assertEqual(row['land_use'],'')
        self.assertEqual(suggested_land_use({'land_use':'1001'}),'')

    def test_city_template_and_all_artwork_preserved(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'MLS_910.xlsx'
            source=calculation_template_path('mls');original=source.read_bytes()
            save_calculation(path,fields(),{'parcels':[parcel()],'sources':[]},{})
            self.assertEqual(source.read_bytes(),original)
            with zipfile.ZipFile(source) as a,zipfile.ZipFile(path) as b:
                paths=sheet_paths(a)
                for name in a.namelist():
                    if name not in paths.values() and name!='xl/workbook.xml':
                        self.assertEqual(a.read(name),b.read(name),name)
                for name in paths.values():
                    x,y=ET.fromstring(a.read(name)),ET.fromstring(b.read(name))
                    for c in x.findall('.//s:c[s:f]',N):
                        ref=c.get('r')
                        if name==paths['Step 2 - Calculation Sheet'] and ref in ('B8','D8'):continue
                        self.assertEqual(c.find('s:f',N).text,y.find(f'.//s:c[@r="{ref}"]/s:f',N).text,(name,ref))
            w=load_workbook(path,read_only=True,data_only=False)
            try:
                s=w['Step 1 - Input Properties Info'];c=w['Case Info'];step=w['Step 2 - Calculation Sheet']
                self.assertEqual(s['AX2'].value,'0200650000017');self.assertEqual(s['Q2'].value,'07708')
                self.assertEqual(s['AO2'].value,6599);self.assertEqual(s['AI2'].value,'SINGLE-Family Residential')
                self.assertNotIn('HCAD_NUM_1',[c.value for c in s[1]])
                self.assertEqual([s.cell(1,n).value for n in (54,55,56)],
                                 ['Property Description','Land Use','Lot Size'])
                self.assertEqual(s['BC2'].data_type,'f');self.assertEqual(s['BD2'].value,'=AO2')
                self.assertIsNone(s['AX3'].value);self.assertIsNone(s['E3'].value)
                self.assertEqual(c['B19'].value,'Applicant Test');self.assertEqual(c['B58'].data_type,'f')
                self.assertEqual(step['B3'].value,datetime(2026,1,20));self.assertIsNone(step['D3'].value)
                self.assertEqual(step['D4'].value,'2000 block of Cortlandt Street')
                self.assertTrue(all(step.cell(n,3).value is None for n in range(17,126)))
                self.assertIn('$D125',step['B8'].value);self.assertIn('$Y125',step['D8'].value)
            finally:w.close()

    def test_excluded_parcel_stays_in_step1_and_support_is_explicit(self):
        row=parcel();row.update(land_use='EXC',support='Y')
        with tempfile.TemporaryDirectory() as tmp:
            path=save_calculation(Path(tmp)/'test.xlsx',fields(),{'parcels':[row]}, {})
            w=load_workbook(path,read_only=True)
            try:
                self.assertEqual(w.worksheets[0]['AX2'].value,row['account'])
                self.assertEqual(w.worksheets[0]['AI2'].value,'Excluded')
                self.assertEqual(w.worksheets[1]['C17'].value,'Y')
            finally:w.close()

    def test_no_overwrite_or_duplicate_accounts(self):
        with tempfile.TemporaryDirectory() as tmp:
            dest=Path(tmp)/'existing.xlsx';dest.write_bytes(b'original')
            with self.assertRaises(ValueError):save_calculation(dest,fields(),{'parcels':[parcel()]},{})
            self.assertEqual(dest.read_bytes(),b'original')
            with self.assertRaises(ValueError):save_calculation(Path(tmp)/'duplicate.xlsx',fields(),{'parcels':[parcel(),parcel()]},{})
            self.assertFalse((Path(tmp)/'duplicate.xlsx').exists())

    def test_source_text_cannot_become_excel_formula(self):
        f=fields();f['applicant_name']='=HYPERLINK("https://example.invalid")'
        with tempfile.TemporaryDirectory() as tmp:
            dest=save_calculation(Path(tmp)/'literal.xlsx',f,{'parcels':[parcel()]},{})
            w=load_workbook(dest,read_only=True)
            try:self.assertEqual(w['Case Info']['B19'].data_type,'s')
            finally:w.close()

    def test_missing_required_data_cannot_produce_misleading_workbook(self):
        with tempfile.TemporaryDirectory() as tmp:
            for row in [parcel(area=0),parcel(state='UNKNOWN')]:
                with self.assertRaises(ValueError):save_calculation(Path(tmp)/'bad.xlsx',fields(),{'parcels':[row]},{})
            self.assertFalse((Path(tmp)/'bad.xlsx').exists())


if __name__=='__main__':unittest.main(verbosity=2)
