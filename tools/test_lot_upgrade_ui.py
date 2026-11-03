import copy,tempfile,unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import tkinter as tk
from PIL import Image
from shapely.geometry import box
from core.lot_map_ui import LotMapEditor
from core.lot_review_ui import LotReview
from core.lot_divisions import lot_key,interior_lines

class LotUITests(unittest.TestCase):
    def setUp(self):
        from tools.test_designation_lookup import no_lookup
        external=patch('core.designation_lookup.lookup',side_effect=no_lookup);external.start();self.addCleanup(external.stop)
        self.root=tk.Tk();self.root.withdraw();self.addCleanup(self.root.destroy)
    def test_plat_alignment_split_assignment_and_save_in_preview(self):
        p=tk.Toplevel(self.root);p.withdraw();p.settings={};p._loading=False;p.changed=lambda:None;p.redraw=lambda:None
        rows=[{'id':lot_key('123',f'LT {n} BLK 1'),'legal_description':f'LT {n} BLK 1'} for n in (1,2)]
        p.data={'case':{},'matched':[{'property':{'address':'100 TEST ST','account':'123','lot_rows':rows},'feature':{'geometry':{'rings':[[[0,0],[100,0],[100,60],[0,60],[0,0]]]}}}]}
        editor=LotMapEditor(p);self.root.update_idletasks()
        editor.reference=Image.new('RGB',(400,300),'white');editor.draw_reference();editor.start_alignment()
        for world,pixel in [([0,0],[10,290]),([100,0],[390,290]),([0,60],[10,10])]:
            x,y=editor.xy(world);editor.map_click(SimpleNamespace(x=x,y=y));ox,oy=editor.reference_offset
            editor.plat_click(SimpleNamespace(x=ox+pixel[0]*editor.reference_scale,y=oy+pixel[1]*editor.reference_scale))
        self.assertIsNotNone(editor.alignment)
        editor.start_line()
        for q in ([50,-2],[50,62]):x,y=editor.xy(q);editor.map_click(SimpleNamespace(x=x,y=y))
        self.assertEqual(2,len(editor.polys))
        for i in (0,1):editor.piece.current(i);editor.identity.current(i);editor.assign()
        editor.source.set('Test plat, block 1');editor.reviewed.set(True)
        self.assertTrue(editor.apply());lines,notes=interior_lines(p.data,p.settings);self.assertEqual(1,len(lines));self.assertFalse(notes)
        editor.show.set(False);editor.toggle();self.assertFalse(p.settings['show_platted_lot_lines'])
        editor.destroy();p.destroy()
    def test_legacy_lot_editor_preserves_cancel_with_allocated_areas(self):
        rows=[{'account':'123','address':'100 TEST','legal_description':'LTS 1 & 2 BLK 1','land_area':6000,'land_use':'SFR'}]
        editor=LotReview(self.root,rows,lambda:False);self.root.update_idletasks()
        self.assertEqual(2,len(editor.entries));self.assertEqual([3000,3000],[l['area'] for l in editor.entries])
        editor.source.set('')
        with patch('tkinter.messagebox.showwarning') as warning:editor.accept_property();self.assertTrue(warning.called)
        self.assertNotIn('lots',rows[0])
        for l in editor.entries:l['area']=3000
        editor.source.set('Verified source');editor.accept_property();self.assertEqual(2,len(rows[0]['lots']))
        editor.done();self.assertTrue(editor.winfo_exists());editor.destroy()

    def test_disconnected_property_parts_can_be_reviewed(self):
        p=tk.Toplevel(self.root);p.withdraw();p.settings={};p._loading=False;p.changed=lambda:None;p.redraw=lambda:None
        rows=[{'id':str(n),'legal_description':f'LT {n} BLK 1'} for n in (1,2)]
        p.data={'case':{},'matched':[{'property':{'address':'100 TEST','account':'123','lot_rows':rows},'feature':{'geometry':{'rings':[list(box(0,0,20,40).exterior.coords),list(box(30,0,50,40).exterior.coords)]}}}]}
        editor=LotMapEditor(p);self.root.update_idletasks();self.assertEqual(2,len(editor.polys))
        for i in (0,1):editor.piece.current(i);editor.identity.current(i);editor.assign()
        editor.source.set('Verified disjoint lots');editor.reviewed.set(True);self.assertTrue(editor.apply())
        editor.remove();self.assertEqual(2,len(editor.polys));editor.destroy();p.destroy()

    def test_all_four_wizards_use_automatic_lots_and_council(self):
        import importlib
        from core import session
        from core.application_import_ui import ApplicationImport
        def descendants(w):
            for c in w.winfo_children():yield c;yield from descendants(c)
        with tempfile.TemporaryDirectory() as tmp,patch.object(session,'SETTINGS_PATH',Path(tmp)/'session.json'),patch.object(session,'load',return_value={}):
            for key,cls in [('mls','MailoutWizard'),('mbl','MBLWizard'),('mls_renewal','MLSRenewalWizard'),('mbl_renewal','MBLRenewalWizard')]:
                owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),cls)(self.root)
                texts=[w.cget('text') for w in descendants(owner) if 'text' in w.keys()]
                self.assertNotIn('Review Individual Lots',texts);self.assertNotIn('Find District / Councilmember',texts)
                self.assertIn('Create Boundary Map',texts);self.assertIn('Create Tax Map',texts)
                dialog=ApplicationImport(owner);dialog.withdraw();self.root.update_idletasks()
                self.assertIn('council_member',dialog.fields)
                texts=[w.cget('text') for w in descendants(dialog) if 'text' in w.keys()]
                self.assertNotIn('Review individual lots',texts);self.assertNotIn('Council district / member',texts)
                self.assertIn('Save New Calculation Sheet',texts);dialog.close();owner.destroy()

    def test_save_button_splits_three_lots_without_review_or_council_prompt(self):
        import importlib,time
        from core import session
        from core.application_import_ui import ApplicationImport
        from core.lot_divisions import workbook_rows
        from tools.test_application_import import fields,parcel
        def fail(*args,**kwargs):raise AssertionError(str(args))
        with tempfile.TemporaryDirectory() as tmp,patch.object(session,'SETTINGS_PATH',Path(tmp)/'session.json'),patch.object(session,'load',return_value={}):
            for key,cls in [('mls','MailoutWizard'),('mbl','MBLWizard'),('mls_renewal','MLSRenewalWizard'),('mbl_renewal','MBLRenewalWizard')]:
                owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),cls)(self.root)
                if hasattr(owner,'vars'):owner.vars['output_folder'].set(tmp)
                else:owner.files['out']=tmp
                dialog=ApplicationImport(owner);dialog.withdraw()
                parent=parcel(area=15000);parent['legal_description']='LTS 1, 2, 3 BLK 69'
                dialog.data={'parcels':[parent],'fields':fields(),'roads':[],'warnings':[],'sources':[]};dialog.application={}
                for name,value in fields().items():
                    if name in dialog.fields:dialog.fields[name].set(value)
                if key.endswith('renewal'):dialog.fields['renewal_minimum'].set('3000' if key.startswith('mls') else '25')
                dialog.refresh_rows();self.assertIn('3 Step 1 rows',dialog.count.cget('text'))
                self.assertEqual('3',str(dialog.tree.set(parent['account'],'lots')))
                dialog.parcels_reviewed.set(True)
                destination=Path(tmp)/(key+'.xlsx')
                with patch('tkinter.filedialog.asksaveasfilename',return_value=str(destination)),patch('tkinter.messagebox.showinfo'),patch('tkinter.messagebox.showerror',side_effect=fail),patch('core.council_district.resolve_selected',side_effect=AssertionError('Council lookup belongs to map creation')):
                    dialog.save_button.invoke();deadline=time.monotonic()+30
                    while not dialog.closed and time.monotonic()<deadline:self.root.update();time.sleep(.02)
                    self.assertTrue(dialog.closed,dialog.status.get())
                rows,_=workbook_rows(destination);self.assertEqual([5000]*3,[r['land_area'] for r in rows])
                selected=owner.vars['excel'].get() if hasattr(owner,'vars') else owner.files['workbook']
                self.assertEqual(str(destination),selected);owner.destroy()

    def test_manual_boundary_finds_parcels_and_saves_without_application_pdf(self):
        import importlib,json
        from core import session,application_import_ui as ui
        from core.lot_divisions import workbook_rows
        from tools.test_application_import import fields,parcel
        with tempfile.TemporaryDirectory() as tmp,patch.object(session,'SETTINGS_PATH',Path(tmp)/'session.json'),patch.object(session,'load',return_value={}):
            for key,cls in [('mls','MailoutWizard'),('mbl','MBLWizard'),('mls_renewal','MLSRenewalWizard'),('mbl_renewal','MBLRenewalWizard')]:
                with self.subTest(wizard=key):
                    owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),cls)(self.root)
                    dialog=ui.ApplicationImport(owner)
                    def run(work,callback):
                        result=work();dialog._saving=False;callback(result)
                    with patch.object(dialog,'start',side_effect=run),patch.object(ui,'read_application',side_effect=AssertionError('No PDF/OCR/AI required')):
                        dialog.case_reviewed.set(True)
                        with patch.object(ui,'find_parcels') as find,patch('tkinter.messagebox.showerror') as error:
                            dialog.find_button.invoke();find.assert_not_called();error.assert_called_once()
                        for name in ('application_number','street','block_range','cross_1','cross_2','sides'):
                            dialog.fields[name].set(fields()[name])
                        if key.endswith('renewal'):dialog.fields['renewal_minimum'].set('5000' if key.startswith('mls') else '25')
                        parent=parcel(area=15000);parent['legal_description']='LTS 1, 2, 3 BLK 69'
                        parent.update(building_line=25,original_building_line=25,original_lot_area=15000)
                        data={'parcels':[parent],'fields':dialog.values(),'roads':[],'warnings':[],'sources':[]}
                        self.assertFalse(dialog.case_reviewed.get())
                        with patch.object(ui,'find_parcels',return_value=data) as find:
                            dialog.find_button.invoke();find.assert_called_once()
                            self.assertEqual([],find.call_args.args[2])
                        self.assertIsNone(dialog.pdf)
                        self.assertEqual('',dialog.application['source_pdf'])
                        self.assertIn('3 Step 1 rows',dialog.count.cget('text'))
                        destination=Path(tmp)/(key+'_manual.xlsx')
                        dialog.parcels_reviewed.set(True)
                        with patch('tkinter.filedialog.asksaveasfilename',return_value=str(destination)),patch('tkinter.messagebox.showinfo'):
                            dialog.save_button.invoke()
                        self.assertTrue(destination.is_file())
                        rows,_=workbook_rows(destination)
                        self.assertEqual([5000]*3,[row['land_area'] for row in rows])
                        sources=destination.with_name(destination.stem+'_sources')
                        self.assertFalse((sources/'Application.pdf').exists())
                        audit=json.loads((sources/'Import review.json').read_text(encoding='utf-8'))
                        self.assertIn('Manual boundary',audit['input_method'])
                        self.assertEqual('',audit['application_sha256'])
                        selected=owner.vars['excel'].get() if hasattr(owner,'vars') else owner.files['workbook']
                        self.assertEqual(str(destination),selected)
                    if not dialog.closed:dialog.close()
                    owner.destroy()

def main():return 0 if unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(LotUITests)).wasSuccessful() else 1
if __name__=='__main__':raise SystemExit(main())
