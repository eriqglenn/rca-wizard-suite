"""Exercise application selection, reviewed HCAD lookup, save and wizard handoff."""
from pathlib import Path
import argparse
import copy
import importlib
import json
import time
import tkinter as tk
from tkinter import ttk
from unittest.mock import patch

from core import application_import_ui as ui, session
from core.application_import import find_parcels
from core.boundary_map import PublicMapClient


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--cases',type=Path,required=True)
    parser.add_argument('--out',type=Path,required=True)
    parser.add_argument('--cache',type=Path,required=True)
    args=parser.parse_args();args.out=args.out.resolve();args.out.mkdir(parents=True,exist_ok=True)
    cases=json.loads(args.cases.read_text(encoding='utf8'))
    classes={'mls':'MailoutWizard','mbl':'MBLWizard','mls_renewal':'MLSRenewalWizard','mbl_renewal':'MBLRenewalWizard'}
    root=tk.Tk();root.withdraw();results=[];map_cases=[]
    def fail_dialog(title,message,**kwargs):raise AssertionError(f'{title}: {message}')
    def cached_find(fields,cache,hints,progress,cancel):
        return find_parcels(fields,args.cache,hints,progress,cancel,use_cache=True)
    def wait(dialog,ready):
        deadline=time.monotonic()+180
        while not ready() and time.monotonic()<deadline:
            root.update();time.sleep(.02)
        assert ready(), 'Application import operation timed out: '+dialog.status.get()
    try:
        with patch.object(session,'SETTINGS_PATH',args.out/'test-session.json'),patch.object(session,'load',return_value={}), \
             patch.object(ui.messagebox,'showerror',side_effect=fail_dialog),patch.object(ui.messagebox,'showinfo'), \
             patch.object(ui,'find_parcels',side_effect=cached_find), \
             patch.object(PublicMapClient,'request',side_effect=AssertionError('Missing cached public data')):
            for case in cases:
                key=case['wizard'];owner=None;dialog=None
                try:
                    owner=getattr(importlib.import_module('wizards.'+key+'.wizard'),classes[key])(root)
                    owner.pack(fill='both',expand=True)
                    buttons=[w for w in descendants(owner) if isinstance(w,(tk.Button,ttk.Button)) and w.cget('text')=='Import Application / Enter Boundary / Fill Calculation Sheet']
                    assert len(buttons)==1, 'Exactly one importer button must appear in Required setup.'
                    buttons[0].invoke();dialog=owner._application_import_window;dialog.withdraw()
                    assert dialog.wizard==key and ui.wizard_label(key) in dialog.title()
                    with patch.object(ui.filedialog,'askopenfilename',return_value=case['pdf']):dialog.choose_pdf()
                    wait(dialog,lambda:dialog.application is not None and not dialog.busy)
                    assert dialog.application.get('detected_wizard')==key
                    for name,value in case['expected'].items():
                        assert dialog.fields[name].get().casefold()==value.casefold(),(key,name,dialog.fields[name].get())
                    for name,value in case.get('extra_fields',{}).items():dialog.fields[name].set(value)
                    dialog.case_reviewed.set(True);dialog.find_button.invoke()
                    wait(dialog,lambda:dialog.data is not None and not dialog.busy)
                    selected=[p for p in dialog.data['parcels'] if p['include']]
                    assert len(selected)==19 and len(dialog.data['parcels'])==111
                    if case.get('property_inputs'):
                        # Exercise the real property editor once; fill the other
                        # fixture measurements directly so no real surveys are implied.
                        first=selected[0];dialog.tree.selection_set(first['account']);dialog.edit_property()
                        edit=next(w for w in dialog.winfo_children() if isinstance(w,tk.Toplevel))
                        for field,value in case['property_inputs'].items():
                            label=dict(ui.PROPERTY_INPUTS[key]).get(field,'Correction / review note')
                            widget=next(w for w in descendants(edit) if isinstance(w,ttk.Label) and w.cget('text')==label)
                            row=widget.grid_info()['row']
                            entry=next(w for w in widget.master.winfo_children() if isinstance(w,ttk.Entry) and w.grid_info().get('row')==row)
                            entry.delete(0,'end');entry.insert(0,str(value))
                        next(w for w in descendants(edit) if isinstance(w,ttk.Button) and w.cget('text')=='Apply changes').invoke()
                        assert not edit.winfo_exists()
                        assert first['building_line']==case['property_inputs']['building_line']
                        for p in selected[1:]:p.update(case['property_inputs'])
                    dialog.parcels_reviewed.set(True)
                    destination=args.out/case['id']/(case['id']+'_Calculation.xlsx');destination.parent.mkdir(exist_ok=True)
                    attachments=('map','tax_map','subdivision_plat','hcad_records','original_ordinance')
                    if hasattr(owner,'vars'):
                        for k in attachments:
                            if k in owner.vars:owner.vars[k].set('old case attachment')
                    else:
                        for k in attachments:
                            if k in owner.files:owner.files[k]='old case attachment'
                    with patch.object(ui.filedialog,'asksaveasfilename',return_value=str(destination)):
                        dialog.save_button.invoke()
                    wait(dialog,lambda:dialog.closed)
                    assert destination.is_file()
                    got=owner.vars['excel'].get() if hasattr(owner,'vars') else owner.files['workbook']
                    assert Path(got)==destination
                    for k in attachments:
                        value=owner.vars[k].get() if hasattr(owner,'vars') and k in owner.vars else (owner.files.get(k,'') if hasattr(owner,'files') else '')
                        assert not value,(key,'stale attachment',k)
                    assert destination.with_name(destination.stem+'_sources').joinpath('Application.pdf').is_file()
                    result={'wizard':key,'application_read':True,'records':len(dialog.data['parcels']),
                            'selected':len(selected),'saved_to_same_wizard':True,'previous_attachments_cleared':True,'workbook':str(destination)}
                    map_cases.append({'id':case['id'],'wizard':key,'workbook':str(destination)})
                except Exception as e:
                    result={'wizard':key,'error':str(e)}
                finally:
                    if dialog is not None and dialog.winfo_exists():dialog.close()
                    if owner is not None:owner.destroy()
                    root.update()
                results.append(result);print(json.dumps(result),flush=True)
    finally:root.destroy()
    (args.out/'results.json').write_text(json.dumps(results,indent=2),encoding='utf8')
    (args.out/'map_cases.json').write_text(json.dumps(map_cases,indent=2),encoding='utf8')
    return int(any(r.get('error') for r in results))


if __name__=='__main__':raise SystemExit(main())
