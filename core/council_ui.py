"""Geographic council lookup for an existing calculation sheet, all wizards."""
import threading,queue
from pathlib import Path
import tkinter as tk
from tkinter import ttk,messagebox,filedialog

def find_workbook_council(owner):
    from .boundary_map_ui import wizard_paths
    from .lot_divisions import workbook_rows
    from .boundary_map import PublicMapClient,PARCEL_SERVICE,PARCEL_FIELDS,CITY_PARCEL_SERVICE,CITY_PARCEL_FIELDS,fingerprint
    from .council_district import resolve_districts,add_members,write_case_district_copy
    from .lot_review_ui import select_calculation_copy
    source,_,folder=wizard_paths(owner)
    if not source:messagebox.showinfo('Council lookup','Select the calculation sheet first.',parent=owner);return
    window=tk.Toplevel(owner);window.title('Council district and member');window.transient(owner.winfo_toplevel())
    status=tk.StringVar(value='Checking the selected properties against City district boundaries...')
    ttk.Label(window,textvariable=status,wraplength=520,padding=18).pack()
    events=queue.Queue();cancel=threading.Event()
    window.protocol('WM_DELETE_WINDOW',lambda:(cancel.set(),window.destroy()))
    def work():
        try:
            stamp=fingerprint(source);rows,_=workbook_rows(source);accounts=sorted({r['account'] for r in rows});features=[]
            cache=Path(folder or Path(source).parent)/'.council_cache';client=PublicMapClient(cache,cancel=cancel,use_cache=False)
            for i in range(0,len(accounts),30):
                wanted=accounts[i:i+30];where='HCAD_NUM IN ('+','.join("'"+a+"'" for a in wanted)+')'
                found=client.query(PARCEL_SERVICE,where=where,fields=PARCEL_FIELDS);seen={str(f['attributes'].get('HCAD_NUM')) for f in found}
                missing=set(wanted)-seen
                if missing:found+=client.query(CITY_PARCEL_SERVICE,where='HCAD_NUM IN ('+','.join("'"+a+"'" for a in missing)+')',fields=CITY_PARCEL_FIELDS)
                if set(wanted)-{str(f['attributes'].get('HCAD_NUM')) for f in found}:raise ValueError('Some selected accounts could not be located. The district lookup needs every selected property.')
                features.extend(found)
            result=add_members(resolve_districts(features,client),cache)
            if fingerprint(source)!=stamp:raise ValueError('The calculation sheet changed during lookup. Please run the lookup again.')
            events.put(('ready',(result,stamp)))
        except Exception as exc:events.put(('error',str(exc)))
    def poll():
        if not window.winfo_exists():return
        try:kind,value=events.get_nowait()
        except queue.Empty:window.after(100,poll);return
        if kind=='error':status.set(value);ttk.Button(window,text='Close',command=window.destroy).pack(pady=8);return
        result,stamp=value
        status.set('District: '+(result['value'] or 'Not found')+'\nCouncilmember: '+(result.get('member_value') or 'Needs review')+'\n\n'+'\n'.join(result['warnings']))
        def save():
            path=filedialog.asksaveasfilename(parent=window,title='Save calculation with council information',initialdir=folder or str(Path(source).parent),initialfile=Path(source).stem+'_Council.xlsx',defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')])
            if not path:return
            try:
                if fingerprint(source)!=stamp:raise ValueError('The source sheet changed. Run the lookup again.')
                write_case_district_copy(source,path,result);select_calculation_copy(owner,path);window.destroy()
            except Exception as exc:messagebox.showerror('Council information',str(exc),parent=window)
        ttk.Button(window,text='Save new calculation sheet',command=save,state='normal' if result['value'] else 'disabled').pack(pady=10)
    threading.Thread(target=work,daemon=True).start();window.after(100,poll)
