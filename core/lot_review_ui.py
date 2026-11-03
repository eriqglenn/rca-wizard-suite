"""Planner review of actual legal lots before writing a calculation copy."""
import copy, tkinter as tk
from pathlib import Path
from tkinter import ttk,messagebox,filedialog
from .lot_divisions import proposed_lots,validate_lots,positive,workbook_rows,reviewed_workbook_copy,clean

class LotReview(tk.Toplevel):
    def __init__(self,parent,rows,on_done=None):
        super().__init__(parent);self.title('Review individual lots');self.geometry('1000x660');self.transient(parent.winfo_toplevel())
        self.rows=rows;self.on_done=on_done;self.current=None;self.entries=[]
        ttk.Label(self,text='Give each legal lot its own description and area. HCAD account, owner and mailing information stay linked to the property.',wraplength=920).pack(anchor='w',padx=12,pady=10)
        body=ttk.Panedwindow(self,orient='horizontal');body.pack(fill='both',expand=True,padx=12)
        left=ttk.Frame(body);right=ttk.Frame(body);body.add(left,weight=1);body.add(right,weight=3)
        self.list=tk.Listbox(left,width=34,exportselection=False);self.list.pack(fill='both',expand=True);self.list.bind('<<ListboxSelect>>',self.choose)
        self.heading=ttk.Label(right,wraplength=610);self.heading.pack(anchor='w',pady=8)
        self.tree=ttk.Treeview(right,columns=('desc','area','use'),show='headings',height=9,selectmode='browse')
        for c,t,w in [('desc','Individual legal description',350),('area','Square feet',105),('use','Land use',70)]:self.tree.heading(c,text=t);self.tree.column(c,width=w)
        self.tree.pack(fill='both',expand=True);self.tree.bind('<<TreeviewSelect>>',self.select_lot)
        form=ttk.Frame(right);form.pack(fill='x',pady=8)
        self.desc=tk.StringVar();self.area=tk.StringVar();self.use=tk.StringVar();self.source=tk.StringVar()
        for n,(label,var) in enumerate([('Lot description',self.desc),('Lot area (sq ft)',self.area),('Land use',self.use)]):
            ttk.Label(form,text=label).grid(row=n,column=0,sticky='w');ttk.Entry(form,textvariable=var,width=55).grid(row=n,column=1,sticky='ew')
        form.columnconfigure(1,weight=1)
        bar=ttk.Frame(right);bar.pack(fill='x')
        for title,cmd in [('Update lot',self.update_lot),('Add lot',self.add_lot),('Remove lot',self.remove_lot),('Equal areas...',self.equal_areas)]:ttk.Button(bar,text=title,command=cmd).pack(side='left',padx=2)
        ttk.Label(right,text='Area source / plat reference or planner verification').pack(anchor='w',pady=(10,0))
        ttk.Entry(right,textvariable=self.source).pack(fill='x')
        self.draft=tk.BooleanVar();ttk.Checkbutton(right,text='Save as draft with unknown individual lot areas',variable=self.draft).pack(anchor='w',pady=5)
        ttk.Button(right,text='Accept this property’s lot rows',command=self.accept_property).pack(anchor='e',pady=8)
        bottom=ttk.Frame(self,padding=12);bottom.pack(fill='x')
        self.status=tk.StringVar();ttk.Label(bottom,textvariable=self.status,wraplength=780).pack(side='left')
        ttk.Button(bottom,text='Done',command=self.done).pack(side='right')
        self.refresh_list()
        pending=next((i for i,r in enumerate(rows) if proposed_lots(r)),0)
        if rows:self.list.selection_set(pending);self.choose()

    def refresh_list(self):
        self.list.delete(0,'end')
        for r in self.rows:
            lots=proposed_lots(r);state=f"{len(lots)} lots {'reviewed' if r.get('lots') else 'to review'}" if lots else 'one lot / manual review'
            self.list.insert('end',f"{r.get('address') or r['account']} — {state}")

    def choose(self,event=None):
        selection=self.list.curselection()
        if not selection:return
        self.current=selection[0];row=self.rows[self.current]
        self.heading.configure(text=f"{row.get('address','')} | HCAD {row['account']}\n{row.get('legal_description','')}\nProperty total: {row.get('land_area') or 'unknown'} sq ft")
        self.entries=proposed_lots(row) or [{'description':row.get('legal_description',''),'area':row.get('land_area'),'land_use':row.get('land_use',''),'area_source':''}]
        self.source.set(next((l.get('area_source','') for l in self.entries if l.get('area_source')),''));self.draft.set(False);self.refresh_tree()

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i,l in enumerate(self.entries):self.tree.insert('','end',iid=str(i),values=(l['description'],'' if l.get('area') is None else l['area'],l.get('land_use','')))

    def select_lot(self,event=None):
        if not self.tree.selection():return
        l=self.entries[int(self.tree.selection()[0])];self.desc.set(l['description']);self.area.set('' if l.get('area') is None else str(l['area']));self.use.set(l.get('land_use',''))

    def update_lot(self):
        if not self.tree.selection():return
        try:area=positive(self.area.get())
        except ValueError as e:messagebox.showwarning('Lot area',str(e),parent=self);return
        use=self.use.get().strip().upper()
        if use not in ('SFR','MF','COM','VAC','EXC',''):messagebox.showwarning('Land use','Use SFR, MF, COM, VAC or EXC.',parent=self);return
        self.entries[int(self.tree.selection()[0])].update(description=self.desc.get().strip(),area=area,land_use=use);self.refresh_tree()

    def add_lot(self):self.entries.append({'description':'','area':None,'land_use':self.rows[self.current].get('land_use','')});self.refresh_tree();self.tree.selection_set(str(len(self.entries)-1));self.select_lot()
    def remove_lot(self):
        if self.tree.selection():self.entries.pop(int(self.tree.selection()[0]));self.refresh_tree()

    def equal_areas(self):
        total=positive(self.rows[self.current].get('land_area'))
        if not total or not self.entries:return
        if not messagebox.askyesno('Verify equal lot areas','Use equal areas only if the plat or other reliable record confirms equal lot sizes. Have you verified that these lots have equal areas?',parent=self):return
        for l in self.entries:l['area']=total/len(self.entries)
        self.source.set('Planner verified equal areas against plat/record');self.refresh_tree()

    def accept_property(self):
        if self.current is None:return
        try:
            validate_lots(self.entries,self.rows[self.current].get('land_area'),allow_unknown=self.draft.get())
            if not self.source.get().strip() and any(l.get('area') is not None for l in self.entries):raise ValueError('Enter the source or planner verification for these individual lot areas.')
        except ValueError as e:messagebox.showwarning('Review lots',str(e),parent=self);return
        for l in self.entries:l['area_source']=self.source.get().strip()
        self.rows[self.current]['lots']=copy.deepcopy(self.entries)
        self.refresh_list();self.list.selection_set(self.current)
        self.status.set('Accepted. Review another property or choose Done.')

    def done(self):
        if self.on_done and self.on_done() is False:return
        self.destroy()

def review_import_lots(dialog):
    if dialog.busy or not dialog.data:return
    rows=[p for p in dialog.data['parcels'] if p.get('include')]
    if not rows:messagebox.showinfo('Review lots','Select the application properties first.',parent=dialog);return
    return LotReview(dialog,rows,dialog.refresh_rows)

def review_existing_workbook(owner):
    from .boundary_map_ui import wizard_paths
    from .boundary_map import fingerprint
    source,_,folder=wizard_paths(owner)
    if not source:messagebox.showinfo('Review lots','Choose a calculation sheet first.',parent=owner);return
    try:rows,metadata=workbook_rows(source);stamp=fingerprint(source)
    except Exception as e:messagebox.showerror('Read calculation sheet',str(e),parent=owner);return
    def save():
        path=filedialog.asksaveasfilename(parent=dialog,title='Save reviewed lot calculation',initialdir=folder or str(Path(source).parent),initialfile=Path(source).stem+'_Lots.xlsx',defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')])
        if not path:return False
        try:
            if fingerprint(source)!=stamp:raise ValueError('The source sheet changed during review. Reopen Review Individual Lots to use its latest saved values.')
            reviewed_workbook_copy(source,path,rows,metadata,owner._boundary_map_wizard_key)
            select_calculation_copy(owner,path)
            messagebox.showinfo('Lot calculation saved','Saved the new calculation sheet. Load Workbook to recalculate and review Step 2. Existing map attachments are retained for your review.',parent=owner)
        except Exception as e:messagebox.showerror('Calculation not saved',str(e),parent=owner);return False
    dialog=LotReview(owner,rows,save)
    return dialog

def select_calculation_copy(owner,path):
    # Same case: preserve selected attachments and output folder.
    while not hasattr(owner,'_save_session') and hasattr(owner,'owner'):owner=owner.owner
    if hasattr(owner,'vars'):owner.vars['excel'].set(str(path));owner.summary=None;owner._show_step(0)
    else:owner.files['workbook']=str(path);owner._select_files_page()
    owner._save_session()
