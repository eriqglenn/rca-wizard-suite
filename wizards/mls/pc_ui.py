"""MLS-only Planning Commission work area, reached from Generate Letters."""
from __future__ import annotations
from pathlib import Path
import copy
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from core.boundary_map import fingerprint, read_map_case, PublicMapClient, fetch_map_data
from core.boundary_map_ui import MapPreview
from core.workbooks import WorkbookReadCache
from .pc_data import read_pc_facts, attach_pc_rows, load_preferences, save_preferences
from .pc_report import report_context, generate_pc_report
from .pc_map import render_pc_map, pc_defaults, save_pc_map, PAGE
from .pc_landuse import attach_land_use

class PCMapPreview(MapPreview):
    page_size=PAGE
    map_kind='PC Map'

    def render_preview(self,data,settings):return render_pc_map(data,settings)
    def default_layout(self,data):return pc_defaults(data)
    def save_preview(self,data,settings,rendered,output):return save_pc_map(data,settings,rendered,output)

    def start_load(self):
        if self._closed or self._loading:return
        self._loading=True;self.refresh_button.configure(state='disabled')
        source=self.source;output=self.output;use_cache=self._use_cache
        facts=copy.deepcopy(self.owner.facts);protests=list(self.owner.options()['protested_rows'])
        resolved=self.owner.resolved
        def worker():
            try:
                if fingerprint(source)!=facts['workbook_hash']:raise ValueError('The workbook changed. Close this preview and reopen the PC work area to read its latest values.')
                from core.pc_workflow import load_map_data
                data=load_map_data(source,self.owner._boundary_map_wizard_key,output,progress=lambda s:self.events.put(('status',s)),cancel=self.cancel,use_cache=use_cache)
                updated=data['calculation_path']
                cache=WorkbookReadCache();updated_resolved=cache.resolve(updated)
                current_facts=self.owner.read_facts(updated_resolved,updated)
                client=PublicMapClient(Path(output)/'.boundary_map_cache',progress=lambda s:self.events.put(('status',s)),cancel=self.cancel,use_cache=use_cache)
                data=attach_land_use(attach_pc_rows(data,current_facts,protests),client)
                data['pc_calculation']={'resolved':updated_resolved,'facts':current_facts}
                self._loaded_workbook_cache=cache
                self.events.put(('ready',data))
            except Exception as error:self.events.put(('error',str(error)))
        threading.Thread(target=worker,daemon=True).start();self.poll()

class PCWorkspace(tk.Toplevel):
    _boundary_map_wizard_key='mls'

    def __init__(self,owner,start_map=False):
        super().__init__(owner)
        self.owner=owner
        self._boundary_map_wizard_key=getattr(owner,'_boundary_map_wizard_key','mls')
        self.renewal=self._boundary_map_wizard_key.endswith('_renewal')
        self.building=self._boundary_map_wizard_key.startswith('mbl')
        self.case_label=('MBL' if self.building else 'MLS')+(' Renewal' if self.renewal else '')
        self.read_facts=lambda path,source:read_pc_facts(path,source,self._boundary_map_wizard_key)
        self.title(self.case_label+' Planning Commission Draft Report and Map')
        self.geometry('1060x820');self.minsize(850,650);self.transient(owner.winfo_toplevel())
        from core.boundary_map_ui import wizard_paths
        self.source,owner_map,self.output=wizard_paths(owner)
        self.facts=None;self.resolved=None;self.events=queue.Queue();self.closed=False;self.busy=False;self.start_map=start_map
        self.protocol('WM_DELETE_WINDOW',self.close)
        self.fields={}
        self.pc_map=tk.StringVar();self.status=tk.StringVar(value='Reading the calculation sheet...')
        self.vars={'excel':tk.StringVar(value=self.source),'map':self.pc_map,'output_folder':tk.StringVar(value=self.output)}
        self.protested=set()
        ttk.Label(self,text=self.case_label+' Planning Commission',font=('Segoe UI',17,'bold')).pack(anchor='w',padx=16,pady=(14,4))
        ttk.Label(self,text='Review the case wording and map symbols, then generate the draft and attachments.').pack(anchor='w',padx=16,pady=(0,8))
        self.tabs=ttk.Notebook(self);self.tabs.pack(fill='both',expand=True,padx=14,pady=4)
        details=ttk.Frame(self.tabs,padding=12);properties=ttk.Frame(self.tabs,padding=12);files=ttk.Frame(self.tabs,padding=12)
        self.tabs.add(details,text='Report details');self.tabs.add(properties,text='Properties and protests');self.tabs.add(files,text='Attachments')
        labels=[('agenda','Agenda number'),('hearing_date','PC meeting / hearing date'),('protest_count','Written protests received'),
                ('plat_year','Subdivision plat year'),('neighborhood','Neighborhood history (optional)'),
                ('structure_age','Age of structures statement'),('preservation','Preservation finding (optional)')]
        if self.renewal and not self.building:
            labels=[('agenda','Agenda number'),('hearing_date','PC meeting / hearing date'),('protest_count','Timely written protests received'),('ordinance','Previous ordinance number'),('adoption_date','Previous ordinance adoption date'),('condition_statement','Condition of block since original ordinance'),('variance_statement','Lot size variation finding (excluding corners)')]
        if self.building and self.renewal:
            labels.extend([('ordinance','Previous ordinance number'),('adoption_date','Previous ordinance adoption date')])
        details.columnconfigure(1,weight=1)
        for r,(key,title) in enumerate(labels):
            self.fields[key]=tk.StringVar()
            ttk.Label(details,text=title).grid(row=r,column=0,sticky='w',padx=(0,10),pady=5)
            ttk.Entry(details,textvariable=self.fields[key],width=75).grid(row=r,column=1,sticky='ew',pady=5)
        self.notice=tk.BooleanVar(value=False)
        ttk.Checkbutton(details,text='Owner renewal notice and protest period confirmed' if self.renewal else 'Public-hearing notice has been transmitted to all property owners',variable=self.notice).grid(row=len(labels),column=0,columnspan=2,sticky='w',pady=8)
        self.hearing_notice=tk.BooleanVar(value=False)
        if self.renewal:
            ttk.Checkbutton(details,text='Public-hearing notice has been transmitted to all property owners',variable=self.hearing_notice).grid(row=len(labels)+1,column=0,columnspan=2,sticky='w',pady=4)
        self.metrics=tk.Text(details,height=10,wrap='word',font=('Segoe UI',10),state='disabled')
        self.metrics.grid(row=len(labels)+2,column=0,columnspan=2,sticky='nsew',pady=8);details.rowconfigure(len(labels)+2,weight=1)
        ttk.Label(properties,text='Support, lot areas and building lines come from Step 2. Double-click a row to mark or unmark a protesting property.\nA written protest may cover more than one property; review the count separately.').pack(anchor='w',pady=(0,8))
        columns=('row','description','use','area','building_line','support','protest') if self.building else ('row','description','use','area','support','protest')
        self.tree=ttk.Treeview(properties,columns=columns,show='headings',selectmode='extended')
        titles=['Step 2 row','Property','Land use','Lot area']+(['Building line (ft)'] if self.building else [])+['Support','Protest']
        widths=[75,390,75,90]+([110] if self.building else [])+[65,65]
        for key,title,width in zip(columns,titles,widths):
            self.tree.heading(key,text=title);self.tree.column(key,width=width,stretch=key=='description')
        self.tree.pack(fill='both',expand=True);self.tree.bind('<Double-1>',lambda e:self.toggle_protests())
        ttk.Button(properties,text='Toggle protest for selected properties',command=self.toggle_protests).pack(anchor='w',pady=8)
        files.columnconfigure(1,weight=1)
        self.fields['application_pdf']=tk.StringVar();self.fields['boundary_map']=tk.StringVar(value=owner_map);self.fields['pc_map']=self.pc_map
        for row,(key,title) in enumerate([('application_pdf','Application PDF'),('boundary_map','Regular boundary map'),('pc_map','PC land-use / support map')]):
            ttk.Label(files,text=title).grid(row=row,column=0,sticky='w',pady=8)
            ttk.Entry(files,textvariable=self.fields[key],width=70).grid(row=row,column=1,sticky='ew',padx=10)
            ttk.Button(files,text='Browse',command=lambda k=key:self.browse(k)).grid(row=row,column=2)
        self.include_calculation=tk.BooleanVar(value=True)
        ttk.Checkbutton(files,text='Include the Step 2 calculation analysis',variable=self.include_calculation).grid(row=3,column=0,columnspan=3,sticky='w',pady=10)
        ttk.Label(files,text='The report includes the selected attachments in this order: calculation analysis, PC map, application, boundary map.\nMissing attachments and incomplete fields are recorded in the review notes. The editable Word draft can still be created.',wraplength=850).grid(row=4,column=0,columnspan=3,sticky='w',pady=10)
        actions=ttk.Frame(self,padding=12);actions.pack(fill='x')
        self.map_button=ttk.Button(actions,text='Generate PC Map',command=self.make_map,state='disabled');self.map_button.pack(side='left')
        self.report_button=ttk.Button(actions,text='Generate PC Draft Report',command=self.make_report,state='disabled');self.report_button.pack(side='left',padx=8)
        ttk.Button(actions,text='Close',command=self.close).pack(side='right')
        ttk.Label(self,textvariable=self.status,wraplength=1000).pack(anchor='w',padx=16,pady=(0,12))
        self.timer=self.after(100,self.poll)
        def worker():
            try:
                resolved=WorkbookReadCache().resolve(self.source)
                facts=self.read_facts(resolved,self.source)
                prefs=load_preferences(self.source,self.output,facts['workbook_hash'])
                self.events.put(('facts',(resolved,facts,prefs)))
            except Exception as error:self.events.put(('error',str(error)))
        threading.Thread(target=worker,daemon=True).start()

    def options(self):
        return {**{k:v.get() for k,v in self.fields.items()},'notice_sent':self.notice.get(),'hearing_notice_sent':self.hearing_notice.get(),
                'include_calculation':self.include_calculation.get(),'protested_rows':sorted(self.protested)}

    def remember(self):
        save_preferences(self.source,self.output,self.facts['workbook_hash'],self.options())

    def browse(self,key):
        path=filedialog.askopenfilename(parent=self,filetypes=[('PDF or image','*.pdf *.png *.jpg *.tif'),('All files','*.*')])
        if path:self.fields[key].set(path)

    def toggle_protests(self):
        for iid in self.tree.selection():
            row=int(iid)
            if row in self.protested:self.protested.remove(row)
            else:self.protested.add(row)
            self.tree.set(iid,'protest','Yes' if row in self.protested else '')
        self.status.set('Protest markings changed. Generate a new PC Map to include these selections.')

    def poll(self):
        if self.closed:return
        try:
            while True:
                kind,value=self.events.get_nowait()
                if kind=='facts':
                    self.resolved,self.facts,prefs=value;ci=self.facts['case_info']
                    mapping={'agenda':'PC_Agenda_Number','hearing_date':'PC_Hearing_Date','protest_count':'Protest_Count','plat_year':'Subdivision_Plat_Year','neighborhood':'Neighborhood_Age_Statement','structure_age':'Structure_Age_Statement','preservation':'Preservation_Finding'}
                    mapping.update(ordinance='Original_Ordinance_No',adoption_date='Original_Passage_Date',condition_statement='Renewal_Block_Condition_Statement',variance_statement='PC_Lot_Size_Variance_Statement')
                    self.hearing_notice.set(prefs.get('hearing_notice_sent',False))
                    for k,var in self.fields.items():var.set(prefs.get(k,ci.get(mapping.get(k,''),var.get())))
                    self.notice.set(prefs.get('notice_sent',False));self.include_calculation.set(prefs.get('include_calculation',True));self.protested=set(prefs.get('protested_rows',[]))
                    for row in self.facts['rows']:
                        r=row['row'];values=[r,row['description'],row['land_use'],row['area']]+([row.get('building_line')] if self.building else [])+['Yes' if row['signed'] else ('No' if row.get('support_known') else 'Unknown'),'Yes' if r in self.protested else '']
                        self.tree.insert('', 'end',iid=str(r),values=values)
                    f=self.facts
                    lines=[f"{self.case_label} {f['application']}: {f['boundary']}",f"Proposed minimum: {f['minimum_text']} square feet",
                           f"Boundary properties: {f['all_count']}; eligible lots: {f['eligible_count']}; excluded properties: {f['excluded_count']}",
                           f"SFR share by LOT COUNT: {f['sfr_count']} / {f['eligible_count']} = {f['sfr_lot_percent']}%",
                           ("Petition support is not required for renewal. Only explicitly recorded support is marked on the map." if self.renewal else f"Petition support by LAND AREA: {f['support_area_percent']}% ({f['signed_count']} eligible properties signed)"),
                           f"Lots at least the minimum: {f['qualifying_count']} / {f['eligible_count']}; their LAND AREA share: {f['qualifying_area_percent']}%",*f['warnings']]
                    if self.building:
                        lines=[f"{self.case_label} {f['application']}: {f['boundary']}",f"Minimum building line: {f['minimum_text']} feet",f"Eligible lots: {f['eligible_count']}; EXC: {f['excluded_count']}; vacant: {f['vacant_count']}",f"SFR by LAND AREA: {f['sfr_area_percent']}%",f"Support by LAND AREA: {f['support_area_percent']}%",f"Developed lots with a building line at least the minimum: {f['qualifying_count']} / {f['developed_count']} = {f['developed_percent']}%",*f['warnings']]
                    self.metrics.configure(state='normal');self.metrics.insert('1.0','\n\n'.join(lines));self.metrics.configure(state='disabled')
                    self.map_button.configure(state='normal');self.report_button.configure(state='normal');self.status.set('Review the figures and report details. Missing fields remain visible placeholders in the draft.')
                    if self.start_map:self.start_map=False;self.after(100,self.make_map)
                elif kind=='report':
                    self.busy=False;self.report_button.configure(state='normal')
                    self.status.set('Created '+value['docx'])
                    messagebox.showinfo('PC Draft Report','Created editable Word'+(' and PDF' if value['pdf'] else '')+' report.\n\n'+value['folder']+('\n\nReview notes:\n'+'\n'.join(value['notes'][:8]) if value['notes'] else ''),parent=self)
                elif kind=='error':
                    self.busy=False;self.status.set(value)
                    if self.facts:self.report_button.configure(state='normal')
                    messagebox.showerror('MLS Planning Commission',value,parent=self)
        except queue.Empty:pass
        self.timer=self.after(100,self.poll)

    def make_map(self):
        if not self.facts:return
        prior=getattr(self,'preview',None)
        if prior is not None and prior.winfo_exists():prior.lift();return
        self.remember();self.preview=PCMapPreview(self)

    def _set_generated_boundary_map(self,path):
        self.pc_map.set(path);self.remember();self.status.set('PC map saved and selected for this report: '+path)

    def _map_calculation_ready(self,data):
        calculation=data.get('pc_calculation')
        if not calculation:return
        self.source=data['calculation_path'];self.vars['excel'].set(self.source)
        self.resolved=calculation['resolved'];self.facts=calculation['facts']
        self.remember()

    def make_report(self):
        if not self.facts or self.busy:return
        if fingerprint(self.source)!=self.facts['workbook_hash']:
            messagebox.showwarning('Workbook changed','Close this PC work area and reopen it to read the latest saved calculation sheet.',parent=self);return
        self.remember();options=self.options();facts=copy.deepcopy(self.facts);resolved=self.resolved;output=self.output
        self.busy=True;self.report_button.configure(state='disabled');self.status.set('Creating the PC draft and exporting its attachments...')
        def worker():
            try:self.events.put(('report',generate_pc_report(resolved,facts,options,output)))
            except Exception as error:self.events.put(('error',str(error)))
        threading.Thread(target=worker,daemon=True).start()

    def close(self):
        if self.busy:
            self.status.set('The report is still generating. This window will remain open until it finishes.');return
        if self.facts:
            try:self.remember()
            except OSError:pass
        preview=getattr(self,'preview',None)
        if preview is not None and preview.winfo_exists():preview.close()
        self.closed=True;self.after_cancel(self.timer);self.destroy()

def open_pc_workspace(owner,start_map=False):
    from core.pc_workflow import prepare_owner_calculation
    source=prepare_owner_calculation(owner)
    if not source or not Path(source).is_file():raise ValueError('Select a saved calculation sheet first.')
    prior=getattr(owner,'_pc_workspace',None)
    if prior is not None and prior.winfo_exists():
        prior.lift()
        if start_map:prior.make_map()
        return prior
    window=PCWorkspace(owner,start_map);owner._pc_workspace=window;return window
