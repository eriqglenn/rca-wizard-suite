"""Application intake for all four City wizards."""
from __future__ import annotations

import copy
import gc
import io
from pathlib import Path
import queue
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

import fitz
from PIL import Image, ImageTk

from .application_ocr import read_application, parse_boundary
from .application_import import (find_parcels, lookup_account, save_calculation,
    validate_fields, record_pdf_facts, LAND_USES, parcel_polygon, wizard_label,
    PROPERTY_INPUTS, optional_measurement, calculation_review_notes, case_number, prepare_reviewed_save)
from . import launch
from .scrollable import ScrollableContent


FORM_FIELDS = [
 ('application_number','Application number'),('date_received','Date received'),
 ('street','Street name'),('block_range','Block / address range'),('sides','Side(s)'),
 ('cross_1','First cross street'),('cross_2','Second cross street'),
 ('applicant_name','Applicant name'),('applicant_address','Mailing street address'),
 ('applicant_address_2','Apartment / second line'),('applicant_city','City'),
 ('applicant_state','State'),('applicant_zip','ZIP'),('applicant_property_address','Applicant property address'),
 ('applicant_phone','Applicant phone'),('applicant_email','Applicant email'),
 ('council_district','Council district'),('council_member','Councilmember'),('planner_name','Planner assigned'),
 ('application_date','Application signed date'),('date_complete','Date complete')]


def manual_application(fields):
    return {'source_pdf':'', 'fields':dict(fields), 'property_hints':[], 'pages':[],
            'warnings':[], 'input_method':'Manual boundary entry; no application PDF'}


def form_fields(wizard):
    fields=list(FORM_FIELDS)
    if wizard.endswith('_renewal'):
        units='square feet' if wizard.startswith('mls') else 'feet'
        fields += [('renewal_minimum',f'Minimum to renew ({units})'),
            ('original_ordinance_no','Original ordinance number'),
            ('original_passage_date','Original Council passage date'),
            ('original_effective_date','Original effective date'),
            ('original_expiration_date','Original expiration date'),
            ('application_filed_date','Application filed date'),
            ('application_accepted_date','Application accepted date')]
    return fields


def select_imported_calculation(owner, destination):
    """Return the new case to the same wizard and clear previous-case attachments."""
    destination=Path(destination)
    source_dir=destination.with_name(destination.stem+'_sources')
    records=sorted(str(p) for p in source_dir.glob('HCAD_*.pdf'))
    attachments=('map','hcad_records','subdivision_plat','tax_map','original_ordinance')
    message='Imported application and selected HCAD parcels: '+str(destination)+'. Load Workbook to recalculate and review Step 2, then create the boundary map.'
    if hasattr(owner,'vars'):
        for key in attachments:
            if key in owner.vars: owner.vars[key].set('')
        owner.vars['excel'].set(str(destination));owner.vars['output_folder'].set(str(destination.parent))
        owner.vars['hcad_records'].set('||'.join(records))
        owner.summary=None;owner._show_step(0);owner.vars['status'].set(message)
    else:
        for key in attachments:
            if key in owner.files: owner.files[key]=''
        owner.files.update(workbook=str(destination),out=str(destination.parent),hcad_records=';'.join(records))
        owner._select_files_page();owner.status_var.set(message)
    owner._save_session()


class ApplicationImport(tk.Toplevel):
    def __init__(self,owner):
        super().__init__(owner)
        self.owner=owner
        self.wizard=owner._boundary_map_wizard_key
        self.wizard_name=wizard_label(self.wizard)
        self.form_items=form_fields(self.wizard)
        self.title(f'Import {self.wizard_name} Application - HCAD to Calculation Sheet')
        self.geometry(f'{min(1320,self.winfo_screenwidth()-60)}x{min(950,self.winfo_screenheight()-90)}')
        self.minsize(980,650)
        self.transient(owner.winfo_toplevel())
        self.configure(bg='#F8FAFC')
        style=ttk.Style(self)
        style.configure('Import.TLabel',background='#F8FAFC',foreground='#0F172A',font=('Segoe UI',10))
        self.temp=tempfile.TemporaryDirectory(prefix='rca_application_review_')
        self.work=Path(self.temp.name)
        self.cancel=threading.Event();self.events=queue.Queue()
        self.busy=False;self.closed=False;self._saving=False;self.application=None;self.data=None;self.pdf=None
        self.page=0;self.rotation=0;self.image=None;self._process=None;self._canvas_items={}
        self.fields={k:tk.StringVar(value='') for k,_ in self.form_items}
        self.status=tk.StringVar(value='No application PDF is required. Enter the boundary, find HCAD parcels, select properties, then save the calculation sheet and create a map.')
        self.reading_note=tk.StringVar(value='No PDF needed: enter the street, block range, side(s) and both cross streets below. Then click Find HCAD Parcels.')
        self.case_reviewed=tk.BooleanVar(value=False);self.parcels_reviewed=tk.BooleanVar(value=False)
        self.only_selected=tk.BooleanVar(value=False)
        self.book=ttk.Notebook(self);self.book.pack(fill='both',expand=True,padx=12,pady=10)
        self.form_tab=ttk.Frame(self.book);self.parcel_tab=ttk.Frame(self.book)
        self.book.add(self.form_tab,text='1. Application / Boundary');self.book.add(self.parcel_tab,text='2. Parcels and calculation sheet')
        self._build_form();self._build_parcels()
        footer=ttk.Frame(self,padding=10);footer.pack(fill='x')
        ttk.Label(footer,style='Import.TLabel',textvariable=self.status,wraplength=1050).pack(side='left',fill='x',expand=True)
        ttk.Button(footer,text='Close',command=self.close).pack(side='right')
        self.protocol('WM_DELETE_WINDOW',self.close)
        self._timer=self.after(100,self.poll)
        for var in self.fields.values(): var.trace_add('write',self.case_changed)

    def _build_form(self):
        self.form_tab.columnconfigure(1,weight=1);self.form_tab.rowconfigure(1,weight=1)
        bar=ttk.Frame(self.form_tab,padding=8);bar.grid(row=0,column=0,columnspan=2,sticky='ew')
        ttk.Button(bar,text='Choose Application PDF (optional)',command=self.choose_pdf).pack(side='left')
        self.file_label=ttk.Label(bar,style='Import.TLabel',text='Manual boundary entry available');self.file_label.pack(side='left',padx=12)
        scroll=ScrollableContent(self.form_tab,bg='#F8FAFC');scroll.grid(row=1,column=0,sticky='nsew',padx=(0,8))
        form=scroll.interior
        ttk.Label(form,style='Import.TLabel',text='Enter or review case information',font=('Segoe UI',13,'bold')).grid(row=0,column=0,columnspan=2,sticky='w',padx=10,pady=8)
        ttk.Label(form,style='Import.TLabel',textvariable=self.reading_note,wraplength=440).grid(row=1,column=0,columnspan=2,sticky='w',padx=10,pady=6)
        for row,(key,label) in enumerate(self.form_items,2):
            ttk.Label(form,style='Import.TLabel',text=label).grid(row=row,column=0,sticky='w',padx=10,pady=4)
            if key=='sides':
                entry=ttk.Combobox(form,textvariable=self.fields[key],values=['north','south','east','west','north and south','east and west'],width=25,state='readonly')
            else: entry=ttk.Entry(form,textvariable=self.fields[key],width=28)
            entry.grid(row=row,column=1,sticky='ew',padx=8,pady=4)
        n=len(self.form_items)+2
        ttk.Label(form,style='Import.TLabel',text='Boundary description (optional)').grid(row=n,column=0,columnspan=2,sticky='w',padx=10,pady=(10,4))
        self.description=tk.Text(form,height=4,width=43,wrap='word',font=('Segoe UI',10))
        self.description.grid(row=n+1,column=0,columnspan=2,padx=10,sticky='ew')
        ttk.Button(form,text='Use this description to suggest boundary fields',command=self.use_description).grid(row=n+2,column=0,columnspan=2,padx=10,pady=6)
        actions=ttk.Frame(self.form_tab,padding=10);actions.grid(row=2,column=0,columnspan=2,sticky='ew')
        ttk.Checkbutton(actions,text='I reviewed the entered boundary fields',variable=self.case_reviewed).pack(side='left')
        self.find_button=ttk.Button(actions,text='Find HCAD Parcels',command=self.find)
        self.find_button.pack(side='right')
        viewer=ttk.Frame(self.form_tab);viewer.grid(row=1,column=1,sticky='nsew');viewer.rowconfigure(1,weight=1);viewer.columnconfigure(0,weight=1)
        nav=ttk.Frame(viewer);nav.grid(row=0,column=0,sticky='ew')
        ttk.Button(nav,text='Previous page',command=lambda:self.show_page(-1)).pack(side='left')
        ttk.Button(nav,text='Next page',command=lambda:self.show_page(1)).pack(side='left',padx=6)
        ttk.Button(nav,text='Rotate page',command=self.rotate_page).pack(side='left')
        ttk.Button(nav,text='Open PDF',command=self.open_pdf).pack(side='left',padx=6)
        self.page_label=ttk.Label(nav,style='Import.TLabel',text='');self.page_label.pack(side='right')
        self.preview=tk.Canvas(viewer,bg='#DDE3EA',highlightthickness=0)
        self.preview.grid(row=1,column=0,sticky='nsew',pady=8)
        self.preview.bind('<Configure>',lambda e:self.show_page())

    def _build_parcels(self):
        tab=self.parcel_tab;tab.columnconfigure(0,weight=1);tab.rowconfigure(2,weight=1)
        ttk.Label(tab,style='Import.TLabel',text='Confirm every parcel in the application boundary',font=('Segoe UI',13,'bold')).grid(row=0,column=0,sticky='w',padx=10,pady=8)
        bar=ttk.Frame(tab,padding=(10,0));bar.grid(row=1,column=0,sticky='ew')
        for text,cmd in [('Include / remove',self.toggle),('Edit property',self.edit_property),('Add HCAD account',self.add_account),('Download HCAD PDFs',self.download_records),('Read HCAD PDFs',self.read_records)]:
            ttk.Button(bar,text=text,command=cmd).pack(side='left',padx=(0,6),pady=5)
        pane=ttk.Panedwindow(tab,orient='horizontal');pane.grid(row=2,column=0,sticky='nsew',padx=10,pady=6)
        table=ttk.Frame(pane);pane.add(table,weight=3);table.rowconfigure(1,weight=1);table.columnconfigure(0,weight=1)
        ttk.Checkbutton(table,text='Show selected parcels only',variable=self.only_selected,command=self.refresh_rows).grid(row=0,column=0,sticky='w')
        columns=['include','address','account','area','lots','use','year','review']
        self.tree=ttk.Treeview(table,columns=columns,show='headings',selectmode='extended')
        for col,title,width in zip(columns,['In','Property address','HCAD account','Property sq ft','Lots','Use','Year','Details'],[32,155,116,95,42,42,45,210]):
            self.tree.heading(col,text=title);self.tree.column(col,width=width,minwidth=width,stretch=col in ('address','review'))
        self.tree.grid(row=1,column=0,sticky='nsew')
        y=ttk.Scrollbar(table,orient='vertical',command=self.tree.yview);y.grid(row=1,column=1,sticky='ns');self.tree.configure(yscrollcommand=y.set)
        x=ttk.Scrollbar(table,orient='horizontal',command=self.tree.xview);x.grid(row=2,column=0,sticky='ew');self.tree.configure(xscrollcommand=x.set)
        self.tree.bind('<Double-1>',lambda e:self.toggle() if self.tree.identify_column(e.x)=='#1' else self.edit_property())
        self.tree.bind('<<TreeviewSelect>>',lambda e:self.draw_parcels())
        self.tree.tag_configure('outside',foreground='#667085');self.tree.tag_configure('review',foreground='#9A5000')
        right=ttk.Frame(pane);pane.add(right,weight=2);right.rowconfigure(1,weight=1);right.columnconfigure(0,weight=1)
        ttk.Label(right,style='Import.TLabel',text='Blue: included   Purple: EXC   Gray: nearby\nClick a parcel to locate it in the list. North is up.',wraplength=340).grid(row=0,column=0,sticky='w')
        self.map=tk.Canvas(right,bg='white',highlightthickness=1,highlightbackground='#CBD5E1')
        self.map.grid(row=1,column=0,sticky='nsew',pady=6);self.map.bind('<Configure>',lambda e:self.draw_parcels());self.map.bind('<Button-1>',self.map_click)
        self.warnings=tk.Text(tab,height=4,wrap='word',font=('Segoe UI',9),bg='#FFF8EB',relief='flat')
        self.warnings.grid(row=3,column=0,sticky='ew',padx=10,pady=5)
        self.count=ttk.Label(tab,style='Import.TLabel',text='No parcels loaded.');self.count.grid(row=4,column=0,sticky='w',padx=10)
        ttk.Checkbutton(tab,text='I accept this boundary and the selected properties. Save incomplete information with review notes.',variable=self.parcels_reviewed).grid(row=5,column=0,sticky='w',padx=10,pady=8)
        bottom=ttk.Frame(tab);bottom.grid(row=6,column=0,sticky='ew',padx=10,pady=8)
        note='Listed whole lots get separate Step 1 rows with equal shares of the property area. Council information fills automatically when you create a map.'
        if self.wizard.endswith('_renewal'):note+=' Saving also looks up the previous ordinance and available dates. Confirm the existing minimum and original property values.'
        else:note+=' Support stays blank until reviewed.'
        if self.wizard.startswith('mbl'): note+=' Use Edit property for building-line measurements; HCAD does not supply setbacks.'
        ttk.Label(bottom,style='Import.TLabel',text=note+' EXC parcels remain in the boundary.',wraplength=650).pack(side='left')
        self.save_button=ttk.Button(bottom,text='Save New Calculation Sheet',command=self.save);self.save_button.pack(side='right')

    def case_changed(self,*args):
        self.case_reviewed.set(False);self.parcels_reviewed.set(False)

    def values(self):
        values={k:v.get().strip() for k,v in self.fields.items()}
        values['boundary_description']=self.description.get('1.0','end').strip()
        return values

    def start(self,work,callback):
        if self.busy: return False
        gc.collect();self.busy=True;self.cancel.clear()
        self.find_button.configure(state='disabled');self.save_button.configure(state='disabled')
        def run():
            try: self.events.put(('done',callback,work()))
            except Exception as e: self.events.put(('error',str(e)))
            finally:
                if self.closed:
                    try:self.temp.cleanup()
                    except OSError:pass
        threading.Thread(target=run,daemon=True).start()
        return True

    def progress(self,message): self.events.put(('progress',message))

    def poll(self):
        if self.closed: return
        try:
            while True:
                event=self.events.get_nowait()
                if event[0]=='progress': self.status.set(event[1])
                else:
                    self.busy=False;self._saving=False;self.find_button.configure(state='normal');self.save_button.configure(state='normal')
                    if event[0]=='error':
                        self.status.set(event[1]);messagebox.showerror('Application Import',event[1],parent=self)
                    else:
                        try: event[1](event[2])
                        except Exception as e:
                            self.status.set(str(e));messagebox.showerror('Application Import',str(e),parent=self)
        except queue.Empty: pass
        if not self.closed:self._timer=self.after(120,self.poll)

    def choose_pdf(self):
        if self.busy:return
        path=filedialog.askopenfilename(parent=self,title=f'Select {self.wizard_name} application',filetypes=[('Application PDF','*.pdf')])
        if not path:return
        try:
            new=fitz.open(path)
            if new.needs_pass: new.close();raise ValueError('Select an unlocked PDF.')
            if self.pdf:self.pdf.close()
            self.pdf=new;self.page=0;self.rotation=0;self.file_label.configure(text=Path(path).name)
            self.application=None;self.data=None;self.refresh_rows()
            for v in self.fields.values():v.set('')
            self.description.delete('1.0','end');self.show_page()
            self.start(lambda:read_application(path,self.progress,self.cancel),self.application_loaded)
        except Exception as e:messagebox.showerror('Application Import',str(e),parent=self)

    def application_loaded(self,result):
        self.application=result
        for k,v in result['fields'].items():
            if k in self.fields:self.fields[k].set(v)
        if self.wizard.endswith('_renewal'):
            minimum_key='minimum_lot_size' if self.wizard.startswith('mls') else 'minimum_building_line'
            self.fields['renewal_minimum'].set(result['fields'].get(minimum_key,''))
        detected=result.get('detected_wizard')
        notes=list(result.get('warnings',[]))
        if detected and detected!=self.wizard:
            notes.insert(0,f'The PDF appears to be {wizard_label(detected)}. You opened {self.wizard_name}; the saved sheet will use {self.wizard_name}. You can continue testing or open the matching wizard.')
        if self.wizard.endswith('_renewal'):
            notes.append('Review the renewal case/reference number and existing minimum against the prior designation.')
        self.reading_note.set('\n'.join(notes) or f'Review the suggested {self.wizard_name} fields.')
        self.description.delete('1.0','end');self.description.insert('1.0',result['fields'].get('boundary_description',''))
        self.page=result.get('cover_page',1)-1;self.show_page()
        self.status.set('Application read. Check the suggested fields beside the scan; fill any missing details. Phone and email are retained in the saved source record.')

    def use_description(self):
        if self.busy:return
        parsed=parse_boundary(self.description.get('1.0','end'))
        for key in ('street','block_range','sides','cross_1','cross_2'):
            if key in parsed:self.fields[key].set(parsed[key])
        self.status.set('Boundary fields suggested. Review the street, range, sides and both cross streets.')

    def show_page(self,delta=0):
        if self.pdf is None:
            self.preview.delete('all')
            self.preview.create_text(max(100,self.preview.winfo_width()/2),max(70,self.preview.winfo_height()/2),
                text='No application PDF needed\n\nEnter the boundary fields on the left,\nthen click Find HCAD Parcels.',
                fill='#475569',font=('Segoe UI',11),justify='center',width=max(180,self.preview.winfo_width()-40))
            return
        self.page=max(0,min(len(self.pdf)-1,self.page+delta))
        page=self.pdf[self.page];width=max(100,self.preview.winfo_width()-20);height=max(100,self.preview.winfo_height()-20)
        rect=page.rect if self.rotation%180==0 else fitz.Rect(0,0,page.rect.height,page.rect.width)
        scale=min(width/rect.width,height/rect.height)
        pix=page.get_pixmap(matrix=fitz.Matrix(scale,scale).prerotate(self.rotation),alpha=False)
        self.image=ImageTk.PhotoImage(Image.open(io.BytesIO(pix.tobytes('png'))))
        self.preview.delete('all');self.preview.create_image(width/2+10,height/2+10,image=self.image,anchor='center')
        self.page_label.configure(text=f'Page {self.page+1} of {len(self.pdf)}')

    def rotate_page(self):self.rotation=(self.rotation+90)%360;self.show_page()
    def open_pdf(self):
        if self.pdf is not None:
            import os
            os.startfile(self.pdf.name)

    def find(self):
        if self.busy:return
        fields=self.values()
        try:validate_fields(fields,wizard=self.wizard)
        except ValueError as e:messagebox.showerror('Application Import',str(e),parent=self);return
        if not self.application:self.application=manual_application(fields)
        self.start(lambda:find_parcels(fields,self.work/'cache',self.application['property_hints'],self.progress,self.cancel),self.parcels_loaded)

    def parcels_loaded(self,data):
        self.data=data;self.parcels_reviewed.set(False);self.only_selected.set(False)
        self.refresh_rows();self.book.select(self.parcel_tab)
        self.warnings.delete('1.0','end');self.warnings.insert('1.0','\n'.join(data['warnings']) or 'Review the property list against the application boundary. HCAD classifications are suggestions; tax years are shown in the list.')
        self.status.set('Parcels loaded. Review the blue selection and nearby parcels; double-click a row to edit its information.')

    def refresh_rows(self):
        selection=set(self.tree.selection()) if hasattr(self,'tree') else set()
        if not hasattr(self,'tree'):return
        self.tree.delete(*self.tree.get_children())
        rows=self.data['parcels'] if self.data else []
        from .lot_divisions import proposed_lots
        for p in rows:
            if self.only_selected.get() and not p['include']:continue
            lots=proposed_lots(p);details=list(p['notes'])
            if lots and lots[0].get('allocation')=='equal' and lots[0].get('area') is not None:
                details.insert(0,f"{len(lots)} lots at {lots[0]['area']:,.2f} sq ft each")
            self.tree.insert('', 'end',iid=p['account'],values=('Yes' if p['include'] else '',p['address'] or '(no site address)',p['account'],f"{p['land_area']:,.2f}".rstrip('0').rstrip('.') if p['land_area'] else '',len(lots) or 1,p['land_use'],p['tax_year'],'; '.join(details)),tags=('outside' if not p['include'] else ('review' if p['notes'] else 'selected'),))
        existing=[x for x in selection if self.tree.exists(x)]
        if existing:self.tree.selection_set(existing)
        included=[p for p in rows if p['include']]
        lot_count=sum(len(proposed_lots(p)) or 1 for p in included)
        self.count.configure(text=f'{len(included)} selected properties will create {lot_count} Step 1 rows. {sum(bool(p["record_pdfs"]) for p in included)} properties have an HCAD PDF attached.')
        self.draw_parcels()

    def toggle(self):
        if self.busy or not self.data:return
        accounts=set(self.tree.selection())
        for p in self.data['parcels']:
            if p['account'] in accounts:p['include']=not p['include']
        self.parcels_reviewed.set(False);self.refresh_rows()

    def selected_row(self):
        selected=self.tree.selection()
        return next((p for p in (self.data or {}).get('parcels',[]) if selected and p['account']==selected[0]),None)

    def edit_property(self):
        if self.busy:return
        row=self.selected_row()
        if not row:return
        dlg=tk.Toplevel(self);dlg.title('Review Property - '+row['account']);dlg.transient(self);dlg.geometry('720x770')
        frame=ScrollableContent(dlg,bg='#F8FAFC');frame.pack(fill='both',expand=True);form=frame.interior
        ttk.Label(form,style='Import.TLabel',text=f"{row['address']}\nHCAD {row['account']} | Tax year {row['tax_year']}",font=('Segoe UI',12,'bold')).grid(row=0,column=0,columnspan=2,sticky='w',padx=12,pady=10)
        items=[('land_area','Lot area (square feet)'),('land_use','City land use'),('owner','Owner / mailing addressee'),('mail_addr_1','Mailing address'),('mail_addr_2','Second mailing line'),('mail_city','Mailing city'),('mail_state','Mailing state'),('mail_zip','Mailing ZIP'),('legal_description','Legal description'),('subdivision','Subdivision / legal line 2')]+PROPERTY_INPUTS[self.wizard]+[('review_note','Correction / review note')]
        variables={k:tk.StringVar(value='' if row.get(k) is None else str(row.get(k))) for k,_ in items}
        for n,(k,label) in enumerate(items,1):
            ttk.Label(form,style='Import.TLabel',text=label).grid(row=n,column=0,sticky='w',padx=12,pady=6)
            if k in ('land_use','support'):
                entry=ttk.Combobox(form,textvariable=variables[k],values=list(LAND_USES) if k=='land_use' else ['','Y','N'],state='readonly',width=40)
            else:entry=ttk.Entry(form,textvariable=variables[k],width=48)
            entry.grid(row=n,column=1,sticky='ew',padx=10,pady=6)
        included=tk.BooleanVar(value=row['include'])
        ttk.Checkbutton(form,text='Include in the application boundary (EXC remains included)',variable=included).grid(row=len(items)+1,column=0,columnspan=2,sticky='w',padx=12,pady=8)
        note='\n'.join(row['notes']) or 'Review against the application and HCAD record. Support is separate from ownership.'
        ttk.Label(form,style='Import.TLabel',text=note,wraplength=630).grid(row=len(items)+2,column=0,columnspan=2,sticky='w',padx=12,pady=10)
        def apply():
            try:
                area=float(variables['land_area'].get().replace(',',''))
                import math
                if not math.isfinite(area) or area<=0:raise ValueError('Enter a positive lot area.')
                values={k:v.get().strip() for k,v in variables.items()};values['land_area']=area
                for key,label in PROPERTY_INPUTS[self.wizard]:
                    if key!='support': values[key]=optional_measurement(values[key],label)
                changed=any(row.get(k)!=v for k,v in values.items() if k not in ('review_note','support'))
                if changed and not values['review_note']:raise ValueError('Add a short note identifying the source of your correction.')
                if row.get('land_area')!=area:
                    row['area_source']={'method':'Planner correction','note':values['review_note'],'previous_source':row.get('area_source',{})}
                row.update(values);row['include']=included.get();self.parcels_reviewed.set(False);self.refresh_rows();dlg.destroy()
            except ValueError as e:messagebox.showerror('Review Property',str(e),parent=dlg)
        buttons=ttk.Frame(dlg,padding=10);buttons.pack(fill='x')
        ttk.Button(buttons,text='Apply changes',command=apply).pack(side='right')
        ttk.Button(buttons,text='Cancel',command=dlg.destroy).pack(side='right',padx=8)
        if row['record_pdfs']:
            import os
            ttk.Button(buttons,text='Open HCAD PDF',command=lambda:os.startfile(row['record_pdfs'][-1])).pack(side='left')

    def draw_parcels(self):
        if not hasattr(self,'map'):return
        self.map.delete('all');self._canvas_items={}
        if not self.data:return
        rows=self.data['parcels'];bounds=[parcel_polygon(p['feature']).bounds for p in rows]
        if not bounds:return
        lo_x=min(b[0] for b in bounds);lo_y=min(b[1] for b in bounds);hi_x=max(b[2] for b in bounds);hi_y=max(b[3] for b in bounds)
        width=max(100,self.map.winfo_width());height=max(100,self.map.winfo_height())
        scale=min((width-30)/max(hi_x-lo_x,1),(height-30)/max(hi_y-lo_y,1))
        dx=(width-(hi_x-lo_x)*scale)/2;dy=(height-(hi_y-lo_y)*scale)/2
        def xy(x,y):return dx+(x-lo_x)*scale,height-dy-(y-lo_y)*scale
        selected=set(self.tree.selection())
        for p in rows:
            poly=parcel_polygon(p['feature'])
            fill=('#C4B5FD' if p['land_use']=='EXC' else '#93C5FD') if p['include'] else '#F1F5F9'
            parts=list(poly.geoms) if hasattr(poly,'geoms') else [poly]
            for part in parts:
                coords=[v for x,y in part.exterior.coords for v in xy(x,y)]
                item=self.map.create_polygon(coords,fill=fill,outline='#D97706' if p['account'] in selected else '#64748B',width=3 if p['account'] in selected else 1)
                self._canvas_items[item]=p['account']
            if p['include']:
                pt=poly.representative_point();x,y=xy(pt.x,pt.y)
                item=self.map.create_text(x,y,text=str(p['feature']['attributes'].get('site_str_num') or ''),font=('Arial',8))
                self._canvas_items[item]=p['account']
        for road in self.data['roads']:
            for path in road.get('geometry',{}).get('paths',[]):
                if len(path)>1:self.map.create_line(*[v for x,y in path for v in xy(x,y)],fill='#9CA3AF',dash=(3,4))
        from .boundary_map import road_key
        labelled=set()
        wanted={road_key(self.data['fields'][k]) for k in ('street','cross_1','cross_2')}
        for road in self.data['roads']:
            name=road['attributes'].get('FULL_NAME','');key=road_key(name)
            if key not in wanted or key in labelled:continue
            paths=road.get('geometry',{}).get('paths',[])
            if not paths:continue
            from shapely.geometry import LineString
            line=max((LineString(p) for p in paths if len(p)>1),key=lambda p:p.length,default=None)
            if line is None:continue
            pt=line.interpolate(.5,normalized=True);x,y=xy(pt.x,pt.y)
            if 0<x<width and 0<y<height:
                self.map.create_text(x,y,text=name.upper(),font=('Arial',9),fill='#334155')
                labelled.add(key)
        self.map.create_text(width-18,18,text='N ↑',font=('Arial',10,'bold'))

    def map_click(self,event):
        for item in reversed(self.map.find_overlapping(event.x-2,event.y-2,event.x+2,event.y+2)):
            account=self._canvas_items.get(item)
            if account:
                if not self.tree.exists(account):self.only_selected.set(False);self.refresh_rows()
                self.tree.selection_set(account);self.tree.see(account);return

    def add_account(self):
        if self.busy or not self.data:return
        value=simpledialog.askstring('Add parcel','HCAD account number:',parent=self)
        if not value:return
        def done(result):
            row,sources=result
            if any(p['account']==row['account'] for p in self.data['parcels']):
                raise ValueError('That account is already in the parcel list. Select its row and include it.')
            self.data['parcels'].append(row);self.data['sources'].extend(sources);self.parcels_reviewed.set(False);self.refresh_rows()
        self.start(lambda:lookup_account(value,self.work/'cache',self.application['property_hints'],self.progress,self.cancel),done)

    def attach_record_results(self,results):
        matched=0;unmatched=[]
        for record in results:
            row=next((p for p in self.data['parcels'] if p['account']==record['account']),None)
            if row is None:unmatched.append(Path(record['path']).name);continue
            if record['path'] not in row['record_pdfs']:row['record_pdfs'].append(record['path'])
            area=record['land_area']
            if area is not None and area!=row['land_area']:
                note=f"HCAD PDF reports {area:,.2f} sq ft; current worksheet value is {row['land_area']}. Open Edit property to choose the confirmed value."
                if note not in row['notes']:row['notes'].append(note)
            matched+=1
        self.parcels_reviewed.set(False);self.refresh_rows()
        self.status.set(f'Read {matched} HCAD PDF(s). '+('Unmatched files: '+', '.join(unmatched) if unmatched else 'Any lot-area differences are listed for review.'))

    def read_records(self):
        if self.busy or not self.data:return
        paths=filedialog.askopenfilenames(parent=self,title='Read downloaded HCAD records',filetypes=[('HCAD PDFs','*.pdf')])
        if paths:self.start(lambda:[record_pdf_facts(p,self.progress,self.cancel) for p in paths],self.attach_record_results)

    def download_records(self):
        if self.busy or not self.data:return
        selected=[p for p in self.data['parcels'] if p['include']]
        if not selected:messagebox.showinfo('HCAD Records','Select parcels first.',parent=self);return
        if not launch.hcad_available():messagebox.showerror('HCAD Records','The HCAD downloader is missing from this installation.',parent=self);return
        # Account-only searches prevent a different parcel at a shared address
        # from being selected by an address-first search.
        address_file=self.work/'HCAD accounts.txt';address_file.write_text('\n'.join(p['account'] for p in selected)+'\n',encoding='utf8')
        folder=self.work/'HCAD_Records';folder.mkdir(exist_ok=True)
        if not self.owner._is_chrome_cdp_available() and not self.owner._start_hcad_chrome_session(show_message=False):return
        messagebox.showinfo('HCAD Records','Complete HCAD human verification in the Chrome window. When the HCAD search page is visible, click OK to download the selected accounts. You can also use Read HCAD PDFs for records you already have.',parent=self)
        if not self.owner._is_chrome_cdp_available():messagebox.showerror('HCAD Records','Chrome is not ready. Retry, or use Read HCAD PDFs.',parent=self);return
        command=launch.hcad_command()+['--file',str(address_file),'--output',str(folder),'--cdp-url',self.owner._chrome_debug_url(),'--delay-min','3','--delay-max','6']
        def work():
            self.progress('Downloading HCAD records. Complete any verification in Chrome; Close cancels this download.')
            with (folder/'download.log').open('w',encoding='utf8') as log:
                self._process=subprocess.Popen(command,cwd=launch.hcad_cwd(),creationflags=launch.hidden_flags(),stdout=log,stderr=subprocess.STDOUT)
                code=self._process.wait();self._process=None
            if self.cancel.is_set():raise ValueError('HCAD download cancelled.')
            records=[record_pdf_facts(p,self.progress,self.cancel) for p in folder.glob('*.pdf')]
            if not records:raise ValueError(f'No HCAD records were downloaded (exit code {code}). Check Chrome verification or use Read HCAD PDFs.')
            return records
        self.start(work,self.attach_record_results)

    def save(self):
        if self.busy or not self.data:return
        if not self.parcels_reviewed.get():
            messagebox.showinfo('Application Import','Accept the boundary and selected properties using the checkbox above Save New Calculation Sheet.',parent=self);return
        fields=self.values()
        data=copy.deepcopy(self.data);application=copy.deepcopy(self.application or manual_application(fields))
        if not self.case_reviewed.get():
            data.setdefault('warnings',[]).append('Application fields have not been marked reviewed. Confirm them in the saved calculation sheet.')
        # Accept the reviewed selection even when lookup wording or other
        # application facts still need correction in Excel.
        prepared_fields, prepared_data=prepare_reviewed_save(fields,data,self.wizard)
        suggested_number=case_number(prepared_fields['application_number'],self.wizard) if prepared_fields['application_number'] else 'Draft'
        folder=self.owner.vars['output_folder'].get() if hasattr(self.owner,'vars') else self.owner.files.get('out','')
        path=filedialog.asksaveasfilename(parent=self,title=f'Save a new {self.wizard_name} calculation sheet',defaultextension='.xlsx',filetypes=[('Excel calculation sheet','*.xlsx')],initialdir=folder,initialfile=f"{self.wizard[:3].upper()}_{suggested_number}_Calculation.xlsx")
        if not path:return
        def done(destination):
            select_imported_calculation(self.owner,destination)
            notes=prepared_data['warnings']+calculation_review_notes(prepared_fields,[p for p in prepared_data['parcels'] if p['include']],self.wizard)
            if len(notes)>6:notes=notes[:6]+['Additional notes are saved in the calculation sheet’s _sources folder (Review notes.txt).']
            messagebox.showinfo('Calculation Sheet Created',f'Saved {destination.name} and its source records.\n\nThe workbook is selected in the {self.wizard_name} wizard. Load Workbook to recalculate and review Step 2.'+('\n\n'+'\n'.join(notes) if notes else ''),parent=self)
            self.close()
        self._saving=True
        def work():
            if self.wizard.endswith('_renewal'):
                from .designation_lookup import lookup,merge_fields
                result=lookup([p['feature'] for p in data['parcels'] if p.get('include')],self.wizard,fields,
                              Path(path).parent/'.renewal_lookup_cache',self.progress,self.cancel)
                merged,filled,notes=merge_fields(fields,result);fields.update(merged)
                result['filled']=filled;result['warnings'].extend(notes)
                application['renewal_lookup']=result
                data.setdefault('warnings',[]).extend(result['warnings'])
                prepared_data['warnings'].extend(result['warnings'])
            self.progress('Saving the calculation sheet with one row per listed lot...')
            return save_calculation(path,fields,data,application,wizard=self.wizard,reviewed=True)
        self.start(work,done)

    def close(self):
        if self.closed:return
        if self._saving:
            self.status.set('Finishing the workbook save. This window will close when it is complete.')
            return
        self.closed=True;self.cancel.set()
        if self._process is not None:
            try:self._process.terminate()
            except OSError:pass
        if self._timer:self.after_cancel(self._timer)
        if self.pdf:self.pdf.close()
        # A worker may still be finishing a bounded network/OCR call. Its temp
        # directory stays alive until that thread releases this dialog object.
        if not self.busy:
            try:self.temp.cleanup()
            except OSError:pass
        self.destroy()


def open_application_import(owner):
    existing=getattr(owner,'_application_import_window',None)
    if existing is not None and existing.winfo_exists():existing.lift();return existing
    dialog=ApplicationImport(owner);owner._application_import_window=dialog
    return dialog
