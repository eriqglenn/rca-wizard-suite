"""Planner-controlled ILMS hold entry shared by all four application wizards."""
from __future__ import annotations
import copy
import os
import queue
from pathlib import Path
import tkinter as tk
from tkinter import ttk,filedialog,messagebox,simpledialog

from .ilms_holds import load_plan,hold_reason,address_key,application_number
from .ilms_browser import ILMSBrowser

def open_ilms_holds(owner):
    """Open one hold window from the application-import action row."""
    existing=getattr(owner,'_ilms_window',None)
    if existing is not None and existing.winfo_exists():
        existing.lift();existing.focus_set();return existing
    from .scrollable import ScrollableContent
    window=tk.Toplevel(owner)
    window.title('ILMS Holds - 126 Situs Holds')
    window.transient(owner.winfo_toplevel())
    width=min(1160,max(700,window.winfo_screenwidth()-80))
    height=min(860,max(500,window.winfo_screenheight()-100))
    window.geometry(f'{width}x{height}')
    window.columnconfigure(0,weight=1);window.rowconfigure(0,weight=1)
    content=ScrollableContent(window,bg='#F8FAFC')
    content.grid(row=0,column=0,sticky='nsew')
    pane=ilms_page(owner,content.interior)
    owner._ilms_window=window
    def close():
        if pane.running:
            pane.stop()
            pane.status.set('Stopping safely. Wait for the current result, then close this window.')
            return
        window.destroy()
    ttk.Button(window,text='Close',command=close).grid(row=1,column=0,sticky='e',padx=12,pady=10)
    window.protocol('WM_DELETE_WINDOW',close)
    pane.after_idle(pane.start_automatic)
    return window


def ilms_page(owner,parent):
    parent.columnconfigure(0,weight=1)
    page=ILMSPanel(parent,owner);page.grid(row=0,column=0,columnspan=2,sticky='nsew',padx=8,pady=8)
    owner._ilms_panel=page;return page

class ILMSPanel(ttk.Frame):
    def __init__(self,parent,owner):
        super().__init__(parent,padding=12);self.owner=owner;self.wizard=owner._boundary_map_wizard_key
        self.columnconfigure(1,weight=1);self.events=queue.Queue();self.browser=None;self.running=False;self.browser_ready=False
        self.plan=copy.deepcopy(getattr(owner,'_ilms_plan',None));self.controls=[];self.report=''
        def value(*keys):
            for key in keys:
                values=getattr(owner,'vars',{});v=values[key].get() if key in values else getattr(owner,'files',{}).get(key,'')
                if v:return str(v)
            return ''
        self.workbook=tk.StringVar(value=value('excel','workbook'));self.output=tk.StringVar(value=value('output_folder','out'))
        self.application=tk.StringVar(value=self.plan['application'] if self.plan else '')
        self.status=tk.StringVar(value='Chrome will open for sign-in, then application holds will be placed automatically.')
        self.reason=tk.StringVar(value='MLS/MBL [application number] call planning 6637')
        ttk.Label(self,text='ILMS · 126 Situs Holds',font=('Segoe UI',18,'bold')).grid(row=0,column=0,columnspan=3,sticky='w')
        ttk.Label(self,text='Add the application hold to the selected property addresses, including EXC properties. Existing holds remain in place.',wraplength=900).grid(row=1,column=0,columnspan=3,sticky='w',pady=(6,16))
        for r,label,var,command in [(2,'Calculation sheet',self.workbook,self.choose_workbook),(3,'Case output folder',self.output,self.choose_output)]:
            ttk.Label(self,text=label).grid(row=r,column=0,sticky='w',padx=(0,12),pady=5)
            entry=ttk.Entry(self,textvariable=var);entry.grid(row=r,column=1,sticky='ew',pady=5)
            button=ttk.Button(self,text='Browse…',command=command);button.grid(row=r,column=2,padx=8);self.controls.extend([entry,button])
        ttk.Label(self,text='Application number').grid(row=4,column=0,sticky='w',pady=5)
        entry=ttk.Entry(self,textvariable=self.application,width=25);entry.grid(row=4,column=1,sticky='w');self.controls.append(entry)
        ttk.Label(self,textvariable=self.reason,font=('Segoe UI',10,'bold')).grid(row=5,column=0,columnspan=3,sticky='w',pady=8)
        ttk.Label(self,text='GE · General Hold     140 · Planning     H · Hard Hold     Date and Authorized By: supplied by ILMS',wraplength=900).grid(row=6,column=0,columnspan=3,sticky='w')
        actions=ttk.Frame(self);actions.grid(row=7,column=0,columnspan=3,sticky='ew',pady=14)
        for label,command in [('Start / Resume',self.start_automatic)]:
            button=ttk.Button(actions,text=label,command=command);button.pack(side='left',padx=(0,8));self.controls.append(button)
        self.stop_button=ttk.Button(actions,text='Stop',command=self.stop,state='disabled');self.stop_button.pack(side='left')
        ttk.Label(self,text='Properties are loaded automatically, including EXC. Split-lot rows share one hold per address. Stop the run before correcting an address or changing the selection.',wraplength=900).grid(row=8,column=0,columnspan=3,sticky='w')
        frame=ttk.Frame(self);frame.grid(row=9,column=0,columnspan=3,sticky='nsew',pady=8);frame.columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(frame,columns=('selected','address','rows','status'),show='headings',height=12,selectmode='browse')
        for key,label,width in [('selected','Include',65),('address','Property address',340),('rows','Step 1 rows',100),('status','Result',235)]:
            self.tree.heading(key,text=label);self.tree.column(key,width=width,minwidth=50,stretch=key in ('address','status'))
        self.tree.grid(row=0,column=0,sticky='nsew');scroll=ttk.Scrollbar(frame,orient='vertical',command=self.tree.yview);scroll.grid(row=0,column=1,sticky='ns');self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind('<Double-1>',self.toggle);self.tree.bind('<<TreeviewSelect>>',self.selection)
        edits=ttk.Frame(self);edits.grid(row=10,column=0,columnspan=3,sticky='w')
        for label,command in [('Include / Exclude Selected',self.toggle),('Correct Property Address',self.correct_address)]:
            button=ttk.Button(edits,text=label,command=command);button.pack(side='left',padx=(0,8));self.controls.append(button)
        self.details=tk.StringVar(value='');ttk.Label(self,textvariable=self.details,wraplength=900).grid(row=11,column=0,columnspan=3,sticky='w',pady=10)
        ttk.Label(self,textvariable=self.status,wraplength=900).grid(row=12,column=0,columnspan=3,sticky='w',pady=6)
        ttk.Button(self,text='Open Completion Report',command=self.open_report).grid(row=13,column=0,columnspan=3,sticky='w',pady=8)
        ttk.Label(self,text='The case folder receives a CSV report and a resumable JSON log. Entries already completed by this batch are not submitted again.',wraplength=900).grid(row=14,column=0,columnspan=3,sticky='w')
        self.application.trace_add('write',self.update_reason);self.workbook.trace_add('write',self.invalidate)
        self.bind('<Destroy>',self.destroyed,add='+');self.poll_id=self.after(150,self.poll)
        if self.plan and str(Path(self.workbook.get()).resolve())!=self.plan.get('workbook'):self.plan=None
        self.update_reason();self.render()

    def update_reason(self,*_):
        try:self.reason.set(hold_reason(self.application.get()))
        except ValueError:self.reason.set('Enter the application number to preview the hold reason.')

    def invalidate(self,*_):
        if self.running:return
        self.plan=None;self.owner._ilms_plan=None;self.render();self.status.set('Calculation sheet changed. Load Properties again.')

    def choose_workbook(self):
        path=filedialog.askopenfilename(parent=self,title='Choose calculation sheet',filetypes=[('Excel calculation sheet','*.xlsx *.xlsm')])
        if path:self.workbook.set(path)

    def choose_output(self):
        path=filedialog.askdirectory(parent=self,title='Choose the case output folder')
        if path:self.output.set(path)

    def load(self):
        try:
            self.plan=load_plan(self.workbook.get(),self.wizard);self.application.set(self.plan['application']);self.owner._ilms_plan=self.plan
            self.render();self.status.set(f"{len(self.plan['properties'])} property address entries loaded. EXC properties are included.")
        except Exception as error:messagebox.showerror('ILMS Holds',str(error),parent=self)

    def render(self):
        selected=self.tree.selection();self.tree.delete(*self.tree.get_children())
        for p in (self.plan or {}).get('properties',[]):
            self.tree.insert('','end',iid=p['id'],values=('Yes' if p['selected'] else 'No',p['address'] or '(no address)',', '.join(map(str,p['rows'])),p['status']))
        if selected and self.tree.exists(selected[0]):self.tree.selection_set(selected[0])

    def selected_property(self):
        ids=self.tree.selection()
        return next((p for p in (self.plan or {}).get('properties',[]) if ids and p['id']==ids[0]),None)

    def selection(self,*_):
        p=self.selected_property();self.details.set(p.get('detail','') if p else '')

    def toggle(self,*_):
        if self.running:return
        p=self.selected_property()
        if p:
            if not p['key']:self.status.set('Correct this property’s actual address before including it.');return
            p['selected']=not p['selected'];self.render()

    def correct_address(self):
        if self.running:return
        p=self.selected_property()
        if not p:return
        value=simpledialog.askstring('Property address','Enter the exact ILMS property address, including any direction, fraction or unit. This does not edit the calculation sheet.',initialvalue=p['address'],parent=self)
        if value is not None:
            key=address_key(value)
            if not key:messagebox.showerror('Property address','Enter a house number, street name and street type.',parent=self);return
            p.update(address=value.strip(),key=key,selected=True,status='Ready to find',detail='Planner-corrected address for ILMS only.');self.render()

    def start_automatic(self):
        if self.running:return
        try:
            from .workbooks import read_shared_bytes
            import hashlib
            # Preserve explicit address corrections on Resume only while the
            # workbook has not changed; otherwise rebuild the property list.
            path=self.workbook.get().strip()
            digest=hashlib.sha256(read_shared_bytes(path)).hexdigest()
            if not self.plan or self.plan.get('workbook_sha256')!=digest:
                self.plan=load_plan(path,self.wizard)
                self.application.set(self.plan['application']);self.render()
            app=application_number(self.application.get())
            if not any(p['selected'] for p in self.plan['properties']):
                raise ValueError('No complete property addresses were found. Correct an address below, then click Start / Resume.')
            if not self.output.get().strip():
                raise ValueError('Choose the case output folder, then click Start / Resume.')
            self.plan['application']=app;self.owner._ilms_plan=self.plan
            data={'wizard':self.wizard,'application':app,'properties':copy.deepcopy(self.plan['properties']),
                  'output':self.output.get(),'workbook':path,'workbook_sha256':digest}
            if not self.browser:self.browser=ILMSBrowser(self.events)
            self.browser.start();self.browser.send('automatic',data)
            self.busy(True)
            self.status.set('Opening Chrome. Sign in to ILMS; property holds will start automatically.')
        except Exception as error:
            self.status.set(str(error))
            messagebox.showerror('ILMS Holds',str(error),parent=self)

    def busy(self,value):
        self.running=value
        for c in self.controls:c.configure(state='disabled' if value else 'normal')
        self.stop_button.configure(state='normal' if value else 'disabled')

    def stop(self):
        if self.browser:self.browser.stop.set()
        self.status.set('Stopping before the next hold. Any submitted hold will finish verification.')

    def poll(self):
        try:
            while True:
                kind,value=self.events.get_nowait()
                if kind=='browser_ready':self.browser_ready=True;self.status.set(value)
                elif kind=='browser_closed':self.browser_ready=False;self.busy(False)
                elif kind=='progress':self.status.set(value)
                elif kind=='property':
                    p=next((p for p in (self.plan or {}).get('properties',[]) if p['id']==value['id']),None)
                    if p:p.update(value);self.render();self.selection()
                elif kind=='done':self.busy(False);self.report=value['report'];self.status.set(value['message'])
                elif kind=='error':self.busy(False);self.status.set(value)
        except queue.Empty:pass
        self.poll_id=self.after(150,self.poll)

    def open_report(self):
        if self.report and Path(self.report).is_file():os.startfile(self.report)
        else:self.status.set('Start the ILMS run to create the case report.')

    def destroyed(self,event):
        if event.widget is self:
            if self.browser:self.browser.close()
            if self.plan:self.owner._ilms_plan=self.plan
            if getattr(self,'poll_id',None):
                try:self.after_cancel(self.poll_id)
                except tk.TclError:pass
