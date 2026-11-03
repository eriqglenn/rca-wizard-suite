"""PC Speech lives with the MLS Planning Commission document actions."""
from __future__ import annotations
import copy
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from core.workbooks import WorkbookReadCache
from core.scrollable import ScrollableContent
from .pc_data import load_preferences, save_preferences
from .pc_speech import (FIELDS, RECOMMENDATIONS, defaults, read_speech_facts,
                        speech_context, speech_text, generate_speech)

class PCSpeech(tk.Toplevel):
    def __init__(self,owner):
        super().__init__(owner)
        self.owner=owner
        from core.boundary_map_ui import wizard_paths
        self.source,_,self.output=wizard_paths(owner)
        from . import pc_speech as api
        if getattr(owner,'_boundary_map_wizard_key','mls')=='mls_renewal':
            from wizards.mls_renewal import pc_speech as api
        wizard_key=getattr(owner,'_boundary_map_wizard_key','mls')
        if wizard_key=='mbl':
            from wizards.mbl import pc_speech as api
        elif wizard_key=='mbl_renewal':
            from wizards.mbl_renewal import pc_speech as api
        self.api=api;self.renewal=wizard_key.endswith('_renewal');self.building=wizard_key.startswith('mbl')
        self.title('PC Speech — '+('MBL' if self.building else 'MLS')+(' Renewal' if self.renewal else ''));self.geometry('1100x790');self.minsize(920,700)
        self.transient(owner.winfo_toplevel());self.protocol('WM_DELETE_WINDOW',self.close)
        self.facts=None;self.busy=False;self.closed=False;self.events=queue.Queue();self.preview_timer=None
        self.cache=WorkbookReadCache()
        self.columnconfigure(1,weight=1);self.rowconfigure(1,weight=1)
        ttk.Label(self,text='PC Speech',font=('Segoe UI',18,'bold')).grid(row=0,column=0,columnspan=2,sticky='w',padx=16,pady=12)
        # Keep the final confirmation reachable on small screens and at larger DPI.
        holder=ScrollableContent(self,bg=ttk.Style(self).lookup('TFrame','background') or '#F8FAFC',width=330)
        holder.grid(row=1,column=0,sticky='nsew');holder.grid_propagate(False)
        self.input_scroller=holder
        inputs=ttk.Frame(holder.interior,padding=(16,0,12,8));inputs.pack(fill='x');inputs.columnconfigure(0,weight=1)
        ttk.Label(inputs,text='Confirm the presentation details.\nFigures and boundary come from Step 2.',wraplength=295).grid(row=0,column=0,sticky='w',pady=(0,10))
        self.fields={};self.editors=[]
        for r,(key,label,_ci) in enumerate(self.api.FIELDS):
            self.fields[key]=tk.StringVar()
            ttk.Label(inputs,text=label).grid(row=r*2+1,column=0,sticky='w')
            entry=ttk.Entry(inputs,textvariable=self.fields[key],width=37,state='disabled')
            entry.grid(row=r*2+2,column=0,sticky='ew',pady=(0,5));self.editors.append(entry)
            self.fields[key].trace_add('write',self.schedule_preview)
        r=len(self.api.FIELDS)*2+1
        ttk.Label(inputs,text='Staff recommendation').grid(row=r,column=0,sticky='w')
        self.recommendation=tk.StringVar(value='Approval')
        combo=ttk.Combobox(inputs,textvariable=self.recommendation,values=list(RECOMMENDATIONS),state='disabled')
        combo.grid(row=r+1,column=0,sticky='ew',pady=(0,8));self.editors.append(combo)
        self.recommendation.trace_add('write',self.schedule_preview)
        self.notice=tk.BooleanVar(value=False);self.notice.trace_add('write',self.schedule_preview)
        check=ttk.Checkbutton(inputs,text='Owner notifications have been mailed',variable=self.notice,state='disabled')
        check.grid(row=r+2,column=0,sticky='w',pady=(0,6));self.editors.append(check)
        ttk.Label(inputs,text='Confirm that the protest count includes timely written protests. Enter 0 if none.',wraplength=290).grid(row=r+3,column=0,sticky='w')
        right=ttk.Frame(self,padding=(0,0,16,8));right.grid(row=1,column=1,sticky='nsew')
        right.columnconfigure(0,weight=1);right.rowconfigure(1,weight=1)
        ttk.Label(right,text='Speech preview',font=('Segoe UI',12,'bold')).grid(row=0,column=0,sticky='w',pady=(0,6))
        self.preview=ScrolledText(right,wrap='word',font=('Times New Roman',13),width=55,state='disabled',padx=15,pady=12)
        self.preview.grid(row=1,column=0,sticky='nsew')
        self.details=tk.StringVar();details_label=ttk.Label(right,textvariable=self.details,wraplength=650)
        details_label.grid(row=2,column=0,sticky='w',pady=(7,0))
        self.details_label=details_label;self._details_width=None
        def wrap_details(event):
            width=max(100,event.width-24)
            if self._details_width!=width:
                self._details_width=width;details_label.configure(wraplength=width)
        right.bind('<Configure>',wrap_details)
        actions=ttk.Frame(self,padding=14);actions.grid(row=2,column=0,columnspan=2,sticky='ew');actions.columnconfigure(0,weight=1)
        self.status=tk.StringVar(value='Reading the saved calculation sheet…')
        ttk.Label(actions,textvariable=self.status,wraplength=650).grid(row=0,column=0,sticky='w')
        self.save_button=ttk.Button(actions,text='Save PC Speech',command=self.save,state='disabled');self.save_button.grid(row=0,column=1,padx=8)
        ttk.Button(actions,text='Close',command=self.close).grid(row=0,column=2)
        self.timer=self.after(80,self.poll)
        def worker():
            try:
                resolved=self.cache.resolve(self.source);facts=self.api.read_speech_facts(resolved,self.source)
                self.events.put(('facts',(facts,load_preferences(self.source,self.output,facts['workbook_hash']))))
            except Exception as error:self.events.put(('error',str(error)))
        threading.Thread(target=worker,daemon=True).start()

    def options(self):
        return {**{k:v.get() for k,v in self.fields.items()},'notice_mailed':self.notice.get(),'recommendation':self.recommendation.get()}

    def schedule_preview(self,*_):
        if self.closed:return
        if self.preview_timer:self.after_cancel(self.preview_timer)
        self.preview_timer=self.after(150,self.update_preview)

    def update_preview(self):
        if self.preview_timer:self.after_cancel(self.preview_timer)
        self.preview_timer=None
        if self.closed:return
        if not self.facts:return
        ctx,notes=self.api.speech_context(self.facts,self.options())
        self.preview.configure(state='normal');self.preview.delete('1.0','end');self.preview.insert('1.0',speech_text(ctx));self.preview.configure(state='disabled')
        self.details.set(('SFR: lot count • Commercial: land area • Petition support not required for renewal' if self.renewal else 'SFR: lot count • Commercial and support: land area • Split lots count individually')+
                         (' • EXC excluded from these figures' if self.facts['excluded_count'] else ''))
        if self.building:self.details.set('SFR and support: land area • Building line: feet • Developed-lot denominator excludes vacant and EXC lots')
        if not self.busy:
            self.status.set(('Draft — '+str(len(notes))+' item(s) need completion. Missing items stay marked in the Word file.') if notes else 'Ready to save the filled speech as editable Word and PDF.')

    def remember(self):
        # Merge the latest preferences so a speech never discards PC map/report settings.
        values=load_preferences(self.source,self.output,self.facts['workbook_hash'])
        values['speech']=self.options()
        save_preferences(self.source,self.output,self.facts['workbook_hash'],values)

    def save(self):
        if not self.facts or self.busy:return
        try:self.remember()
        except OSError as error:
            messagebox.showerror('PC Speech',str(error),parent=self);return
        self.busy=True;self.save_button.configure(state='disabled');self.status.set('Saving the speech and exporting PDF…')
        facts=copy.deepcopy(self.facts);options=self.options();source=self.source;output=self.output
        def worker():
            try:self.events.put(('saved',generate_speech(source,facts,options,output)))
            except Exception as error:self.events.put(('error',str(error)))
        threading.Thread(target=worker,daemon=True).start()

    def poll(self):
        if self.closed:return
        try:
            while True:
                kind,value=self.events.get_nowait()
                if kind=='facts':
                    self.facts,prefs=value;opts=self.api.defaults(self.facts,prefs)
                    for key,var in self.fields.items():var.set(opts.get(key,''))
                    self.notice.set(opts['notice_mailed']);self.recommendation.set(opts['recommendation'])
                    for editor in self.editors:editor.configure(state='readonly' if isinstance(editor,ttk.Combobox) else 'normal')
                    self.save_button.configure(state='normal');self.update_preview()
                elif kind=='saved':
                    self.busy=False;self.save_button.configure(state='normal');self.status.set('Saved '+value['docx'])
                    self.last_result=value
                    messagebox.showinfo('PC Speech','Saved editable Word'+(' and PDF' if value['pdf'] else '')+' speech.\n\n'+value['folder']+
                                        ('\n\nReview before presenting:\n'+'\n'.join(value['notes'][:8]) if value['notes'] else ''),parent=self)
                elif kind=='error':
                    self.busy=False;self.status.set(value)
                    if self.facts:self.save_button.configure(state='normal')
                    messagebox.showerror('PC Speech',value,parent=self)
        except queue.Empty:pass
        self.timer=self.after(100,self.poll)

    def close(self):
        if self.busy:
            self.status.set('The speech is still saving. This window can close when saving finishes.');return
        if self.facts:
            try:self.remember()
            except OSError:pass
        self.closed=True;self.after_cancel(self.timer)
        if self.preview_timer:self.after_cancel(self.preview_timer)
        self.destroy()

def open_pc_speech(owner):
    # Speech generation is read-only: no map, PDF application or external lookup required.
    from core.boundary_map_ui import wizard_paths
    source,_,output=wizard_paths(owner)
    if not source or not Path(source).is_file():raise ValueError('Select a saved calculation sheet first.')
    if not output:raise ValueError('Select the case output folder first.')
    prior=getattr(owner,'_pc_speech',None)
    if prior is not None and prior.winfo_exists():prior.lift();return prior
    window=PCSpeech(owner);owner._pc_speech=window;return window
