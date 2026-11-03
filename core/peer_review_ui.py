"""Peer Reviewer tab shared by all four wizards; work happens off the UI thread."""
from __future__ import annotations
import os
import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from .peer_review import run_review, save_report


def defaults(owner):
    def get(*keys):
        for key in keys:
            values=getattr(owner,'vars',{})
            value=values[key].get() if key in values else getattr(owner,'files',{}).get(key,'')
            if value:return str(value)
        return ''
    key=owner._boundary_map_wizard_key
    output=Path(get('output_folder','out') or '.')
    stem={'mls':'MLS','mbl':'MBL','mls_renewal':'MLS_Renewal','mbl_renewal':'MBL_Renewal'}[key]
    def existing(name):
        p=output/name
        return str(p.resolve()) if p.is_file() else ''
    return {'calculation':get('excel','workbook'), 'hcad':get('hcad_records'),
            'rca':existing(stem+'_RCA_Packet.pdf'),
            'ordinance':existing(stem+'_Draft_Ordinance.docx'),'output':str(output),'ocr':True}


def reviewer_page(owner,parent):
    parent.columnconfigure(0,weight=1)
    review=PeerReviewer(parent,owner)
    review.grid(row=0,column=0,columnspan=2,sticky='nsew',padx=6,pady=4)
    owner._peer_reviewer=review
    return review


class PeerReviewer(ttk.Frame):
    def __init__(self,parent,owner):
        super().__init__(parent,padding=16)
        self.owner=owner;self.wizard=owner._boundary_map_wizard_key
        self.state=getattr(owner,'_peer_review_state',None) or defaults(owner)
        owner._peer_review_state=self.state
        self.result=self.state.get('result');self.cancel=None;self.messages=queue.Queue();self.running=False
        self.columnconfigure(1,weight=1)
        ttk.Label(self,text='Review the finished case',font=('Segoe UI',17,'bold')).grid(row=0,column=0,columnspan=3,sticky='w')
        ttk.Label(self,text='Choose the saved files the peer reviewer should check. Maps are read from the RCA packet. Findings are advisory; your files stay unchanged.',wraplength=930).grid(row=1,column=0,columnspan=3,sticky='w',pady=(5,16))
        self.vars={};self.controls=[]
        for row,(key,label) in enumerate([('rca','RCA packet'),('hcad','HCAD records'),('ordinance','Draft ordinance'),('calculation','Calculation sheet')],2):
            ttk.Label(self,text=label).grid(row=row,column=0,sticky='w',padx=(0,14),pady=5)
            var=tk.StringVar(value=self.state.get(key,''));self.vars[key]=var
            entry=ttk.Entry(self,textvariable=var);entry.grid(row=row,column=1,sticky='ew',pady=5)
            button=ttk.Button(self,text='Choose files…' if key=='hcad' else 'Choose file…',command=lambda k=key:self.choose(k))
            button.grid(row=row,column=2,padx=(10,0));self.controls.extend([entry,button])
            var.trace_add('write',lambda *_args,k=key:self.changed(k))
        self.ocr=tk.BooleanVar(value=self.state.get('ocr',True))
        ocr=ttk.Checkbutton(self,text='Read scanned pages and embedded maps with local text recognition',variable=self.ocr,command=self.ocr_changed)
        ocr.grid(row=6,column=0,columnspan=3,sticky='w',pady=(8,4));self.controls.append(ocr)
        ttk.Label(self,text='HCAD: select individual records or a combined PDF. Multiple paths are separated by ||. Scanned values are marked for confirmation.',wraplength=930).grid(row=7,column=0,columnspan=3,sticky='w',pady=(0,12))
        actions=ttk.Frame(self);actions.grid(row=8,column=0,columnspan=3,sticky='ew')
        self.run=ttk.Button(actions,text='Run Peer Review',command=self.start);self.run.pack(side='left')
        self.stop=ttk.Button(actions,text='Cancel',command=self.stop_work,state='disabled');self.stop.pack(side='left',padx=8)
        self.save=ttk.Button(actions,text='Save Review Report',command=self.export,state='disabled');self.save.pack(side='right')
        self.status=tk.StringVar(value='Choose the files, then run the review.')
        ttk.Label(self,textvariable=self.status,wraplength=930).grid(row=9,column=0,columnspan=3,sticky='w',pady=(10,6))
        self.progress=ttk.Progressbar(self,mode='indeterminate');self.progress.grid(row=10,column=0,columnspan=3,sticky='ew',pady=(0,10))
        result_frame=ttk.Frame(self);result_frame.grid(row=11,column=0,columnspan=3,sticky='nsew');result_frame.columnconfigure(0,weight=1)
        self.tree=ttk.Treeview(result_frame,columns=('result','finding','location'),show='headings',height=9,selectmode='browse')
        for key,label,width in [('result','Result',110),('finding','Finding',340),('location','Location',360)]:
            self.tree.heading(key,text=label);self.tree.column(key,width=width,minwidth=90,stretch=key!='result')
        self.tree.grid(row=0,column=0,sticky='nsew');scroll=ttk.Scrollbar(result_frame,orient='vertical',command=self.tree.yview)
        scroll.grid(row=0,column=1,sticky='ns');self.tree.configure(yscrollcommand=scroll.set)
        for tag,color in [('Error','#a52424'),('Needs review','#8a5700'),('Passed','#236244')]:self.tree.tag_configure(tag,foreground=color)
        self.tree.bind('<<TreeviewSelect>>',self.select)
        self.details=tk.Text(self,height=6,wrap='word',font=('Segoe UI',10),background='#f8fafc',relief='solid',borderwidth=1,state='disabled')
        self.details.grid(row=12,column=0,columnspan=3,sticky='ew',pady=(10,5))
        ttk.Button(self,text='Open selected source',command=self.open_source).grid(row=13,column=0,sticky='w')
        ttk.Label(self,text='Save Review Report creates a printable HTML file that opens in your browser.',wraplength=690).grid(row=13,column=1,columnspan=2,sticky='e')
        self.bind('<Destroy>',self.destroyed,add='+')
        if self.result:self.display(self.result)
        self.poll_id=self.after(100,self.poll)

    def changed(self,key):
        self.state[key]=self.vars[key].get()
        if self.result:
            self.status.set('File selection changed. Run again to review these files.');self.save.configure(state='disabled')
            self.result=None;self.state.pop('result',None);self.tree.delete(*self.tree.get_children())
            self.details.configure(state='normal');self.details.delete('1.0','end');self.details.configure(state='disabled')

    def ocr_changed(self):
        self.state['ocr']=self.ocr.get()
        if self.result:self.changed('calculation')

    def choose(self,key):
        types=[('Calculation sheet','*.xlsx *.xlsm')] if key=='calculation' else [('Documents','*.pdf *.docx *.doc')]
        if key=='hcad':
            paths=filedialog.askopenfilenames(parent=self,title='Choose HCAD records',filetypes=[('PDF records','*.pdf')])
            if paths:self.vars[key].set('||'.join(paths))
        else:
            path=filedialog.askopenfilename(parent=self,title='Choose '+key,filetypes=types)
            if path:self.vars[key].set(path)

    def start(self):
        if self.running:return
        if not self.vars['calculation'].get().strip():
            messagebox.showinfo('Peer Reviewer','Choose the calculation sheet to establish the case information.',parent=self);return
        inputs={k:v.get().strip() for k,v in self.vars.items()}
        inputs['hcad']=[p.strip() for p in inputs['hcad'].split('||') if p.strip()]
        use_ocr=self.ocr.get();self.cancel=threading.Event();self.running=True
        self.result=None;self.state.pop('result',None);self.tree.delete(*self.tree.get_children());self.save.configure(state='disabled')
        self.run.configure(state='disabled');self.stop.configure(state='normal')
        for c in self.controls:c.configure(state='disabled')
        self.progress.start();self.status.set('Reading the selected files…')
        self.details.configure(state='normal');self.details.delete('1.0','end');self.details.configure(state='disabled')
        def worker():
            try:
                result=run_review(inputs,self.wizard,lambda t:self.messages.put(('progress',t)),self.cancel,use_ocr)
                self.messages.put(('complete',result))
            except InterruptedError:self.messages.put(('error','Review cancelled. No source files were changed.'))
            except Exception as error:self.messages.put(('error','Review could not finish: '+str(error)))
        threading.Thread(target=worker,daemon=True).start()

    def poll(self):
        try:
            while True:
                kind,value=self.messages.get_nowait()
                if kind=='progress':self.status.set(value)
                else:
                    self.running=False;self.progress.stop();self.run.configure(state='normal',text='Run Again');self.stop.configure(state='disabled')
                    for c in self.controls:c.configure(state='normal')
                    if kind=='complete':self.display(value)
                    else:self.status.set(value)
        except queue.Empty:pass
        if self.winfo_exists():self.poll_id=self.after(100,self.poll)

    def display(self,result):
        self.result=result;self.state['result']=result;self.tree.delete(*self.tree.get_children())
        counts=result['counts'];self.status.set(('Files changed during review — run again. ' if result.get('status')=='stale' else 'Review complete. ')+ ' · '.join(f'{counts.get(k,0)} {k.lower()}' for k in ('Error','Needs review','Passed'))+'. Saved-file review: '+result['created_utc'][:19].replace('T',' ')+' UTC. Run again after edits.')
        for level in ('Error','Needs review','Passed'):
            for i,f in enumerate(result['findings']):
                if f['severity']==level:self.tree.insert('','end',iid=str(i),values=(level,f['title'],'; '.join(f['locations'])),tags=(level,))
        self.save.configure(state='normal');self.run.configure(text='Run Again')
        if self.tree.get_children():self.tree.selection_set(self.tree.get_children()[0])

    def select(self,_event=None):
        if not self.result or not self.tree.selection():return
        finding=self.result['findings'][int(self.tree.selection()[0])]
        text=finding['title']+'\n\n'+finding['details']+'\n\n'+'\n'.join(finding['locations'])
        self.details.configure(state='normal');self.details.delete('1.0','end');self.details.insert('1.0',text);self.details.configure(state='disabled')

    def open_source(self):
        if not self.result or not self.tree.selection():return
        f=self.result['findings'][int(self.tree.selection()[0])];locations=' '.join(f['locations'])
        matches=[p['path'] for p in self.result['files'] if Path(p['path']).name in locations]
        if not matches and ('Step ' in locations or 'Case Info' in locations):matches=[p['path'] for p in self.result['files'] if p['role']=='Calculation sheet']
        if not matches:
            self.status.set('Use the file selectors above to identify the source; this finding refers to the whole packet.');return
        try:os.startfile(matches[0])
        except Exception as error:messagebox.showerror('Peer Reviewer',str(error),parent=self)

    def export(self):
        if not self.result:return
        path=filedialog.asksaveasfilename(parent=self,title='Save Review Report',defaultextension='.html',initialfile='Peer_Review_Report.html',initialdir=self.state.get('output'),filetypes=[('Printable HTML report','*.html')])
        if path:
            try:
                save_report(self.result,path);self.status.set('Review report saved: '+path)
            except Exception as error:messagebox.showerror('Peer Reviewer',str(error),parent=self)

    def stop_work(self):
        if self.cancel:self.cancel.set()
        self.status.set('Cancelling review…')

    def destroyed(self,event):
        if event.widget is self:
            if self.cancel:self.cancel.set()
            if getattr(self,'poll_id',None):
                try:self.after_cancel(self.poll_id)
                except tk.TclError:pass
