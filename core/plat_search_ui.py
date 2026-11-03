"""Find, review, download, and attach recorded subdivision plats."""
from __future__ import annotations
from datetime import datetime
import hashlib
import gc
import io
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog,messagebox,ttk
import webbrowser

from PIL import Image,ImageTk
from .boundary_map_ui import wizard_paths
from .plat_browser import PlatBrowser
from .plat_search import SEARCH_URL,PlatError,PlatSearchClient,read_plat_context,search_name,validated_pdf,save_plats


class PlatFinder(tk.Toplevel):
    def __init__(self,owner):
        super().__init__(owner)
        self.owner=owner
        self.source,_,self.output=wizard_paths(owner)
        self.title('Find Subdivision Plats')
        self.geometry(f'{min(1350,self.winfo_screenwidth()-60)}x{min(870,self.winfo_screenheight()-80)}')
        self.minsize(950,650)
        self.transient(owner.winfo_toplevel())
        self.protocol('WM_DELETE_WINDOW',self.close)
        self.events=queue.Queue()
        self.cancel=threading.Event()
        self.browser=PlatBrowser(self.events)
        self.records=[];self.items=[];self.current=None;self.page=0
        self.busy=False;self.browser_busy=False;self.closed=False;self._image=None
        self.columnconfigure(0,weight=1);self.rowconfigure(2,weight=1)
        top=ttk.Frame(self,padding=12);top.grid(row=0,column=0,sticky='ew')
        top.columnconfigure(1,weight=1)
        ttk.Label(top,text='Find recorded subdivision plats',font=('Segoe UI',15,'bold')).grid(row=0,column=0,columnspan=6,sticky='w')
        ttk.Label(top,text='Worksheet subdivisions').grid(row=1,column=0,sticky='w',pady=(10,5))
        self.names=ttk.Combobox(top,state='readonly')
        self.names.grid(row=1,column=1,columnspan=5,sticky='ew',padx=(8,0),pady=(10,5))
        self.names.bind('<<ComboboxSelected>>',lambda e:self.choose_name())
        self.fields={k:tk.StringVar() for k in ('name','file_number','volume','page')}
        filters=ttk.Frame(top);filters.grid(row=2,column=0,columnspan=6,sticky='ew')
        filters.columnconfigure(0,weight=1)
        for col,(key,label,width) in enumerate([('name','Subdivision / description',42),('file_number','File number',22),('volume','Volume',8),('page','Page',8)]):
            ttk.Label(filters,text=label).grid(row=0,column=col,sticky='w',padx=4)
            ttk.Entry(filters,textvariable=self.fields[key],width=width).grid(row=1,column=col,sticky='ew',padx=4)
        self.search_button=ttk.Button(filters,text='Search County Records',command=self.search)
        self.search_button.grid(row=1,column=4,padx=8)
        self.note=tk.StringVar(value='Reading subdivision names from the calculation sheet...')
        ttk.Label(self,textvariable=self.note,wraplength=1250,padding=(12,2),justify='left').grid(row=1,column=0,sticky='ew')
        body=ttk.Panedwindow(self,orient='horizontal');body.grid(row=2,column=0,sticky='nsew',padx=12,pady=8)
        left=ttk.Frame(body);right=ttk.Frame(body);body.add(left,weight=3);body.add(right,weight=2)
        left.columnconfigure(0,weight=1);left.rowconfigure(0,weight=1)
        table=ttk.Frame(left);table.grid(row=0,column=0,columnspan=2,sticky='nsew');table.columnconfigure(0,weight=1);table.rowconfigure(0,weight=1)
        self.results=ttk.Treeview(table,columns=('name','date','file','volume','pages'),show='headings',selectmode='browse',height=12)
        for key,label,width in [('name','Subdivision / replat',255),('date','Recorded',85),('file','File number',140),('volume','Vol./page',75),('pages','Pages',45)]:
            self.results.heading(key,text=label);self.results.column(key,width=width,minwidth=40,stretch=key=='name')
        self.results.grid(row=0,column=0,sticky='nsew')
        scroll=ttk.Scrollbar(table,orient='vertical',command=self.results.yview);scroll.grid(row=0,column=1,sticky='ns');self.results.configure(yscrollcommand=scroll.set)
        horizontal=ttk.Scrollbar(table,orient='horizontal',command=self.results.xview);horizontal.grid(row=1,column=0,sticky='ew');self.results.configure(xscrollcommand=horizontal.set)
        self.results.bind('<<TreeviewSelect>>',lambda e:self.selected_record())
        self.detail=tk.StringVar(value='Choose a record to see its recording details.')
        ttk.Label(left,textvariable=self.detail,wraplength=660,justify='left').grid(row=1,column=0,sticky='ew',pady=7)
        actions=ttk.Frame(left);actions.grid(row=2,column=0,sticky='w',pady=4)
        self.download_button=ttk.Button(actions,text='Download Selected Plat',command=self.download,state='disabled');self.download_button.pack(side='left')
        ttk.Button(actions,text='Open County Website',command=self.open_county).pack(side='left',padx=6)
        self.cancel_browser=ttk.Button(actions,text='Cancel Download',command=self.browser.close,state='disabled');self.cancel_browser.pack(side='left')
        ttk.Label(left,text='County images require your free Clerk login. Sign in only in the browser.\nReview the plat name, section, recording date and all pages before attaching.',wraplength=670,justify='left').grid(row=3,column=0,sticky='ew',pady=7)
        ttk.Label(left,text='Downloaded plats — select the ones to attach',font=('Segoe UI',10,'bold')).grid(row=4,column=0,sticky='w',pady=(8,3))
        self.downloaded=tk.Listbox(left,selectmode='extended',exportselection=False,height=5)
        self.downloaded.grid(row=5,column=0,sticky='ew');self.downloaded.bind('<<ListboxSelect>>',lambda e:self.preview_selection())
        ttk.Button(left,text='Add Downloaded Plat…',command=self.add_files).grid(row=6,column=0,sticky='w',pady=6)
        right.columnconfigure(0,weight=1);right.rowconfigure(1,weight=1)
        self.preview_title=tk.StringVar(value='Download or add a plat to preview it here.')
        ttk.Label(right,textvariable=self.preview_title,wraplength=440,justify='left').grid(row=0,column=0,sticky='ew',pady=4)
        self.canvas=tk.Canvas(right,bg='#E2E8F0',highlightthickness=0);self.canvas.grid(row=1,column=0,sticky='nsew',padx=(8,0))
        self.canvas.bind('<Configure>',lambda e:self.display())
        paging=ttk.Frame(right);paging.grid(row=2,column=0,pady=6)
        ttk.Button(paging,text='Previous page',command=lambda:self.turn_page(-1)).pack(side='left')
        self.page_text=tk.StringVar(value='');ttk.Label(paging,textvariable=self.page_text).pack(side='left',padx=8)
        ttk.Button(paging,text='Next page',command=lambda:self.turn_page(1)).pack(side='left')
        bottom=ttk.Frame(self,padding=12);bottom.grid(row=3,column=0,sticky='ew');bottom.columnconfigure(0,weight=1)
        self.status=tk.StringVar(value='Loading...')
        ttk.Label(bottom,textvariable=self.status,wraplength=600,justify='left').grid(row=0,column=0,sticky='w')
        self.export_button=ttk.Button(bottom,text='Save Selected Plats As…',command=self.export,state='disabled');self.export_button.grid(row=0,column=1,padx=6)
        self.use_button=ttk.Button(bottom,text='Use Selected Plats',command=self.accept,state='disabled');self.use_button.grid(row=0,column=2,padx=6)
        ttk.Button(bottom,text='Close',command=self.close).grid(row=0,column=3)
        self._timer=self.after(100,self.poll)
        self._initial=self.after(50,self.load_context)

    def worker(self,kind,operation):
        gc.collect()
        def work():
            try: self.events.put((kind,operation()))
            except Exception as exc:self.events.put(('error',str(exc)))
        threading.Thread(target=work,daemon=True).start()

    def load_context(self):
        self.worker('context',lambda:read_plat_context(self.source))

    def choose_name(self):
        self.fields['name'].set(search_name(self.names.get()))
        for k in ('file_number','volume','page'):self.fields[k].set('')

    def search(self):
        if self.busy:return
        fields={k:v.get() for k,v in self.fields.items()}
        self.busy=True;self.search_button.configure(state='disabled')
        self.status.set('Searching County recorded plats...')
        self.worker('results',lambda:PlatSearchClient(self.cancel).search(**fields))

    def selected_record(self):
        selected=self.results.selection()
        record=self.records[int(selected[0])] if selected else None
        if record:
            self.detail.set(f"{record['name']}\nRecorded {record['file_date']} | File {record['file_number']} | Volume/page {record['volume_page']} | {record['map_pages'] or '?'} map pages")
        self.download_button.configure(state='normal' if record and record['viewer_url'] and not self.browser_busy else 'disabled')
        return record

    def download(self):
        record=self.selected_record()
        if not record:return
        try:
            self.browser.start(record);self.browser_busy=True
            self.download_button.configure(state='disabled');self.cancel_browser.configure(state='normal')
            self.status.set('Opening the County browser. Sign in there and download the complete plat PDF.')
        except Exception as exc:messagebox.showerror('Download Plat',str(exc),parent=self)

    def open_county(self):
        record=self.selected_record()
        webbrowser.open(record['viewer_url'] if record and record['viewer_url'] else SEARCH_URL)

    def add_files(self):
        paths=filedialog.askopenfilenames(parent=self,title='Add downloaded subdivision plat PDFs',filetypes=[('PDF plats','*.pdf')])
        for path in paths:
            try:
                content=Path(path).read_bytes();pages=validated_pdf(content)
                self.add_item({'record':{'name':Path(path).stem,'file_number':'','map_pages':0},'pdf':content,'pages':pages,'source_url':'','local_source':str(path)})
            except Exception as exc:messagebox.showerror('Add Downloaded Plat',f'{Path(path).name}: {exc}',parent=self)

    def add_item(self,item):
        item['sha256']=hashlib.sha256(item['pdf']).hexdigest()
        if any(old['sha256']==item['sha256'] for old in self.items):
            self.status.set('That exact PDF is already in the downloaded plats list.');return
        item['downloaded_at']=datetime.now().astimezone().isoformat()
        self.items.append(item);r=item['record']
        self.downloaded.insert('end',f"{r['name']} | {r.get('file_number','')} | {item['pages']} pages")
        self.downloaded.selection_set(0,'end');self.current=len(self.items)-1;self.page=0
        expected=r.get('map_pages',0)
        self.status.set(f"Review needed: the County lists {expected} map pages, but this PDF has {item['pages']}. Download any missing pages before attaching." if expected and item['pages']!=expected else 'Plat downloaded. Review the pages, then select the plats to attach.')
        self.update_buttons();self.display()

    def preview_selection(self):
        selected=self.downloaded.curselection()
        if selected:self.current=selected[-1];self.page=0
        self.update_buttons();self.display()

    def update_buttons(self):
        state='normal' if self.downloaded.curselection() else 'disabled'
        self.export_button.configure(state=state);self.use_button.configure(state=state)

    def display(self):
        if self.closed or self.current is None:return
        import fitz
        item=self.items[self.current]
        with fitz.open(stream=item['pdf'],filetype='pdf') as doc:
            self.page=max(0,min(self.page,len(doc)-1));page=doc[self.page]
            width,height=max(100,self.canvas.winfo_width()-20),max(100,self.canvas.winfo_height()-20)
            factor=min(width/page.rect.width,height/page.rect.height,2)
            pix=page.get_pixmap(matrix=fitz.Matrix(factor,factor),alpha=False)
            self._image=ImageTk.PhotoImage(Image.open(io.BytesIO(pix.tobytes('png'))))
            self.canvas.delete('all');self.canvas.create_image(width/2+10,height/2+10,image=self._image,anchor='center')
            self.page_text.set(f'{self.page+1} / {len(doc)}')
        self.preview_title.set(item['record']['name'])

    def turn_page(self,delta):
        if self.current is not None:self.page+=delta;self.display()

    def chosen(self):
        return [self.items[i] for i in self.downloaded.curselection()]

    def export(self):
        selected=self.chosen()
        if not selected:return
        path=filedialog.asksaveasfilename(parent=self,title='Save subdivision plats',initialdir=self.output or None,initialfile='Subdivision_Plats.pdf',defaultextension='.pdf',filetypes=[('PDF','*.pdf')])
        if not path:return
        try:
            import tempfile
            with tempfile.TemporaryDirectory() as folder:
                result=save_plats(selected,folder)
                Path(path).write_bytes(result.read_bytes())
            self.status.set('Plats saved: '+path)
        except Exception as exc:messagebox.showerror('Save Plats',str(exc),parent=self)

    def accept(self):
        try:
            _,_,output=wizard_paths(self.owner)
            if not output:
                output=filedialog.askdirectory(parent=self,title='Choose the case output folder')
            if not output:return
            path=save_plats(self.chosen(),output)
            self.owner._set_downloaded_subdivision_plat(str(path));self.close()
        except Exception as exc:messagebox.showerror('Use Subdivision Plats',str(exc),parent=self)

    def poll(self):
        if self.closed:return
        try:
            while True:
                kind,value=self.events.get_nowait()
                if kind=='context':
                    names=[n['name'] for n in value['names']]
                    self.names.configure(values=names)
                    self.note.set(' '.join(value['warnings']) or 'Subdivision names came from the calculation sheet. Review the recorded plat before using it.')
                    if names:
                        self.names.current(0);self.choose_name();self.search()
                    else:self.status.set('Enter a subdivision name or recording information to search.')
                elif kind=='results':
                    self.busy=False;self.search_button.configure(state='normal')
                    self.records=value['records'];self.results.delete(*self.results.get_children())
                    for i,r in enumerate(self.records):self.results.insert('','end',iid=str(i),values=(r['name'],r['file_date'],r['file_number'],r['volume_page'],r['map_pages'] or ''))
                    self.download_button.configure(state='disabled')
                    self.status.set(f"Showing {len(self.records)} of {value['total']} records. Narrow the search to see the remaining records." if value['truncated'] else f"{len(self.records)} recorded plats found. Choose a record to download.")
                elif kind=='downloaded':self.add_item(value)
                elif kind in ('browser_status','browser_error'):self.status.set(value)
                elif kind=='browser_done':
                    self.browser_busy=False;self.cancel_browser.configure(state='disabled');self.selected_record()
                elif kind=='error':
                    self.busy=False;self.search_button.configure(state='normal');self.note.set(value);self.status.set('You can retry the search or add a downloaded plat.')
        except queue.Empty:pass
        self._timer=self.after(100,self.poll)

    def close(self):
        if self.closed:return
        self.closed=True;self.cancel.set();self.browser.close()
        for timer in (getattr(self,'_timer',None),getattr(self,'_initial',None)):
            if timer:
                try:self.after_cancel(timer)
                except tk.TclError:pass
        self.destroy()


def open_plat_finder(owner):
    prior=getattr(owner,'_plat_finder',None)
    if prior is not None and prior.winfo_exists():prior.lift();return prior
    owner._plat_finder=PlatFinder(owner)
    return owner._plat_finder
