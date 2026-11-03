"""Actual internal lot boundaries from reviewed geometry or an aligned plat."""
import copy,io,json
from pathlib import Path
import tkinter as tk
from tkinter import ttk,filedialog,messagebox
from PIL import Image,ImageTk
from shapely.geometry import shape,mapping
from .lot_divisions import geometry_lots,split_geometry,map_lot_records,read_metadata,put_metadata,write_parts
from .boundary_map import parcel_polygon

def open_lot_map_editor(preview):
    if not preview.data or getattr(preview,'_loading',False):return
    old=getattr(preview,'lot_editor',None)
    if old is not None and old.winfo_exists():old.lift();return
    preview.lot_editor=LotMapEditor(preview)
    return preview.lot_editor

class LotMapEditor(tk.Toplevel):
    def __init__(self,preview):
        super().__init__(preview);self.p=preview;self.title('Platted lot divisions');self.geometry(f'{min(1240,self.winfo_screenwidth()-60)}x{min(850,self.winfo_screenheight()-90)}');self.transient(preview)
        self.matches=preview.data['matched'];self.reference=None;self.reference_path='';self.page=0;self.pdf=None;self.alignment=None;self.pairs=[];self.mode=None;self.first=None;self.polys=[];self.assignments=[];self.history=[];self.reference_zoom=1.;self.reference_center=None;self.reference_drag=None
        bar=ttk.Frame(self,padding=8);bar.pack(fill='x')
        self.account=ttk.Combobox(bar,state='readonly',width=52,values=[m['property']['address']+' | '+m['property']['account'] for m in self.matches]);self.account.pack(side='left');self.account.bind('<<ComboboxSelected>>',self.select_property)
        self.show=tk.BooleanVar(value=preview.settings.get('show_platted_lot_lines',True));ttk.Checkbutton(bar,text='Show platted lot lines on map',variable=self.show,command=self.toggle).pack(side='left',padx=10)
        tools=ttk.Frame(self,padding=8);tools.pack(fill='x')
        for title,cmd in [('Load plat',self.load_plat),('Previous page',lambda:self.change_page(-1)),('Next page',lambda:self.change_page(1)),('Align plat: 3 points',self.start_alignment),('Draw division',self.start_line),('Undo division',self.undo),('Import lot GeoJSON',self.import_geometry)]:ttk.Button(tools,text=title,command=cmd).pack(side='left',padx=2)
        self.status=tk.StringVar(value='Select a property. Split its Step 1 rows before assigning individual lots.');ttk.Label(self,textvariable=self.status,wraplength=1170,padding=8).pack(fill='x')
        panes=ttk.Panedwindow(self,orient='horizontal');panes.pack(fill='both',expand=True,padx=8)
        self.map=tk.Canvas(panes,bg='white',width=600,height=250);self.plat=tk.Canvas(panes,bg='#eeeeee',width=560,height=250);panes.add(self.map,weight=1);panes.add(self.plat,weight=1)
        self.map.bind('<Configure>',lambda e:self.draw());self.plat.bind('<Configure>',lambda e:self.draw_reference());self.map.bind('<Button-1>',self.map_click);self.plat.bind('<Button-1>',self.plat_click)
        self.plat.bind('<MouseWheel>',self.zoom_reference);self.plat.bind('<ButtonPress-3>',lambda e:setattr(self,'reference_drag',(e.x,e.y)))
        self.plat.bind('<B3-Motion>',self.pan_reference)
        ttk.Label(self,text='Plat: use the mouse wheel to zoom and drag with the right mouse button to move. Division labels are for review only.').pack(anchor='w',padx=10)
        lower=ttk.Frame(self,padding=8);lower.pack(fill='x')
        self.piece=ttk.Combobox(lower,state='readonly',width=12);self.piece.pack(side='left');self.piece.bind('<<ComboboxSelected>>',self.piece_selected)
        self.identity=ttk.Combobox(lower,state='readonly',width=65);self.identity.pack(side='left',padx=6)
        ttk.Button(lower,text='Assign lot to division',command=self.assign).pack(side='left')
        self.source=tk.StringVar();ttk.Label(self,text='Plat / geometry source and recording reference').pack(anchor='w',padx=12);ttk.Entry(self,textvariable=self.source).pack(fill='x',padx=12)
        self.reviewed=tk.BooleanVar();ttk.Checkbutton(self,text='I checked every division and its lot identity against the plat or verified geometry.',variable=self.reviewed).pack(anchor='w',padx=10,pady=5)
        bottom=ttk.Frame(self,padding=8);bottom.pack(fill='x')
        for title,cmd in [('Apply to this map',self.apply),('Save divisions in calculation copy',self.save_workbook),('Remove this property’s divisions',self.remove),('Close',self.destroy)]:ttk.Button(bottom,text=title,command=cmd).pack(side='left',padx=4)
        self.account.current(0);self.select_property()

    def select_property(self,event=None):
        match=self.matches[self.account.current()];self.parent_polygon=parcel_polygon(match['feature']);self.prop=match['property'];self.rows=self.prop.get('lot_rows',[])
        self.alignment=None;self.pairs=[];self.mode=None;self.history=[];self.polys=self.parent_parts();self.assignments=[None]*len(self.polys);self.reviewed.set(False)
        self.identity['values']=[r['legal_description'] for r in self.rows]
        record=map_lot_records(self.p.data,self.p.settings).get(self.prop['account'])
        if record:
            try:
                lots=geometry_lots(record,self.parent_polygon,[r['id'] for r in self.rows]);self.polys=[p for _,p in lots];self.assignments=[l['id'] for l,_ in lots];self.source.set(record['source']);self.reviewed.set(bool(lots))
            except Exception as exc:self.status.set(str(exc))
        if not self.polys:self.polys=self.parent_parts();self.assignments=[None]*len(self.polys)
        self.update_pieces();self.draw()

    def parent_parts(self):
        return list(self.parent_polygon.geoms) if self.parent_polygon.geom_type=='MultiPolygon' else [self.parent_polygon]

    def transform(self):
        a,b,c,d=self.parent_polygon.bounds;w=max(self.map.winfo_width(),200);h=max(self.map.winfo_height(),200);s=min((w-90)/max(c-a,1),(h-90)/max(d-b,1))
        return s,(w-s*(a+c))/2,(h+s*(b+d))/2
    def xy(self,q):s,x,y=self.transform();return q[0]*s+x,y-q[1]*s
    def world(self,x,y):s,ox,oy=self.transform();return [(x-ox)/s,(oy-y)/s]

    def draw(self):
        if not hasattr(self,'parent_polygon'):return
        self.map.delete('all')
        if self.reference is not None and self.alignment is not None:
            import numpy as np
            s,x,y=self.transform();page=np.array([[s,0,x],[0,-s,y],[0,0,1.]])
            inverse=np.linalg.inv(page@self.alignment);coeff=tuple(inverse[:2,:].flatten())
            image=self.reference.transform((max(self.map.winfo_width(),200),max(self.map.winfo_height(),200)),Image.Transform.AFFINE,coeff,Image.Resampling.BILINEAR,fillcolor='white').convert('RGBA');image.putalpha(135)
            self.overlay=ImageTk.PhotoImage(image);self.map.create_image(0,0,image=self.overlay,anchor='nw')
        for i,p in enumerate(self.polys):
            self.map.create_polygon(*[v for q in p.exterior.coords for v in self.xy(q)],fill='',outline='#334155',width=2)
            x,y=self.xy(p.representative_point().coords[0]);lot=next((r['legal_description'] for r in self.rows if r['id']==self.assignments[i]),'Unassigned')
            self.map.create_text(x,y,text=f'Division {i+1}\n{lot}',fill='#1d4ed8',font=('Segoe UI',10,'bold'),width=180)
        for i,pair in enumerate(self.pairs):
            x,y=self.xy(pair[0]);self.map.create_oval(x-4,y-4,x+4,y+4,fill='red');self.map.create_text(x+10,y-10,text=str(i+1),fill='red')
        if self.first:
            x,y=self.xy(self.first);self.map.create_oval(x-4,y-4,x+4,y+4,fill='red')

    def load_plat(self):
        path=filedialog.askopenfilename(parent=self,title='Select subdivision plat',filetypes=[('Plat PDF or image','*.pdf *.png *.jpg *.jpeg *.tif *.tiff')])
        if not path:return
        try:
            self.reference_path=path;self.page=0
            if self.pdf:self.pdf.close();self.pdf=None
            if Path(path).suffix.lower()=='.pdf':
                import fitz
                self.pdf=fitz.open(path);self.reference=Image.open(io.BytesIO(self.pdf[0].get_pixmap(dpi=130).tobytes('png'))).convert('RGB')
            else:self.reference=Image.open(path).convert('RGB')
            self.source.set(Path(path).name);self.alignment=None;self.reference_zoom=1.;self.reference_center=None;self.draw_reference();self.draw()
        except Exception as exc:messagebox.showerror('Load plat',str(exc),parent=self)

    def change_page(self,step):
        if not self.pdf:return
        self.page=max(0,min(len(self.pdf)-1,self.page+step));self.reference=Image.open(io.BytesIO(self.pdf[self.page].get_pixmap(dpi=130).tobytes('png'))).convert('RGB');self.alignment=None;self.reference_center=None;self.reference_zoom=1.;self.source.set(Path(self.reference_path).name+f', page {self.page+1}');self.draw_reference();self.draw()

    def draw_reference(self):
        self.plat.delete('all')
        if self.reference is None:return
        w,h=self.reference.size;cw,ch=max(self.plat.winfo_width(),200),max(self.plat.winfo_height(),200);s=min((cw-12)/w,(ch-12)/h)*self.reference_zoom
        center=self.reference_center or (w/2,h/2);self.reference_scale=s;self.reference_offset=(cw/2-center[0]*s,ch/2-center[1]*s)
        ox,oy=self.reference_offset
        view=self.reference.transform((cw,ch),Image.Transform.AFFINE,(1/s,0,-ox/s,0,1/s,-oy/s),Image.Resampling.BILINEAR,fillcolor='white')
        self.reference_tk=ImageTk.PhotoImage(view);self.plat.create_image(0,0,image=self.reference_tk,anchor='nw')

    def zoom_reference(self,event):
        if self.reference is None:return
        ox,oy=self.reference_offset;px,py=(event.x-ox)/self.reference_scale,(event.y-oy)/self.reference_scale
        self.reference_zoom=max(.5,min(25,self.reference_zoom*(1.3 if event.delta>0 else 1/1.3)))
        w,h=self.reference.size;cw,ch=max(self.plat.winfo_width(),200),max(self.plat.winfo_height(),200);s=min((cw-12)/w,(ch-12)/h)*self.reference_zoom
        self.reference_center=(px-(event.x-cw/2)/s,py-(event.y-ch/2)/s);self.draw_reference()
    def pan_reference(self,event):
        if self.reference is None or self.reference_drag is None:return
        dx,dy=event.x-self.reference_drag[0],event.y-self.reference_drag[1];center=self.reference_center or (self.reference.width/2,self.reference.height/2)
        self.reference_center=(center[0]-dx/self.reference_scale,center[1]-dy/self.reference_scale);self.reference_drag=(event.x,event.y);self.draw_reference()

    def start_alignment(self):
        if self.reference is None:messagebox.showinfo('Align plat','Load the plat first.',parent=self);return
        self.mode='align_map';self.pairs=[];self.first=None;self.alignment=None;self.reviewed.set(False);self.status.set('Point 1: click a recognizable corner on the parcel map, then the same corner on the plat. Use three non-collinear points.');self.draw()
    def start_line(self):
        self.mode='line';self.first=None;self.status.set('Click two endpoints spanning the parcel. Place both endpoints at or beyond its edges. Check the line against the plat.');self.reviewed.set(False)

    def map_click(self,event):
        point=self.world(event.x,event.y)
        if self.mode=='align_map':self.first=point;self.mode='align_plat';self.status.set('Click the same corner on the plat at right.');self.draw()
        elif self.mode=='line':
            if self.first is None:self.first=point;self.draw();return
            points=[self.first,point];self.first=None
            # Extend a small amount to accommodate clicking on the border stroke.
            dx,dy=point[0]-points[0][0],point[1]-points[0][1]
            points=[[points[0][0]-.03*dx,points[0][1]-.03*dy],[point[0]+.03*dx,point[1]+.03*dy]]
            try:
                new=split_geometry(self.polys,points);self.history.append((self.polys,self.assignments));self.polys=new;self.assignments=[None]*len(new);self.update_pieces();self.draw()
            except Exception as exc:messagebox.showwarning('Division unchanged',str(exc),parent=self)

    def plat_click(self,event):
        if self.mode!='align_plat':return
        ox,oy=self.reference_offset;pixel=[(event.x-ox)/self.reference_scale,(event.y-oy)/self.reference_scale]
        self.pairs.append((self.first,pixel));self.first=None
        if len(self.pairs)<3:self.mode='align_map';self.status.set(f'Point {len(self.pairs)+1}: click the next corner on the parcel map.');self.draw();return
        import numpy as np
        try:
            pixel=np.array([[*p,1.] for _,p in self.pairs]);world=np.array([w for w,_ in self.pairs])
            if abs(np.linalg.det(pixel))<1:raise ValueError('Choose three points that do not lie on a single line.')
            coeff=np.linalg.solve(pixel,world);self.alignment=np.vstack((coeff.T,[0,0,1.]));self.mode=None
            self.status.set('Plat aligned. Check other corners for agreement, then draw the internal divisions and assign each to its Step 1 lot.');self.draw()
        except Exception as exc:self.mode=None;messagebox.showwarning('Alignment',str(exc),parent=self)

    def update_pieces(self):self.piece['values']=[f'Division {i+1}' for i in range(len(self.polys))];self.piece.current(0);self.piece_selected()
    def piece_selected(self,event=None):
        i=self.piece.current();key=self.assignments[i] if i>=0 else None
        self.identity.set(next((r['legal_description'] for r in self.rows if r['id']==key),''))
    def assign(self):
        i,j=self.piece.current(),self.identity.current()
        if i<0 or j<0:return
        self.assignments[i]=self.rows[j]['id'];self.reviewed.set(False);self.draw()
    def undo(self):
        if self.history:self.polys,self.assignments=self.history.pop();self.reviewed.set(False);self.update_pieces();self.draw()

    def import_geometry(self):
        path=filedialog.askopenfilename(parent=self,title='Verified lot polygons in EPSG:2278',filetypes=[('GeoJSON','*.geojson *.json')])
        if not path:return
        try:
            content=json.loads(Path(path).read_text(encoding='utf8'));crs=content.get('crs',{}).get('properties',{}).get('name','')
            if not str(crs).endswith(('2278','102740')):raise ValueError('The GeoJSON must declare EPSG:2278. Longitude/latitude and unknown coordinates cannot be used as survey feet.')
            polys=[shape(f['geometry']) for f in content['features']];ids=[]
            if not polys or any(p.geom_type!='Polygon' or p.is_empty or not p.is_valid for p in polys):
                raise ValueError('Import one valid, nonempty polygon per lot.')
            for feature in content['features']:
                desc=feature.get('properties',{}).get('legal_description','');ids.append(next((r['id'] for r in self.rows if r['legal_description'].upper()==desc.upper()),None))
            self.history.append((self.polys,self.assignments));self.polys=polys;self.assignments=ids;self.source.set(Path(path).name);self.reviewed.set(False);self.update_pieces();self.draw()
            self.status.set('Imported geometry. Assign any unmatched lot identities, verify the complete parcel coverage and check the review box.')
        except Exception as exc:messagebox.showerror('Import geometry',str(exc),parent=self)

    def record(self):
        if not self.reviewed.get():raise ValueError('Check each lot identity and division against the source, then check the review box.')
        if len(self.rows)<2:raise ValueError('This property needs separate Step 1 lot rows first. Use Review Individual Lots on the wizard’s setup page.')
        record={'crs':'EPSG:2278','reviewed':True,'source':self.source.get().strip(),'lots':[{'id':key,'geometry':mapping(p)} for key,p in zip(self.assignments,self.polys)]}
        if self.reference_path:record['plat_file']=Path(self.reference_path).name;record['plat_page']=self.page+1;record['alignment_points']=self.pairs
        if not record['source']:raise ValueError('Enter the plat or geometry source reference.')
        geometry_lots(record,self.parent_polygon,[r['id'] for r in self.rows]);return record

    def refresh_map(self):
        self.p.settings['show_platted_lot_lines']=self.show.get()
        if getattr(self.p,'is_tax',False):self.p.changed();self.p.start_load()
        else:self.p.changed();self.p.redraw()
    def toggle(self):self.refresh_map()
    def apply(self):
        try:record=self.record()
        except Exception as exc:messagebox.showwarning('Review lot divisions',str(exc),parent=self);return False
        self.p.settings.setdefault('lot_divisions',{})[self.prop['account']]=record;self.refresh_map();self.status.set('Applied. Use This Map or Download Map saves these divisions with the map. Save a calculation copy to reuse them in other map types.');return True
    def remove(self):
        self.p.settings.setdefault('lot_divisions',{})[self.prop['account']]={};self.polys=self.parent_parts();self.assignments=[None]*len(self.polys);self.reviewed.set(False);self.update_pieces();self.draw();self.refresh_map()
    def save_workbook(self):
        if not self.apply():return
        path=filedialog.asksaveasfilename(parent=self,title='Save calculation with reviewed lot divisions',initialdir=str(Path(self.p.source).parent),initialfile=Path(self.p.source).stem+'_LotMap.xlsx',defaultextension='.xlsx',filetypes=[('Excel','*.xlsx')])
        if not path:return
        try:
            import zipfile
            from .workbooks import read_shared_bytes
            from .lot_review_ui import select_calculation_copy
            from openpyxl import load_workbook
            from .boundary_map import fingerprint
            if fingerprint(self.p.source)!=self.p.data['case']['workbook_hash']:raise ValueError('The source sheet changed. Reload the map before saving these divisions.')
            raw=read_shared_bytes(self.p.source)
            with zipfile.ZipFile(io.BytesIO(raw)) as z:parts={n:z.read(n) for n in z.namelist()}
            book=load_workbook(io.BytesIO(raw),read_only=True,data_only=True);meta=read_metadata(book);book.close()
            meta.setdefault('groups',{}).update(self.p.settings.get('lot_divisions',{}));put_metadata(parts,meta);write_parts(parts,path)
            select_calculation_copy(self.p.owner,path)
            messagebox.showinfo('Lot divisions saved','Saved and selected the new calculation sheet. Future boundary, PC and tax maps can reuse its reviewed lot divisions.',parent=self)
        except Exception as exc:messagebox.showerror('Calculation not saved',str(exc),parent=self)
    def destroy(self):
        if self.pdf:self.pdf.close();self.pdf=None
        super().destroy()
