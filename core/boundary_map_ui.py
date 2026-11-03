"""Shared Tk map creation, preview, placement controls and attachment selection."""
from __future__ import annotations

import io
import copy
import gc
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

from PIL import Image, ImageTk

from .boundary_map import (MapError, PublicMapClient, assert_map_current, fetch_map_data,
                           fingerprint, read_map_case)
from .boundary_map_render import default_settings, render_map, pdf_png, save_map, export_map
from .workbooks import WorkbookReadCache
from .map_text import saved_layout
from .street_label_ui import StreetAppearanceEditor


def wizard_paths(owner):
    if hasattr(owner, 'vars'):
        return (owner.vars['excel'].get(), owner.vars['map'].get(), owner.vars['output_folder'].get())
    return (owner.files.get('workbook', ''), owner.files.get('map', ''), owner.files.get('out', ''))


class MapControls(tk.Frame):
    def __init__(self, parent, owner, row):
        super().__init__(parent, bg='white')
        self.owner = owner
        self.grid(row=row, column=0, columnspan=3, sticky='ew', padx=12, pady=(2, 8))
        self.columnconfigure(1, weight=1)
        tk.Button(self, text='Create Boundary Map', command=owner._create_boundary_map,
                  bg='#1D4ED8', fg='white', activebackground='#1E40AF', activeforeground='white',
                  font=('Segoe UI', 10, 'bold'), padx=12, pady=8, relief='flat').grid(row=0, column=0, sticky='w')
        tk.Button(self, text='Create Tax Map', command=owner._create_tax_map,
                  bg='#1D4ED8', fg='white', activebackground='#1E40AF', activeforeground='white',
                  font=('Segoe UI', 10, 'bold'), padx=12, pady=8, relief='flat').grid(row=1, column=0, sticky='w', pady=(8, 0))
        tk.Label(self, text='HCAD Parcel Viewer with a yellow application boundary.', bg='white', fg='#475569',
                 font=('Segoe UI', 9), anchor='w', wraplength=500).grid(row=1, column=1, sticky='w', padx=(12, 0), pady=(8, 0))
        tk.Button(self, text='Find Subdivision Plats', command=owner._find_subdivision_plats,
                  bg='#1D4ED8', fg='white', activebackground='#1E40AF', activeforeground='white',
                  font=('Segoe UI', 10, 'bold'), padx=12, pady=8, relief='flat').grid(row=2, column=0, sticky='w', pady=(8, 0))
        tk.Label(self, text='Search recorded plats, download, review and attach.', bg='white', fg='#475569',
                 font=('Segoe UI', 9), anchor='w', wraplength=500).grid(row=2, column=1, sticky='w', padx=(12, 0), pady=(8, 0))
        tk.Label(self,text='Map creation fills the council district and councilmember in Case Info.',
                 bg='white',fg='#475569',font=('Segoe UI',9),wraplength=650,
                 justify='left').grid(row=3,column=0,columnspan=2,sticky='w',pady=(8,0))
        self.status = tk.Label(self, text='', bg='white', fg='#475569', font=('Segoe UI', 9),
                               anchor='w', justify='left', wraplength=500)
        self.status.grid(row=0, column=1, sticky='w', padx=(12, 0))
        self._stamp = None
        self._timer = None
        self.bind('<Destroy>', self._destroyed, add='+')
        self.refresh()

    def _destroyed(self, event):
        if event.widget == self and self._timer:
            self.after_cancel(self._timer)
            self._timer = None

    def refresh(self):
        if not self.winfo_exists():
            return
        source, selected, _ = wizard_paths(self.owner)
        def stamp(p):
            try:
                s = Path(p).stat()
                return (p, s.st_mtime_ns, s.st_size)
            except OSError:
                return (p, None)
        updates=getattr(self.owner,'_calculation_update_notes',[])
        current = (stamp(source), stamp(selected),tuple(updates))
        if current != self._stamp:
            self._stamp = current
            try:
                warning = assert_map_current(selected, source, self.owner._boundary_map_wizard_key)
                self.status.configure(
                    text=(warning or 'Create a City-format map from the saved calculation sheet.')+('\n'+'\n'.join(updates) if updates else ''),
                    fg='#B45309' if warning else '#475569')
            except Exception as e:
                self.status.configure(text=str(e), fg='#B45309')
        self._timer = self.after(2000, self.refresh)


class MapPreview(tk.Toplevel):
    page_size = (612, 792)
    map_kind = 'Boundary Map'

    def render_preview(self, data, settings):
        return render_map(data, settings)

    def default_layout(self, data):
        return default_settings(data)

    def save_preview(self, data, settings, rendered, output):
        return save_map(data, settings, rendered, output)

    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner
        self.source, self.selected_map, self.output = wizard_paths(owner)
        self.wizard = owner._boundary_map_wizard_key
        self.title('Create ' + self.map_kind)
        width = min(1100, self.winfo_screenwidth() - 60)
        height = min(900, self.winfo_screenheight() - 80)
        self.geometry(f'{width}x{height}')
        self.minsize(min(880, width), min(660, height))
        self.transient(owner.winfo_toplevel())
        self.configure(bg='#F8FAFC')
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self.cancel = threading.Event()
        self.events = queue.Queue()
        self.data = None
        self.rendered = None
        self.settings = None
        self._layout_case = None
        self.drag = None
        self._closed = False
        self._loading = False
        self._use_cache = True
        self._image = None
        self._render_timer = None
        self._load_timer = None
        self.protocol('WM_DELETE_WINDOW', self.close)
        controls = tk.Frame(self, bg='white', padx=14, pady=14, width=315)
        controls.grid(row=0, column=0, sticky='ns')
        controls.grid_propagate(False)
        controls.columnconfigure(0, weight=1)
        tk.Label(controls, text=self.map_kind + ' preview', bg='white', fg='#0F172A',
                 font=('Segoe UI', 15, 'bold')).grid(row=0, column=0, sticky='w', pady=(0, 10))
        self.summary = tk.Label(controls, text='Reading the saved calculation sheet...', bg='white',
                                fg='#334155', wraplength=285, justify='left', anchor='w', font=('Segoe UI', 10))
        self.summary.grid(row=1, column=0, sticky='ew', pady=(0, 12))
        tk.Label(controls, text='Orientation', bg='white', font=('Segoe UI', 10, 'bold')).grid(row=2, column=0, sticky='w')
        row = tk.Frame(controls, bg='white')
        row.grid(row=3, column=0, sticky='ew', pady=(4, 10))
        self.rotation = tk.DoubleVar(value=0)
        spin = ttk.Spinbox(row, from_=-180, to=180, increment=5, textvariable=self.rotation, width=8,
                           command=self.changed)
        spin.pack(side='left')
        spin.bind('<Return>', lambda e: self.changed())
        spin.bind('<FocusOut>', lambda e: self.changed())
        ttk.Button(row, text='North up', command=lambda: self.set_rotation(0)).pack(side='left', padx=6)
        self.zoom = tk.DoubleVar(value=1)
        tk.Label(controls, text='Framing: more area / closer', bg='white', font=('Segoe UI', 10, 'bold')).grid(row=4, column=0, sticky='w')
        ttk.Scale(controls, from_=.6, to=1.5, variable=self.zoom, command=lambda x: self.changed()).grid(row=5, column=0, sticky='ew')
        pan = tk.Frame(controls, bg='white')
        pan.grid(row=6, column=0, sticky='w', pady=(0, 8))
        for caption, dx, dy in [('Left', -20, 0), ('Right', 20, 0), ('Up', 0, 20), ('Down', 0, -20)]:
            ttk.Button(pan, text=caption, width=6, command=lambda x=dx, y=dy: self.pan(x, y)).pack(side='left', padx=1)
        self.include_street = tk.BooleanVar(value=False)
        tk.Checkbutton(controls, text='Enclose the street between parcel rows', variable=self.include_street,
                       command=self.changed, bg='white', anchor='w', wraplength=280).grid(row=7, column=0, sticky='w')
        self.vertical_addresses = tk.BooleanVar(value=False)
        tk.Checkbutton(controls, text='Rotate address labels 90 degrees', variable=self.vertical_addresses,
                       command=self.changed, bg='white', anchor='w').grid(row=8, column=0, sticky='w')
        self.interior_lot_lines = tk.BooleanVar(value=True)
        tk.Checkbutton(controls, text='Show tax-parcel boundaries', variable=self.interior_lot_lines,
                       command=self.changed, bg='white', anchor='w').grid(row=9, column=0, sticky='w')
        actions = tk.Frame(controls, bg='white')
        actions.grid(row=10, column=0, sticky='w', pady=8)
        ttk.Button(actions, text='Reset layout', command=self.reset).pack(side='left')
        self.refresh_button = ttk.Button(actions, text='Refresh parcel data', command=self.refresh_data, state='disabled')
        self.refresh_button.pack(side='left', padx=6)
        ttk.Button(controls, text='Reset map text', command=self.reset_text).grid(row=11, column=0, sticky='w')
        from .lot_map_ui import open_lot_map_editor
        ttk.Button(controls,text='Platted lots...',command=lambda:open_lot_map_editor(self)).grid(row=13,column=0,sticky='w')
        self.street_appearance_button = ttk.Button(controls, text='Street name size / rotation...', command=self.open_street_appearance)
        self.street_appearance_button.grid(row=12, column=0, sticky='w', pady=(4, 0))
        tk.Label(controls, text='Click an address, street name, or legend line to edit its text. Drag addresses or street names to move them.',
                 wraplength=280, bg='white', fg='#475569', justify='left', font=('Segoe UI', 9)).grid(row=14, column=0, sticky='w')
        self.review = tk.Text(controls, width=35, height=12, wrap='word', font=('Segoe UI', 9),
                              bg='#F8FAFC', relief='solid', bd=1)
        self.review.grid(row=15, column=0, sticky='nsew', pady=(12, 0))
        controls.rowconfigure(15, weight=1)
        self.canvas = tk.Canvas(self, bg='#E2E8F0', highlightthickness=0)
        self.canvas.grid(row=0, column=1, sticky='nsew', padx=12, pady=12)
        self.canvas.bind('<Configure>', lambda e: self.display())
        self.canvas.bind('<ButtonPress-1>', self.drag_start)
        self.canvas.bind('<ButtonRelease-1>', self.drag_end)
        self.canvas.bind('<Button-3>', self.street_context)
        bottom = tk.Frame(self, bg='white', padx=14, pady=10)
        bottom.grid(row=1, column=0, columnspan=2, sticky='ew')
        bottom.columnconfigure(0, weight=1)
        self.status = tk.StringVar(value='Loading properties...')
        tk.Label(bottom, textvariable=self.status, bg='white', fg='#334155', anchor='w',
                 wraplength=400).grid(row=0, column=0, sticky='w')
        self.use = ttk.Button(bottom, text='Use This Map', command=self.accept, state='disabled')
        self.download = ttk.Button(bottom, text='Download Map', command=self.download_map, state='disabled')
        self.download.grid(row=0, column=1, padx=(8, 0))
        self.use.grid(row=0, column=2, padx=10)
        ttk.Button(bottom, text='Close', command=self.close).grid(row=0, column=3)
        self.after(50, self.start_load)

    def start_load(self):
        if self._closed or self._loading:
            return
        # Tk objects from a previously closed wizard can form Python cycles.
        # Finalize them here on Tk's thread, before workbook parsing allocates
        # enough objects to trigger garbage collection on the worker thread.
        gc.collect()
        self._loading = True
        self.refresh_button.configure(state='disabled')
        source, output, wizard = self.source, self.output, self.wizard
        use_cache = self._use_cache
        def work():
            try:
                from .automatic_case import load_map_data
                data=load_map_data(source,wizard,output,
                                   progress=lambda s:self.events.put(('status',s)),
                                   cancel=self.cancel,use_cache=use_cache)
                self.events.put(('ready', data))
            except Exception as e:
                self.events.put(('error', str(e)))
        threading.Thread(target=work, daemon=True).start()
        self.poll()

    def refresh_data(self):
        if self._loading:
            return
        if self._render_timer:
            self.after_cancel(self._render_timer)
            self._render_timer = None
        self._use_cache = False
        self.data = None
        self.rendered = None
        self.use.configure(state='disabled')
        self.download.configure(state='disabled')
        self.status.set('Refreshing public parcel and street data...')
        self.start_load()

    def poll(self):
        if self._closed:
            return
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == 'status':
                    self.status.set(value)
                elif kind == 'error':
                    self._loading = False
                    self.refresh_button.configure(state='normal')
                    self.status.set('Map not created. Correct the issue below and try again.')
                    self.show_review(value)
                    self.summary.configure(text='The map could not be loaded. Review the message below, then use Refresh parcel data to retry.')
                    return
                elif kind == 'ready':
                    self._loading = False
                    self.refresh_button.configure(state='normal')
                    self.data = value
                    from .automatic_case import select_updated_calculation,council_summary
                    updated=value.get('calculation_path',self.source)
                    select_updated_calculation(self.owner,updated,self.source)
                    self.source=updated
                    if hasattr(self.owner,'_map_calculation_ready'):self.owner._map_calculation_ready(value)
                    c = value['case']
                    self.summary.configure(text=f"{c['reference']}\n{c['title']}\n{c['block']}\n{c['sides']}\n{c['minimum']}\n\n{len(value['matched'])} matched properties"+council_summary(value))
                    identity = (c['workbook_hash'], c['wizard'])
                    if self._layout_case != identity or self.settings is None:
                        self.settings = saved_layout(self.selected_map, value, self.default_layout(value))
                    self._layout_case = identity
                    self.sync_controls()
                    self.changed()
                    return
        except queue.Empty:
            pass
        self._load_timer = self.after(100, self.poll)

    def show_review(self, message):
        self.review.configure(state='normal')
        self.review.delete('1.0', 'end')
        self.review.insert('1.0', message)
        self.review.configure(state='disabled')

    def reset(self):
        if self.data is None:
            return
        overrides = copy.deepcopy((self.settings or {}).get('text_overrides', {}))
        self.settings = self.default_layout(self.data)
        self.settings['text_overrides'] = overrides
        self.sync_controls()
        self.changed()

    def sync_controls(self):
        self.rotation.set(self.settings['rotation'])
        self.zoom.set(self.settings['zoom'])
        self.include_street.set(self.settings['include_street'])
        self.interior_lot_lines.set(self.settings['show_interior_lot_lines'])
        self.vertical_addresses.set(bool(self.settings.get('address_rotation', 0)))

    def set_rotation(self, value):
        self.rotation.set(value)
        self.changed()

    def pan(self, dx, dy):
        if self.settings is not None:
            self.settings['pan_x'] += dx
            self.settings['pan_y'] += dy
            self.changed()

    def changed(self):
        if self.data is None or self.settings is None:
            return
        self.use.configure(state='disabled')
        self.download.configure(state='disabled')
        if self._render_timer:
            self.after_cancel(self._render_timer)
        self._render_timer = self.after(150, self.redraw)

    def redraw(self):
        if self._render_timer:
            self.after_cancel(self._render_timer)
        self._render_timer = None
        try:
            self.settings.update(rotation=float(self.rotation.get()), zoom=float(self.zoom.get()),
                                 include_street=self.include_street.get(),
                                 show_interior_lot_lines=self.interior_lot_lines.get(),
                                 address_rotation=90 if self.vertical_addresses.get() else 0)
            self.rendered = self.render_preview(self.data, self.settings)
            for key in self.rendered['rejected_offsets']:
                self.settings['label_offsets'].pop(key, None)
            self._image = Image.open(io.BytesIO(pdf_png(self.rendered['pdf']))).copy()
            self.display()
            warnings = self.rendered.get('issues', []) + self.rendered.get('blocking', [])
            notes = '\n'.join(warnings)
            if not notes:
                notes = 'All properties matched. Review the boundary and labels, then click Use This Map.'
            updates=self.data.get('calculation_updates',[])
            if updates:notes+='\n\n'+'\n'.join(updates)
            sources = self.data['sources']
            dates = sorted({s['retrieved_utc'][:10] for s in sources})
            notes += '\n\nParcel/street data retrieved: ' + ', '.join(dates)
            if any(s.get('cached') for s in sources):
                notes += '\nSaved public data reused from this case folder.'
            self.show_review(notes)
            self.status.set('Map created with warnings. Review the notes; Download Map and Use This Map are available.' if warnings else
                            'Ready to download or use this map.')
            self.use.configure(state='normal')
            self.download.configure(state='normal')
            editor = getattr(self, '_street_editor', None)
            if editor is not None and editor.winfo_exists():
                editor.select_key(editor.selected_key())
        except Exception as e:
            self.rendered = None
            self.status.set('Could not draw this layout.')
            self.show_review(str(e))
            self.use.configure(state='disabled')
            self.download.configure(state='disabled')

    def display(self):
        if self._image is None or self._closed:
            return
        w, h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if w < 10 or h < 10:
            return
        factor = min((w - 16) / self._image.width, (h - 16) / self._image.height)
        size = (max(1, int(self._image.width * factor)), max(1, int(self._image.height * factor)))
        self.photo = ImageTk.PhotoImage(self._image.resize(size, Image.Resampling.LANCZOS), master=self)
        self.image_x, self.image_y = (w - size[0]) / 2, (h - size[1]) / 2
        self.page_scale = size[0] / self.page_size[0]
        self.canvas.delete('all')
        self.canvas.create_image(self.image_x, self.image_y, image=self.photo, anchor='nw')

    def page_point(self, event):
        return ((event.x - self.image_x) / self.page_scale,
                self.page_size[1] - (event.y - self.image_y) / self.page_scale)

    def label_at(self, x, y):
        from shapely.geometry import Polygon, Point, box
        for item in reversed(self.rendered.get('editable_labels', self.rendered['labels'])):
            area = Polygon(item['corners']) if 'corners' in item else box(*item['bounds'])
            if area.buffer(3).covers(Point(x,y)):
                return item
        return None

    def drag_start(self, event):
        if not self.rendered or not hasattr(self, 'page_scale') or self._render_timer:
            return
        x, y = self.page_point(event)
        item = self.label_at(x,y)
        if item:
            self.drag = (item['key'], x, y)
            self.canvas.configure(cursor='hand2' if item['key'].startswith('legend:') else 'fleur')

    def drag_end(self, event):
        if self.drag:
            key, x, y = self.drag
            ex, ey = self.page_point(event)
            self.drag = None
            self.canvas.configure(cursor='')
            if abs(ex-x)+abs(ey-y) < .75:
                self.edit_text(key)
            elif not key.startswith('legend:'):
                old = self.settings['label_offsets'].get(key, [0, 0])
                self.settings['label_offsets'][key] = [old[0]+ex-x, old[1]+ey-y]
                self.changed()

    def street_context(self, event):
        if self.rendered and not self._render_timer:
            item = self.label_at(*self.page_point(event))
            if item and item['key'].startswith('street:'):
                self.open_street_appearance(item['key'])

    def open_street_appearance(self, key=None):
        if not self.rendered or self._loading or self._render_timer:
            return
        prior = getattr(self, '_street_editor', None)
        if prior is not None and prior.winfo_exists():
            prior.select_key(key)
            prior.lift()
        else:
            self._street_editor = StreetAppearanceEditor(self, key)

    def apply_street_style(self, key, size=None, angle=None):
        if not self.rendered or not any(i['key'] == key for i in self.rendered['labels'] if i['key'].startswith('street:')):
            return False
        candidate = copy.deepcopy(self.settings)
        styles = candidate.setdefault('street_styles', {})
        if size is None:
            styles.pop(key, None)
        else:
            styles[key] = {'size': size, 'angle': angle}
        try:
            self.render_preview(self.data, candidate)
        except Exception as error:
            messagebox.showwarning('Street name was not changed', str(error), parent=self)
            return False
        self.settings = candidate
        self.changed()
        self.redraw()
        return True

    def edit_text(self, key):
        item = next((i for i in self.rendered['editable_labels'] if i['key'] == key), None)
        if item is None:
            return
        value = simpledialog.askstring('Edit map text',
            'Enter the wording for this map. City fonts and spacing are preserved.',
            initialvalue=item['text'], parent=self)
        if value is None:
            return
        candidate = copy.deepcopy(self.settings)
        candidate.setdefault('text_overrides', {})[key] = value
        try:
            self.render_preview(self.data, candidate)
        except Exception as error:
            messagebox.showwarning('Map text was not changed', str(error), parent=self)
            return
        self.settings = candidate
        self.changed()
        self.redraw()

    def reset_text(self):
        if self.settings is not None:
            self.settings['text_overrides'] = {}
            self.changed()

    def snapshot_warnings(self):
        """Retain the displayed case when its workbook changes during review."""
        notes = []
        try:
            if fingerprint(self.source) != self.data['case']['workbook_hash']:
                notes.append('The calculation sheet changed after this preview. The saved map keeps the displayed case and property data.')
        except OSError:
            notes.append('The calculation sheet could not be rechecked. The saved map keeps the displayed case and property data.')
        current_source, _, _ = wizard_paths(self.owner)
        if Path(current_source).resolve() != Path(self.source).resolve():
            notes.append('A different calculation sheet is now selected. The saved map keeps the case shown in this preview.')
        for note in notes:
            if note not in self.rendered['issues']:
                self.rendered['issues'].append(note)
        return notes

    def download_map(self):
        if not self.rendered or self._loading or self._render_timer:
            return
        destination = filedialog.asksaveasfilename(
            parent=self, title='Download ' + self.map_kind, initialdir=self.output,
            initialfile=self.data['case']['reference'].replace(' ', '_') + '_' + self.map_kind.replace(' ', '_') + '.pdf',
            defaultextension='.pdf', filetypes=[('PDF map', '*.pdf'), ('PNG image', '*.png')])
        if not destination:
            return
        try:
            notes = self.snapshot_warnings()
            path = export_map(self.rendered, destination)
            if notes:
                self.show_review(self.review.get('1.0', 'end').strip() + '\n' + '\n'.join(notes))
            self.status.set(f'Map downloaded to {path}')
        except Exception as e:
            messagebox.showerror('Download map', str(e), parent=self)

    def accept(self):
        if not self.rendered or self._loading or self._render_timer:
            return
        try:
            self.snapshot_warnings()
            _, _, current_output = wizard_paths(self.owner)
            self.use.configure(state='disabled')
            self.download.configure(state='disabled')
            path = self.save_preview(self.data, self.settings, self.rendered, current_output or self.output)
            self.owner._set_generated_boundary_map(str(path))
            self.close()
        except Exception as e:
            messagebox.showerror('Boundary map', str(e), parent=self)
            self.use.configure(state='normal')
            self.download.configure(state='normal')

    def close(self):
        if self._closed:
            return
        self._closed = True
        self.cancel.set()
        for timer in (self._render_timer, self._load_timer):
            if timer:
                self.after_cancel(timer)
        self.destroy()


def open_map_preview(owner):
    source, _, output = wizard_paths(owner)
    if not source or not Path(source).is_file():
        raise MapError('Select and save the calculation workbook before creating a boundary map.')
    if not output:
        raise MapError('Select an output folder before creating a boundary map.')
    prior = getattr(owner, '_boundary_map_preview', None)
    if prior is not None and prior.winfo_exists():
        prior.lift()
        return prior
    preview = MapPreview(owner)
    owner._boundary_map_preview = preview
    return preview
