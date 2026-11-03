"""Shared HCAD tax-map preview and attachment controls for all four wizards."""
from __future__ import annotations

import gc
import io
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

from .boundary_map import MapError, PublicMapClient, fetch_map_data, fingerprint, read_map_case
from .boundary_map_ui import wizard_paths
from .boundary_map_render import export_map, pdf_png
from .tax_map import HCADPrintClient, default_settings, map_extent, render_tax_map, save_tax_map
from .workbooks import WorkbookReadCache


class TaxMapPreview(tk.Toplevel):
    is_tax=True
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner
        self.source, _, self.output = wizard_paths(owner)
        self.wizard = owner._boundary_map_wizard_key
        self.title('Create HCAD Tax Map')
        self.geometry(f'{min(1280, self.winfo_screenwidth()-60)}x{min(840, self.winfo_screenheight()-80)}')
        self.transient(owner.winfo_toplevel())
        self.configure(bg='#F8FAFC')
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self.data = self.rendered = self._image = None
        self.settings = default_settings()
        self.rendered_settings = None
        self.events = queue.Queue()
        self.cancel = threading.Event()
        self._closed = self._loading = False
        self._poll_timer = None
        self._controls = []
        self.protocol('WM_DELETE_WINDOW', self.close)

        left = tk.Frame(self, bg='white', width=300, padx=14, pady=14)
        left.grid(row=0, column=0, sticky='ns')
        left.grid_propagate(False)
        left.columnconfigure(0, weight=1)
        left.rowconfigure(8, weight=1)
        tk.Label(left, text='HCAD tax map', font=('Segoe UI', 15, 'bold'), bg='white').grid(row=0, sticky='w')
        self.summary = tk.Label(left, text='Reading the calculation sheet...', wraplength=270, justify='left', bg='white', fg='#334155')
        self.summary.grid(row=1, sticky='ew', pady=12)
        tk.Label(left, text='North up - yellow application boundary', wraplength=270, justify='left', bg='white').grid(row=2, sticky='w')
        tk.Label(left, text='Framing', font=('Segoe UI', 10, 'bold'), bg='white').grid(row=3, sticky='w', pady=(16, 4))
        framing = tk.Frame(left, bg='white')
        framing.grid(row=4, sticky='w')
        self._button(framing, 'More area', lambda: self.zoom(-.15)).pack(side='left')
        self._button(framing, 'Closer', lambda: self.zoom(.15)).pack(side='left', padx=6)
        pan = tk.Frame(left, bg='white')
        pan.grid(row=5, sticky='w', pady=8)
        for caption, dx, dy in [('Left', -1, 0), ('Right', 1, 0), ('Up', 0, 1), ('Down', 0, -1)]:
            self._button(pan, caption, lambda x=dx, y=dy: self.pan(x, y), width=6).pack(side='left', padx=1)
        tk.Label(left, text='One continuous boundary around all properties, including streets and gaps.',
                 wraplength=270, justify='left', bg='white').grid(row=6, sticky='w', pady=(0, 8))
        actions = tk.Frame(left, bg='white')
        actions.grid(row=7, sticky='w')
        self.update_button = self._button(actions, 'Update Preview', self.start_load)
        self.update_button.grid(row=0, column=0, padx=(0, 4), pady=4)
        self._button(actions, 'Reset framing', self.reset).grid(row=0, column=1, pady=4)
        self.refresh_button = self._button(actions, 'Refresh HCAD data', lambda: self.start_load(refresh=True))
        self.refresh_button.grid(row=1, column=0, columnspan=2, sticky='w', pady=4)
        from .lot_map_ui import open_lot_map_editor
        self._button(actions,'Platted lot divisions...',lambda:open_lot_map_editor(self)).grid(row=3,column=0,columnspan=2,sticky='w',pady=4)
        self.review = tk.Text(left, width=32, height=12, wrap='word', font=('Segoe UI', 9), bg='#F8FAFC', relief='solid', bd=1)
        self.review.grid(row=8, sticky='nsew', pady=(12, 0))
        self.canvas = tk.Canvas(self, bg='#E2E8F0', highlightthickness=0)
        self.canvas.grid(row=0, column=1, sticky='nsew', padx=12, pady=12)
        self.canvas.bind('<Configure>', lambda e: self.display())
        bottom = tk.Frame(self, bg='white', padx=14, pady=10)
        bottom.grid(row=1, column=0, columnspan=2, sticky='ew')
        bottom.columnconfigure(0, weight=1)
        self.status = tk.StringVar(value='Loading properties...')
        tk.Label(bottom, textvariable=self.status, bg='white', wraplength=440, justify='left').grid(row=0, column=0, sticky='w')
        self.download = ttk.Button(bottom, text='Download Tax Map', command=self.download_map, state='disabled')
        self.download.grid(row=0, column=1, padx=6)
        self.use = ttk.Button(bottom, text='Use This Tax Map', command=self.accept, state='disabled')
        self.use.grid(row=0, column=2, padx=6)
        ttk.Button(bottom, text='Close', command=self.close).grid(row=0, column=3)
        self._start_timer = self.after(50, self.start_load)

    def _button(self, parent, text, command, **kwargs):
        button = ttk.Button(parent, text=text, command=command, **kwargs)
        self._controls.append(button)
        return button

    def changed(self):
        if self._loading:
            return
        self.use.configure(state='disabled')
        self.download.configure(state='disabled')
        self.status.set('Click Update Preview to apply the framing changes.')

    def zoom(self, amount):
        self.settings['zoom'] = max(.5, min(2, self.settings['zoom'] + amount))
        self.changed()

    def pan(self, dx, dy):
        if self.data is None:
            return
        _, extent = map_extent(self.data, self.settings)
        distance = (extent[2] - extent[0]) * .08
        self.settings['pan_x'] += dx * distance
        self.settings['pan_y'] += dy * distance
        self.changed()

    def reset(self):
        self.settings = default_settings()
        self.changed()
        self.start_load()

    def start_load(self, refresh=False):
        if self._closed or self._loading:
            return
        gc.collect()
        self._loading = True
        for widget in self._controls + [self.download, self.use]:
            widget.configure(state='disabled')
        settings = dict(self.settings)
        existing = None if refresh else self.data
        self.status.set('Loading HCAD tax map...')

        def work():
            try:
                data = existing
                if data is None:
                    from .automatic_case import load_map_data
                    data=load_map_data(self.source,self.wizard,self.output,
                        progress=lambda s:self.events.put(('status',s)),cancel=self.cancel,use_cache=not refresh)
                client = HCADPrintClient(Path(self.output) / '.tax_map_cache',
                    progress=lambda s: self.events.put(('status', s)), cancel=self.cancel, use_cache=not refresh)
                rendered = render_tax_map(data, settings, client)
                self.events.put(('ready', (data, settings, rendered)))
            except Exception as exc:
                self.events.put(('error', str(exc)))
        threading.Thread(target=work, daemon=True).start()
        self._poll_timer = self.after(100, self.poll)

    def poll(self):
        self._poll_timer = None
        if self._closed:
            return
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == 'status':
                    self.status.set(value)
                    continue
                self._loading = False
                for widget in self._controls:
                    widget.configure(state='normal')
                if kind == 'error':
                    self.show_review(value)
                    if self.rendered and self.rendered_settings == self.settings:
                        self.status.set('Could not refresh HCAD. The previous tax map is still available to download or use.')
                        self.download.configure(state='normal')
                        self.use.configure(state='normal')
                    else:
                        self.status.set('Could not create the tax map. Retry with Update Preview or Refresh HCAD data.')
                    return
                self.data, self.rendered_settings, self.rendered = value
                from .automatic_case import select_updated_calculation,council_summary
                updated=self.data.get('calculation_path',self.source)
                select_updated_calculation(self.owner,updated,self.source)
                self.source=updated
                self._image = Image.open(io.BytesIO(pdf_png(self.rendered['pdf'], 110))).copy()
                c = self.data['case']
                self.summary.configure(text=f"{c['reference']}\n{c['block']}\n{c['sides']}\n\n{len(self.data['matched'])} properties outlined in yellow"+council_summary(self.data))
                self.snapshot_warnings()
                self.show_review('\n\n'.join(self.rendered['issues']+self.data.get('calculation_updates',[])))
                self.display()
                self.status.set('Review the yellow boundary, then download or use this tax map.')
                self.download.configure(state='normal')
                self.use.configure(state='normal')
                return
        except queue.Empty:
            pass
        self._poll_timer = self.after(100, self.poll)

    def show_review(self, text):
        self.review.configure(state='normal')
        self.review.delete('1.0', 'end')
        self.review.insert('1.0', text)
        self.review.configure(state='disabled')

    def display(self):
        if self._image is None or self._closed:
            return
        w, h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if w < 10 or h < 10:
            return
        ratio = min((w-16)/self._image.width, (h-16)/self._image.height)
        size = (max(1, int(self._image.width*ratio)), max(1, int(self._image.height*ratio)))
        self.photo = ImageTk.PhotoImage(self._image.resize(size, Image.Resampling.LANCZOS), master=self)
        self.canvas.delete('all')
        self.canvas.create_image(w/2, h/2, image=self.photo)

    def snapshot_warnings(self):
        try:
            changed = fingerprint(self.source) != self.data['case']['workbook_hash']
        except OSError:
            changed = True
        source, _, _ = wizard_paths(self.owner)
        if changed or Path(source).resolve() != Path(self.source).resolve():
            note = 'The calculation sheet has changed. This tax map uses the case and properties displayed in this preview.'
            if note not in self.rendered['issues']:
                self.rendered['issues'].append(note)

    def download_map(self):
        if self._loading or not self.rendered or self.rendered_settings != self.settings:
            return
        destination = filedialog.asksaveasfilename(parent=self, title='Download HCAD Tax Map', initialdir=self.output,
            initialfile=self.data['case']['reference'].replace(' ', '_')+'_Tax_Map.pdf', defaultextension='.pdf',
            filetypes=[('PDF map', '*.pdf'), ('PNG image', '*.png')])
        if destination:
            try:
                self.snapshot_warnings()
                export_map(self.rendered, destination)
                self.status.set('Tax map downloaded to ' + destination)
            except Exception as exc:
                messagebox.showerror('Download Tax Map', str(exc), parent=self)

    def accept(self):
        if self._loading or not self.rendered or self.rendered_settings != self.settings:
            return
        try:
            self.snapshot_warnings()
            _, _, output = wizard_paths(self.owner)
            path = save_tax_map(self.data, self.rendered_settings, self.rendered, output or self.output)
            self.owner._set_generated_tax_map(str(path))
            self.close()
        except Exception as exc:
            messagebox.showerror('Use Tax Map', str(exc), parent=self)

    def close(self):
        if self._closed:
            return
        self._closed = True
        self.cancel.set()
        for timer in (self._start_timer, self._poll_timer):
            if timer:
                self.after_cancel(timer)
        self.destroy()


def open_tax_map_preview(owner):
    source, _, output = wizard_paths(owner)
    if not source or not Path(source).is_file():
        raise MapError('Select and save the calculation workbook before creating a tax map.')
    if not output:
        raise MapError('Select an output folder before creating a tax map.')
    prior = getattr(owner, '_tax_map_preview', None)
    if prior is not None and prior.winfo_exists():
        prior.lift()
        return prior
    preview = TaxMapPreview(owner)
    owner._tax_map_preview = preview
    return preview
