"""Individual street-name size and rotation controls for the map preview."""
import tkinter as tk
from tkinter import ttk, messagebox
from .boundary_map_render import STREET_FONT_SIZE


class StreetAppearanceEditor(tk.Toplevel):
    def __init__(self, preview, key=None):
        super().__init__(preview)
        self.preview = preview
        self.title('Street name size and rotation')
        self.transient(preview)
        self.resizable(False, False)
        body = ttk.Frame(self, padding=14)
        body.pack(fill='both', expand=True)
        self.items = [i for i in preview.rendered['labels'] if i['key'].startswith('street:')]
        ttk.Label(body, text='Street name').grid(row=0, column=0, sticky='w', pady=5)
        self.choice = ttk.Combobox(body, state='readonly', width=38, values=[i['text'] for i in self.items])
        self.choice.grid(row=0, column=1, sticky='ew', pady=5)
        self.choice.bind('<<ComboboxSelected>>', lambda e: self.load_values())
        self.size = tk.DoubleVar(value=STREET_FONT_SIZE)
        self.angle = tk.DoubleVar(value=0)
        self.automatic = tk.BooleanVar(value=True)
        ttk.Label(body, text='Font size (points)').grid(row=1, column=0, sticky='w', pady=5)
        ttk.Spinbox(body, from_=6, to=72, increment=1, textvariable=self.size, width=10).grid(row=1, column=1, sticky='w', pady=5)
        ttk.Label(body, text='Rotation (degrees)').grid(row=2, column=0, sticky='w', pady=5)
        self.rotation_input = ttk.Spinbox(body, from_=-180, to=180, increment=5, textvariable=self.angle, width=10)
        self.rotation_input.grid(row=2, column=1, sticky='w', pady=5)
        ttk.Checkbutton(body, text='Follow street direction automatically', variable=self.automatic,
                        command=self.toggle_automatic).grid(row=3, column=0, columnspan=2, sticky='w', pady=5)
        ttk.Label(body, text='0 degrees is horizontal; 90 degrees is vertical.\nApply updates the map while this window stays open.').grid(row=4,column=0,columnspan=2,sticky='w',pady=8)
        actions=ttk.Frame(body);actions.grid(row=5,column=0,columnspan=2,sticky='e')
        ttk.Button(actions,text='Reset this street',command=self.reset_selected).pack(side='left',padx=4)
        ttk.Button(actions,text='Apply',command=self.apply).pack(side='left',padx=4)
        ttk.Button(actions,text='Close',command=self.destroy).pack(side='left',padx=4)
        self.select_key(key)

    def select_key(self, key=None):
        key = key or self.selected_key()
        self.items = [i for i in self.preview.rendered['labels'] if i['key'].startswith('street:')]
        self.choice.configure(values=[i['text'] for i in self.items])
        index = next((n for n,i in enumerate(self.items) if i['key']==key),0)
        if self.items:
            self.choice.current(index)
            self.load_values()

    def selected_key(self):
        index=self.choice.current()
        return self.items[index]['key'] if index>=0 else None

    def load_values(self):
        key=self.selected_key()
        style=self.preview.settings.get('street_styles',{}).get(key,{})
        item=next((i for i in self.preview.rendered['labels'] if i['key']==key),{})
        self.size.set(style.get('size',STREET_FONT_SIZE))
        self.angle.set(style.get('angle') if style.get('angle') is not None else item.get('angle',0))
        self.automatic.set(style.get('angle') is None)
        self.toggle_automatic()

    def toggle_automatic(self):
        self.rotation_input.configure(state='disabled' if self.automatic.get() else 'normal')

    def apply(self):
        try:
            size=self.size.get()
            angle=None if self.automatic.get() else self.angle.get()
        except (ValueError, tk.TclError):
            messagebox.showwarning('Street name was not changed','Enter numbers for font size and rotation.',parent=self)
            return
        if self.preview.apply_street_style(self.selected_key(),size,angle):
            self.load_values()

    def reset_selected(self):
        if self.preview.apply_street_style(self.selected_key()):
            self.load_values()
