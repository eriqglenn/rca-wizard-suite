"""Shared embedding support for the RCA Wizard Suite.

Each City Planning wizard (MLS, MBL, MLS Renewal, MBL Renewal) was originally
written as its own top-level ``tk.Tk`` window.  Inside the suite they run as
``tk.Frame`` instances hosted by one shared root window (see ``app.py``).

``tk.Frame`` does not implement the window-manager methods the wizards call
(``title``, ``geometry``, ``minsize``), so this mixin forwards those calls to
the real toplevel window.  This keeps the original wizard code byte-for-byte
identical everywhere except its class declaration, which preserves each
wizard's tested behavior (a core rule of the master brief: do not rewrite
working modules during a packaging change).

The title forward also prefixes the suite name/version so the title bar keeps
serving as the build-version diagnostic required by the master brief
(section 14 / 20.12.4): the planner can read both the suite build and the
individual wizard build from one place.
"""

from __future__ import annotations

SUITE_NAME = "City Planning RCA Wizard Suite"
SUITE_VERSION = "v1.30"


class EmbeddedWizardMixin:
    """Mixin that lets a wizard written as a ``tk.Tk`` window run as a
    ``tk.Frame`` embedded in the shared suite window.

    Use as::

        class MyWizard(EmbeddedWizardMixin, tk.Frame):
            def __init__(self, parent=None):
                super().__init__(parent)
                ...

    Only window-manager calls are forwarded; everything else behaves exactly
    as it did when the wizard owned its own root window.
    """

    def _add_boundary_map_controls(self, parent, row=3):
        from .boundary_map_ui import MapControls
        return MapControls(parent, self, row)

    def _add_application_import_controls(self, parent, row=4):
        from tkinter import ttk
        from .application_import_ui import open_application_import
        from .ilms_ui import open_ilms_holds
        actions=ttk.Frame(parent)
        actions.grid(row=row,column=0,columnspan=3,sticky='w',padx=12,pady=(2,10))
        button=ttk.Button(actions,text='Import Application / Enter Boundary / Fill Calculation Sheet',
                         command=lambda:open_application_import(self))
        button.pack(side='left')
        ttk.Button(actions,text='ILMS Holds',command=lambda:open_ilms_holds(self)).pack(side='left',padx=(8,0))
        return button

    def _create_boundary_map(self):
        from tkinter import messagebox
        try:
            from .boundary_map_ui import open_map_preview
            open_map_preview(self)
        except Exception as e:
            messagebox.showerror('Create Boundary Map', str(e), parent=self.winfo_toplevel())

    def _find_subdivision_plats(self):
        from tkinter import messagebox
        try:
            from .plat_search_ui import open_plat_finder
            open_plat_finder(self)
        except Exception as e:
            messagebox.showerror('Find Subdivision Plats', str(e), parent=self.winfo_toplevel())

    def _set_downloaded_subdivision_plat(self, path):
        if hasattr(self, 'vars'):
            self.vars['subdivision_plat'].set(path)
            self.vars['status'].set('Subdivision plats selected: ' + path)
        else:
            self.files['subdivision_plat'] = path
            self.status_var.set('Subdivision plats selected: ' + path)
            self._select_files_page()
        self._save_session()

    def _create_tax_map(self):
        from tkinter import messagebox
        try:
            from .tax_map_ui import open_tax_map_preview
            open_tax_map_preview(self)
        except Exception as e:
            messagebox.showerror('Create Tax Map', str(e), parent=self.winfo_toplevel())

    def _set_generated_tax_map(self, path):
        if hasattr(self, 'vars'):
            self.vars['tax_map'].set(path)
            self.vars['status'].set('HCAD tax map created and selected: ' + path)
        else:
            self.files['tax_map'] = path
            self.status_var.set('HCAD tax map created and selected: ' + path)
            self._select_files_page()
        self._save_session()

    def _check_boundary_map(self):
        from .boundary_map_ui import wizard_paths
        from .boundary_map import assert_map_current
        source, selected, _ = wizard_paths(self)
        return assert_map_current(selected, source, self._boundary_map_wizard_key)

    def _set_generated_boundary_map(self, path):
        if hasattr(self, 'vars'):
            self.vars['map'].set(path)
            self.vars['status'].set('Boundary map created and selected: ' + path)
        else:
            self.files['map'] = path
            self.status_var.set('Boundary map created and selected: ' + path)
            self._select_files_page()
        self._save_session()

    def _document_output(self, filename):
        """An empty folder must not silently save into the application directory."""
        from pathlib import Path
        folder = str(self.files.get('out', '')).strip()
        if not folder:
            raise ValueError('Select an output folder before generating documents.')
        return Path(folder) / filename

    def title(self, text: str | None = None):
        top = self.winfo_toplevel()
        if text is None:
            return top.title()
        return top.title(f"{SUITE_NAME} {SUITE_VERSION} \u2014 {text}")

    def geometry(self, spec: str | None = None):
        top = self.winfo_toplevel()
        if spec is None:
            return top.geometry()
        return top.geometry(spec)

    def minsize(self, width: int, height: int):
        return self.winfo_toplevel().minsize(width, height)

    def maxsize(self, width: int, height: int):
        return self.winfo_toplevel().maxsize(width, height)

    def resizable(self, *args):
        return self.winfo_toplevel().resizable(*args)

    def iconbitmap(self, *args, **kwargs):
        return self.winfo_toplevel().iconbitmap(*args, **kwargs)
