"""Scrollable content area for the RCA wizards.

Problem this solves (master brief, Pitfall 7 "Hidden bottom buttons"): on
smaller screens the Generate page's lower buttons — Draft Ordinance, RCA Packet
Builder — can fall below the window with no way to reach them.

Design goal: add scrolling *without* touching any page-building code.  Every
wizard builds its pages into a frame it calls ``self.content``.  Here we make
``self.content`` be the *interior* of a vertically scrolling canvas.  Pages keep
attaching to it with ``grid`` or ``pack`` exactly as before; clearing it with
``for w in self.content.winfo_children(): w.destroy()`` still works because the
interior frame itself is never destroyed.

The scrollbar auto-hides when everything fits, so pages that already fit look
unchanged.  Mouse-wheel scrolling is bound only while the pointer is over the
area, so it never fights an inner widget (e.g. an Owners Treeview) elsewhere.
"""

from __future__ import annotations

import sys
import tkinter as tk


class ScrollableContent(tk.Frame):
    """A vertically scrolling region.  Build page content into ``.interior``.

    Usage inside a wizard's ``_build_ui``::

        holder = ScrollableContent(main, bg=self.colors["bg"])
        holder.grid(row=1, column=0, sticky="nsew")
        self.content = holder.interior        # pages attach here, unchanged
    """

    def __init__(self, parent, bg: str = "#FFFFFF", **kwargs):
        super().__init__(parent, **kwargs)
        self._bg = bg
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self._canvas.grid(row=0, column=0, sticky="nsew")

        self._vbar = tk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._on_scrollset)

        #: Build all page content into this frame (it is the wizard's
        #: ``self.content``).  It is never destroyed, only its children are.
        self.interior = tk.Frame(self._canvas, bg=bg)
        self._win = self._canvas.create_window((0, 0), window=self.interior, anchor="nw")

        self.interior.bind("<Configure>", self._on_interior_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)

        # Wheel support only while the pointer is over this region.
        self._canvas.bind("<Enter>", self._bind_wheel)
        self._canvas.bind("<Leave>", self._unbind_wheel)

    # -- keep the interior the canvas's width, and refresh the scroll region --
    def _on_interior_configure(self, _event=None):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))
        self._sync_scrollbar()

    def _on_canvas_configure(self, event):
        self._canvas.itemconfigure(self._win, width=event.width)
        self._sync_scrollbar()

    # -- auto-hide the scrollbar when everything already fits --
    def _on_scrollset(self, first, last):
        self._vbar.set(first, last)
        self._sync_scrollbar()

    def _sync_scrollbar(self):
        try:
            first, last = self._canvas.yview()
        except Exception:
            return
        needed = not (float(first) <= 0.0 and float(last) >= 1.0)
        mapped = bool(self._vbar.winfo_ismapped())
        if needed and not mapped:
            self._vbar.grid(row=0, column=1, sticky="ns")
        elif not needed and mapped:
            self._vbar.grid_remove()

    # -- mouse wheel (cross-platform) --
    def _bind_wheel(self, _event=None):
        self._canvas.bind_all("<MouseWheel>", self._on_wheel)      # Windows / macOS
        self._canvas.bind_all("<Button-4>", self._on_wheel_linux)  # X11 up
        self._canvas.bind_all("<Button-5>", self._on_wheel_linux)  # X11 down

    def _unbind_wheel(self, _event=None):
        self._canvas.unbind_all("<MouseWheel>")
        self._canvas.unbind_all("<Button-4>")
        self._canvas.unbind_all("<Button-5>")

    def _on_wheel(self, event):
        first, last = self._canvas.yview()
        if float(first) <= 0.0 and float(last) >= 1.0:
            return  # nothing to scroll
        delta = event.delta
        if sys.platform == "darwin":
            step = -1 * delta
        else:
            step = -1 * int(delta / 120)
        self._canvas.yview_scroll(step, "units")

    def _on_wheel_linux(self, event):
        first, last = self._canvas.yview()
        if float(first) <= 0.0 and float(last) >= 1.0:
            return
        self._canvas.yview_scroll(-1 if event.num == 4 else 1, "units")
