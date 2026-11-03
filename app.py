"""City Planning RCA Wizard Suite — launcher and home screen.

One desktop application hosting all four City Planning RCA document wizards:

    MLS Wizard          Special Minimum Lot Size block applications
    MBL Wizard          Special Minimum Building Line block applications
    MLS Renewal Wizard  SMLS block renewals
    MBL Renewal Wizard  SMBL block renewals

Each wizard keeps its own tested code, templates, and sample input; the suite
only provides the shared window, the home screen, and the shared core modules
(``core/``).  Opening a wizard creates a fresh instance; returning Home
destroys it completely, so every case starts from a clean state.

Packaged installs run ``RCA Wizard Suite.exe`` without a separate Python install.
For a source install, run ``Run Wizard.bat`` or invoke Python directly::

    python app.py                # home screen
    python app.py --wizard mls   # jump straight into one wizard
                                 # keys: mls | mbl | mls_renewal | mbl_renewal
"""

from __future__ import annotations

import argparse
import importlib
import os
import sys
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox

# Make absolute imports (core.*, wizards.*) work no matter how app.py is run.
_WIZARD_FILES = Path(__file__).resolve().parent
if str(_WIZARD_FILES) not in sys.path:
    sys.path.insert(0, str(_WIZARD_FILES))

from core.embedding import SUITE_NAME, SUITE_VERSION  # noqa: E402
from core.workbooks import calculation_template_path, save_calculation_sheet  # noqa: E402

HOME_GEOMETRY = "1280x860"
HOME_MIN = (1100, 760)

# ----------------------------------------------------------------------------
# Wizard registry.
#
# Each entry is data only; the wizard module is imported lazily when the
# planner opens it, so a problem in one wizard can never block the others.
# ``version`` mirrors the build stamp each wizard reports in its own title bar
# (the title-bar diagnostic from the master brief still works: the suite title
# shows BOTH the suite build and the active wizard build).
# ----------------------------------------------------------------------------
WIZARDS = [
    {
        "key": "mls",
        "name": "MLS Wizard",
        "version": "v15.0",
        "module": "wizards.mls.wizard",
        "class": "MailoutWizard",
        "color": "#1E3A8A",
        "color_mid": "#1D4ED8",
        "color_light": "#DBEAFE",
        "tagline": "Special Minimum Lot Size — new block applications",
        "details": "Owner notification mail-out \u2022 Applicant sign letter \u2022 Director / PC approvals \u2022 Routing form \u2022 Draft ordinance \u2022 RCA packet PDF \u2022 HCAD records",
    },
    {
        "key": "mbl",
        "name": "MBL Wizard",
        "version": "v9.0",
        "module": "wizards.mbl.wizard",
        "class": "MBLWizard",
        "color": "#14532D",
        "color_mid": "#16A34A",
        "color_light": "#DCFCE7",
        "tagline": "Special Minimum Building Line — block applications",
        "details": "Owner notification mail-out \u2022 Applicant sign letter \u2022 Director / PC approvals \u2022 Routing form \u2022 Draft ordinance \u2022 RCA packet PDF \u2022 HCAD records",
    },
    {
        "key": "mls_renewal",
        "name": "MLS Renewal Wizard",
        "version": "v3.0",
        "module": "wizards.mls_renewal.wizard",
        "class": "MLSRenewalWizard",
        "color": "#581C2B",
        "color_mid": "#991B1B",
        "color_light": "#FEE2E2",
        "tagline": "Special Minimum Lot Size — block renewals",
        "details": "SFR / No-SFR renewal letters \u2022 Spanish notice fit \u2022 Renewal approvals \u2022 Renewal routing form \u2022 Original ordinance attachment \u2022 Draft renewal ordinance \u2022 RCA packet PDF",
    },
    {
        "key": "mbl_renewal",
        "name": "MBL Renewal Wizard",
        "version": "v2.0",
        "module": "wizards.mbl_renewal.wizard",
        "class": "MBLRenewalWizard",
        "color": "#7C2D12",
        "color_mid": "#EA580C",
        "color_light": "#FFEDD5",
        "tagline": "Special Minimum Building Line — block renewals",
        "details": "Renewal owner letters \u2022 Building-line renewal approvals \u2022 Renewal routing form \u2022 Original ordinance attachment \u2022 Draft renewal ordinance \u2022 RCA packet PDF",
    },
]

_BG = "#F1F5F9"
_CARD = "#FFFFFF"
_TEXT = "#0F172A"
_MUTED = "#64748B"
_HEADER = "#0F172A"


class SuiteApp(tk.Tk):
    """Single shared root window: shows the home screen, hosts one wizard at
    a time, and provides the persistent "All Wizards" navigation bar."""

    def __init__(self, hidden=False):
        super().__init__()
        if hidden:
            self.withdraw()
        self.title(f"{SUITE_NAME} {SUITE_VERSION}")
        self.geometry(HOME_GEOMETRY)
        self.minsize(*HOME_MIN)
        self.configure(bg=_BG)
        self._container: tk.Frame | None = None
        self.show_home()

    # ------------------------------------------------------------------ home
    def _clear(self):
        if self._container is not None:
            self._container.destroy()
            self._container = None

    def show_home(self):
        self._clear()
        self.title(f"{SUITE_NAME} {SUITE_VERSION}")
        c = tk.Frame(self, bg=_BG)
        c.pack(fill="both", expand=True)
        self._container = c

        header = tk.Frame(c, bg=_HEADER)
        header.pack(fill="x")
        tk.Label(
            header,
            text="City of Houston \u2014 Planning & Development",
            bg=_HEADER, fg="#CBD5E1", font=("Segoe UI", 10),
        ).pack(anchor="w", padx=34, pady=(20, 0))
        tk.Label(
            header,
            text="RCA Document Wizard Suite",
            bg=_HEADER, fg="white", font=("Segoe UI", 26, "bold"),
        ).pack(anchor="w", padx=32, pady=(2, 2))
        tk.Label(
            header,
            text="Select the application you are processing. Each wizard fills the official City "
                 "templates from the calculation workbook and assembles the RCA packet.",
            bg=_HEADER, fg="#94A3B8", font=("Segoe UI", 11),
            wraplength=900, justify="left",
        ).pack(anchor="w", padx=34, pady=(0, 14))

        tools = tk.Frame(header, bg=_HEADER)
        tools.pack(fill="x", padx=34, pady=(0, 18))
        tk.Button(
            tools, text="\u2699  Manage Templates",
            bg="#334155", fg="white", activebackground="#475569", activeforeground="white",
            relief="flat", font=("Segoe UI", 10, "bold"), padx=14, pady=6, cursor="hand2",
            command=self.open_template_manager,
        ).pack(side="left")
        tk.Button(
            tools, text="Check PDF setup", bg="#334155", fg="white",
            activebackground="#475569", activeforeground="white", relief="flat",
            font=("Segoe UI", 10, "bold"), padx=12, pady=6, cursor="hand2",
            command=self.open_pdf_setup,
        ).pack(side="left", padx=(8, 0))
        tk.Label(
            tools, text=self._setup_summary(), bg=_HEADER, fg="#94A3B8",
            font=("Segoe UI", 9), wraplength=560, justify="left",
        ).pack(side="left", padx=16)

        grid = tk.Frame(c, bg=_BG)
        grid.pack(fill="both", expand=True, padx=30, pady=24)
        grid.grid_columnconfigure(0, weight=1, uniform="col")
        grid.grid_columnconfigure(1, weight=1, uniform="col")
        grid.grid_rowconfigure(0, weight=1, uniform="row")
        grid.grid_rowconfigure(1, weight=1, uniform="row")

        for i, spec in enumerate(WIZARDS):
            self._build_card(grid, spec, row=i // 2, col=i % 2)

        footer = tk.Frame(c, bg=_BG)
        footer.pack(fill="x", side="bottom")
        tk.Label(
            footer,
            text=f"{SUITE_NAME} {SUITE_VERSION}  \u2022  "
                 + "  \u2022  ".join(f"{w['name']} {w['version']}" for w in WIZARDS)
                 + "  \u2022  Templates live in Wizard_Files\\templates",
            bg=_BG, fg=_MUTED, font=("Segoe UI", 9),
        ).pack(pady=(0, 12))

    def _build_card(self, parent: tk.Frame, spec: dict, row: int, col: int):
        outer = tk.Frame(parent, bg="#CBD5E1", bd=0)
        outer.grid(row=row, column=col, sticky="nsew", padx=12, pady=12)
        card = tk.Frame(outer, bg=_CARD)
        card.pack(fill="both", expand=True, padx=1, pady=1)

        stripe = tk.Frame(card, bg=spec["color"], height=8)
        stripe.pack(fill="x")

        body = tk.Frame(card, bg=_CARD)
        body.pack(fill="both", expand=True, padx=22, pady=(14, 16))

        top = tk.Frame(body, bg=_CARD)
        top.pack(fill="x")
        tk.Label(top, text=spec["name"], bg=_CARD, fg=_TEXT,
                 font=("Segoe UI", 17, "bold")).pack(side="left")
        tk.Label(top, text=spec["version"], bg=spec["color_light"], fg=spec["color"],
                 font=("Segoe UI", 9, "bold"), padx=8, pady=2).pack(side="right")

        tk.Label(body, text=spec["tagline"], bg=_CARD, fg=spec["color_mid"],
                 font=("Segoe UI", 11, "bold"), anchor="w", justify="left",
                 wraplength=440).pack(fill="x", pady=(4, 6))
        tk.Label(body, text=spec["details"], bg=_CARD, fg=_MUTED,
                 font=("Segoe UI", 9), anchor="w", justify="left",
                 wraplength=460).pack(fill="x")

        actions = tk.Frame(body, bg=_CARD)
        actions.pack(side="bottom", anchor="w", fill="x", pady=(10, 0))
        btn = tk.Button(
            actions, text="Open wizard  \u2192",
            bg=spec["color"], fg="white",
            activebackground=spec["color_mid"], activeforeground="white",
            relief="flat", font=("Segoe UI", 10, "bold"), padx=12, pady=8,
            cursor="hand2",
            command=lambda k=spec["key"]: self.open_wizard(k),
        )
        btn.pack(side="left")
        tk.Button(
            actions, text="Download calculation sheet",
            bg=spec["color_light"], fg=spec["color"],
            activebackground=spec["color_light"], activeforeground=spec["color"],
            relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=8,
            cursor="hand2", command=lambda k=spec["key"]: self.download_calculation_sheet(k),
        ).pack(side="left", padx=(8, 0))

    def download_calculation_sheet(self, key: str):
        """Save an editable copy of the matching instructional workbook."""
        spec = next(w for w in WIZARDS if w["key"] == key)
        source = calculation_template_path(key)
        destination = filedialog.asksaveasfilename(
            parent=self, title=f"Save calculation sheet — {spec['name']}",
            initialfile=source.name, defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")], confirmoverwrite=True,
        )
        if not destination:
            return
        try:
            saved = save_calculation_sheet(key, destination)
        except Exception as exc:
            messagebox.showerror("Calculation sheet", str(exc), parent=self)
            return
        try:
            os.startfile(str(saved))
        except OSError:
            messagebox.showinfo("Calculation sheet saved",
                                f"Saved to:\n{saved}\n\nOpen this file in Excel. "
                                "The example entries and formulas are included.", parent=self)

    # ---------------------------------------------------------------- wizard
    def open_wizard(self, key: str):
        spec = next((w for w in WIZARDS if w["key"] == key), None)
        if spec is None:
            messagebox.showerror(SUITE_NAME, f"Unknown wizard: {key}")
            return

        # Import lazily so one wizard's problem never blocks the others.
        try:
            module = importlib.import_module(spec["module"])
            wizard_cls = getattr(module, spec["class"])
        except Exception:
            messagebox.showerror(
                SUITE_NAME,
                f"Could not start {spec['name']}.\n\n"
                f"{traceback.format_exc(limit=3)}\n"
                "If this is the first run, double-click install_requirements.bat once, "
                "then try again.",
            )
            return

        self._clear()
        c = tk.Frame(self, bg=_BG)
        c.pack(fill="both", expand=True)
        self._container = c

        # Persistent navigation bar (suite-level, wizard code untouched).
        bar = tk.Frame(c, bg=spec["color"], height=40)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Button(
            bar, text="\u2190  All Wizards",
            bg=spec["color"], fg="white",
            activebackground=spec["color_mid"], activeforeground="white",
            relief="flat", font=("Segoe UI", 10, "bold"), padx=14,
            cursor="hand2", command=self.go_home,
        ).pack(side="left", fill="y")
        tk.Label(
            bar, text=f"{SUITE_NAME} {SUITE_VERSION}",
            bg=spec["color"], fg="#E2E8F0", font=("Segoe UI", 9),
        ).pack(side="right", padx=14)

        host = tk.Frame(c, bg=_BG)
        host.pack(fill="both", expand=True)

        try:
            wizard = wizard_cls(host)
        except Exception:
            messagebox.showerror(
                SUITE_NAME,
                f"{spec['name']} failed to open.\n\n{traceback.format_exc(limit=3)}",
            )
            self.show_home()
            return
        wizard.pack(fill="both", expand=True)

    # ------------------------------------------------------ templates
    def open_pdf_setup(self):
        from core.pdf_setup import show_dialog
        show_dialog(self)

    def _setup_summary(self) -> str:
        """One-line diagnostic for the home screen (build type, PDF engine)."""
        try:
            from core.office_convert import describe_engine
            engine = describe_engine()
        except Exception:
            engine = "unknown"
        build = "packaged build" if getattr(sys, "frozen", False) else "running from source"
        return f"{build}  \u2022  PDF engine: {engine}"

    def open_template_manager(self):
        """Show the Template Manager with the same navigation bar as a wizard."""
        try:
            from core.template_manager import TemplateManager
        except Exception:
            messagebox.showerror(
                SUITE_NAME,
                f"Could not open the Template Manager.\n\n{traceback.format_exc(limit=3)}")
            return
        self._clear()
        self.title(f"{SUITE_NAME} {SUITE_VERSION} \u2014 Template Manager")
        c = tk.Frame(self, bg=_BG)
        c.pack(fill="both", expand=True)
        self._container = c
        bar = tk.Frame(c, bg="#334155", height=40)
        bar.pack(fill="x")
        bar.pack_propagate(False)
        tk.Button(
            bar, text="\u2190  All Wizards", bg="#334155", fg="white",
            activebackground="#475569", activeforeground="white", relief="flat",
            font=("Segoe UI", 10, "bold"), padx=14, cursor="hand2", command=self.go_home,
        ).pack(side="left", fill="y")
        tk.Label(bar, text=f"{SUITE_NAME} {SUITE_VERSION}", bg="#334155", fg="#E2E8F0",
                 font=("Segoe UI", 9)).pack(side="right", padx=14)
        host = tk.Frame(c, bg=_BG)
        host.pack(fill="both", expand=True)
        try:
            TemplateManager(host).pack(fill="both", expand=True)
        except Exception:
            messagebox.showerror(
                SUITE_NAME, f"Template Manager failed to open.\n\n{traceback.format_exc(limit=3)}")
            self.show_home()

    def go_home(self):
        """Return to the wizard picker.  The active wizard is destroyed
        entirely, so the next case always starts from a clean state."""
        self.show_home()


def main():
    parser = argparse.ArgumentParser(description=f"{SUITE_NAME} {SUITE_VERSION}")
    parser.add_argument(
        "--wizard",
        choices=[w["key"] for w in WIZARDS],
        help="open one wizard directly (skip the home screen)",
    )
    parser.add_argument(
        "--templates", action="store_true",
        help="open the Template Manager directly (skip the home screen)",
    )
    parser.add_argument('--startup-check', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args()

    app = SuiteApp(hidden=args.startup_check)
    if args.startup_check:
        try:
            app.update_idletasks()
            app.update()
            print(f'STARTUP CHECK PASSED: {SUITE_NAME} {SUITE_VERSION}; home screen initialized.')
        finally:
            app.destroy()
        return
    if args.wizard:
        app.open_wizard(args.wizard)
    elif args.templates:
        app.open_template_manager()
    app.mainloop()


if __name__ == "__main__":
    main()
