from __future__ import annotations

import os
import subprocess
import sys
import threading
import time
import webbrowser
import urllib.request
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core.embedding import EmbeddedWizardMixin
from core.scrollable import ScrollableContent
from core.workbooks import calculation_template_path, WorkbookReadCache
from core import session as core_session
from core.casefolder import maybe_case_folder
from core.errorlog import log_exception
from core.preflight import run_preflight, build_preflight_panel
from core.paths import templates_dir, HCAD_DOWNLOADER
from core.templates import template_path
from core import launch as core_launch

from .mailout_engine import (
    generate_planning_commission_approval,
    summarize_workbook,
    validate_applicant_inputs,
    build_director_context,
    validate_director_inputs,
    build_routing_context,
    validate_routing_inputs,
    build_pc_context,
    validate_pc_inputs,
    build_ordinance_context,
    validate_ordinance_inputs,
    export_hcad_address_list,
)
# 2026 City templates (wording-based) -- see templates_2026.py
from .templates_2026 import (
    generate_owner_letters,
    generate_applicant_letter,
    generate_director_approval,
    generate_routing_form,
    generate_staff_report,
    generate_draft_ordinance,
    generate_rca_packet_pdf,
    build_context,
)

APP_TITLE = "MLS/RCA Document Wizard"
_WIZARD_KEY = "mls"
_SESSION_KEYS = ("excel", "template", "protest_template", "hearing_template", "applicant_template",
                 "director_template", "routing_template", "pc_template", "staff_report_template",
                 "ordinance_template", "map", "subdivision_plat", "tax_map", "hcad_records",
                 "include_staff_analysis", "output_folder")
_TEMPLATES = templates_dir("mls")
STEPS = [
    "1. Select Files",
    "2. Review Case Info",
    "3. Review Owners",
    "4. Generate Letters",
    "5. Peer Reviewer",
]

class MailoutWizard(EmbeddedWizardMixin, tk.Frame):
    _boundary_map_wizard_key = _WIZARD_KEY
    def __init__(self, parent=None):
        super().__init__(parent)
        self.title(APP_TITLE + " v15.0 (2026 City templates)")
        self.geometry("1220x860")
        self.minsize(1100, 780)
        self.summary = None
        self.current_step = 0
        self.vars = {
            "excel": tk.StringVar(value=str(calculation_template_path("mls"))),
            "template": tk.StringVar(value=str(template_path("mls", "template", "MLS_Owner_Notification_Letter.docx"))),
            "protest_template": tk.StringVar(value=str(template_path("mls", "protest_template", "MLS_Protest_All_Owners.docx"))),
            "hearing_template": tk.StringVar(value=str(template_path("mls", "hearing_template", "MLS_Notice_Public_Hearing_Less_51.docx"))),
            "applicant_template": tk.StringVar(value=str(template_path("mls", "applicant_template", "MLS_Applicant_Complete_Signs.docx"))),
            "director_template": tk.StringVar(value=str(template_path("mls", "director_template", "MLS_Director_Approval.docx"))),
            "routing_template": tk.StringVar(value=str(template_path("mls", "routing_template", "MLS_Routing Form.docx"))),
            "pc_template": tk.StringVar(value=str(template_path("mls", "pc_template", "PC_Approval_NEW.docx"))),
            "staff_report_template": tk.StringVar(value=str(template_path("mls", "staff_report_template", "MLS_OneMeeting_Staff_Report.docx"))),
            "ordinance_template": tk.StringVar(value=str(template_path("mls", "ordinance_template", "MLS_Ordinance_Template.docx"))),
            "map": tk.StringVar(),
            "subdivision_plat": tk.StringVar(),
            "tax_map": tk.StringVar(),
            "hcad_records": tk.StringVar(),
            "include_staff_analysis": tk.StringVar(value="no"),
            "output_folder": tk.StringVar(value=str(Path.home() / "Documents")),
            "status": tk.StringVar(value="Select the calculation workbook, notification template, and map file."),
        }
        # Restore the planner's last file selections for this wizard
        # (paths that no longer exist are skipped automatically).
        for _k, _v in core_session.load(_WIZARD_KEY).items():
            if _k in self.vars and _k != "status" and not self._is_foreign_bundled_template(_k, _v):
                self.vars[_k].set(_v)
        self._configure_style()
        self._build_ui()
        self._show_step(0)

    def _is_foreign_bundled_template(self, key, saved_path):
        """True when a remembered template path is another suite version's
        bundled copy (.../templates/mls/<same file name>).  The session file is
        shared by every installed version, so without this rule a newer build
        would keep merging the previous version's templates.  A template the
        planner browsed to elsewhere (e.g. a City download) is still restored."""
        if not (key == "template" or key.endswith("_template")) or not saved_path:
            return False
        try:
            saved = Path(saved_path)
            default = Path(self.vars[key].get())
            if saved.resolve() == default.resolve():
                return False
            return (saved.parent.name.lower() == "mls"
                    and saved.parent.parent.name.lower() == "templates")
        except Exception:
            return False

    def _configure_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        self.colors = {
            "nav": "#0F172A",
            "nav2": "#1E293B",
            "blue": "#2563EB",
            "blue_dark": "#1D4ED8",
            "bg": "#F8FAFC",
            "card": "#FFFFFF",
            "border": "#CBD5E1",
            "muted": "#64748B",
            "text": "#0F172A",
            "green": "#16A34A",
            "amber": "#D97706",
        }
        self.configure(bg=self.colors["bg"])
        style.configure("TFrame", background=self.colors["bg"])
        style.configure("Card.TFrame", background=self.colors["card"], relief="flat")
        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"), foreground=self.colors["text"], background=self.colors["bg"])
        style.configure("Hero.TLabel", font=("Segoe UI", 12), foreground="#334155", background=self.colors["bg"])
        style.configure("Section.TLabel", font=("Segoe UI", 13, "bold"), foreground=self.colors["text"], background=self.colors["card"])
        style.configure("Sub.TLabel", font=("Segoe UI", 9), foreground=self.colors["muted"], background=self.colors["bg"])
        style.configure("CardSub.TLabel", font=("Segoe UI", 9), foreground=self.colors["muted"], background=self.colors["card"])
        style.configure("Field.TLabel", font=("Segoe UI", 9, "bold"), foreground="#334155", background=self.colors["card"])
        style.configure("Value.TLabel", font=("Segoe UI", 9), foreground=self.colors["text"], background=self.colors["card"])
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 9))
        style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(12, 8))
        style.configure("Card.TLabelframe", background=self.colors["card"], bordercolor=self.colors["border"], relief="solid")
        style.configure("Card.TLabelframe.Label", font=("Segoe UI", 11, "bold"), foreground=self.colors["text"], background=self.colors["card"])

    def _build_ui(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        sidebar = tk.Frame(self, bg=self.colors["nav"], width=255)
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.grid_propagate(False)

        tk.Label(sidebar, text="MLS Planner\nWizard", bg=self.colors["nav"], fg="white", font=("Segoe UI", 20, "bold"), justify="left").pack(anchor="nw", padx=22, pady=(24, 8))
        tk.Label(sidebar, text="Document automation for RCA packets", bg=self.colors["nav"], fg="#CBD5E1", font=("Segoe UI", 9), justify="left", wraplength=200).pack(anchor="nw", padx=22, pady=(0, 24))

        self.step_buttons = []
        for idx, label in enumerate(STEPS):
            b = tk.Button(
                sidebar,
                text=label,
                anchor="w",
                bg=self.colors["nav2"],
                fg="white",
                activebackground=self.colors["blue_dark"],
                activeforeground="white",
                relief="flat",
                font=("Segoe UI", 10, "bold" if idx == 0 else "normal"),
                padx=14,
                pady=12,
                command=lambda i=idx: self._show_step(i),
            )
            b.pack(fill="x", padx=12, pady=5)
            self.step_buttons.append(b)

        modules = "Built modules\n• Notifications\n• Applicant letter\n• Approvals\n• RCA forms\n• RCA packet PDF\n• Draft ordinance\n• HCAD record finder"
        tk.Label(sidebar, text=modules, bg=self.colors["nav"], fg="#BFDBFE", font=("Segoe UI", 9), justify="left").pack(side="bottom", anchor="sw", padx=22, pady=22)

        main = ttk.Frame(self, padding=(22, 18))
        main.grid(row=0, column=1, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        header = ttk.Frame(main)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="MLS/RCA Document Generator", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Load one calculation workbook, review the Case Info, then generate the exact City templates and packet attachments.", style="Hero.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Button(header, text="Load Workbook", style="Accent.TButton", command=self._load_summary).grid(row=0, column=1, rowspan=2, padx=(14, 0), sticky="e")

        _holder = ScrollableContent(main, bg=self.colors["bg"])
        _holder.grid(row=1, column=0, sticky="nsew")
        self.content = _holder.interior
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)

        status = ttk.LabelFrame(main, text="Status", style="Card.TLabelframe")
        status.grid(row=2, column=0, sticky="ew", pady=(14, 0))
        ttk.Label(status, textvariable=self.vars["status"], wraplength=930, justify="left", style="CardSub.TLabel").pack(anchor="w", padx=14, pady=10)

    def _clear_content(self):
        for child in self.content.winfo_children():
            child.destroy()

    def _show_step(self, idx: int):
        self.current_step = idx
        for i, b in enumerate(self.step_buttons):
            b.configure(bg="#1e40af" if i == idx else "#123b82")
        self._clear_content()
        if idx == 0:
            self._step_files()
        elif idx == 1:
            self._step_case_info()
        elif idx == 2:
            self._step_owners()
        elif idx == 3:
            self._step_generate()
        else:
            from core.peer_review_ui import reviewer_page
            reviewer_page(self, self.content)

    def _row(self, parent, label, key, row, filetypes=None, folder=False):
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=0, sticky="w", padx=12, pady=8)
        entry = ttk.Entry(parent, textvariable=self.vars[key], width=70)
        entry.grid(row=row, column=1, sticky="ew", padx=8, pady=8)
        ttk.Button(parent, text="Browse", style="Secondary.TButton", command=lambda: self._browse(key, filetypes, folder)).grid(row=row, column=2, padx=12, pady=8)

    def _multi_file_row(self, parent, label, key, row, filetypes=None):
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=0, sticky="w", padx=12, pady=8)
        ttk.Entry(parent, textvariable=self.vars[key], width=70).grid(row=row, column=1, sticky="ew", padx=8, pady=8)
        ttk.Button(parent, text="Browse Multiple", style="Secondary.TButton", command=lambda: self._browse_multiple(key, filetypes)).grid(row=row, column=2, padx=12, pady=8)

    def _section_card(self, parent, title: str, row: int, col: int = 0, colspan: int = 1, pady=8):
        card = ttk.LabelFrame(parent, text=title, style="Card.TLabelframe")
        card.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=8, pady=pady)
        card.columnconfigure(1, weight=1)
        return card

    def _info_pair(self, parent, label: str, value: str, row: int, col: int = 0, wrap=300):
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=col, sticky="nw", padx=10, pady=4)
        ttk.Label(parent, text=value or "—", style="Value.TLabel", wraplength=wrap, justify="left").grid(row=row, column=col+1, sticky="w", padx=10, pady=4)

    def _action_card(self, parent, title: str, description: str, buttons: list[tuple[str, object]], row: int, col: int, colspan: int = 1):
        card = self._section_card(parent, title, row, col, colspan, pady=5)
        ttk.Label(card, text=description, style="CardSub.TLabel", wraplength=430, justify="left").grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 5))
        btnrow = ttk.Frame(card)
        btnrow.grid(row=1, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 8))
        for i, (text, command) in enumerate(buttons):
            ttk.Button(btnrow, text=text, style="Accent.TButton", command=command).pack(side="left", padx=(0 if i == 0 else 8, 0))
        return card

    def _hcad_count_text(self):
        paths = self._split_multi_paths(self.vars["hcad_records"].get())
        if not paths:
            return "No HCAD records selected."
        return f"{len(paths)} HCAD record PDF(s) selected. They will be merged in the order selected, with each record condensed to one page."

    def _step_files(self):
        outer = ttk.Frame(self.content)
        outer.grid(sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.columnconfigure(1, weight=1)

        ttk.Label(outer, text="Step 1: Select the files for this case", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 6))
        ttk.Label(outer, text="Only select what the wizard cannot read from the calculation sheet. Template files stay external so they can be replaced without changing code.", style="Sub.TLabel", wraplength=920).grid(row=1, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 10))

        required = self._section_card(outer, "Required setup", 2, 0, 2)
        self._row(required, "Calculation workbook", "excel", 0, [("Excel files", "*.xlsx")])
        self._row(required, "Boundary map", "map", 1, [("Map files", "*.pdf *.png *.jpg *.jpeg")])
        self._row(required, "Output folder", "output_folder", 2, folder=True)
        self._add_boundary_map_controls(required)
        self._add_application_import_controls(required)

        templates = self._section_card(outer, "City templates (2026)", 3, 0)
        self._row(templates, "Owner notification letter", "template", 0, [("Word documents", "*.docx")])
        self._row(templates, "Protest letter to owners", "protest_template", 1, [("Word documents", "*.docx")])
        self._row(templates, "Public hearing notice (<51%)", "hearing_template", 2, [("Word documents", "*.docx")])
        self._row(templates, "Applicant sign letter", "applicant_template", 3, [("Word documents", "*.docx")])
        self._row(templates, "Director approval", "director_template", 4, [("Word documents", "*.docx")])
        self._row(templates, "Planning Commission approval", "pc_template", 5, [("Word documents", "*.docx")])
        self._row(templates, "MLS routing form", "routing_template", 6, [("Word documents", "*.docx")])
        self._row(templates, "OneMeeting staff report", "staff_report_template", 7, [("Word documents", "*.docx")])
        self._row(templates, "Draft ordinance", "ordinance_template", 8, [("Word documents", "*.docx")])

        attachments = self._section_card(outer, "Packet attachments", 3, 1)
        ttk.Label(attachments, text="The 2026 routing form lists four attachments: routing form, approval, map of the area and tax map. Only the tax map must be uploaded; plat, HCAD records and the Step 2 staff analysis are optional extras.", style="CardSub.TLabel", wraplength=430).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(6, 4))
        self._row(attachments, "Tax map (required)", "tax_map", 1, [("Attachment files", "*.pdf *.png *.jpg *.jpeg")])
        self._row(attachments, "Subdivision plat (optional)", "subdivision_plat", 2, [("Attachment files", "*.pdf *.png *.jpg *.jpeg")])
        self._multi_file_row(attachments, "HCAD records (optional)", "hcad_records", 3, [("PDF files", "*.pdf")])
        self._staff_var = tk.BooleanVar(value=str(self.vars["include_staff_analysis"].get()).lower() in ("yes", "true", "1"))
        ttk.Checkbutton(attachments, text="Also include the Step 2 staff analysis export in the packet", variable=self._staff_var,
                        command=lambda: self.vars["include_staff_analysis"].set("yes" if self._staff_var.get() else "no")).grid(row=4, column=0, columnspan=3, sticky="w", padx=12, pady=(2, 4))
        hcad_buttons = ttk.Frame(attachments)
        hcad_buttons.grid(row=5, column=1, sticky="w", padx=8, pady=(0, 6))
        ttk.Button(hcad_buttons, text="Download HCAD Records", style="Accent.TButton", command=self._download_hcad_records).pack(side="left", padx=(0, 6))
        ttk.Button(hcad_buttons, text="Open HCAD Folder", style="Secondary.TButton", command=self._open_hcad_folder).pack(side="left")
        ttk.Label(attachments, text=self._hcad_count_text(), style="CardSub.TLabel").grid(row=6, column=1, sticky="w", padx=8, pady=(0, 6))
        ttk.Label(attachments, text="HCAD workflow: after the workbook is loaded, click Download HCAD Records. The wizard will read Step 1, open Google Chrome, wait for you to complete HCAD human verification, then download/import the HCAD PDFs for the RCA packet.", style="CardSub.TLabel", wraplength=430).grid(row=7, column=0, columnspan=3, sticky="w", padx=12, pady=(4, 10))

        footer = ttk.Frame(outer)
        footer.grid(row=4, column=0, columnspan=2, sticky="ew", padx=8, pady=(8, 0))
        ttk.Button(footer, text="Load & Validate Workbook", style="Accent.TButton", command=self._load_summary).pack(side="left")
        ttk.Button(footer, text="Next: Review Case Info", style="Secondary.TButton", command=lambda: self._show_step(1)).pack(side="left", padx=8)

    def _step_case_info(self):
        frame = ttk.Frame(self.content)
        frame.grid(sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        ttk.Label(frame, text="Step 2: Review Case Info", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 6))
        ttk.Label(frame, text="This screen shows what the wizard will use. Yellow workbook fields are planner-entered; GIS labels stay unchanged on Step 1.", style="Sub.TLabel", wraplength=920).grid(row=1, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 10))
        top = ttk.Frame(frame)
        top.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        ttk.Button(top, text="Refresh from Workbook", style="Accent.TButton", command=self._load_summary).pack(side="left")
        ttk.Button(top, text="Next: Review Owners", style="Secondary.TButton", command=lambda: self._show_step(2)).pack(side="left", padx=8)

        if not self.summary:
            card = self._section_card(frame, "Workbook not loaded", 3, 0, 2)
            ttk.Label(card, text="Select a calculation workbook, then click Load & Validate Workbook.", style="CardSub.TLabel").pack(anchor="w", padx=12, pady=12)
            return

        ci = self.summary.case_info
        required = self._section_card(frame, "Required case fields", 3, 0)
        for i, f in enumerate(["Letter_Date", "Protest_Deadline", "Planner_Name", "Planner_Phone", "Planner_Email", "Approval_Path", "Council_Deadline"]):
            self._info_pair(required, f, ci.get(f, ""), i, wrap=300)

        auto = self._section_card(frame, "Auto-read from Step 2", 3, 1)
        for i, f in enumerate(["Application_Number", "Application_Type", "Block_Description", "Side_Of_Street", "Cross_Streets", "Minimum_Lot_Size"]):
            self._info_pair(auto, f, ci.get(f, ""), i, wrap=300)

        applicant = self._section_card(frame, "Applicant sign-letter fields", 4, 0)
        for i, f in enumerate(["Applicant_Name", "Applicant_Address_1", "Applicant_Address_2", "Applicant_City", "Applicant_State", "Applicant_Zip", "Sign_Count_Number", "Sign_Placement"]):
            self._info_pair(applicant, f, ci.get(f, ""), i, wrap=300)

        approvals = self._section_card(frame, "Approval narrative fields", 4, 1)
        for i, f in enumerate(["Director_Name", "PC_Signer", "Subdivision_Plat_Year", "Neighborhood_Age_Statement", "Structure_Age_Statement", "Preservation_Finding"]):
            self._info_pair(approvals, f, ci.get(f, ""), i, wrap=300)

        subj = self._section_card(frame, "Subject line preview", 5, 0, 2)
        ttk.Label(subj, text=self.summary.subject_line or "—", style="Value.TLabel", wraplength=780, justify="left").pack(anchor="w", padx=12, pady=12)

        pf = self._section_card(frame, "Preflight checks (informational — nothing here blocks generation)", 6, 0, 2)
        results = run_preflight(self.vars["excel"].get(), _WIZARD_KEY, self.vars["map"].get(), self.vars["output_folder"].get())
        build_preflight_panel(pf, results, card_bg=self.colors["card"], text_color=self.colors["text"]).pack(anchor="w", fill="x", padx=12, pady=(4, 12))

    def _step_owners(self):
        frame = ttk.Frame(self.content)
        frame.grid(sticky="nsew")
        frame.rowconfigure(3, weight=1)
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text="Step 3: Review Owner Mail-Out Data", style="Section.TLabel").grid(row=0, column=0, sticky="w", padx=8, pady=(0, 6))
        ttk.Label(frame, text="These records come directly from Step 1 GIS data. The wizard does not rename or alter the Step 1 columns.", style="Sub.TLabel", wraplength=900).grid(row=1, column=0, sticky="w", padx=8, pady=(0, 10))
        top = ttk.Frame(frame)
        top.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 8))
        ttk.Button(top, text="Refresh from Workbook", style="Accent.TButton", command=self._load_summary).pack(side="left")
        ttk.Button(top, text="Next: Generate", style="Secondary.TButton", command=lambda: self._show_step(3)).pack(side="left", padx=8)

        if not self.summary:
            card = self._section_card(frame, "Workbook not loaded", 3, 0)
            ttk.Label(card, text="Load a workbook first.", style="CardSub.TLabel").pack(anchor="w", padx=12, pady=12)
            return
        ttk.Label(frame, text=f"Valid owner records found: {len(self.summary.owners)}", font=("Segoe UI", 11, "bold"), foreground=self.colors["green"], background=self.colors["bg"]).grid(row=3, column=0, sticky="w", padx=8, pady=(0, 6))
        columns = ("Excel Row", "Owner", "Mailing Address", "Land Area", "Land Use")
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=18)
        for c in columns:
            tree.heading(c, text=c)
        tree.column("Excel Row", width=80, anchor="center")
        tree.column("Owner", width=230)
        tree.column("Mailing Address", width=360)
        tree.column("Land Area", width=95, anchor="center")
        tree.column("Land Use", width=190)
        tree.grid(row=4, column=0, sticky="nsew", padx=8)
        scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        scroll.grid(row=4, column=1, sticky="ns")
        tree.configure(yscrollcommand=scroll.set)
        for owner in self.summary.owners:
            addr = f"{owner.get('Mail_Addr_1','')} {owner.get('Mail_Addr_2','')}, {owner.get('Mail_City','')}, {owner.get('Mail_State','')} {owner.get('Mail_Zip','')}"
            tree.insert("", "end", values=(owner.get("Excel_Row"), owner.get("Mailto"), addr, owner.get("Total_Land_Area"), owner.get("group_dscr")))

    def _step_generate(self):
        frame = ttk.Frame(self.content)
        frame.grid(sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Step 4: Generate Documents", style="Section.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 6))
        ttk.Label(frame, text="Pick the document you need. The wizard preserves the uploaded City templates and only fills the case-specific fields.", style="Sub.TLabel", wraplength=940).grid(row=1, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 10))

        qa = self._section_card(frame, "Quality checks", 2, 0, 2)
        warnings_box = tk.Text(qa, height=3, wrap="word", bg="#F8FAFC", relief="flat", font=("Segoe UI", 9))
        warnings_box.pack(fill="x", padx=12, pady=10)
        if self.summary:
            all_warnings = list(self.summary.warnings) + validate_applicant_inputs(self.summary.case_info)
            checks = [
                ("Director Approval", validate_director_inputs, build_director_context),
                ("Routing Form", validate_routing_inputs, build_routing_context),
                ("Planning Commission Approval", validate_pc_inputs, build_pc_context),
                ("Draft Ordinance", validate_ordinance_inputs, build_ordinance_context),
            ]
            for label, validator, builder in checks:
                try:
                    all_warnings += validator(builder(self._wb()))
                except Exception as e:
                    all_warnings += [f"{label} validation could not run: {e}"]
            if all_warnings:
                warnings_box.insert("end", "Warnings found:\n" + "\n".join(f"- {w}" for w in all_warnings))
            else:
                warnings_box.insert("end", "No warnings found. Ready to generate outputs.")
        else:
            warnings_box.insert("end", "Load and validate a workbook first.")
        warnings_box.configure(state="disabled")

        self._action_card(frame, "Owner Notification Mail-Out", "Create a one-owner preview or one combined Word document with all owner letters and the selected map after each letter (2026 letter with HCAD legal description and land-use confirmation).", [("Preview One Notification Letter", lambda: self._run(preview=True)), ("Generate All Notification Letters", lambda: self._run(preview=False))], 3, 0)
        self._action_card(frame, "Protest / Public Hearing Letters", "Protest filed: the protest letter to all owners (uses Protest_Letter_Date and PC_Hearing_Date). Support under 51%: the notice of public hearing letter to all owners (uses Letter_Date, Support_Percentage and PC_Hearing_Date).", [("Preview Protest Letter", lambda: self._run_protest(preview=True)), ("Generate All Protest Letters", lambda: self._run_protest(preview=False)), ("Preview Hearing Notice", lambda: self._run_hearing(preview=True)), ("Generate All Hearing Notices", lambda: self._run_hearing(preview=False))], 3, 1)
        self._action_card(frame, "Applicant Letter", "Generate the applicant complete/sign-posting letter. Blank Applicant_Address_2 lines are suppressed automatically.", [("Generate Applicant Sign Letter", self._run_applicant)], 4, 0)
        self._action_card(frame, "Approvals", "Generate the approval document that matches the selected approval path. These use the exact uploaded City templates.", [("Generate Director Approval", self._run_director), ("Generate PC Approval", self._run_pc)], 4, 1)
        self._action_card(frame, "RCA Packet Forms", "Routing form (2026: checks Lot Size and SMLS Block Application, fills the council district, lot size and Council_Deadline) and the OneMeeting staff report for the Director path.", [("Generate Routing Form", self._run_routing), ("Generate OneMeeting Staff Report", self._run_staff_report)], 5, 0)
        self._action_card(frame, "Draft Ordinance", "Create the planner draft ordinance using the 2026 City ordinance template, workbook facts, Exhibit A property descriptions, Exhibit B map, and Exhibit C property list.", [("Generate Draft Ordinance", self._run_ordinance)], 5, 1)
        self._action_card(frame, "RCA Packet Builder", "Assemble the RCA packet PDF in the 2026 order: routing form, Director or PC approval, boundary map, tax map (plus plat, Step 2 staff analysis and HCAD records when selected).", [("Generate RCA Packet PDF", self._run_packet)], 6, 0, 2)

        self._action_card(frame, "Planning Commission Draft Report", "Review the MLS report wording and calculation figures. Create the separate PC map showing land use, support, protests and lots below the proposed minimum, then assemble an editable Word draft and PDF with attachments.", [("Generate PC Map", lambda: self._open_pc_draft(start_map=True)), ("Generate PC Draft Report", self._open_pc_draft), ("PC Speech", self._open_pc_speech)], 7, 0, 2)
        ttk.Label(frame, text="Export rules: Step 2 staff analysis exports only the meaningful staff-analysis range. HCAD records can be selected as multiple PDFs. The packet builder does not add blank pages after the final item.", style="Sub.TLabel", wraplength=930).grid(row=8, column=0, columnspan=2, sticky="w", padx=8, pady=(10, 0))

    def _open_pc_speech(self):
        try:
            from .pc_speech_ui import open_pc_speech
            self._suggest_case_folder()
            open_pc_speech(self)
        except Exception as error:
            messagebox.showerror('PC Speech', str(error), parent=self)

    def _open_pc_draft(self, start_map=False):
        try:
            from .pc_ui import open_pc_workspace
            self._suggest_case_folder()
            open_pc_workspace(self, start_map=start_map)
        except Exception as error:
            messagebox.showerror('MLS Planning Commission', str(error), parent=self)

    def _browse(self, key, filetypes=None, folder=False):
        if folder:
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=filetypes or [("All files", "*.*")])
        if path:
            self.vars[key].set(path)

    def _browse_multiple(self, key, filetypes=None):
        paths = filedialog.askopenfilenames(filetypes=filetypes or [("All files", "*.*")])
        if paths:
            # Store multiple paths in one text field using a delimiter unlikely to appear in Windows paths.
            self.vars[key].set("||".join(paths))
            self.vars["status"].set(f"Selected {len(paths)} HCAD record PDF(s).")
            if self.current_step == 0:
                self._show_step(0)

    def _split_multi_paths(self, value: str):
        if not value:
            return []
        if "||" in value:
            return [p for p in value.split("||") if p]
        return [value]



    def _import_application(self):
        from core.application_import_ui import open_application_import
        open_application_import(self)

    def _hcad_folder(self) -> Path:
        """Return the case output folder used for downloaded/imported HCAD PDFs."""
        base = Path(self.vars["output_folder"].get() or (Path.home() / "Documents"))
        return base / "HCAD_Records"

    def _chrome_debug_profile(self) -> Path:
        """Dedicated real-Chrome profile for HCAD CDP connection mode.

        This is not a Playwright-owned browser profile. The wizard starts a normal
        Chrome window with a debugging port so the downloader can attach after the
        planner has completed HCAD's human verification.
        """
        return self._hcad_folder() / "Chrome_Debug_Profile"

    def _chrome_debug_url(self) -> str:
        return "http://127.0.0.1:9222"

    def _is_chrome_cdp_available(self) -> bool:
        try:
            with urllib.request.urlopen(self._chrome_debug_url() + "/json/version", timeout=2) as resp:
                return resp.status == 200
        except Exception:
            return False

    def _find_chrome_executable(self) -> str:
        """Best-effort Chrome lookup for Windows and common non-Windows installs."""
        candidates = []
        local = os.environ.get("LOCALAPPDATA")
        pf = os.environ.get("PROGRAMFILES")
        pf86 = os.environ.get("PROGRAMFILES(X86)")
        if local:
            candidates.append(Path(local) / "Google" / "Chrome" / "Application" / "chrome.exe")
        if pf:
            candidates.append(Path(pf) / "Google" / "Chrome" / "Application" / "chrome.exe")
        if pf86:
            candidates.append(Path(pf86) / "Google" / "Chrome" / "Application" / "chrome.exe")
        for c in candidates:
            if c.exists():
                return str(c)
        # Let the OS PATH resolve it if possible.
        return "chrome"

    def _export_hcad_list(self, show_message: bool = True):
        """Export Step 1 site addresses/TAX_ID values to a text file for HCAD."""
        try:
            self._validate_paths(
                require_notification_template=False,
                require_applicant_template=False,
                require_director_template=False,
                require_map=False,
            )
            address_file, count = export_hcad_address_list(
                self._wb(),
                self.vars["output_folder"].get(),
            )
            self.vars["status"].set(f"Prepared {count} unique HCAD property record(s) from Step 1.")
            if show_message:
                messagebox.showinfo(APP_TITLE, f"Prepared {count} unique HCAD property record(s). One record is used for all lot rows sharing an HCAD account.\n\n{address_file}")
            return address_file, count
        except Exception as e:
            self.vars["status"].set(f"HCAD address export error: {e}")
            messagebox.showerror(APP_TITLE, str(e))
            return None, 0

    def _open_hcad_in_chrome(self):
        """Open HCAD in the user's default Chrome/browser for manual verification."""
        try:
            webbrowser.open("https://search.hcad.org/", new=2)
            self.vars["status"].set("Opened HCAD in your default browser. Complete verification there, then use Connect & Download or Import Downloaded PDFs.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Could not open HCAD in browser:\n{e}")

    def _start_hcad_chrome_session(self, show_message: bool = True) -> bool:
        """Start Google Chrome with the wizard's connection port and open HCAD.

        The planner sees this only as part of the single Download HCAD Records
        workflow. Chrome is launched explicitly instead of using the default
        browser, so the button does not accidentally open Edge.
        """
        hcad_folder = self._hcad_folder()
        hcad_folder.mkdir(parents=True, exist_ok=True)
        profile = self._chrome_debug_profile()
        profile.mkdir(parents=True, exist_ok=True)
        chrome = self._find_chrome_executable()
        url = "https://search.hcad.org/"
        try:
            subprocess.Popen([
                chrome,
                "--remote-debugging-port=9222",
                f"--user-data-dir={profile}",
                "--new-window",
                url,
            ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

            # Give Chrome a moment to expose the debugging endpoint.
            for _ in range(20):
                if self._is_chrome_cdp_available():
                    break
                time.sleep(0.25)

            self.vars["status"].set("Opened HCAD in Google Chrome. Complete human verification, then return to the wizard and click OK on the prompt.")
            if show_message:
                messagebox.showinfo(
                    APP_TITLE,
                    "Google Chrome was opened for HCAD.\n\n"
                    "1. Complete the HCAD human verification in that Chrome window.\n"
                    "2. Wait until the HCAD property search page is visible.\n"
                    "3. Return to this wizard and click OK/Continue.\n\n"
                    "The wizard will then continue the download automatically."
                )
            return True
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Could not start Google Chrome for HCAD:\n{e}\n\nMake sure Google Chrome is installed.")
            self.vars["status"].set(f"Could not start Google Chrome for HCAD: {e}")
            return False

    def _import_hcad_pdfs(self):
        """Scan the HCAD folder and select downloaded PDFs for the RCA packet."""
        folder = self._hcad_folder()
        folder.mkdir(parents=True, exist_ok=True)
        pdfs = sorted(str(p) for p in folder.glob("*.pdf") if p.is_file())
        if not pdfs:
            messagebox.showinfo(APP_TITLE, f"No HCAD PDFs were found in:\n{folder}\n\nSave or move downloaded HCAD PDFs into this folder, then click Import Downloaded PDFs again.")
            self.vars["status"].set(f"No HCAD PDFs found in {folder}.")
            return []
        self.vars["hcad_records"].set("||".join(pdfs))
        msg = f"Imported {len(pdfs)} HCAD PDF(s) from {folder}."
        self.vars["status"].set(msg)
        messagebox.showinfo(APP_TITLE, msg)
        if self.current_step == 0:
            self._show_step(0)
        return pdfs

    def _download_hcad_records(self):
        """One-button HCAD workflow for planners.

        The wizard reads Step 1, opens Google Chrome explicitly for HCAD human
        verification, waits for planner confirmation, runs the connected downloader,
        imports any resulting PDFs, and selects them for the RCA packet.
        """
        address_file, count = self._export_hcad_list(show_message=False)
        if not address_file:
            return

        hcad_folder = self._hcad_folder()
        hcad_folder.mkdir(parents=True, exist_ok=True)
        script_path = HCAD_DOWNLOADER
        if not core_launch.hcad_available():
            messagebox.showerror(APP_TITLE, f"Could not find bundled HCAD downloader:\n{script_path}")
            return

        if count <= 0:
            messagebox.showwarning(APP_TITLE, "No HCAD addresses or TAX_ID values were found in Step 1.")
            self.vars["status"].set("No HCAD records found to download.")
            return

        # Start Chrome automatically if the connection is not already available.
        if not self._is_chrome_cdp_available():
            if not self._start_hcad_chrome_session(show_message=False):
                return

        messagebox.showinfo(
            APP_TITLE,
            f"HCAD will open in Google Chrome for {count} record(s).\n\n"
            "Complete the human verification in Chrome if prompted.\n"
            "Do not close the Chrome window.\n\n"
            "When the HCAD search page is visible, click OK here and the wizard will continue downloading the HCAD PDFs."
        )

        if not self._is_chrome_cdp_available():
            messagebox.showwarning(
                APP_TITLE,
                "The wizard still cannot connect to Chrome.\n\n"
                "Make sure the Chrome window opened by the wizard is still open. If Chrome did not open, install Google Chrome or open the HCAD records manually and place the PDFs in the HCAD_Records folder."
            )
            self.vars["status"].set("Chrome connection was not available. HCAD PDFs can still be added manually.")
            return

        self.vars["status"].set(f"Downloading {count} unique HCAD property record(s). Complete HCAD verification in Chrome if prompted.")

        def worker():
            try:
                result = subprocess.run(
                    core_launch.hcad_command() + [
                        "--file",
                        str(address_file),
                        "--output",
                        str(hcad_folder),
                        "--cdp-url",
                        self._chrome_debug_url(),
                        "--delay-min",
                        "3",
                        "--delay-max",
                        "6",
                    ],
                    cwd=core_launch.hcad_cwd(),
                    creationflags=core_launch.console_flags(),
                    text=True,
                )
                pdfs = sorted(str(p) for p in hcad_folder.glob("*.pdf") if p.is_file())
                report = hcad_folder / "HCAD_Download_Report.csv"
                selected_pdfs = "||".join(pdfs) if pdfs else ""

                if pdfs:
                    msg = f"HCAD download finished. {len(pdfs)} PDF(s) found in {hcad_folder}."
                    if report.exists():
                        msg += f"\n\nDownload report: {report}"
                else:
                    msg = (
                        f"HCAD download finished, but no PDFs were found in {hcad_folder}.\n\n"
                        "If HCAD blocked the automated download, save the HCAD PDFs manually into that folder. "
                        "The RCA packet builder can still use any PDFs placed there."
                    )
                    if report.exists():
                        msg += f"\n\nDownload report: {report}"

                if result.returncode != 0:
                    msg += f"\n\nDownloader exit code: {result.returncode}."

                def finish_hcad_download():
                    if selected_pdfs:
                        self.vars["hcad_records"].set(selected_pdfs)
                    self.vars["status"].set(msg)
                    messagebox.showinfo(APP_TITLE, msg)
                    self._show_step(self.current_step)

                self.after(0, finish_hcad_download)
            except Exception as e:
                self.after(0, lambda: self.vars["status"].set(f"HCAD download error: {e}"))
                self.after(0, lambda: messagebox.showerror(APP_TITLE, str(e)))

        threading.Thread(target=worker, daemon=True).start()

    def _open_hcad_folder(self):
        """Open the folder that stores downloaded/imported HCAD PDFs."""
        folder = self._hcad_folder()
        folder.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Could not open HCAD folder:\n{e}")

    def _validate_paths(self, require_notification_template=True, require_applicant_template=True, require_director_template=True, require_routing_template=False, require_pc_template=False, require_ordinance_template=False, require_map=True, require_packet_attachments=False):
        required = [("excel", "calculation workbook"), ("output_folder", "output folder")]
        if require_notification_template:
            required.append(("template", "notification template"))
        if require_applicant_template:
            required.append(("applicant_template", "applicant sign-letter template"))
        if require_director_template:
            required.append(("director_template", "Director approval template"))
        if require_routing_template:
            required.append(("routing_template", "MLS routing form template"))
        if require_pc_template:
            required.append(("pc_template", "Planning Commission approval template"))
        if require_ordinance_template:
            required.append(("ordinance_template", "draft ordinance template"))
        if require_map:
            required.append(("map", "boundary map"))
            self._check_boundary_map()
        if require_packet_attachments:
            required.append(("tax_map", "tax map"))

        for key, label in required:
            if not self.vars[key].get():
                raise ValueError(f"Please select a {label}.")

        for key, label in required:
            if key == "output_folder":
                continue
            if key == "hcad_records":
                paths = self._split_multi_paths(self.vars[key].get())
                if not paths:
                    raise ValueError("Please select at least one HCAD record PDF.")
                missing = [p for p in paths if not Path(p).exists()]
                if missing:
                    raise ValueError("One or more selected HCAD record PDFs do not exist:\n" + "\n".join(missing))
                continue
            if not Path(self.vars[key].get()).exists():
                raise ValueError(f"The selected {label} does not exist.")
        Path(self.vars["output_folder"].get()).mkdir(parents=True, exist_ok=True)

    def _wb(self):
        """Read the latest saved workbook; invalidate older recalculated copies."""
        if not hasattr(self, "_workbook_read_cache"):
            self._workbook_read_cache = WorkbookReadCache()
        from core.automatic_case import prepare_owner_calculation
        return self._workbook_read_cache.resolve(prepare_owner_calculation(self))

    def _save_session(self):
        try:
            core_session.save(_WIZARD_KEY, {k: self.vars[k].get() for k in _SESSION_KEYS})
        except Exception:
            pass

    def _suggest_case_folder(self):
        """Roll the output folder forward to Documents/RCA_Cases/<case> when
        the current value is the plain default or a previous case folder.
        A custom folder the planner picked is never overridden."""
        try:
            new_out = maybe_case_folder(self.vars["output_folder"].get(), self.vars["excel"].get())
            if new_out:
                new_out.mkdir(parents=True, exist_ok=True)
                self.vars["output_folder"].set(str(new_out))
                return new_out
        except Exception:
            pass
        return None

    def _load_summary(self):
        try:
            self._validate_paths(require_notification_template=False, require_applicant_template=False, require_director_template=False, require_map=False)
            self.summary = summarize_workbook(self._wb(), self.vars["map"].get() or None)
            msg = f"Loaded workbook. Found {len(self.summary.owners)} valid owner record(s)."
            if self.summary.warnings:
                msg += f" {len(self.summary.warnings)} warning(s) found."
            new_out = self._suggest_case_folder()
            if new_out:
                msg += f" Output folder set to case folder: {new_out}"
            self.vars["status"].set(msg)
            self._save_session()
            self._show_step(self.current_step)
        except Exception as e:
            self.vars["status"].set(f"Error: {e}")
            messagebox.showerror(APP_TITLE, str(e))

    def _run(self, preview: bool):
        try:
            self._validate_paths(require_applicant_template=False, require_director_template=False, require_map=True)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        mode_name = "Preview_One_Letter" if preview else "All_Owner_Notification_Letters"
        self._run_owner_letters(self.vars["template"].get(), mode_name, preview)

    def _run_owner_letters(self, template, mode_name, preview):
        output = Path(self.vars["output_folder"].get()) / f"{mode_name}.docx"
        self.vars["status"].set(f"Working... creating {output.name}")
        def worker():
            try:
                count = generate_owner_letters(
                    self._wb(),
                    template,
                    self.vars["map"].get(),
                    output,
                    preview=preview,
                )
                self.after(0, lambda: self._done(output, count))
            except Exception as e:
                self.after(0, lambda error=e: self._error(error))
        threading.Thread(target=worker, daemon=True).start()

    def _run_protest(self, preview: bool):
        try:
            self._validate_paths(require_notification_template=False, require_applicant_template=False, require_director_template=False, require_map=True)
            if not Path(self.vars["protest_template"].get()).exists():
                raise ValueError("Please select the protest letter template.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        self._run_owner_letters(self.vars["protest_template"].get(), "Preview_Protest_Letter" if preview else "All_Owner_Protest_Letters", preview)

    def _run_hearing(self, preview: bool):
        try:
            self._validate_paths(require_notification_template=False, require_applicant_template=False, require_director_template=False, require_map=True)
            if not Path(self.vars["hearing_template"].get()).exists():
                raise ValueError("Please select the public hearing notice template.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        self._run_owner_letters(self.vars["hearing_template"].get(), "Preview_Public_Hearing_Notice" if preview else "All_Owner_Public_Hearing_Notices", preview)

    def _run_staff_report(self):
        try:
            self._validate_paths(require_notification_template=False, require_applicant_template=False, require_director_template=False, require_map=False)
            if not Path(self.vars["staff_report_template"].get()).exists():
                raise ValueError("Please select the OneMeeting staff report template.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        output = Path(self.vars["output_folder"].get()) / "MLS_OneMeeting_Staff_Report.docx"
        self.vars["status"].set(f"Working... creating {output.name}")
        def worker():
            try:
                count = generate_staff_report(self._wb(), self.vars["staff_report_template"].get(), output)
                self.after(0, lambda: self._done(output, count))
            except Exception as e:
                self.after(0, lambda error=e: self._error(error))
        threading.Thread(target=worker, daemon=True).start()

    def _run_applicant(self):
        try:
            self._validate_paths(require_notification_template=False, require_applicant_template=True, require_director_template=False, require_map=False)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        output = Path(self.vars["output_folder"].get()) / "Applicant_Complete_Sign_Letter.docx"
        self.vars["status"].set(f"Working... creating {output.name}")
        def worker():
            try:
                count = generate_applicant_letter(
                    self._wb(),
                    self.vars["applicant_template"].get(),
                    output,
                )
                self.after(0, lambda: self._done(output, count))
            except Exception as e:
                self.after(0, lambda error=e: self._error(error))
        threading.Thread(target=worker, daemon=True).start()

    def _run_director(self):
        try:
            self._validate_paths(require_notification_template=False, require_applicant_template=False, require_director_template=True, require_map=False)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        output = Path(self.vars["output_folder"].get()) / "Director_Approval.docx"
        self.vars["status"].set(f"Working... creating {output.name}")
        def worker():
            try:
                count = generate_director_approval(
                    self._wb(),
                    self.vars["director_template"].get(),
                    output,
                )
                self.after(0, lambda: self._done(output, count))
            except Exception as e:
                self.after(0, lambda error=e: self._error(error))
        threading.Thread(target=worker, daemon=True).start()

    def _run_routing(self):
        try:
            self._validate_paths(require_notification_template=False, require_applicant_template=False, require_director_template=False, require_routing_template=True, require_map=False)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        output = Path(self.vars["output_folder"].get()) / "MLS_Routing_Form.docx"
        self.vars["status"].set(f"Working... creating {output.name}")
        def worker():
            try:
                count = generate_routing_form(
                    self._wb(),
                    self.vars["routing_template"].get(),
                    output,
                )
                self.after(0, lambda: self._done(output, count))
            except Exception as e:
                self.after(0, lambda error=e: self._error(error))
        threading.Thread(target=worker, daemon=True).start()


    def _run_pc(self):
        try:
            self._validate_paths(require_notification_template=False, require_applicant_template=False, require_director_template=False, require_routing_template=False, require_pc_template=True, require_map=False)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        output = Path(self.vars["output_folder"].get()) / "Planning_Commission_Approval.docx"
        self.vars["status"].set(f"Working... creating {output.name}")
        def worker():
            try:
                count = generate_planning_commission_approval(
                    self._wb(),
                    self.vars["pc_template"].get(),
                    output,
                )
                self.after(0, lambda: self._done(output, count))
            except Exception as e:
                self.after(0, lambda error=e: self._error(error))
        threading.Thread(target=worker, daemon=True).start()

    def _run_ordinance(self):
        try:
            self._validate_paths(require_notification_template=False, require_applicant_template=False, require_director_template=False, require_routing_template=False, require_pc_template=False, require_ordinance_template=True, require_map=True)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        output = Path(self.vars["output_folder"].get()) / "Draft_Ordinance.docx"
        self.vars["status"].set(f"Working... creating {output.name}")
        def worker():
            try:
                count = generate_draft_ordinance(
                    self._wb(),
                    self.vars["ordinance_template"].get(),
                    self.vars["map"].get(),
                    output,
                )
                self.after(0, lambda: self._done(output, count))
            except Exception as e:
                self.after(0, lambda error=e: self._error(error))
        threading.Thread(target=worker, daemon=True).start()

    def _run_packet(self):
        try:
            self._validate_paths(require_notification_template=False, require_applicant_template=False, require_director_template=True, require_routing_template=True, require_pc_template=True, require_map=True, require_packet_attachments=True)
        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))
            return
        output = Path(self.vars["output_folder"].get()) / "RCA_Packet.pdf"
        self.vars["status"].set(f"Working... creating {output.name}")
        def worker():
            try:
                count = generate_rca_packet_pdf(
                    self._wb(),
                    {
                        "routing": self.vars["routing_template"].get(),
                        "director": self.vars["director_template"].get(),
                        "pc": self.vars["pc_template"].get(),
                    },
                    self.vars["map"].get(),
                    self.vars["tax_map"].get(),
                    output,
                    subdivision_plat_path=self.vars["subdivision_plat"].get(),
                    hcad_records=self._split_multi_paths(self.vars["hcad_records"].get()),
                    include_staff_analysis=str(self.vars["include_staff_analysis"].get()).lower() in ("yes", "true", "1"),
                )
                self.after(0, lambda: self._done(output, count))
            except Exception as e:
                self.after(0, lambda error=e: self._error(error))
        threading.Thread(target=worker, daemon=True).start()

    def _done(self, output: Path, count: int):
        self._save_session()
        self.vars["status"].set(f"Done. Generated {count} letter(s): {output}")
        answer = messagebox.askyesno(APP_TITLE, f"Generated {count} letter(s).\n\n{output}\n\nOpen the output folder?")
        if answer:
            self._open_folder(output.parent)

    def _error(self, e: Exception):
        self.vars["status"].set(f"Error: {e}")
        log_path = log_exception(APP_TITLE, self.vars["output_folder"].get(), e)
        detail = f"\n\nDetails were saved to:\n{log_path}" if log_path else ""
        messagebox.showerror(APP_TITLE, str(e) + detail)

    def _open_folder(self, folder: Path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", str(folder)])
            else:
                subprocess.run(["xdg-open", str(folder)])
        except Exception:
            pass

if __name__ == "__main__":
    print("This wizard now runs inside the RCA Wizard Suite.")
    print("Start it with Run Wizard.bat, or: python app.py --wizard mls")
