"""Template Manager screen for the RCA Wizard Suite.

Opened from the home screen.  Lets the wizard maintainer, without touching any
code:

  * see which .docx file each wizard uses for each document, and whether it
    looks healthy (file present, {{tokens}} recognised, no split-run tokens)
  * replace a template with a revised City file (the old one is archived in
    ``templates/<wizard>/_previous`` and can be restored with one click)
  * run the validation harness for one wizard, or all of them, and read the
    result here
  * generate the sample documents into a folder and open it, to eyeball the
    wording of templates that are filled by wording match rather than tokens

All file bookkeeping lives in :mod:`core.templates`; this module is UI only.
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import launch as core_launch
from . import templates as tpl
from .office_convert import describe_engine
from .paths import TEMPLATES_ROOT, templates_dir

_BG = "#F1F5F9"
_CARD = "#FFFFFF"
_TEXT = "#0F172A"
_MUTED = "#64748B"
_ACCENT = "#334155"
_ACCENT_MID = "#475569"
_OK = "#166534"
_WARN = "#B45309"
_ERR = "#B91C1C"

_LEVEL_TEXT = {"ok": "OK", "warn": "Check", "error": "Missing"}


class TemplateManager(tk.Frame):
    def __init__(self, parent=None):
        super().__init__(parent, bg=_BG)
        self.wizard_key = "mls"
        self._running = False
        self._tab_buttons: dict[str, tk.Button] = {}
        self._status_rows: list[dict] = []
        self._build()
        self._select_wizard("mls")

    # ------------------------------------------------------------------ UI
    def _build(self):
        head = tk.Frame(self, bg=_BG)
        head.pack(fill="x", padx=28, pady=(20, 8))
        tk.Label(head, text="Template Manager", bg=_BG, fg=_TEXT,
                 font=("Segoe UI", 20, "bold")).pack(anchor="w")
        tk.Label(
            head,
            text="Replace an official City template here when the City revises it. "
                 "The old file is kept, the change is recorded, and you can run the "
                 "built-in check right away. No code changes are needed.",
            bg=_BG, fg=_MUTED, font=("Segoe UI", 10), wraplength=1000, justify="left",
        ).pack(anchor="w", pady=(2, 0))

        info = tk.Frame(head, bg=_BG)
        info.pack(fill="x", pady=(8, 0))
        tk.Label(info, text=f"Templates folder: {TEMPLATES_ROOT}", bg=_BG, fg=_MUTED,
                 font=("Segoe UI", 9)).pack(side="left")
        tk.Button(info, text="Open folder", command=self._open_templates_root,
                  relief="flat", bg="#E2E8F0", fg=_TEXT, font=("Segoe UI", 9),
                  padx=10, cursor="hand2").pack(side="left", padx=12)
        tk.Label(info, text=f"PDF engine: {describe_engine()}", bg=_BG, fg=_MUTED,
                 font=("Segoe UI", 9)).pack(side="right")

        tabs = tk.Frame(self, bg=_BG)
        tabs.pack(fill="x", padx=28, pady=(6, 0))
        for key in tpl.wizard_keys():
            b = tk.Button(tabs, text=tpl.WIZARD_NAMES[key], relief="flat",
                          font=("Segoe UI", 10, "bold"), padx=16, pady=8, cursor="hand2",
                          command=lambda k=key: self._select_wizard(k))
            b.pack(side="left", padx=(0, 6))
            self._tab_buttons[key] = b

        card = tk.Frame(self, bg=_CARD, highlightbackground="#CBD5E1", highlightthickness=1)
        card.pack(fill="both", expand=True, padx=28, pady=(0, 10))

        cols = ("document", "file", "fill", "status", "updated")
        self.tree = ttk.Treeview(card, columns=cols, show="headings", height=8,
                                 selectmode="browse")
        for cid, text, width in (
            ("document", "Document", 230), ("file", "Template file", 360),
            ("fill", "Filled by", 110), ("status", "Status", 320), ("updated", "Last replaced", 130),
        ):
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=width, anchor="w", stretch=(cid in ("file", "status")))
        self.tree.tag_configure("ok", foreground=_OK)
        self.tree.tag_configure("warn", foreground=_WARN)
        self.tree.tag_configure("error", foreground=_ERR)
        self.tree.pack(fill="x", padx=12, pady=(12, 6))
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._show_details())

        self.details = tk.Label(card, text="", bg=_CARD, fg=_MUTED, font=("Segoe UI", 9),
                                anchor="w", justify="left", wraplength=1100)
        self.details.pack(fill="x", padx=14, pady=(0, 8))

        btns = tk.Frame(card, bg=_CARD)
        btns.pack(fill="x", padx=12, pady=(0, 12))

        def button(text, cmd, primary=False, side="left"):
            b = tk.Button(btns, text=text, command=cmd, relief="flat", cursor="hand2",
                          bg=_ACCENT if primary else "#E2E8F0",
                          fg="white" if primary else _TEXT,
                          activebackground=_ACCENT_MID if primary else "#CBD5E1",
                          activeforeground="white" if primary else _TEXT,
                          font=("Segoe UI", 10, "bold" if primary else "normal"),
                          padx=14, pady=7)
            b.pack(side=side, padx=(0, 8) if side == "left" else (8, 0))
            return b

        self.btn_replace = button("Replace selected template…", self._replace, primary=True)
        self.btn_restore = button("Restore previous version", self._restore)
        self.btn_open = button("Open selected in Word", self._open_selected)
        self.btn_check = button("Check this wizard", self._check_wizard, primary=True, side="right")
        self.btn_check_all = button("Check all wizards", self._check_all, side="right")
        self.btn_samples = button("Generate sample documents…", self._samples, side="right")

        logcard = tk.Frame(self, bg=_CARD, highlightbackground="#CBD5E1", highlightthickness=1)
        logcard.pack(fill="both", expand=True, padx=28, pady=(0, 20))
        top = tk.Frame(logcard, bg=_CARD)
        top.pack(fill="x", padx=12, pady=(8, 0))
        tk.Label(top, text="Check results", bg=_CARD, fg=_TEXT,
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        self.progress = ttk.Progressbar(top, mode="indeterminate", length=160)
        self.progress.pack(side="right")
        self.log = tk.Text(logcard, height=12, bg="#0F172A", fg="#E2E8F0", relief="flat",
                           font=("Consolas", 9), wrap="word", state="disabled")
        self.log.pack(fill="both", expand=True, padx=12, pady=(6, 12))
        self.log.tag_configure("pass", foreground="#4ADE80")
        self.log.tag_configure("fail", foreground="#F87171")
        self.log.tag_configure("head", foreground="#93C5FD")

    # ------------------------------------------------------------- helpers
    def _select_wizard(self, key: str):
        self.wizard_key = key
        for k, b in self._tab_buttons.items():
            active = k == key
            b.configure(bg=_ACCENT if active else "#E2E8F0",
                        fg="white" if active else _TEXT,
                        activebackground=_ACCENT_MID if active else "#CBD5E1",
                        activeforeground="white" if active else _TEXT)
        self._refresh()

    def _refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        try:
            self._status_rows = tpl.wizard_status(self.wizard_key)
        except Exception as e:
            self._status_rows = []
            self._append(f"Could not read the template manifest: {e}\n", "fail")
        for st in self._status_rows:
            self.tree.insert("", "end", iid=st["key"], tags=(st["level"],), values=(
                st["label"], st["file"], "tokens" if st["fill"] == "tokens" else "wording",
                f"{_LEVEL_TEXT.get(st['level'], '')} — {st['message']}",
                st["updated"],
            ))
        if self._status_rows:
            self.tree.selection_set(self._status_rows[0]["key"])
        self._show_details()

    def _selected(self) -> dict | None:
        sel = self.tree.selection()
        if not sel:
            return None
        return next((s for s in self._status_rows if s["key"] == sel[0]), None)

    def _show_details(self):
        st = self._selected()
        if not st:
            self.details.configure(text="")
            return
        lines = [f"{st['label']}: {st['path']}"]
        if st["fill"] == "tokens":
            lines.append("Filled by replacing {{tokens}}. A revised template must keep the same "
                         "token names; a new token the wizard does not know prints literally.")
            if st["unexpected_tokens"]:
                lines.append("New tokens: " + ", ".join(st["unexpected_tokens"]))
            if st["missing_tokens"]:
                lines.append("No longer used: " + ", ".join(st["missing_tokens"]))
        else:
            lines.append("Filled by matching the template's own placeholder wording "
                         "(for example 'XX Block of XX' or 'Planner Name'). Keep that wording "
                         "in a revised template, then use Generate sample documents to confirm "
                         "every field filled.")
        if st["split_run_risk"]:
            lines.append("Warning: a token is split across Word runs. Retype it in one motion in Word.")
        self.details.configure(text="\n".join(lines))
        self.btn_restore.configure(state="normal" if st["can_restore"] else "disabled")

    def _append(self, text: str, tag: str | None = None):
        self.log.configure(state="normal")
        self.log.insert("end", text, tag) if tag else self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _open_path(self, path: Path):
        try:
            if sys.platform == "win32":
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showerror("Template Manager", f"Could not open:\n{path}\n\n{e}")

    def _open_templates_root(self):
        self._open_path(templates_dir(self.wizard_key))

    def _open_selected(self):
        st = self._selected()
        if st and st["exists"]:
            self._open_path(st["path"])

    # ------------------------------------------------------------- actions
    def _replace(self):
        st = self._selected()
        if not st:
            return
        chosen = filedialog.askopenfilename(
            title=f"Choose the new {st['label']} template ({tpl.WIZARD_NAMES[self.wizard_key]})",
            filetypes=[("Word documents", "*.docx")])
        if not chosen:
            return
        if not messagebox.askyesno(
            "Replace template",
            f"Replace the {st['label']} template for the {tpl.WIZARD_NAMES[self.wizard_key]}?\n\n"
            f"Current: {st['file']}\nNew: {Path(chosen).name}\n\n"
            "The current file is kept in the _previous folder and can be restored.",
        ):
            return
        try:
            result = tpl.replace_template(self.wizard_key, st["key"], chosen)
        except Exception as e:
            messagebox.showerror("Replace template", str(e))
            return
        self._clear_log()
        self._append(f"Installed {result['installed'].name} as the {st['label']} template.\n", "head")
        if result["archived"]:
            self._append(f"Previous file archived as {result['archived'].name}\n")
        if result["split_run_risk"]:
            self._append("Warning: at least one {{token}} is split across Word runs and will not fill. "
                         "Open the template in Word, retype the token in one motion, and replace again.\n", "fail")
        self._refresh()
        self.tree.selection_set(st["key"])
        if messagebox.askyesno("Run check", "Run the built-in check for this wizard now? (about a minute)"):
            self._run_check([self.wizard_key], after=lambda ok: self._after_replace_check(st, ok))

    def _after_replace_check(self, st: dict, ok: bool):
        if ok:
            tpl.set_baseline(self.wizard_key, st["key"])
            self._refresh()
            messagebox.showinfo(
                "Check passed",
                "The new template generated every sample document with no unfilled placeholders.\n\n"
                "Use 'Generate sample documents' and open the results in Word to confirm the wording "
                "and layout before planners use it.")
        else:
            if messagebox.askyesno(
                "Check failed",
                "The check found problems with the new template (see the results panel).\n\n"
                "Restore the previous template now?"):
                self._restore(confirm=False)

    def _restore(self, confirm: bool = True):
        st = self._selected()
        if not st:
            return
        if confirm and not messagebox.askyesno(
            "Restore previous version",
            f"Put back the previous {st['label']} template for the {tpl.WIZARD_NAMES[self.wizard_key]}?"):
            return
        try:
            result = tpl.restore_previous(self.wizard_key, st["key"])
        except Exception as e:
            messagebox.showerror("Restore previous version", str(e))
            return
        self._append(f"Restored {result['restored'].name}.\n", "head")
        self._refresh()
        self.tree.selection_set(st["key"])

    def _check_wizard(self):
        self._run_check([self.wizard_key])

    def _check_all(self):
        self._run_check(None)

    def _samples(self):
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        out = Path.home() / "Documents" / "RCA_Template_Checks" / f"{self.wizard_key}_{stamp}"
        self._run_check([self.wizard_key], keep_dir=out,
                        after=lambda ok: self._open_path(out) if out.exists() else None)

    # ------------------------------------------------------------ checking
    def _set_running(self, running: bool):
        self._running = running
        state = "disabled" if running else "normal"
        for b in (self.btn_replace, self.btn_restore, self.btn_check, self.btn_check_all,
                  self.btn_samples):
            b.configure(state=state)
        if running:
            self.progress.start(12)
        else:
            self.progress.stop()
        if not running:
            self._show_details()

    def _run_check(self, wizards: list[str] | None, keep_dir: Path | None = None, after=None):
        if self._running:
            return
        self._clear_log()
        self._set_running(True)
        label = ", ".join(tpl.WIZARD_NAMES[w] for w in wizards) if wizards else "all wizards"
        self._append(f"Checking {label}…\n", "head")
        if keep_dir:
            self._append(f"Sample documents will be saved to {keep_dir}\n")

        def worker():
            ok = True
            try:
                if wizards:
                    for w in wizards:
                        ok = self._run_one(w, keep_dir) and ok
                else:
                    ok = self._run_one(None, keep_dir)
            except Exception as e:
                ok = False
                self.after(0, self._append, f"Check could not run: {e}\n", "fail")
            self.after(0, self._finish_check, ok, after)

        threading.Thread(target=worker, daemon=True).start()

    def _run_one(self, wizard: str | None, keep_dir: Path | None) -> bool:
        cmd = core_launch.validate_command(wizard, keep_dir)
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, encoding="utf-8", errors="replace",
                                creationflags=core_launch.hidden_flags(),
                                cwd=str(Path(cmd[0]).parent) if core_launch.FROZEN else None)
        assert proc.stdout is not None
        for line in proc.stdout:
            tag = "fail" if "[FAIL]" in line or "RESULT: FAIL" in line else (
                "pass" if "RESULT: PASS" in line else None)
            self.after(0, self._append, line, tag)
        return proc.wait() == 0

    def _finish_check(self, ok: bool, after):
        self._append("\nCHECK PASSED\n" if ok else "\nCHECK FAILED — see the [FAIL] lines above\n",
                     "pass" if ok else "fail")
        self._set_running(False)
        if after:
            after(ok)
