# -*- coding: utf-8 -*-
r"""
WinRecent Explorer – Historique des éléments consultés (Windows) – GUI Tkinter
- Parcourt %AppData%\Microsoft\Windows\Recent
- Résout les .lnk via COM (ctypes) -> pas de pywin32
- Stocke en SQLite (historique persistant, 2 ans par défaut)
- Double-clic/bouton "Ouvrir (dans 'Recent')" : ouvre l'Explorateur sur le .lnk sélectionné
- Bouton "Activer la sauvegarde Auto (tâche planifiée)" : crée une tâche hebdomadaire (--weekly-scan)
- Clic droit sur une ligne -> copie le champ "Nom" dans le presse-papiers
- Mode --weekly-scan : écrit un autoscan.log (50 dernières insertions de ce run)
"""

import os
import sys
import csv
import shutil
import sqlite3
import subprocess
import ctypes
import ctypes.wintypes as wt
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional, List, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ===================== CONFIG =====================
APP_NAME = "WinRecent Explorer"
TASK_NAME = "RecentHistory_AutoScanWeekly"
DEFAULT_LOOKBACK_DAYS = 730  # 2 ans
RECENT_DIR = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Recent"
DB_DIR = Path(os.environ.get("LOCALAPPDATA", "")) / "RecentHistory"
DB_PATH = DB_DIR / "history.db"

# ===================== COMPAT: GUID / HRESULT / LPCOLESTR =====================
if not hasattr(wt, "GUID"):
    class GUID(ctypes.Structure):
        _fields_ = [
            ("Data1", ctypes.c_ulong),
            ("Data2", ctypes.c_ushort),
            ("Data3", ctypes.c_ushort),
            ("Data4", ctypes.c_ubyte * 8),
        ]
        def __init__(self, guid_str):
            import uuid
            u = uuid.UUID(guid_str)
            super().__init__()
            self.Data1 = u.time_low
            self.Data2 = u.time_mid
            self.Data3 = u.time_hi_version
            for i, b in enumerate(u.bytes[8:]):
                self.Data4[i] = b
    wt.GUID = GUID

if not hasattr(wt, "HRESULT"):
    class HRESULT(ctypes.c_long):
        pass
    wt.HRESULT = HRESULT

if not hasattr(wt, "LPCOLESTR"):
    wt.LPCOLESTR = wt.LPWSTR  # Unicode OK

# ===================== COM / ShellLink (.lnk) =====================
CLSID_ShellLink = wt.GUID("{00021401-0000-0000-C000-000000000046}")
IID_IShellLinkW = wt.GUID("{000214F9-0000-0000-C000-000000000046}")
IID_IPersistFile = wt.GUID("{0000010b-0000-0000-C000-000000000046}")

class IShellLinkW(ctypes.Structure): pass
class IPersistFile(ctypes.Structure): pass

LPIShellLinkW = ctypes.POINTER(IShellLinkW)
LPIPersistFile = ctypes.POINTER(IPersistFile)

class IShellLinkW_VTable(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        *[(f"slot{i}", ctypes.c_void_p) for i in range(3, 20)],
        ("GetPath", ctypes.c_void_p),
        ("GetIDList", ctypes.c_void_p),
        ("SetIDList", ctypes.c_void_p),
        ("GetDescription", ctypes.c_void_p),
    ]
IShellLinkW._fields_ = [("lpVtbl", ctypes.POINTER(IShellLinkW_VTable))]

class IPersistFile_VTable(ctypes.Structure):
    _fields_ = [
        ("QueryInterface", ctypes.c_void_p),
        ("AddRef", ctypes.c_void_p),
        ("Release", ctypes.c_void_p),
        ("GetClassID", ctypes.c_void_p),
        ("IsDirty", ctypes.c_void_p),
        ("Load", ctypes.c_void_p),
        ("Save", ctypes.c_void_p),
        ("SaveCompleted", ctypes.c_void_p),
        ("GetCurFile", ctypes.c_void_p),
    ]
IPersistFile._fields_ = [("lpVtbl", ctypes.POINTER(IPersistFile_VTable))]

GetPathProto = ctypes.WINFUNCTYPE(wt.HRESULT, LPIShellLinkW, wt.LPWSTR, ctypes.c_int, ctypes.c_void_p, ctypes.c_uint)
PFLoadProto  = ctypes.WINFUNCTYPE(wt.HRESULT, LPIPersistFile, wt.LPCOLESTR, ctypes.c_uint)

ole32 = ctypes.OleDLL("ole32")
CoCreateInstance = ole32.CoCreateInstance
CoInitialize     = ole32.CoInitialize
CoUninitialize   = ole32.CoUninitialize
CoCreateInstance.argtypes = [
    ctypes.POINTER(wt.GUID), ctypes.c_void_p, ctypes.c_uint,
    ctypes.POINTER(wt.GUID), ctypes.POINTER(ctypes.c_void_p)
]
CLSCTX_INPROC_SERVER = 1

def resolve_lnk(lnk_path: Path) -> str:
    """Résout un .lnk vers sa cible (chemin réel). Retourne '' en cas d'échec."""
    if not lnk_path.exists() or lnk_path.suffix.lower() != ".lnk":
        return ""
    CoInitialize(None)
    try:
        psl = ctypes.c_void_p()
        hr = CoCreateInstance(
            ctypes.byref(CLSID_ShellLink), None, CLSCTX_INPROC_SERVER,
            ctypes.byref(IID_IShellLinkW), ctypes.byref(psl)
        )
        if (getattr(hr, "value", hr) != 0) or not psl:
            return ""
        shell_link = ctypes.cast(psl, LPIShellLinkW)
        QI = ctypes.WINFUNCTYPE(
            wt.HRESULT, LPIShellLinkW, ctypes.POINTER(wt.GUID), ctypes.POINTER(ctypes.c_void_p)
        )(shell_link.lpVtbl.contents.QueryInterface)
        ppv = ctypes.c_void_p()
        if QI(shell_link, ctypes.byref(IID_IPersistFile), ctypes.byref(ppv)) != 0 or not ppv:
            return ""
        persist_file = ctypes.cast(ppv, LPIPersistFile)
        pf_load = PFLoadProto(persist_file.lpVtbl.contents.Load)
        if pf_load(persist_file, str(lnk_path), 0) != 0:
            return ""
        get_path = GetPathProto(shell_link.lpVtbl.contents.GetPath)
        buf = ctypes.create_unicode_buffer(1024)
        SLGP_RAWPATH = 0x0000
        if get_path(shell_link, buf, 1024, None, SLGP_RAWPATH) != 0:
            return ""
        return buf.value or ""
    except Exception:
        return ""
    finally:
        try: CoUninitialize()
        except Exception: pass

# ===================== UTIL =====================
def file_mtime_dt(path: Path) -> datetime:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime)
    except Exception:
        return datetime.min

def _app_dir() -> Path:
    """Dossier où écrire le log : dossier de l'exe (si packagé) sinon du script .py."""
    try:
        if getattr(sys, "frozen", False) and Path(sys.executable).exists():
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent
    except Exception:
        return Path.cwd()

# ===================== DB =====================
def ensure_db():
    DB_DIR.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS items (
            id INTEGER PRIMARY KEY,
            target_path TEXT,
            display_name TEXT,
            source TEXT,
            opened_at TEXT, -- ISO
            exists_now INTEGER
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_items_opened_at ON items(opened_at)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_items_target ON items(target_path)")
    con.commit()
    return con

def upsert_item(con, target_path: str, display_name: str, source: str, opened_at: datetime) -> bool:
    """Insère si nouveau couple (target_path, opened_at). Retourne True si INSERT, False sinon."""
    cur = con.cursor()
    cur.execute("SELECT id FROM items WHERE target_path=? AND opened_at=?", (target_path, opened_at.isoformat()))
    row = cur.fetchone()
    exists_now = 1 if (target_path and Path(target_path).exists()) else 0
    if row:
        cur.execute("UPDATE items SET display_name=?, source=?, exists_now=? WHERE id=?",
                    (display_name, source, exists_now, row[0]))
        con.commit()
        return False
    else:
        cur.execute("""INSERT INTO items(target_path, display_name, source, opened_at, exists_now)
                       VALUES(?,?,?,?,?)""",
                    (target_path, display_name, source, opened_at.isoformat(), exists_now))
        con.commit()
        return True

# ===================== SCAN =====================
def scan_recent(con) -> Tuple[int, List[Tuple[str, str]]]:
    """
    Parcourt le dossier 'Recent'. Retourne:
      - count: nb de .lnk parcourus
      - added: liste [(display_name, opened_at_iso)] pour les entrées insérées pendant CE scan
    """
    if not RECENT_DIR.exists():
        return 0, []
    count = 0
    added: List[Tuple[str, str]] = []
    for lnk in RECENT_DIR.glob("*.lnk"):
        opened_at = file_mtime_dt(lnk)  # ≈ date dernière ouverture
        target = resolve_lnk(lnk)
        display = lnk.stem
        inserted = upsert_item(con, target, display, "Recent(.lnk)", opened_at)
        count += 1
        if inserted:
            added.append((display, opened_at.isoformat()))
    return count, added

# ===================== TÂCHE PLANIFIÉE =====================
def _norm_name(s: str) -> str:
    return "".join(ch.lower() for ch in s if ch.isalnum())

def _preferred_exe_for_this_script() -> Optional[Path]:
    """Retourne l'exe correspondant au nom du script (même dossier puis ./dist/), variantes nom acceptées."""
    try:
        script = Path(__file__).resolve()
    except Exception:
        return None
    stem = script.stem
    cand = script.with_suffix(".exe")
    if cand.exists():
        return cand
    dist_dir = script.parent / "dist"
    if dist_dir.is_dir():
        exact = dist_dir / f"{stem}.exe"
        if exact.exists():
            return exact
        wanted = _norm_name(stem)
        for exe in dist_dir.glob("*.exe"):
            if _norm_name(exe.stem) == wanted:
                return exe
    return None

def create_weekly_task():
    """
    Crée/MAJ la tâche planifiée Windows pour lancer **l’EXE** correspondant
    avec --weekly-scan (lundi 09:00). Si aucun EXE trouvé, repli sur le .py.
    AUCUNE boîte de dialogue.
    """
    if getattr(sys, "frozen", False) and str(sys.executable).lower().endswith(".exe"):
        run_target = Path(sys.executable).resolve()
        used_exe = True
    else:
        run_target = _preferred_exe_for_this_script()
        used_exe = run_target is not None

    if used_exe:
        tr_cmd = f'"{run_target}" --weekly-scan'
    else:
        try:
            script_path = Path(__file__).resolve()
        except Exception:
            return False, "Impossible de déterminer le chemin du script."
        py = Path(sys.executable).resolve()
        tr_cmd = f'"{py}" "{script_path}" --weekly-scan'

    try:
        subprocess.run([
            "schtasks", "/Create",
            "/TN", TASK_NAME,
            "/F",
            "/SC", "WEEKLY",
            "/D", "MON",
            "/ST", "09:00",
            "/RL", "LIMITED",
            "/TR", tr_cmd
        ], check=True)

        if used_exe:
            msg = (f"Tâche planifiée créée/actualisée : {TASK_NAME}\n"
                   f"Exécutera : {run_target}\n"
                   f"Argument : --weekly-scan\n"
                   f"Déclenchement : chaque lundi à 09:00.")
        else:
            msg = (f"Tâche planifiée créée/actualisée : {TASK_NAME}\n"
                   f"Aucun .exe correspondant trouvé : repli sur le script Python.\n"
                   f"Commande : {tr_cmd}\n"
                   f"Déclenchement : chaque lundi à 09:00.")
        return True, msg
    except subprocess.CalledProcessError as e:
        return False, f"Échec de la création de la tâche (code {e.returncode})."
    except FileNotFoundError:
        return False, "Commande 'schtasks' introuvable (Windows requis)."

def run_weekly_scan_once():
    """
    Mode silencieux lancé par la tâche planifiée : scan + maj DB puis sortie.
    Écrit un autoscan.log dans le dossier de l'exe/.py avec les 50 dernières INSERTIONS de CE run.
    """
    con = ensure_db()
    count, added = scan_recent(con)

    # 50 dernières insertions du run
    added_sorted = sorted(added, key=lambda t: t[1], reverse=True)[:50]

    log_path = _app_dir() / "autoscan.log"
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"\n[{ts}] Autoscan: {count} .lnk parcourus, {len(added)} insertion(s).\n")
            if added_sorted:
                f.write("Dernières insertions (Nom | Date d'ouverture):\n")
                for name, opened_at_iso in added_sorted:
                    try:
                        dt = datetime.fromisoformat(opened_at_iso).strftime("%Y-%m-%d %H:%M:%S")
                    except Exception:
                        dt = opened_at_iso
                    f.write(f" - {name} | {dt}\n")
            else:
                f.write("Aucune nouvelle entrée insérée sur ce run.\n")
    except Exception:
        # Log de secours dans %LOCALAPPDATA%
        try:
            DB_DIR.mkdir(parents=True, exist_ok=True)
            with open(DB_DIR / "autoscan_fallback.log", "a", encoding="utf-8") as f:
                f.write(f"[{ts}] Échec d'écriture {log_path}\n")
        except Exception:
            pass

# ===================== GUI =====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1100x660")
        self.minsize(900, 560)

        self.con = ensure_db()

        # --- Barre du haut ---
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Label(top, text="Recherche:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(top, textvariable=self.search_var, width=40)
        self.search_entry.pack(side="left", padx=6)
        self.search_entry.bind("<Return>", lambda e: self.refresh_table())

        ttk.Label(top, text="Du (YYYY-MM-DD):").pack(side="left", padx=(20,2))
        self.from_var = tk.StringVar(value=(datetime.now() - timedelta(days=DEFAULT_LOOKBACK_DAYS)).strftime("%Y-%m-%d"))
        ttk.Entry(top, textvariable=self.from_var, width=12).pack(side="left")

        ttk.Label(top, text="Au:").pack(side="left", padx=(10,2))
        self.to_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(top, textvariable=self.to_var, width=12).pack(side="left")

        ttk.Button(top, text="Appliquer filtres", command=self.refresh_table).pack(side="left", padx=8)
        ttk.Button(top, text="Scanner maintenant", command=self.scan_now).pack(side="left", padx=8)
        ttk.Button(top, text="Exporter CSV", command=self.export_csv).pack(side="left", padx=8)

        # --- Tableau ---
        table = ttk.Frame(self)
        table.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)

        cols = ("opened_at", "display_name", "target_path", "source", "exists_now")
        self.tree = ttk.Treeview(table, columns=cols, show="headings")

        yscroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.heading("opened_at", text="Date d'ouverture")
        self.tree.heading("display_name", text="Nom")
        self.tree.heading("target_path", text="Chemin (cible)")
        self.tree.heading("source", text="Source")
        self.tree.heading("exists_now", text="Existe")

        self.tree.column("opened_at", width=160, anchor="w")
        self.tree.column("display_name", width=240, anchor="w")
        self.tree.column("target_path", width=520, anchor="w")
        self.tree.column("source", width=140, anchor="w")
        self.tree.column("exists_now", width=70, anchor="center")

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        # Double-clic (ouvrir le .lnk dans Explorer)
        self.tree.bind("<Double-1>", self._on_double_click_row)
        # Clic droit -> copier le Nom
        self.tree.bind("<Button-3>", self._on_right_click_copy_name)

        # --- Bas (avec barre de statut) ---
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=(0,10))
        ttk.Button(bottom, text="Ouvrir (dans 'Recent')", command=self.open_file).pack(side="left")
        ttk.Button(bottom, text="Activer la sauvegarde Auto (tâche planifiée)", command=self.enable_autoscan).pack(side="left", padx=8)
        ttk.Button(bottom, text="Sauvegarder la base (copie)", command=self.backup_db).pack(side="right")

        self.status_var = tk.StringVar(value="")
        status = ttk.Label(self, textvariable=self.status_var, anchor="w")
        status.pack(fill="x", side="bottom", padx=10, pady=(0,6))

        self.refresh_table()

    # ----- Helpers -----
    def parse_dates(self):
        def parse(s, default):
            try:
                return datetime.strptime(s.strip(), "%Y-%m-%d")
            except Exception:
                return default
        dfrom = parse(self.from_var.get(), datetime.now() - timedelta(days=DEFAULT_LOOKBACK_DAYS))
        dto = parse(self.to_var.get(), datetime.now())
        return dfrom, dto.replace(hour=23, minute=59, second=59, microsecond=999999)

    def query_rows(self, dfrom, dto):
        cur = self.con.cursor()
        cur.execute("""
            SELECT opened_at, display_name, target_path, source, exists_now
            FROM items
            WHERE datetime(opened_at) BETWEEN ? AND ?
            ORDER BY datetime(opened_at) DESC
        """, (dfrom.isoformat(), dto.isoformat()))
        return cur.fetchall()

    # ----- UI actions -----
    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        dfrom, dto = self.parse_dates()
        q = self.search_var.get().strip().lower()
        for opened_at, name, path, source, exists_now in self.query_rows(dfrom, dto):
            if q and q not in (name or "").lower() and q not in (path or "").lower():
                continue
            exists_label = "Oui" if exists_now else "Non"
            self.tree.insert("", "end", values=(opened_at.replace("T", " ")[:19], name, path, source, exists_label))

    def scan_now(self):
        count, _ = scan_recent(self.con)  # on ignore la liste des ajouts en mode GUI
        self.refresh_table()
        messagebox.showinfo(APP_NAME, f"Scan terminé.\n{count} élément(s) importé(s) depuis 'Recent'.")

    def get_selected(self):
        sel = self.tree.selection()
        if not sel:
            return None
        vals = self.tree.item(sel[0], "values")
        return {"opened_at": vals[0], "name": vals[1], "path": vals[2], "source": vals[3], "exists": vals[4]}

    def _recent_lnk_path_for_row(self, item) -> Path:
        name = (item or {}).get("name") or ""
        return RECENT_DIR / f"{name}.lnk"

    def _on_double_click_row(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            self.tree.focus(row_id)
        self.open_file()

    def _on_right_click_copy_name(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self.tree.selection_set(row_id)
        self.tree.focus(row_id)
        vals = self.tree.item(row_id, "values")
        if not vals or len(vals) < 2:
            return
        name = vals[1]
        try:
            self.clipboard_clear()
            self.clipboard_append(name)
            self.update()
            self._flash_status(f'Nom copié : "{name}"')
        except Exception:
            self._flash_status("Échec de la copie du nom.")

    def _flash_status(self, msg: str, delay_ms: int = 2000):
        self.status_var.set(msg)
        self.after(delay_ms, lambda: self.status_var.set(""))

    def open_file(self):
        item = self.get_selected()
        if not item:
            messagebox.showwarning(APP_NAME, "Aucune sélection.")
            return
        lnk = self._recent_lnk_path_for_row(item)
        if not lnk.exists():
            messagebox.showwarning(APP_NAME, f"Raccourci introuvable dans 'Recent':\n{lnk}")
            return
        try:
            subprocess.Popen(["explorer", "/select,", str(lnk.resolve())])
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Impossible d’ouvrir dans l’Explorateur.\n{e}")

    def enable_autoscan(self):
        ok, msg = create_weekly_task()
        if ok:
            messagebox.showinfo(APP_NAME, msg)
        else:
            messagebox.showerror(APP_NAME, msg)

    def export_csv(self):
        path = filedialog.asksaveasfilename(
            title="Exporter CSV", defaultextension=".csv",
            filetypes=[("CSV", "*.csv")], initialfile="historique.csv"
        )
        if not path:
            return
        dfrom, dto = self.parse_dates()
        rows = self.query_rows(dfrom, dto)
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["opened_at", "name", "path", "source", "exists"])
            for opened_at, name, path, source, exists_now in rows:
                w.writerow([opened_at, name, path, source, "1" if exists_now else "0"])
        messagebox.showinfo(APP_NAME, f"Exporté vers:\n{path}")

    def backup_db(self):
        path = filedialog.asksaveasfilename(
            title="Sauvegarder la base", defaultextension=".db",
            filetypes=[("SQLite DB", "*.db")], initialfile="history_backup.db"
        )
        if not path:
            return
        try:
            self.con.commit()
            shutil.copy2(DB_PATH, path)
            messagebox.showinfo(APP_NAME, f"Copie réalisée:\n{path}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Échec de la sauvegarde:\n{e}")

# ===================== Bootstrap =====================
def first_scan_if_needed(con):
    c = con.cursor()
    c.execute("SELECT COUNT(*) FROM items")
    n = c.fetchone()[0]
    if n == 0:
        scan_recent(con)

def main_gui():
    con = ensure_db()
    first_scan_if_needed(con)
    app = App()
    app.mainloop()

if __name__ == "__main__":
    # Mode tâche planifiée : scan silencieux puis sortie
    if "--weekly-scan" in sys.argv:
        run_weekly_scan_once()
        sys.exit(0)

    if not RECENT_DIR.exists():
        root = tk.Tk(); root.withdraw()
        messagebox.showerror(APP_NAME, f"Dossier 'Recent' introuvable:\n{RECENT_DIR}")
        sys.exit(1)

    main_gui()
