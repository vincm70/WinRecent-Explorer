# -*- coding: utf-8 -*-
r"""
WinRecent Explorer – Historique des éléments consultés (Windows)
- Parcourt %AppData%\Microsoft\Windows\Recent
- Stocke en SQLite (2 ans d’historique) avec colonne target_path conservée
- UI sans colonne chemin (on garde target_path en DB pour compat)
- Double-clic / bouton "Ouvrir (dans 'Recent')" : ouvre le .lnk dans Recent
- Clic droit : copie le "Nom" dans le presse-papiers
- Tâche planifiée (--weekly-scan) : met à jour la base et écrit autoscan.log
- NOUVEAU : chaque .lnk est sauvegardé dans lnk_backup (à côté du script/exe).
  Si le .lnk a disparu de Recent, on le restaure depuis lnk_backup avant ouverture.
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
from datetime import datetime
from typing import Optional, List, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ===================== CONFIG =====================
APP_NAME = "WinRecent Explorer"
TASK_NAME = "RecentHistory_AutoScanWeekly"
DEFAULT_LOOKBACK_DAYS = 730
RECENT_DIR = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Recent"
DB_DIR = Path(os.environ.get("LOCALAPPDATA", "")) / "RecentHistory"
DB_PATH = DB_DIR / "history.db"

# ===================== COM / ShellLink =====================
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
    class HRESULT(ctypes.c_long): pass
    wt.HRESULT = HRESULT
if not hasattr(wt, "LPCOLESTR"):
    wt.LPCOLESTR = wt.LPWSTR

CLSID_ShellLink = wt.GUID("{00021401-0000-0000-C000-000000000046}")
IID_IShellLinkW = wt.GUID("{000214F9-0000-0000-C000-000000000046}")
IID_IPersistFile = wt.GUID("{0000010b-0000-0000-C000-000000000046}")

class IShellLinkW(ctypes.Structure): pass
class IPersistFile(ctypes.Structure): pass

LPIShellLinkW = ctypes.POINTER(IShellLinkW)
LPIPersistFile = ctypes.POINTER(IPersistFile)

ole32 = ctypes.OleDLL("ole32")
CoInitialize = ole32.CoInitialize
CoUninitialize = ole32.CoUninitialize

# ===================== UTIL =====================
def file_mtime_dt(path: Path) -> datetime:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime)
    except Exception:
        return datetime.min

def _app_dir() -> Path:
    """Dossier du script .py ou de l'exe packagé."""
    try:
        if getattr(sys, "frozen", False) and Path(sys.executable).exists():
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent
    except Exception:
        return Path.cwd()

def get_backup_dir() -> Path:
    """
    Dossier de sauvegarde des .lnk :
    lnk_backup à côté du script/exe.
    """
    d = _app_dir() / "lnk_backup"
    d.mkdir(parents=True, exist_ok=True)
    return d

# ===================== DB =====================
def ensure_db():
    """Crée la DB et assure la présence de toutes les colonnes nécessaires."""
    DB_DIR.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS items (
            id INTEGER PRIMARY KEY,
            target_path TEXT,
            display_name TEXT,
            source TEXT,
            opened_at TEXT,
            exists_now INTEGER
        )
    """)
    cur.execute("PRAGMA table_info(items)")
    cols = {row[1] for row in cur.fetchall()}
    changed = False
    if "target_path" not in cols:
        cur.execute("ALTER TABLE items ADD COLUMN target_path TEXT"); changed = True
    if "display_name" not in cols:
        cur.execute("ALTER TABLE items ADD COLUMN display_name TEXT"); changed = True
    if "source" not in cols:
        cur.execute("ALTER TABLE items ADD COLUMN source TEXT"); changed = True
    if "opened_at" not in cols:
        cur.execute("ALTER TABLE items ADD COLUMN opened_at TEXT"); changed = True
    if "exists_now" not in cols:
        cur.execute("ALTER TABLE items ADD COLUMN exists_now INTEGER"); changed = True
    if changed:
        con.commit()
    cur.execute("CREATE INDEX IF NOT EXISTS idx_items_opened_at ON items(opened_at)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_items_target ON items(target_path)")
    con.commit()
    return con

def upsert_item(con, target_path: str, display_name: str, source: str, opened_at: datetime) -> bool:
    """Ajoute si nouveau (target_path, opened_at), sinon met à jour. Retourne True si INSERT."""
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

# ===================== SCAN + BACKUP LNK =====================
def scan_recent(con) -> Tuple[int, List[Tuple[str, str]]]:
    """
    Parcourt 'Recent' et ajoute les entrées dans la DB.
    Sauvegarde chaque .lnk dans lnk_backup (à côté du script/exe).
    Retourne (nb_lus, [(display_name, opened_at_iso)] insérés).
    """
    if not RECENT_DIR.exists():
        return 0, []
    count = 0
    added: List[Tuple[str, str]] = []
    backup_dir = get_backup_dir()

    for lnk in RECENT_DIR.glob("*.lnk"):
        opened_at = file_mtime_dt(lnk)
        display = lnk.stem
        # target_path laissé vide (on garde la colonne pour compat, mais pas besoin pour l’UI)
        target_path = ""

        # Sauvegarde du .lnk dans lnk_backup
        try:
            shutil.copy2(lnk, backup_dir / lnk.name)
        except Exception:
            # on ne bloque pas le scan si la copie échoue
            pass

        inserted = upsert_item(con, target_path, display, "Recent(.lnk)", opened_at)
        count += 1
        if inserted:
            added.append((display, opened_at.isoformat()))
    return count, added

# ===================== TACHE PLANIFIEE =====================
def _norm_name(s: str) -> str:
    return "".join(ch.lower() for ch in s if ch.isalnum())

def _preferred_exe_for_this_script() -> Optional[Path]:
    """Tente de trouver l'EXE qui correspond au nom du script (même dossier puis ./dist/)."""
    try:
        script = Path(__file__).resolve()
    except Exception:
        return None
    stem = _norm_name(script.stem)
    exe_here = script.with_suffix(".exe")
    if exe_here.exists() and _norm_name(exe_here.stem) == stem:
        return exe_here
    dist = script.parent / "dist"
    if dist.is_dir():
        for exe in dist.glob("*.exe"):
            if _norm_name(exe.stem) == stem:
                return exe
    return None

def create_weekly_task():
    """Crée la tâche planifiée hebdomadaire (lundi 09:00). Priorité à l'EXE correspondant, sinon repli .py."""
    if getattr(sys, "frozen", False) and str(sys.executable).lower().endswith(".exe"):
        run_target = Path(sys.executable).resolve()
    else:
        run_target = _preferred_exe_for_this_script() or Path(__file__).resolve()
    tr_cmd = f'"{run_target}" --weekly-scan'
    try:
        subprocess.run([
            "schtasks", "/Create", "/TN", TASK_NAME, "/F",
            "/SC", "WEEKLY", "/D", "MON", "/ST", "09:00",
            "/RL", "LIMITED", "/TR", tr_cmd
        ], check=True)
        return True, f"Tâche planifiée créée/MAJ : {TASK_NAME}\n{tr_cmd}"
    except subprocess.CalledProcessError as e:
        return False, f"Erreur schtasks (code {e.returncode})."
    except Exception as e:
        return False, f"Erreur création tâche : {e}"

def run_weekly_scan_once():
    """Scan silencieux et écriture d’un autoscan.log (50 dernières insertions de CE run) dans le dossier de l'exe/.py."""
    con = ensure_db()
    count, added = scan_recent(con)
    log_path = _app_dir() / "autoscan.log"
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    added_sorted = sorted(added, key=lambda t: t[1], reverse=True)[:50]
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
                f.write("Aucune nouvelle entrée insérée.\n")
    except Exception:
        try:
            with open(DB_DIR / "autoscan_fallback.log", "a", encoding="utf-8") as f:
                f.write(f"[{ts}] Échec d'écriture {log_path}\n")
        except Exception:
            pass

# ===================== GUI =====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("950x620")
        self.minsize(820, 540)
        self.con = ensure_db()

        # --- Barre haut ---
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Label(top, text="Recherche:").pack(side="left")
        self.search_var = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.search_var, width=40)
        ent.pack(side="left", padx=6)
        ent.bind("<Return>", lambda e: self.refresh_table())
        ttk.Button(top, text="Appliquer filtres", command=self.refresh_table).pack(side="left", padx=8)
        ttk.Button(top, text="Scanner maintenant", command=self.scan_now).pack(side="left", padx=8)
        ttk.Button(top, text="Exporter CSV", command=self.export_csv).pack(side="left", padx=8)

        # --- Tableau (sans la colonne chemin) ---
        frame = ttk.Frame(self); frame.pack(fill="both", expand=True, padx=10, pady=(0,8))
        frame.rowconfigure(0, weight=1); frame.columnconfigure(0, weight=1)
        cols = ("opened_at", "display_name", "source", "exists_now")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings")
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.heading("opened_at", text="Date d'ouverture")
        self.tree.heading("display_name", text="Nom")
        self.tree.heading("source", text="Source")
        self.tree.heading("exists_now", text="Existe")
        self.tree.column("opened_at", width=170)
        self.tree.column("display_name", width=360)
        self.tree.column("source", width=160)
        self.tree.column("exists_now", width=70, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        # Interactions
        self.tree.bind("<Double-1>", self._on_double_click_row)      # ouvre le .lnk
        self.tree.bind("<Button-3>", self._on_right_click_copy_name) # copie le nom

        # --- Bas + statut ---
        bottom = ttk.Frame(self); bottom.pack(fill="x", padx=10, pady=(0,10))
        ttk.Button(bottom, text="Ouvrir (dans 'Recent')", command=self.open_file).pack(side="left")
        ttk.Button(bottom, text="Activer la sauvegarde Auto", command=self.enable_autoscan).pack(side="left", padx=8)
        ttk.Button(bottom, text="Sauvegarder la base", command=self.backup_db).pack(side="right")

        self.status_var = tk.StringVar()
        ttk.Label(self, textvariable=self.status_var, anchor="w").pack(fill="x", padx=10, pady=(0,6))

        self.refresh_table()

    # ---- Data ----
    def query_rows(self):
        cur = self.con.cursor()
        cur.execute("""
            SELECT opened_at, display_name, source, exists_now
            FROM items
            ORDER BY datetime(opened_at) DESC
        """)
        return cur.fetchall()

    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        q = self.search_var.get().lower().strip()
        for opened_at, name, source, exists_now in self.query_rows():
            if q and q not in (name or "").lower():
                continue
            exists_label = "Oui" if exists_now else "Non"
            opened_at_disp = (opened_at or "").replace("T", " ")[:19]
            self.tree.insert("", "end", values=(opened_at_disp, name, source, exists_label))

    # ---- Actions UI ----
    def scan_now(self):
        count, _ = scan_recent(self.con)
        self.refresh_table()
        messagebox.showinfo(APP_NAME, f"Scan terminé : {count} éléments parcourus.")

    def _get_selected_name(self) -> Optional[str]:
        sel = self.tree.selection()
        if not sel: return None
        vals = self.tree.item(sel[0], "values")
        return vals[1] if vals and len(vals) >= 2 else None

    def _on_double_click_row(self, event):
        # sélectionne la ligne sous le curseur avant d’ouvrir
        row_id = self.tree.identify_row(event.y)
        if row_id:
            self.tree.selection_set(row_id)
            self.tree.focus(row_id)
        self.open_file()

    def _on_right_click_copy_name(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id: return
        self.tree.selection_set(row_id)
        self.tree.focus(row_id)
        vals = self.tree.item(row_id, "values")
        if not vals or len(vals) < 2: return
        name = vals[1]
        self.clipboard_clear(); self.clipboard_append(name); self.update()
        self._flash_status(f'Nom copié : "{name}"')

    def _flash_status(self, msg: str, delay_ms: int = 2000):
        self.status_var.set(msg)
        self.after(delay_ms, lambda: self.status_var.set(""))

    def open_file(self):
        """Ouvre le .lnk dans Recent.
        Si le .lnk a été nettoyé, on tente de le restaurer depuis lnk_backup.
        """
        name = self._get_selected_name()
        if not name:
            messagebox.showwarning(APP_NAME, "Aucune sélection.")
            return

        lnk = RECENT_DIR / f"{name}.lnk"

        # Si le .lnk n'existe plus dans Recent -> tentative de restauration
        if not lnk.exists():
            backup_dir = get_backup_dir()
            backup_lnk = backup_dir / f"{name}.lnk"
            if backup_lnk.exists():
                try:
                    RECENT_DIR.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(backup_lnk, lnk)
                    # On continue ensuite comme si de rien n'était
                except Exception as e:
                    messagebox.showerror(
                        APP_NAME,
                        f"Le raccourci n'existe plus dans 'Recent' et la restauration a échoué :\n{e}"
                    )
                    return
            else:
                messagebox.showwarning(
                    APP_NAME,
                    f"Raccourci introuvable dans 'Recent' et aucune sauvegarde trouvée :\n{lnk}"
                )
                return

        try:
            subprocess.Popen(["explorer", "/select,", str(lnk.resolve())])
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Impossible d’ouvrir dans l’Explorateur.\n{e}")

    def enable_autoscan(self):
        ok, msg = create_weekly_task()
        if ok: messagebox.showinfo(APP_NAME, msg)
        else: messagebox.showerror(APP_NAME, msg)

    def export_csv(self):
        path = filedialog.asksaveasfilename(
            title="Exporter CSV",
            defaultextension=".csv",
            initialfile=f"historique_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            filetypes=[("CSV", "*.csv"), ("Tous les fichiers", "*.*")]
        )
        if not path: return
        rows = self.query_rows()
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["opened_at", "name", "source", "exists"])
            for opened_at, name, source, exists_now in rows:
                w.writerow([opened_at, name, source, "1" if exists_now else "0"])
        messagebox.showinfo(APP_NAME, f"Exporté : {path}")

    def backup_db(self):
        """Sauvegarder la base SQLite avec nom + extension par défaut."""
        default_name = f"history_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        path = filedialog.asksaveasfilename(
            title="Sauvegarder la base de données",
            defaultextension=".db",
            initialfile=default_name,
            filetypes=[("Base SQLite (*.db)", "*.db"), ("Tous les fichiers", "*.*")]
        )
        if not path:
            return
        try:
            shutil.copy2(DB_PATH, path)
            messagebox.showinfo(APP_NAME, f"Base sauvegardée avec succès :\n{path}")
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Erreur lors de la sauvegarde :\n{e}")

# ===================== BOOTSTRAP =====================
def first_scan_if_needed(con):
    c = con.cursor(); c.execute("SELECT COUNT(*) FROM items")
    if c.fetchone()[0] == 0:
        scan_recent(con)

def main_gui():
    con = ensure_db()
    first_scan_if_needed(con)
    App().mainloop()

if __name__ == "__main__":
    # Mode planifié
    if "--weekly-scan" in sys.argv:
        run_weekly_scan_once()
        sys.exit(0)

    if not RECENT_DIR.exists():
        tk.Tk().withdraw()
        messagebox.showerror(APP_NAME, f"Dossier 'Recent' introuvable:\n{RECENT_DIR}")
        sys.exit(1)

    main_gui()
