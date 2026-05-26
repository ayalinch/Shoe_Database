#!/usr/bin/env python3
"""
KITE UHN Shoe MAA Database
================================
Requires: pip install PySide6 openpyxl
Run:  python3 fin.py
"""

import sys
import os
import re
import json
import sqlite3
import shutil
import csv
import subprocess

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QScrollArea, QLabel, QPushButton, QLineEdit, QFrame,
    QFileDialog, QMessageBox, QDialog, QTextEdit, QProgressBar,
    QSizePolicy, QStackedWidget, QFormLayout, QLayout, QSplitter,
    QTabWidget, QMenu
)
from PySide6.QtCore import Qt, QSize, QThread, Signal, QTimer, QRect, QPoint, QMutex, QWaitCondition, QPointF
from PySide6.QtGui import (
    QPixmap, QColor, QFont,
    QPainter, QPainterPath, QBrush, QPen, QImage
)

# ── Config ──────────────────────────────────────────────────────────────────
BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
DB_PATH      = os.path.join(BASE_DIR, "shoes.db")
IMG_DIR      = os.path.join(BASE_DIR, "images")
LOGO_PATH    = os.path.join(BASE_DIR, "68556ca78f14ebbed4120b97_Blue-KITE.png")
EXCEL_PATH   = os.path.join(BASE_DIR, "WinterLab Master list of footwear.xlsx")
REPORTS_DIR  = os.path.join(BASE_DIR, "Photos and Reports")
IMG_EXTS     = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff"}

THUMB_W  = 160
THUMB_H  = 130
CARD_W   = THUMB_W + 20
CARD_H   = THUMB_H + 50

# ── Global pixmap cache ──────────────────────────────────────────────────────
_PIXMAP_CACHE = {}

# ── Excel Lab Fields ────────────────────────────────────────────────────────
LAB_FIELDS = [
    "Safety?", "MAA for wet ice", "MAA for cold ice", "Final MAA score",
    "Listed on the RMT website (Y/N)", "RMT/Yes/Pass", "Prototype", "Client",
    "Type", "technology", "Size weighed", "Weight (g)", "Upper", "Shoe",
    "Date Manufactured", "Date Received", "Date returned", "Date to complete test",
    "Date Tested", "Report date", "Date sent to client", "Report #",
    "Test surface", "Protocol", "weight of left shoe", "weight of right shoe",
    "test id", "Hardness R1", "Hardness R2", "Hardness R3", "HardnessL1",
    "HardnessL2", "HardnessL3", "Variation Within Spot <5", "Hardness R1.1",
    "Hardness R2.1", "Hardness R3.1", "Hardness L1", "Hardness L2", "Hardness L3",
    "Variation Within Spot <5.1", "Hardness measurement date", "SATRA test date",
    "Website", "Size Range", "Features/Upper", "Insulation", "Height",
    "Sole (Inner/Midsole/Outersole)", "Standard", "To test", "Repeated test production",
    "Repeated test prototype", "MAA Received 1", "Received 2", "Received 3", "Received 4"
]

# ── Themes ──────────────────────────────────────────────────────────────────
THEMES = {
    "dark": {
        "BG": "#09090b", "SURFACE": "#111113", "SURFACE2": "#18181c", "BORDER": "#1e1e2c",
        "TEXT": "#ededf0", "TEXT2": "#9898ab", "DIM": "#55556a", "FAINT": "#2a2a38",
        "ACCENT": "#3b82f6", "GREEN": "#34d399", "AMBER": "#fbbf24", "RED": "#f87171",
        "CARD": "#111115", "CARD_HOVER": "#1a1a22", "SIDEBAR": "#0d0d11",
        "INPUT_BG": "#1a1a22", "INPUT_BORDER": "#3b82f6",
    },
    "light": {
        "BG": "#f4f4f5", "SURFACE": "#ffffff", "SURFACE2": "#e4e4e7", "BORDER": "#d4d4d8",
        "TEXT": "#18181b", "TEXT2": "#52525b", "DIM": "#71717a", "FAINT": "#e4e4e7",
        "ACCENT": "#2563eb", "GREEN": "#10b981", "AMBER": "#f59e0b", "RED": "#ef4444",
        "CARD": "#ffffff", "CARD_HOVER": "#f0f0f3", "SIDEBAR": "#fafafa",
        "INPUT_BG": "#ffffff", "INPUT_BORDER": "#a1a1aa",
    }
}

CURRENT_MODE = "light"

def get_c():
    return THEMES[CURRENT_MODE]

def generate_css():
    c = get_c()
    return f"""
    * {{ font-family: "Helvetica Neue", Helvetica, Arial; }}

    QMainWindow, QDialog, QWidget {{
        background-color: {c['BG']}; color: {c['TEXT']};
    }}
    QWidget#db_page, QWidget#grid_widget {{ background-color: {c['BG']}; }}
    QWidget#sidebar, QWidget#sidebar_inner {{ background-color: {c['SIDEBAR']}; }}
    QScrollArea#strip_scroll, QScrollArea#strip_scroll > QWidget > QWidget {{ background-color: {c['SIDEBAR']}; }}
    QFrame#gallery {{ background-color: {c['SIDEBAR']}; }}
    QFrame#topbar {{ background-color: {c['SURFACE']}; border-bottom: 1px solid {c['BORDER']}; }}
    QLabel#topbar_title {{ font-size:16px; font-weight:700; color:{c['TEXT']}; }}
    QScrollArea {{ border:none; background:transparent; }}
    QScrollBar:vertical {{ background:transparent; width:8px; margin:0; }}
    QScrollBar::handle:vertical {{ background:{c['FAINT']}; border-radius:4px; min-height:20px; }}
    QScrollBar::handle:vertical:hover {{ background:{c['DIM']}; }}

    QLineEdit, QTextEdit {{
        background-color: {c['INPUT_BG']}; color: {c['TEXT']};
        border: 1px solid {c['INPUT_BORDER']}; border-radius: 8px;
        padding: 10px; font-size: 13px; selection-background-color: {c['ACCENT']};
    }}
    QLineEdit:hover, QTextEdit:hover {{ border: 1px solid {c['ACCENT']}; }}
    QLineEdit:focus, QTextEdit:focus {{ border: 2px solid {c['ACCENT']}; }}

    QTabWidget::pane {{ border: 1px solid {c['BORDER']}; background: {c['BG']}; border-radius: 6px; }}
    QTabBar::tab {{ background: {c['SURFACE2']}; color: {c['TEXT2']}; padding: 8px 16px; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 2px; }}
    QTabBar::tab:selected {{ background: {c['ACCENT']}; color: #ffffff; font-weight: bold; }}

    QMenu {{ background-color: {c['SURFACE']}; color: {c['TEXT']}; border: 1px solid {c['BORDER']}; padding: 4px; }}
    QMenu::item {{ padding: 6px 24px; border-radius: 4px; }}
    QMenu::item:selected {{ background-color: {c['FAINT']}; }}

    QProgressBar {{ background:{c['SURFACE2']}; border:none; border-radius:4px; height:6px; }}
    QProgressBar::chunk {{ background:{c['ACCENT']}; border-radius:4px; }}
    QLabel {{ background:transparent; color:{c['TEXT']}; }}

    QSplitter::handle {{ background:{c['BORDER']}; width:1px; }}
    QSplitter::handle:horizontal:hover {{ background:{c['ACCENT']}; }}

    QWidget#home_widget {{ background: #ffffff; }}
    QLabel#home_title {{ font-size:38px; font-weight:800; color:#18181b; letter-spacing:1px; }}
    QLabel#home_subtitle {{ font-size:16px; color:#52525b; font-weight:500; }}

    QPushButton#home_btn {{
        background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 {c['ACCENT']},stop:1 #1d4ed8);
        color:#ffffff; font-size:16px; font-weight:bold; border-radius:25px;
    }}
    QPushButton#home_btn:hover {{ background:#1d4ed8; }}

    QPushButton.icon_btn {{
        background:{c['SURFACE2']}; color:{c['TEXT']}; border-radius:16px;
        min-width:50px; min-height:32px; font-size:16px; border:none;
    }}
    QPushButton.icon_btn:hover {{ background:{c['FAINT']}; }}

    QPushButton.primary_icon_btn {{
        background:{c['ACCENT']}; color:#ffffff; border-radius:16px;
        min-width:60px; min-height:32px; font-size:18px; font-weight:bold; border:none;
    }}
    QPushButton.primary_icon_btn:hover {{ background:#2563eb; }}

    QPushButton.action_btn {{ background:{c['SURFACE2']}; color:{c['TEXT']}; border-radius:8px; padding:8px 16px; font-weight:600; border:none; font-size:13px; }}
    QPushButton.action_btn:hover {{ background:{c['FAINT']}; }}
    QPushButton.primary_btn {{ background:{c['ACCENT']}; color:#ffffff; border-radius:8px; padding:8px 16px; font-weight:600; border:none; font-size:13px; }}
    QPushButton.primary_btn:hover {{ background:#2563eb; }}
    QPushButton.danger_btn {{ background:#fef2f2; color:{c['RED']}; border-radius:8px; padding:8px 16px; font-weight:600; border:1px solid {c['RED']}; font-size:13px; }}
    QPushButton.danger_btn:hover {{ background:{c['RED']}; color:#ffffff; }}
    QPushButton.report_btn {{ background:#f0fdf4; color:#16a34a; border-radius:8px; padding:8px 16px; font-weight:600; border:1px solid #16a34a; font-size:13px; }}
    QPushButton.report_btn:hover {{ background:#16a34a; color:#ffffff; }}

    QWidget#splash_widget {{ background:{c['BG']}; }}
    
    ShoeCard[state="normal"] {{ background: {c['CARD']}; border-radius: 10px; border: 2px solid transparent; }}
    ShoeCard[state="normal"]:hover {{ background: {c['CARD_HOVER']}; border: 2px solid {c['FAINT']}; }}
    ShoeCard[state="pinned"] {{ background: {c['CARD_HOVER']}; border-radius: 10px; border: 2px solid {c['AMBER']}; }}
    ShoeCard[state="selected"] {{ background: {c['CARD_HOVER']}; border-radius: 10px; border: 2px solid {c['ACCENT']}; }}
    
    ShoeCard QLabel#card_brand {{ color: {c['DIM']}; font-size: 10px; }}
    ShoeCard QLabel#card_name {{ color: {c['TEXT']}; font-weight: 600; font-size: 11px; }}
    """

# ── Splash Screen ─────────────────────────────────────────────────────────────
class SplashScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self.setFixedSize(480, 300)
        self.setStyleSheet("background:#ffffff; border-radius:18px;")
        screen = QApplication.primaryScreen().geometry()
        self.move((screen.width()-self.width())//2, (screen.height()-self.height())//2)
        root = QVBoxLayout(self); root.setContentsMargins(0,0,0,0)
        card = QFrame(); card.setObjectName("splash_card")
        card.setStyleSheet("""
            QFrame#splash_card{background:#ffffff;border-radius:18px;border:1px solid #e4e4e7;}
            QFrame#splash_card QLabel{border:none;background:transparent;color:#18181b;}
        """)
        cl = QVBoxLayout(card); cl.setContentsMargins(40,40,40,36); cl.setSpacing(14); cl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.logo_lbl = QLabel(); self.logo_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        if os.path.exists(LOGO_PATH):
            pix = QPixmap(LOGO_PATH).scaled(260,100,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
            self.logo_lbl.setPixmap(pix)
        else:
            self.logo_lbl.setText("🍁 UHN KITE"); self.logo_lbl.setStyleSheet("font-size:32px;font-weight:bold;color:#e11d48;")
        cl.addWidget(self.logo_lbl)
        self.status_lbl = QLabel("Initialising database…"); self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_lbl.setStyleSheet("color:#52525b;font-size:13px;font-weight:500;background:transparent;")
        cl.addWidget(self.status_lbl)
        self.bar = QProgressBar(); self.bar.setRange(0,100); self.bar.setValue(0); self.bar.setTextVisible(False); self.bar.setFixedHeight(6)
        self.bar.setStyleSheet("QProgressBar{background:#e4e4e7;border:none;border-radius:3px;}QProgressBar::chunk{background:#2563eb;border-radius:3px;}")
        cl.addWidget(self.bar)
        ver = QLabel("Footwear Evaluation Database · KITE Research Institute"); ver.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ver.setStyleSheet("color:#a1a1aa;font-size:10px;background:transparent;"); cl.addWidget(ver)
        root.addWidget(card)
        self._val=0; self._timer=QTimer(self); self._timer.timeout.connect(self._tick); self._timer.start(20)

    def _tick(self):
        self._val+=2
        if self._val<=30: self.status_lbl.setText("Initialising database…")
        elif self._val<=60: self.status_lbl.setText("Loading shoe records…")
        elif self._val<=85: self.status_lbl.setText("Preparing interface…")
        else: self.status_lbl.setText("Ready!")
        self.bar.setValue(min(self._val,100))
        if self._val>=100: self._timer.stop()

    def set_status(self, text): self.status_lbl.setText(text)


# ── Custom Widgets ───────────────────────────────────────────────────────────
class MoreOptionsButton(QPushButton):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(32, 32)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        
    def paintEvent(self, event):
        super().paintEvent(event)
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        c = get_c()
        color = QColor(c['TEXT'])
        p.setBrush(color)
        p.setPen(Qt.PenStyle.NoPen)
        cx = self.width() / 2
        cy = self.height() / 2
        r = 2.0
        p.drawEllipse(QPointF(cx, cy - 6), r, r)
        p.drawEllipse(QPointF(cx, cy), r, r)
        p.drawEllipse(QPointF(cx, cy + 6), r, r)
        p.end()

class ClickableLabel(QLabel):
    clicked = Signal(int)
    def __init__(self, idx, parent=None):
        super().__init__(parent); self.idx=idx
    def mousePressEvent(self, e):
        if e.button()==Qt.MouseButton.LeftButton: self.clicked.emit(self.idx)
        super().mousePressEvent(e)

class DynamicScrollArea(QScrollArea):
    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.widget(): self.widget().setFixedWidth(self.viewport().width())

# ── Database ─────────────────────────────────────────────────────────────────
def init_db():
    con=sqlite3.connect(DB_PATH); cur=con.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS shoes(
        id INTEGER PRIMARY KEY AUTOINCREMENT,idapt_id TEXT,brand TEXT,model TEXT,year TEXT,size TEXT,
        sole_type TEXT,tread TEXT,maa_mean REAL,maa_sd REAL,notes TEXT,image_path TEXT,all_images TEXT
    )""")
    for col,t in [("idapt_id","TEXT"),("all_images","TEXT"),("lab_data","TEXT"),("report_path","TEXT"),("all_reports","TEXT")]:
        try: cur.execute(f"ALTER TABLE shoes ADD COLUMN {col} {t}")
        except: pass
    con.commit(); con.close()

def get_all_shoes(q="", search_mode="all"):
    con=sqlite3.connect(DB_PATH); con.row_factory=sqlite3.Row
    order="ORDER BY LENGTH(idapt_id) ASC, idapt_id ASC, brand ASC"
    if q:
        like=f"%{q}%"
        rows=con.execute(f"SELECT * FROM shoes WHERE brand LIKE ? OR model LIKE ? OR sole_type LIKE ? OR idapt_id LIKE ? OR lab_data LIKE ? {order}",(like,like,like,like,like)).fetchall()
    else:
        rows=con.execute(f"SELECT * FROM shoes {order}").fetchall()
    con.close(); return [dict(r) for r in rows]

def save_shoe(data, shoe_id=None):
    def f(k): return data.get(k) or None
    def fn(k):
        v=data.get(k,"")
        try: return float(v) if v not in (None,"") else None
        except: return None
    vals=(f("idapt_id"),f("brand"),f("model"),f("year"),f("size"),f("sole_type"),f("tread"),fn("maa_mean"),fn("maa_sd"),f("notes"),f("image_path"),f("lab_data"))
    con=sqlite3.connect(DB_PATH)
    if shoe_id: con.execute("UPDATE shoes SET idapt_id=?,brand=?,model=?,year=?,size=?,sole_type=?,tread=?,maa_mean=?,maa_sd=?,notes=?,image_path=?,lab_data=? WHERE id=?",(*vals,shoe_id))
    else: con.execute("INSERT INTO shoes(idapt_id,brand,model,year,size,sole_type,tread,maa_mean,maa_sd,notes,image_path,lab_data)VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",vals)
    con.commit(); con.close()

def delete_shoe(shoe_id):
    con=sqlite3.connect(DB_PATH); con.execute("DELETE FROM shoes WHERE id=?",(shoe_id,)); con.commit(); con.close()

# ── Sync Worker ───────────────────────────────────────────────────────────────
class SyncWorker(QThread):
    progress = Signal(str)
    done     = Signal(int, int, int)

    def run(self):
        added=updated=errors=0
        con=sqlite3.connect(DB_PATH)

        # 1. Read Excel
        excel_rows={}
        if os.path.exists(EXCEL_PATH):
            try:
                import openpyxl
                wb=openpyxl.load_workbook(EXCEL_PATH,data_only=True); ws=wb.active
                HEADERS=[
                    None,"Safety?","MAA for wet ice","MAA for cold ice","Final MAA score",
                    "Listed on the RMT website (Y/N)","RMT/Yes/Pass","Prototype","Client",
                    "Brand","Name","Model#","Type","technology","Size","Size weighed",
                    "Weight (g)","Upper","Date Manufactured","Date Received","Date returned",
                    "Date to complete test","Date Tested","Report date","Date sent to client",
                    "Report #",None,"Test surface","Protocol","weight of left shoe",
                    "weight of right shoe","test id","Hardness R1","Hardness R2","Hardness R3",
                    "HardnessL1","HardnessL2","HardnessL3","Variation Within Spot <5",
                    "Hardness R1.1","Hardness R2.1","Hardness R3.1","Hardness L1","Hardness L2",
                    "Hardness L3","Variation Within Spot <5.1","Hardness measurement date",
                    "SATRA test date","Website","Size Range","Features/Upper","Insulation",
                    "Height","Sole (Inner/Midsole/Outersole)","Standard","To test",
                    "Repeated test production","Repeated test prototype","MAA",
                    "Received 1","Received 2","Received 3","Received 4"
                ]
                for row in ws.iter_rows(min_row=2,values_only=True):
                    raw_id=row[0]
                    if not raw_id: continue
                    idapt_id=str(raw_id).strip().upper()
                    if not idapt_id.startswith("IDAPT"): continue
                    lab={}
                    for i,h in enumerate(HEADERS):
                        if h and i<len(row) and row[i] not in (None,""):
                            lab[h]=str(row[i]).strip()
                    brand=str(row[9] or "").strip()
                    model=str(row[10] or "").strip()
                    size=str(row[14] or "").strip()
                    maa_raw=row[58]
                    try: maa=float(maa_raw) if maa_raw not in (None,"") else None
                    except: maa=None
                    excel_rows[idapt_id]={"brand":brand,"model":model,"size":size,"maa":maa,"lab":lab}
                self.progress.emit(f"📊 Read {len(excel_rows)} rows from Excel")
            except Exception as e:
                self.progress.emit(f"⚠️ Excel error: {e}"); errors+=1
        else:
            self.progress.emit("⚠️ Excel not found — skipping")

        # 2. Scan Photos and Reports
        folder_data={}
        if os.path.exists(REPORTS_DIR):
            for entry in os.listdir(REPORTS_DIR):
                path=os.path.join(REPORTS_DIR,entry)
                if not os.path.isdir(path): continue
                m=re.match(r"(iDAPT\d+)",entry,re.IGNORECASE)
                if not m: continue
                idapt_id=m.group(1).upper()
                imgs=sorted([os.path.join(path,f) for f in os.listdir(path) if os.path.splitext(f)[1].lower() in IMG_EXTS])
                pdfs=sorted([os.path.join(path,f) for f in os.listdir(path) if f.lower().endswith(".pdf")])
                best=next((i for i in imgs if "side" in os.path.basename(i).lower()),imgs[0] if imgs else "")
                folder_data[idapt_id]={"imgs":imgs,"pdfs":pdfs,"best":best}
            self.progress.emit(f"📁 Found {len(folder_data)} iDAPT folders")
        else:
            self.progress.emit("⚠️ Photos and Reports folder not found — skipping")

        # 3. Merge
        existing={r[0]:r[1] for r in con.execute("SELECT idapt_id,id FROM shoes WHERE idapt_id IS NOT NULL")}
        all_ids=set(excel_rows.keys())|set(folder_data.keys())

        for idapt_id in sorted(all_ids):
            ex=excel_rows.get(idapt_id,{})
            fld=folder_data.get(idapt_id,{})
            brand=ex.get("brand",""); model=ex.get("model",""); size=ex.get("size","")
            maa=ex.get("maa"); lab_json=json.dumps(ex.get("lab",{})) if ex.get("lab") else ""
            imgs=fld.get("imgs",[]); pdfs=fld.get("pdfs",[]); best_img=fld.get("best","")

            if idapt_id in existing:
                con.execute("""UPDATE shoes SET brand=?,model=?,size=?,maa_mean=?,
                    lab_data=?,image_path=?,all_images=?,report_path=?,all_reports=?
                    WHERE idapt_id=?""",
                    (brand,model,size,maa,lab_json,best_img,json.dumps(imgs),
                     pdfs[0] if pdfs else "",json.dumps(pdfs),idapt_id))
                updated+=1; self.progress.emit(f"↻  Updated {idapt_id} — {brand} {model}".strip())
            else:
                con.execute("""INSERT INTO shoes(idapt_id,brand,model,size,maa_mean,
                    lab_data,image_path,all_images,report_path,all_reports)VALUES(?,?,?,?,?,?,?,?,?,?)""",
                    (idapt_id,brand,model,size,maa,lab_json,best_img,json.dumps(imgs),
                     pdfs[0] if pdfs else "",json.dumps(pdfs)))
                added+=1; self.progress.emit(f"✚  Added {idapt_id} — {brand} {model}".strip())

        con.commit(); con.close()
        self.done.emit(added,updated,errors)


class SyncDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Sync"); self.setMinimumSize(500,400); self.resize(500,460)
        self.setStyleSheet("""
            QDialog { background:#ffffff; color:#18181b; }
            QLabel  { background:transparent; color:#18181b; }
            QTextEdit { background:#f8f8f8; color:#18181b; border:1px solid #e4e4e7; border-radius:6px; font-size:12px; }
            QProgressBar { background:#e4e4e7; border:none; border-radius:3px; height:6px; }
            QProgressBar::chunk { background:#2563eb; border-radius:3px; }
            QPushButton { background:#2563eb; color:#ffffff; border-radius:8px; padding:8px 24px; font-weight:600; border:none; font-size:13px; }
            QPushButton:hover { background:#1d4ed8; }
            QPushButton:disabled { background:#e4e4e7; color:#a1a1aa; }
        """)
        lay = QVBoxLayout(self); lay.setContentsMargins(20,20,20,20); lay.setSpacing(12)

        self.title_lbl = QLabel("Syncing from Excel + Photos folder…")
        self.title_lbl.setStyleSheet("font-size:15px; font-weight:700; color:#18181b;")
        lay.addWidget(self.title_lbl)

        self.bar = QProgressBar(); self.bar.setRange(0,0); self.bar.setFixedHeight(6); lay.addWidget(self.bar)

        self.log = QTextEdit(); self.log.setReadOnly(True); lay.addWidget(self.log)

        self.ok_btn = QPushButton("OK — Close"); self.ok_btn.setEnabled(False)
        self.ok_btn.clicked.connect(self.accept)
        lay.addWidget(self.ok_btn, alignment=Qt.AlignmentFlag.AlignRight)

        self._worker = SyncWorker()
        self._worker.progress.connect(self._on_progress)
        self._worker.done.connect(self._on_done)
        self._worker.start()

    def _on_progress(self, msg):
        self.log.append(msg)
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def _on_done(self, added, updated, errors):
        self.bar.setRange(0,100); self.bar.setValue(100)
        summary = f"✅  Done — {added} new shoes added, {updated} updated"
        if errors: summary += f", {errors} errors"
        self.title_lbl.setText(summary)
        self.log.append(f"\n{summary}")
        self.ok_btn.setEnabled(True)


# ── Helpers ──────────────────────────────────────────────────────────────────
def parse_idapt_folder(name):
    m=re.match(r"(iDAPT\d+)\s*(.*)",name,re.IGNORECASE)
    if not m: return None,name,[]
    idapt_id=m.group(1).upper(); rest=m.group(2).strip(); flags=[]
    for flag in ["NO REPORT","NO PHOTOS"]:
        if rest.upper().endswith(flag): flags.append(flag); rest=rest[:-(len(flag))].strip()
    return idapt_id,rest.strip(" -_"),flags

def find_best_image(folder):
    try: files=os.listdir(folder)
    except: return None
    imgs=[f for f in files if os.path.splitext(f)[1].lower() in IMG_EXTS]
    if not imgs: return None
    for f in imgs:
        if "side" in f.lower(): return os.path.join(folder,f)
    return os.path.join(folder,imgs[0])

def get_all_images(folder):
    try: files=os.listdir(folder)
    except: return []
    return [os.path.join(folder,f) for f in sorted(files) if os.path.splitext(f)[1].lower() in IMG_EXTS]

def load_pixmap(path, w, h):
    key=(path,w,h)
    if key in _PIXMAP_CACHE: return _PIXMAP_CACHE[key]
    if not path or not os.path.isfile(path):
        result=_placeholder(w,h)
    else:
        px=QPixmap(path)
        if px.isNull(): result=_placeholder(w,h)
        else:
            scaled=px.scaled(w,h,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
            canvas=QPixmap(w,h); canvas.fill(QColor("#ffffff"))
            p=QPainter(canvas); x=(w-scaled.width())//2; y=(h-scaled.height())//2
            p.drawPixmap(x,y,scaled); p.end()
            result=canvas
    _PIXMAP_CACHE[key]=result
    return result

def _placeholder(w,h):
    c=get_c()
    px=QPixmap(w,h); px.fill(QColor("#ffffff"))
    p=QPainter(px); p.setRenderHint(QPainter.RenderHint.Antialiasing); p.setPen(Qt.PenStyle.NoPen)
    p.setBrush(QColor("#f0f0f0")); p.drawRoundedRect(6,6,w-12,h-12,8,8)
    p.setFont(QFont("Arial",24)); p.setPen(QColor(c['DIM'])); p.drawText(QRect(0,0,w,h),Qt.AlignmentFlag.AlignCenter,"👟"); p.end()
    return px

def score_color(v):
    if v is None: return get_c()['DIM']
    try: v=float(v)
    except: return get_c()['DIM']
    return get_c()['GREEN'] if v>=30 else get_c()['AMBER'] if v>=20 else get_c()['RED']

def mk_btn(text, cls, parent=None):
    b=QPushButton(text,parent); b.setProperty("class",cls); return b

def open_pdf(path):
    if not path or not os.path.isfile(path):
        QMessageBox.warning(None,"Report Not Found",f"Could not find:\n{path}"); return
    
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])

# ── Home Widget ───────────────────────────────────────────────────────────────
class HomeWidget(QWidget):
    def __init__(self, enter_callback, parent=None):
        super().__init__(parent); self.setObjectName("home_widget")
        layout=QVBoxLayout(self); layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_lbl=QLabel()
        if os.path.exists(LOGO_PATH):
            pix=QPixmap(LOGO_PATH).scaled(500,250,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation)
            logo_lbl.setPixmap(pix)
        else:
            logo_lbl.setText("🍁 UHN KITE"); logo_lbl.setStyleSheet("font-size:50px;font-weight:bold;color:#e11d48;")
        logo_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(logo_lbl); layout.addSpacing(30)
        t1=QLabel("Footwear Evaluation Database"); t1.setObjectName("home_title"); t1.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(t1)
        t2=QLabel("Designed & Engineered by Ayalinch Jonathan and Yevgeniy Korolyov"); t2.setObjectName("home_subtitle"); t2.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(t2)
        layout.addSpacing(50)
        btn=QPushButton("Access Database →"); btn.setFixedSize(240,50); btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setObjectName("home_btn"); btn.clicked.connect(enter_callback); layout.addWidget(btn,alignment=Qt.AlignmentFlag.AlignCenter)

    def paintEvent(self,event):
        painter=QPainter(self); painter.fillRect(self.rect(),QColor("#ffffff")); painter.end()

# ── FlowLayout ────────────────────────────────────────────────────────────────
class FlowLayout(QLayout):
    def __init__(self,parent=None,margin=0,spacing=10):
        super().__init__(parent); self.setContentsMargins(margin,margin,margin,margin); self.setSpacing(spacing); self._items=[]
    def __del__(self):
        item=self.takeAt(0)
        while item: item=self.takeAt(0)
    def addItem(self,item): self._items.append(item)
    def count(self): return len(self._items)
    def itemAt(self,index): return self._items[index] if 0<=index<len(self._items) else None
    def takeAt(self,index): return self._items.pop(index) if 0<=index<len(self._items) else None
    def expandingDirections(self): return Qt.Orientation(0)
    def hasHeightForWidth(self): return True
    def heightForWidth(self,width): return self._do_layout(QRect(0,0,width,0),True)
    def setGeometry(self,rect): super().setGeometry(rect); self._do_layout(rect,False)
    def sizeHint(self): return self.minimumSize()
    def minimumSize(self):
        size=QSize()
        for item in self._items: size=size.expandedTo(item.minimumSize())
        m=self.contentsMargins(); size+=QSize(m.left()+m.right(),m.top()+m.bottom()); return size
    def _do_layout(self,rect,test_only):
        x=rect.x(); y=rect.y(); line_height=0; spacing=self.spacing()
        for item in self._items:
            next_x=x+item.sizeHint().width()+spacing
            if next_x-spacing>rect.right() and line_height>0:
                x=rect.x(); y=y+line_height+spacing; next_x=x+item.sizeHint().width()+spacing; line_height=0
            if not test_only: item.setGeometry(QRect(QPoint(x,y),item.sizeHint()))
            x=next_x; line_height=max(line_height,item.sizeHint().height())
        return y+line_height-rect.y()

# ── Shoe Card ─────────────────────────────────────────────────────────────────
class ShoeCard(QFrame):
    clicked = Signal(dict)
    pin_requested = Signal(dict)

    def __init__(self, shoe, parent=None):
        super().__init__(parent); self.shoe=shoe; self._selected=False
        self.setFixedSize(CARD_W,CARD_H); self.setCursor(Qt.CursorShape.PointingHandCursor); self._pinned=False; self._build()

    def _build(self):
        lay=QVBoxLayout(self); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        img_container=QWidget(); img_container.setFixedSize(CARD_W,THUMB_H)
        img_container.setStyleSheet("background:#ffffff;border-radius:8px 8px 0 0;")
        ic_lay=QVBoxLayout(img_container); ic_lay.setContentsMargins(0,0,0,0)
        self.img_lbl=QLabel(); self.img_lbl.setFixedSize(CARD_W,THUMB_H)
        self.img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter); self.img_lbl.setStyleSheet("background:transparent;")
        self.img_lbl.setPixmap(_placeholder(CARD_W,THUMB_H))
        ic_lay.addWidget(self.img_lbl); lay.addWidget(img_container)
        nw=QWidget(); nw.setStyleSheet("background:transparent;"); nl=QVBoxLayout(nw)
        nl.setContentsMargins(8,5,8,8); nl.setSpacing(2)
        idapt,brand,model=self.shoe.get("idapt_id",""),self.shoe.get("brand",""),self.shoe.get("model","")
        display=(f"{brand} {model}".strip() if brand and brand!=idapt else f"{idapt} {model}".strip()) or idapt or "?"
        self.brand_lbl=QLabel(brand if brand and brand!=idapt else "")
        self.brand_lbl.setObjectName("card_brand")
        self.brand_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.name_lbl=QLabel(display)
        self.name_lbl.setObjectName("card_name")
        self.name_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.name_lbl.setWordWrap(True)
        nl.addWidget(self.brand_lbl); nl.addWidget(self.name_lbl); lay.addWidget(nw)
        sc=self.shoe.get("maa_mean"); self._score_txt=f"{float(sc):.1f}°" if sc not in (None,"") else ""
        self.setProperty("state", "normal")

    def update_shoe(self, shoe):
        self.shoe = shoe
        idapt,brand,model=shoe.get("idapt_id",""),shoe.get("brand",""),shoe.get("model","")
        display=(f"{brand} {model}".strip() if brand and brand!=idapt else f"{idapt} {model}".strip()) or idapt or "?"
        self.brand_lbl.setText(brand if brand and brand!=idapt else "")
        self.name_lbl.setText(display)
        sc=shoe.get("maa_mean")
        self._score_txt=f"{float(sc):.1f}°" if sc not in (None,"") else ""
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self._score_txt: return
        p=QPainter(self); p.setRenderHint(QPainter.RenderHint.Antialiasing); p.setFont(QFont("Courier New",9,QFont.Weight.Bold))
        tw=p.fontMetrics().horizontalAdvance(self._score_txt)+14; rx,ry,th=self.width()-tw-5,5,18
        pill_bg=QColor(0,0,0,150) if CURRENT_MODE=="dark" else QColor(255,255,255,200)
        
        score_c = score_color(self.shoe.get("maa_mean"))
        
        p.setBrush(QBrush(pill_bg)); p.setPen(QPen(QColor(score_c),1.5))
        path=QPainterPath(); path.addRoundedRect(rx,ry,tw,th,9,9); p.drawPath(path)
        p.setPen(QColor(score_c)); p.drawText(QRect(int(rx),int(ry),int(tw),int(th)),Qt.AlignmentFlag.AlignCenter,self._score_txt); p.end()

    def set_selected(self, sel):
        if self._selected == sel: return
        self._selected=sel; self._update_state()

    def set_pinned(self, pinned):
        if self._pinned == pinned: return
        self._pinned=pinned; self._update_state()

    def _update_state(self):
        if self._selected:
            self.setProperty("state", "selected")
        elif self._pinned:
            self.setProperty("state", "pinned")
        else:
            self.setProperty("state", "normal")
        self.style().unpolish(self)
        self.style().polish(self)

    def mousePressEvent(self, e):
        if e.button()==Qt.MouseButton.LeftButton: self.clicked.emit(self.shoe)
        elif e.button()==Qt.MouseButton.RightButton: self.pin_requested.emit(self.shoe)
        super().mousePressEvent(e)

# ── Sidebar ───────────────────────────────────────────────────────────────────
class SidebarEmpty(QWidget):
    def __init__(self):
        super().__init__(); self.setObjectName("sidebar")
        l=QVBoxLayout(self); l.setAlignment(Qt.AlignmentFlag.AlignCenter); l.setSpacing(8); self.lbls=[]
        for txt in ["ℹ","Select a shoe","to see details"]:
            lbl=QLabel(txt); lbl.setAlignment(Qt.AlignmentFlag.AlignCenter); l.addWidget(lbl); self.lbls.append(lbl)
        self.refresh_theme()

    def refresh_theme(self):
        c=get_c()
        self.lbls[0].setStyleSheet(f"color:{c['DIM']};font-size:32px;")
        self.lbls[1].setStyleSheet(f"color:{c['TEXT2']};font-weight:600;font-size:14px;")
        self.lbls[2].setStyleSheet(f"color:{c['DIM']};font-size:12px;")

class DetailLoadWorker(QThread):
    finished_load = Signal(dict, list, list, list)
    
    def __init__(self, shoe, parent=None):
        super().__init__(parent)
        self.shoe = shoe
        self._is_stopped = False
        
    def stop(self):
        self._is_stopped = True
        
    def run(self):
        imgs = []
        reports = []
        shoe = self.shoe
        try: imgs = [p for p in json.loads(shoe.get("all_images") or "[]") if os.path.isfile(p)]
        except: pass
        if not imgs and shoe.get("image_path") and os.path.isfile(shoe["image_path"]):
            imgs = [shoe["image_path"]]
            
        try: reports = [p for p in json.loads(shoe.get("all_reports") or "[]") if os.path.isfile(p)]
        except: pass
        if not reports and shoe.get("report_path") and os.path.isfile(shoe["report_path"]):
            reports = [shoe["report_path"]]
            
        if self._is_stopped: return
        
        qimages_data = []
        if imgs:
            path = imgs[0]
            img = QImage(path)
            if not img.isNull() and not self._is_stopped:
                scaled = img.scaled(330, 206, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                canvas = QImage(330, 206, QImage.Format.Format_ARGB32_Premultiplied)
                canvas.fill(QColor("#ffffff"))
                p = QPainter(canvas)
                p.drawImage((330 - scaled.width()) // 2, (206 - scaled.height()) // 2, scaled)
                p.end()
                qimages_data.append((path, 330, 206, canvas))
            
            for path in imgs[:8]:
                if self._is_stopped: return
                img = QImage(path)
                if not img.isNull():
                    scaled = img.scaled(42, 38, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                    canvas = QImage(42, 38, QImage.Format.Format_ARGB32_Premultiplied)
                    canvas.fill(QColor("#ffffff"))
                    p = QPainter(canvas)
                    p.drawImage((42 - scaled.width()) // 2, (38 - scaled.height()) // 2, scaled)
                    p.end()
                    qimages_data.append((path, 42, 38, canvas))
                    
        if not self._is_stopped:
            self.finished_load.emit(shoe, imgs, reports, qimages_data)

class SidebarDetail(QWidget):
    edit_requested=Signal(dict); delete_requested=Signal(dict)
    def __init__(self):
        super().__init__(); self.setObjectName("sidebar"); self._shoe=None; self._imgs=[]; self._reports=[]; self._idx=0; self._strip_lbls=[]; self._build()

    def _build(self):
        root=QVBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)
        self.stack = QStackedWidget(); root.addWidget(self.stack)
        
        self.loading_widget = QWidget()
        ll = QVBoxLayout(self.loading_widget)
        self.loading_lbl = QLabel("Loading details...")
        self.loading_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        ll.addWidget(self.loading_lbl)
        self.stack.addWidget(self.loading_widget)
        
        self.content_widget = QWidget()
        cl = QVBoxLayout(self.content_widget); cl.setContentsMargins(0,0,0,0); cl.setSpacing(0)
        
        self.gallery=QFrame(); self.gallery.setFixedHeight(210); self.gallery.setObjectName("gallery")
        gl=QHBoxLayout(self.gallery); gl.setContentsMargins(4,0,4,0); gl.setSpacing(4)
        self.prev_btn=mk_btn("‹","action_btn"); self.prev_btn.setFixedSize(28,28); self.prev_btn.clicked.connect(self._prev)
        self.gal_img=QLabel(); self.gal_img.setAlignment(Qt.AlignmentFlag.AlignCenter); self.gal_img.setSizePolicy(QSizePolicy.Policy.Expanding,QSizePolicy.Policy.Expanding)
        self.next_btn=mk_btn("›","action_btn"); self.next_btn.setFixedSize(28,28); self.next_btn.clicked.connect(self._next)
        gl.addWidget(self.prev_btn,0,Qt.AlignmentFlag.AlignVCenter); gl.addWidget(self.gal_img); gl.addWidget(self.next_btn,0,Qt.AlignmentFlag.AlignVCenter); cl.addWidget(self.gallery)
        self.counter=QLabel(); self.counter.setAlignment(Qt.AlignmentFlag.AlignCenter); cl.addWidget(self.counter)
        self.strip_scroll=QScrollArea(); self.strip_scroll.setFixedHeight(50)
        self.strip_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.strip_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.strip_scroll.setObjectName("strip_scroll")
        sw=QWidget(); self.strip_layout=QHBoxLayout(sw); self.strip_layout.setContentsMargins(6,4,6,4); self.strip_layout.setSpacing(4); self.strip_layout.addStretch()
        self.strip_scroll.setWidget(sw); self.strip_scroll.setWidgetResizable(True); cl.addWidget(self.strip_scroll)
        self.bar=QFrame(); self.bar.setFixedHeight(3); cl.addWidget(self.bar)
        self.scroll=QScrollArea(); self.scroll.setWidgetResizable(True)
        self.inner=QWidget(); self.inner.setObjectName("sidebar_inner")
        self.inner_l=QVBoxLayout(self.inner); self.inner_l.setContentsMargins(0,0,0,16); self.inner_l.setSpacing(0)
        self.scroll.setWidget(self.inner); cl.addWidget(self.scroll)
        
        self.stack.addWidget(self.content_widget)
        self.stack.setCurrentIndex(1)
        self._worker = None

    def refresh_theme(self):
        c=get_c()
        self.counter.setStyleSheet(f"color:{c['DIM']};font-size:11px;padding:2px;background:{c['SIDEBAR']};")
        self.bar.setStyleSheet(f"background:{c['ACCENT']};border:none;")
        self.gallery.setStyleSheet(f"QFrame#gallery{{background-color:{c['SIDEBAR']};}}")
        self.strip_scroll.setStyleSheet(f"QScrollArea{{background:{c['SIDEBAR']};border:none;}}QWidget{{background:{c['SIDEBAR']};}}")
        if hasattr(self, 'loading_lbl'): self.loading_lbl.setStyleSheet(f"color:{c['DIM']};font-size:14px;")
        if self._shoe: self._refresh_detail()

    def load(self, shoe):
        if self._worker and self._worker.isRunning():
            try:
                self._worker.finished_load.disconnect()
            except:
                pass
            self._worker.stop()
            self._worker = None
            
        self._shoe=shoe; self._imgs=[]; self._reports=[]; self._idx=0
        if not shoe:
            self._refresh_gallery(); self._refresh_strip(); self._refresh_detail(); self.stack.setCurrentIndex(1)
            return
            
        self.stack.setCurrentIndex(0)
        brand = shoe.get("brand", "")
        model = shoe.get("model", "")
        self.loading_lbl.setText(f"Loading {brand} {model}...")
        
        self._worker = DetailLoadWorker(shoe)
        self._worker.finished_load.connect(self._on_load_finished)
        self._worker.start()

    def _on_load_finished(self, shoe, imgs, reports, qimages_data):
        if self._shoe != shoe: return
        for path, w, h, qimg in qimages_data:
            _PIXMAP_CACHE[(path, w, h)] = QPixmap.fromImage(qimg)
        self._imgs = imgs; self._reports = reports; self._idx = 0
        self._refresh_gallery(); self._refresh_strip(); self._refresh_detail(); self.stack.setCurrentIndex(1)

    def _refresh_gallery(self):
        if self._imgs:
            px=load_pixmap(self._imgs[self._idx],330,206)
            self.gal_img.setPixmap(px.scaled(330,206,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation))
            self.counter.setText(f"{self._idx+1} / {len(self._imgs)}")
        else:
            self.gal_img.setPixmap(_placeholder(330,206)); self.counter.setText("No images" if self._shoe else "")
        self.prev_btn.setVisible(len(self._imgs)>1); self.next_btn.setVisible(len(self._imgs)>1)

    def _refresh_strip(self):
        while self.strip_layout.count()>1:
            item=self.strip_layout.takeAt(0)
            if item and item.widget(): item.widget().deleteLater()
        self._strip_lbls=[]; self.strip_scroll.setVisible(len(self._imgs)>1)
        for i,path in enumerate(self._imgs[:8]):
            lbl=ClickableLabel(i); lbl.setPixmap(load_pixmap(path,42,38)); lbl.setFixedSize(42,38); lbl.setCursor(Qt.CursorShape.PointingHandCursor)
            lbl.setProperty("active",i==0); lbl.clicked.connect(self._thumb_click)
            self.strip_layout.insertWidget(i,lbl); self._strip_lbls.append(lbl)

    def _thumb_click(self,idx):
        self._idx=idx; self._refresh_gallery()
        for i,lbl in enumerate(self._strip_lbls): lbl.setProperty("active",i==idx); lbl.style().unpolish(lbl); lbl.style().polish(lbl)
    def _prev(self):
        if self._imgs: self._idx=(self._idx-1)%len(self._imgs); self._refresh_gallery()
    def _next(self):
        if self._imgs: self._idx=(self._idx+1)%len(self._imgs); self._refresh_gallery()

    def _open_report(self):
        if not self._reports: return
        if len(self._reports)==1: open_pdf(self._reports[0])
        else:
            menu=QMenu(self)
            for path in self._reports:
                action=menu.addAction(os.path.basename(path)); action.setData(path)
            chosen=menu.exec(self.sender().mapToGlobal(self.sender().rect().bottomLeft()))
            if chosen: open_pdf(chosen.data())

    def _refresh_detail(self):
        c=get_c(); shoe=self._shoe
        while self.inner_l.count():
            item=self.inner_l.takeAt(0)
            if item and item.widget(): item.widget().deleteLater()
        if not shoe: return

        def mk(text,color,size=13,bold=False,wrap=False):
            l=QLabel(text); l.setStyleSheet(f"color:{color};font-size:{size}px;font-weight:{'700' if bold else '400'};background:transparent;")
            if wrap: l.setWordWrap(True)
            l.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            return l
        def divider():
            f=QFrame(); f.setFixedHeight(1); f.setStyleSheet(f"background:{c['FAINT']};border:none;margin:0 14px;"); self.inner_l.addWidget(f)
        def section(title):
            w=QWidget(); l=QHBoxLayout(w); l.setContentsMargins(14,10,14,4); l.setSpacing(8)
            t=QLabel(title.upper()); t.setStyleSheet(f"color:{c['TEXT2']};font-size:10px;font-weight:700;letter-spacing:1px;")
            line=QFrame(); line.setFrameShape(QFrame.Shape.HLine); line.setStyleSheet(f"color:{c['FAINT']};")
            l.addWidget(t); l.addWidget(line,1); self.inner_l.addWidget(w)
        def row(label,value):
            if not value or str(value).strip() in ("nan","None",""): return
            w=QWidget(); l=QHBoxLayout(w); l.setContentsMargins(14,2,14,2)
            lb=QLabel(label+":"); lb.setFixedWidth(130); lb.setStyleSheet(f"color:{c['DIM']};font-size:12px;")
            vl=QLabel(str(value)); vl.setStyleSheet(f"color:{c['TEXT']};font-size:12px;font-weight:600;"); vl.setWordWrap(True)
            vl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            l.addWidget(lb); l.addWidget(vl,1); self.inner_l.addWidget(w)

        nw=QWidget(); nl=QVBoxLayout(nw); nl.setContentsMargins(14,14,14,4); nl.setSpacing(2)
        idapt,brand,model=shoe.get("idapt_id",""),shoe.get("brand",""),shoe.get("model","")
        display=(f"{brand} {model}".strip() if brand and brand!=idapt else f"{idapt} {model}".strip()) or "Untitled"
        name_lbl=QLabel(display); name_lbl.setWordWrap(True); name_lbl.setStyleSheet(f"color:{c['TEXT']};font-weight:700;font-size:15px;")
        name_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        brand_lbl=QLabel(brand if brand and brand!=idapt else ""); brand_lbl.setStyleSheet(f"color:{c['TEXT2']};font-size:12px;")
        brand_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        nl.addWidget(name_lbl); nl.addWidget(brand_lbl); self.inner_l.addWidget(nw)

        sc=shoe.get("maa_mean")
        if sc not in (None,""):
            sd=shoe.get("maa_sd"); sc_txt=f"{float(sc):.1f}°"; sd_txt=f"σ = {float(sd):.2f}°" if sd not in (None,"") else ""
            mw=QWidget(); ml=QHBoxLayout(mw); ml.setContentsMargins(14,10,14,4)
            val_lbl=QLabel(sc_txt); val_lbl.setStyleSheet(f"color:{score_color(sc)};font-size:28px;font-weight:700;font-family:'Courier New';")
            val_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            meta=QWidget(); mm=QVBoxLayout(meta); mm.setContentsMargins(12,0,0,4); mm.setSpacing(2)
            mm.addWidget(mk("MAA Mean",c['TEXT2'],11,bold=True))
            if sd_txt: mm.addWidget(mk(sd_txt,c['DIM'],10))
            ml.addWidget(val_lbl); ml.addWidget(meta); ml.addStretch(); self.inner_l.addWidget(mw); divider()

        section("Details")
        row("iDAPT ID",shoe.get("idapt_id")); row("Year",shoe.get("year")); row("Size",shoe.get("size"))
        row("Sole Type",shoe.get("sole_type")); row("Tread Pattern",shoe.get("tread"))
        row("Images",f"{len(self._imgs)} found")
        if self._reports: row("Reports",f"{len(self._reports)} available")

        try:
            lab_data=json.loads(shoe.get("lab_data") or "{}")
            if lab_data:
                section("Lab Data")
                for k,v in lab_data.items(): row(k,v)
        except: pass

        if shoe.get("notes"):
            section("Notes")
            nw2=QWidget(); nl2=QVBoxLayout(nw2); nl2.setContentsMargins(14,0,14,8)
            note=QLabel(shoe["notes"]); note.setWordWrap(True)
            note.setStyleSheet(f"color:{c['TEXT2']};font-size:12px;background:{c['SURFACE2']};padding:8px;border-radius:6px;")
            note.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            nl2.addWidget(note); self.inner_l.addWidget(nw2)

        self.inner_l.addStretch(); divider()
        bw=QWidget(); bl=QHBoxLayout(bw); bl.setContentsMargins(14,8,14,0); bl.setSpacing(8)
        edit_btn=mk_btn("Edit","action_btn"); edit_btn.clicked.connect(lambda: self.edit_requested.emit(self._shoe))
        del_btn=mk_btn("Delete","danger_btn"); del_btn.clicked.connect(lambda: self.delete_requested.emit(self._shoe))
        bl.addWidget(edit_btn,1); bl.addWidget(del_btn,1); self.inner_l.addWidget(bw)

        if self._reports:
            rw=QWidget(); rl=QHBoxLayout(rw); rl.setContentsMargins(14,4,14,8)
            label=f"Open Report" if len(self._reports)==1 else f"Open Report ({len(self._reports)})"
            report_btn=mk_btn(label,"report_btn"); report_btn.clicked.connect(self._open_report)
            rl.addWidget(report_btn,1); self.inner_l.addWidget(rw)


# ── Form Dialog ───────────────────────────────────────────────────────────────
class ShoeForm(QDialog):
    CORE_FIELDS=[("iDAPT ID","idapt_id"),("Brand","brand"),("Model","model"),("Year","year"),("Size","size"),
                 ("Sole Type","sole_type"),("Tread Pattern","tread"),("MAA Mean (°)","maa_mean"),("MAA Std Dev","maa_sd")]

    def __init__(self, parent, shoe=None, on_save=None):
        super().__init__(parent)
        self.shoe=shoe or {}; self.on_save=on_save
        self._img_path=shoe.get("image_path","") if shoe else ""
        self._core_entries={}; self._lab_entries={}
        self.setWindowTitle("Edit Shoe" if shoe else "Add Shoe"); self.setFixedSize(500,650)
        self.setStyleSheet(generate_css()); self._build()
        self.progress_timer=QTimer(self); self.progress_timer.timeout.connect(self.update_progress); self.progress_timer.start(30); self.progress_val=0

    def update_progress(self):
        self.progress_val+=5; self.progressBar.setValue(self.progress_val)
        if self.progress_val>=100: self.progress_timer.stop(); QTimer.singleShot(500,self.progressBar.hide)

    def _build(self):
        c=get_c()
        root=QVBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)
        hdr=QFrame(); hdr.setFixedHeight(50); hdr.setStyleSheet(f"background:{c['SURFACE']};border-bottom:1px solid {c['BORDER']};")
        hl=QHBoxLayout(hdr); hl.setContentsMargins(20,0,20,0)
        t=QLabel("Edit Shoe" if self.shoe else "Add Shoe"); t.setStyleSheet(f"color:{c['TEXT']};font-size:15px;font-weight:700;")
        hl.addWidget(t); root.addWidget(hdr)
        self.progressBar=QProgressBar(self); self.progressBar.setTextVisible(False); self.progressBar.setFixedHeight(6); root.addWidget(self.progressBar)
        self.tabs=QTabWidget(); self.tabs.setContentsMargins(10,10,10,10); root.addWidget(self.tabs)

        core_scroll=QScrollArea(); core_scroll.setWidgetResizable(True)
        core_body=QWidget(); core_body.setStyleSheet(f"background:{c['BG']};")
        core_form=QFormLayout(core_body); core_form.setContentsMargins(20,16,20,16); core_form.setSpacing(15)
        core_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        def flbl(text): l=QLabel(text); l.setStyleSheet(f"color:{c['TEXT2']};font-size:13px;font-weight:600;background:transparent;"); return l
        for label,key in self.CORE_FIELDS:
            e=QLineEdit(); v=self.shoe.get(key,"")
            if v not in (None,""): e.setText(str(v))
            core_form.addRow(flbl(label),e); self._core_entries[key]=e
        self._notes=QTextEdit(); self._notes.setFixedHeight(80)
        if self.shoe.get("notes"): self._notes.setPlainText(self.shoe["notes"])
        core_form.addRow(flbl("Notes"),self._notes)
        img_row=QWidget(); img_row.setStyleSheet("background:transparent;"); il=QHBoxLayout(img_row); il.setContentsMargins(0,0,0,0); il.setSpacing(8)
        self._img_lbl=QLabel(os.path.basename(self._img_path) if self._img_path else "No image selected")
        self._img_lbl.setStyleSheet(f"color:{c['DIM']};font-size:11px;")
        browse=mk_btn("Browse…","action_btn"); browse.setFixedWidth(100); browse.clicked.connect(self._pick_img)
        il.addWidget(self._img_lbl,1); il.addWidget(browse); core_form.addRow(flbl("Image"),img_row)
        core_scroll.setWidget(core_body); self.tabs.addTab(core_scroll,"Core Details")

        lab_scroll=QScrollArea(); lab_scroll.setWidgetResizable(True)
        lab_body=QWidget(); lab_body.setStyleSheet(f"background:{c['BG']};")
        lab_form=QFormLayout(lab_body); lab_form.setContentsMargins(20,16,20,16); lab_form.setSpacing(15)
        lab_form.setLabelAlignment(Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
        try: existing_lab_data=json.loads(self.shoe.get("lab_data") or "{}")
        except: existing_lab_data={}
        for field in LAB_FIELDS:
            e=QLineEdit(); v=existing_lab_data.get(field,"")
            if v not in (None,"nan",""): e.setText(str(v))
            lab_form.addRow(flbl(field),e); self._lab_entries[field]=e
        lab_scroll.setWidget(lab_body); self.tabs.addTab(lab_scroll,"Lab Data")

        foot=QFrame(); foot.setFixedHeight(56); foot.setStyleSheet(f"background:{c['SURFACE']};border-top:1px solid {c['BORDER']};")
        fl=QHBoxLayout(foot); fl.setContentsMargins(20,0,20,0); fl.setSpacing(8)
        cancel=mk_btn("Cancel","action_btn"); cancel.clicked.connect(self.reject)
        save=mk_btn("Save","primary_btn"); save.clicked.connect(self._save)
        fl.addWidget(cancel); fl.addStretch(); fl.addWidget(save); root.addWidget(foot)

    def _pick_img(self):
        p,_=QFileDialog.getOpenFileName(self,"Select Image","","Images (*.jpg *.jpeg *.png *.webp *.bmp)")
        if p: self._img_path=p; self._img_lbl.setText(os.path.basename(p))

    def _save(self):
        data={k: e.text().strip() for k,e in self._core_entries.items()}
        data["notes"]=self._notes.toPlainText().strip()
        lab_dict={k: e.text().strip() for k,e in self._lab_entries.items() if e.text().strip()}
        data["lab_data"]=json.dumps(lab_dict) if lab_dict else ""
        if self._img_path and os.path.isfile(self._img_path):
            dst=os.path.join(IMG_DIR,os.path.basename(self._img_path))
            if os.path.abspath(self._img_path)!=os.path.abspath(dst): shutil.copy2(self._img_path,dst)
            data["image_path"]=dst
        else: data["image_path"]=self.shoe.get("image_path","")
        if self.on_save: self.on_save(data)
        self.accept()


# ── Import ────────────────────────────────────────────────────────────────────
class ImportWorker(QThread):
    progress=Signal(int,int,str); done=Signal(int,int)
    def __init__(self,folder): super().__init__(); self.folder=folder
    def run(self):
        con=sqlite3.connect(DB_PATH)
        existing={r[0] for r in con.execute("SELECT idapt_id FROM shoes WHERE idapt_id IS NOT NULL")}
        entries=sorted(os.listdir(self.folder)); imp=skip=0
        for i,entry in enumerate(entries):
            path=os.path.join(self.folder,entry)
            if not os.path.isdir(path): continue
            idapt_id,name,flags=parse_idapt_folder(entry)
            if not idapt_id: continue
            self.progress.emit(i,len(entries),entry)
            if idapt_id in existing: skip+=1; continue
            bi,imgs,notes=find_best_image(path),get_all_images(path),[]
            if "NO REPORT" in flags: notes.append("No report available")
            if "NO PHOTOS" in flags: notes.append("No photos available")
            pdfs=sorted([os.path.join(path,f) for f in os.listdir(path) if f.lower().endswith(".pdf")])
            con.execute("INSERT INTO shoes(idapt_id,brand,model,notes,image_path,all_images,report_path,all_reports)VALUES(?,?,?,?,?,?,?,?)",
                        (idapt_id,idapt_id,name or "","\n".join(notes),bi or "",json.dumps(imgs),pdfs[0] if pdfs else "",json.dumps(pdfs)))
            imp+=1
        con.commit(); con.close(); self.done.emit(imp,skip)

class ImportDialog(QDialog):
    def __init__(self,parent,folder):
        super().__init__(parent); self.setWindowTitle("Importing…"); self.setFixedSize(380,110); self.setStyleSheet(generate_css())
        l=QVBoxLayout(self); l.setContentsMargins(24,20,24,20); l.setSpacing(12)
        self.lbl=QLabel("Starting import…"); self.bar=QProgressBar(); self.bar.setRange(0,100); self.bar.setValue(0)
        l.addWidget(self.lbl); l.addWidget(self.bar); self._result=(0,0); self._w=ImportWorker(folder)
        self._w.progress.connect(self._on_progress); self._w.done.connect(lambda i,s:(setattr(self,'_result',(i,s)),self.accept())); self._w.start()
    def _on_progress(self,i,total,name): self.lbl.setText(f"Importing: {name}"); self.bar.setValue(int((i+1)/total*100) if total else 0)
    def result_counts(self): return self._result



# ── Pinned Tray ───────────────────────────────────────────────────────────────
# ── Compare Panel ────────────────────────────────────────────────────────────
class CompareColumn(QFrame):
    """One shoe column inside the ComparePanel."""
    unpin_requested = Signal(dict)
    open_report_requested = Signal(dict)

    COL_W     = 220
    IMG_H     = 180

    def __init__(self, shoe, parent=None):
        super().__init__(parent)
        self.shoe = shoe
        self.setFixedWidth(self.COL_W)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)
        self._imgs = []
        self._reports = []
        self._img_idx = 0
        self._build()
        self._load_media()
        self.refresh_theme()

    def _build(self):
        root = QVBoxLayout(self); root.setContentsMargins(0, 0, 0, 0); root.setSpacing(0)

        # ── image area ──
        img_frame = QFrame(); img_frame.setFixedHeight(self.IMG_H)
        img_frame.setStyleSheet("background:#ffffff;")
        il = QVBoxLayout(img_frame); il.setContentsMargins(0,0,0,0)
        self.img_lbl = QLabel(); self.img_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.img_lbl.setFixedSize(self.COL_W, self.IMG_H)
        self.img_lbl.setStyleSheet("background:transparent;")
        il.addWidget(self.img_lbl)

        # prev/next nav overlaid
        nav_w = QWidget(img_frame); nav_w.setStyleSheet("background:transparent;")
        nav_w.setGeometry(0, 0, self.COL_W, self.IMG_H)
        nav_l = QHBoxLayout(nav_w); nav_l.setContentsMargins(2,0,2,0)
        self.prev_btn = QPushButton("‹"); self.prev_btn.setFixedSize(22,22)
        self.prev_btn.clicked.connect(self._prev_img)
        self.next_btn = QPushButton("›"); self.next_btn.setFixedSize(22,22)
        self.next_btn.clicked.connect(self._next_img)
        nav_l.addWidget(self.prev_btn, 0, Qt.AlignmentFlag.AlignVCenter)
        nav_l.addStretch()
        nav_l.addWidget(self.next_btn, 0, Qt.AlignmentFlag.AlignVCenter)
        root.addWidget(img_frame)

        # ── thin accent bar ──
        self.accent_bar = QFrame(); self.accent_bar.setFixedHeight(3); root.addWidget(self.accent_bar)

        # ── scrollable detail ──
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        inner = QWidget(); self.detail_lay = QVBoxLayout(inner)
        self.detail_lay.setContentsMargins(10,10,10,10); self.detail_lay.setSpacing(4)
        scroll.setWidget(inner); root.addWidget(scroll, 1)

        # ── bottom button bar ──
        btn_w = QWidget(); btn_l = QHBoxLayout(btn_w)
        btn_l.setContentsMargins(8,6,8,8); btn_l.setSpacing(6)
        self.report_btn = QPushButton("Report")
        self.report_btn.setFixedHeight(26)
        self.report_btn.clicked.connect(lambda: self.open_report_requested.emit(self.shoe))
        self.unpin_btn = QPushButton("Unpin")
        self.unpin_btn.setFixedHeight(26)
        self.unpin_btn.clicked.connect(lambda: self.unpin_requested.emit(self.shoe))
        btn_l.addWidget(self.report_btn, 1); btn_l.addWidget(self.unpin_btn, 1)
        root.addWidget(btn_w)

        self._fill_detail()

    def _load_media(self):
        shoe = self.shoe
        try: self._imgs = [p for p in json.loads(shoe.get("all_images") or "[]") if os.path.isfile(p)]
        except: pass
        if not self._imgs and shoe.get("image_path") and os.path.isfile(shoe["image_path"]):
            self._imgs = [shoe["image_path"]]
        try: self._reports = [p for p in json.loads(shoe.get("all_reports") or "[]") if os.path.isfile(p)]
        except: pass
        if not self._reports and shoe.get("report_path") and os.path.isfile(shoe["report_path"]):
            self._reports = [shoe["report_path"]]
        self._show_img()
        self.prev_btn.setVisible(len(self._imgs) > 1)
        self.next_btn.setVisible(len(self._imgs) > 1)
        self.report_btn.setVisible(bool(self._reports))

    def _show_img(self):
        if self._imgs:
            px = load_pixmap(self._imgs[self._img_idx], self.COL_W, self.IMG_H)
            self.img_lbl.setPixmap(px.scaled(self.COL_W, self.IMG_H,
                Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        else:
            self.img_lbl.setPixmap(_placeholder(self.COL_W, self.IMG_H))

    def _prev_img(self):
        if self._imgs: self._img_idx = (self._img_idx - 1) % len(self._imgs); self._show_img()
    def _next_img(self):
        if self._imgs: self._img_idx = (self._img_idx + 1) % len(self._imgs); self._show_img()

    def _fill_detail(self):
        c = get_c(); shoe = self.shoe
        while self.detail_lay.count():
            item = self.detail_lay.takeAt(0)
            if item and item.widget(): item.widget().deleteLater()

        def lbl(text, color, size=12, bold=False, wrap=True):
            l = QLabel(text); l.setWordWrap(wrap)
            l.setStyleSheet(f"color:{color};font-size:{size}px;font-weight:{'700' if bold else '400'};background:transparent;")
            l.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            return l
        def divider():
            f = QFrame(); f.setFixedHeight(1); f.setStyleSheet(f"background:{c['FAINT']};border:none;")
            self.detail_lay.addWidget(f)
        def section(title):
            w = QWidget(); l = QHBoxLayout(w); l.setContentsMargins(0,8,0,2); l.setSpacing(6)
            t = QLabel(title.upper()); t.setStyleSheet(f"color:{c['TEXT2']};font-size:9px;font-weight:700;letter-spacing:1px;background:transparent;")
            line = QFrame(); line.setFrameShape(QFrame.Shape.HLine); line.setStyleSheet(f"color:{c['FAINT']};")
            l.addWidget(t); l.addWidget(line,1); self.detail_lay.addWidget(w)
        def row(label, value):
            if not value or str(value).strip() in ("nan","None",""): return
            w = QWidget(); l = QHBoxLayout(w); l.setContentsMargins(0,1,0,1); l.setSpacing(6)
            lb = QLabel(label+":"); lb.setFixedWidth(100); lb.setWordWrap(True)
            lb.setStyleSheet(f"color:{c['DIM']};font-size:11px;")
            vl = QLabel(str(value)); vl.setStyleSheet(f"color:{c['TEXT']};font-size:11px;font-weight:600;"); vl.setWordWrap(True)
            vl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            l.addWidget(lb); l.addWidget(vl,1); self.detail_lay.addWidget(w)

        # ── Name / brand ──
        idapt = shoe.get("idapt_id",""); brand = shoe.get("brand",""); model = shoe.get("model","")
        display = (f"{brand} {model}".strip() if brand and brand!=idapt else f"{idapt} {model}".strip()) or "Untitled"
        self.detail_lay.addWidget(lbl(display, c['TEXT'], 13, bold=True))
        if brand and brand != idapt:
            self.detail_lay.addWidget(lbl(brand, c['TEXT2'], 11))

        # ── MAA score ──
        sc = shoe.get("maa_mean")
        if sc not in (None,""):
            sd = shoe.get("maa_sd")
            maa_lbl = QLabel(f"{float(sc):.1f}°")
            maa_lbl.setStyleSheet(f"color:{score_color(sc)};font-size:24px;font-weight:700;font-family:'Courier New';background:transparent;")
            maa_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            maa_lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            sub_txt = "MAA Mean"
            if sd not in (None,""):
                try: sub_txt += f"  ·  σ={float(sd):.2f}°"
                except: pass
            sub = QLabel(sub_txt); sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
            sub.setStyleSheet(f"color:{c['TEXT2']};font-size:10px;font-weight:600;background:transparent;")
            self.detail_lay.addSpacing(6); self.detail_lay.addWidget(maa_lbl)
            self.detail_lay.addWidget(sub); self.detail_lay.addSpacing(4)
            divider()

        # ── Core details ──
        section("Details")
        row("iDAPT ID", idapt); row("Year", shoe.get("year")); row("Size", shoe.get("size"))
        row("Sole Type", shoe.get("sole_type")); row("Tread", shoe.get("tread"))
        row("Notes", shoe.get("notes"))
        imgs_count = 0
        try: imgs_count = len([p for p in json.loads(shoe.get("all_images") or "[]") if os.path.isfile(p)])
        except: pass
        if imgs_count: row("Images", f"{imgs_count} found")
        reports_count = 0
        try: reports_count = len([p for p in json.loads(shoe.get("all_reports") or "[]") if os.path.isfile(p)])
        except: pass
        if reports_count: row("Reports", f"{reports_count} available")

        # ── Full lab data ──
        try:
            lab = json.loads(shoe.get("lab_data") or "{}")
            if lab:
                section("Lab Data")
                for k, v in lab.items():
                    row(k, v)
        except: pass

        self.detail_lay.addStretch()

    def refresh_theme(self):
        c = get_c()
        self.setStyleSheet(f"""
            CompareColumn {{ background:{c['SIDEBAR']}; border-right:1px solid {c['BORDER']}; }}
        """)
        self.accent_bar.setStyleSheet(f"background:{c['AMBER']};border:none;")
        self.prev_btn.setStyleSheet(f"QPushButton{{background:{c['SURFACE2']};color:{c['TEXT']};border-radius:4px;border:none;font-size:14px;font-weight:bold;}}QPushButton:hover{{background:{c['FAINT']};}}")
        self.next_btn.setStyleSheet(f"QPushButton{{background:{c['SURFACE2']};color:{c['TEXT']};border-radius:4px;border:none;font-size:14px;font-weight:bold;}}QPushButton:hover{{background:{c['FAINT']};}}")
        self.unpin_btn.setStyleSheet(f"QPushButton{{background:{c['SURFACE2']};color:{c['TEXT2']};border-radius:6px;border:none;font-size:11px;padding:2px 6px;}}QPushButton:hover{{background:{c['RED']};color:#ffffff;}}")
        self.report_btn.setStyleSheet(f"QPushButton{{background:#f0fdf4;color:#16a34a;border-radius:6px;border:1px solid #16a34a;font-size:11px;padding:2px 6px;}}QPushButton:hover{{background:#16a34a;color:#ffffff;}}")
        self._fill_detail()


class ComparePanel(QWidget):
    """Replaces sidebar when 1+ shoes are pinned. Shows columns side by side."""
    unpin_all_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent); self.setObjectName("sidebar")
        self._columns = {}  # shoe_id -> CompareColumn
        self._open_report_cb = None
        self._build()

    def set_report_callback(self, cb): self._open_report_cb = cb

    def _build(self):
        root = QVBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)

        # header bar
        self.header = QFrame(); self.header.setFixedHeight(36); self.header.setObjectName("compare_header")
        hl = QHBoxLayout(self.header); hl.setContentsMargins(12,0,12,0); hl.setSpacing(8)
        self.title_lbl = QLabel("📌 Comparing")
        self.title_lbl.setStyleSheet("font-size:13px;font-weight:700;background:transparent;")
        self.hint_lbl = QLabel("Right-click cards to pin/unpin")
        self.hint_lbl.setStyleSheet("font-size:10px;background:transparent;")
        self.clear_btn = QPushButton("Clear all"); self.clear_btn.setFixedHeight(24)
        self.clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_btn.clicked.connect(self.unpin_all_requested.emit)
        hl.addWidget(self.title_lbl); hl.addWidget(self.hint_lbl, 1); hl.addWidget(self.clear_btn)
        root.addWidget(self.header)

        # horizontal scroll area for columns
        self.col_scroll = QScrollArea(); self.col_scroll.setWidgetResizable(True)
        self.col_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.col_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.col_inner = QWidget()
        self.col_lay = QHBoxLayout(self.col_inner)
        self.col_lay.setContentsMargins(0,0,0,0); self.col_lay.setSpacing(0)
        self.col_lay.addStretch()
        self.col_scroll.setWidget(self.col_inner)
        root.addWidget(self.col_scroll, 1)

        # empty hint
        self.empty_lbl = QLabel("Right-click any shoe card\nto add it here for comparison")
        self.empty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_lbl.setWordWrap(True)
        root.addWidget(self.empty_lbl)

    def refresh_theme(self):
        c = get_c()
        self.header.setStyleSheet(f"QFrame#compare_header{{background:{c['SURFACE']};border-bottom:1px solid {c['BORDER']};}}")
        self.title_lbl.setStyleSheet(f"font-size:13px;font-weight:700;color:{c['TEXT']};background:transparent;")
        self.hint_lbl.setStyleSheet(f"font-size:10px;color:{c['DIM']};background:transparent;")
        self.clear_btn.setStyleSheet(f"QPushButton{{background:{c['SURFACE2']};color:{c['TEXT2']};border-radius:6px;border:none;font-size:11px;padding:2px 10px;}}QPushButton:hover{{background:{c['RED']};color:#ffffff;}}")
        self.col_scroll.setStyleSheet(f"QScrollArea{{background:{c['SIDEBAR']};border:none;}}QWidget{{background:{c['SIDEBAR']};}}")
        self.empty_lbl.setStyleSheet(f"color:{c['DIM']};font-size:13px;background:transparent;")
        for col in self._columns.values(): col.refresh_theme()

    def is_pinned(self, shoe): return shoe["id"] in self._columns

    def pin(self, shoe):
        if shoe["id"] in self._columns: return
        col = CompareColumn(shoe)
        col.unpin_requested.connect(self._on_unpin)
        col.open_report_requested.connect(self._on_open_report)
        self._columns[shoe["id"]] = col
        self.col_lay.insertWidget(self.col_lay.count()-1, col)
        self.empty_lbl.hide()
        col.refresh_theme()

    def unpin(self, shoe):
        sid = shoe["id"]
        if sid not in self._columns: return
        col = self._columns.pop(sid)
        self.col_lay.removeWidget(col); col.deleteLater()
        if not self._columns: self.empty_lbl.show()

    def clear_all(self):
        for col in list(self._columns.values()):
            self.col_lay.removeWidget(col); col.deleteLater()
        self._columns.clear(); self.empty_lbl.show()

    def update_shoe(self, shoe):
        if shoe["id"] in self._columns:
            self._columns[shoe["id"]].shoe = shoe
            self._columns[shoe["id"]]._load_media()
            self._columns[shoe["id"]]._fill_detail()

    def _on_unpin(self, shoe): self.unpin(shoe)
    def _on_open_report(self, shoe):
        if self._open_report_cb: self._open_report_cb(shoe)

    def pinned_ids(self): return set(self._columns.keys())

    def count(self): return len(self._columns)


class AsyncImageLoader(QThread):
    loaded = Signal(int, QImage, str, int, int)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.queue = []
        self.mutex = QMutex()
        self.cond = QWaitCondition()
        self.running = True

    def run(self):
        while self.running:
            self.mutex.lock()
            if not self.queue:
                self.cond.wait(self.mutex)
            if not self.queue:
                self.mutex.unlock()
                continue
            idx, path, w, h = self.queue.pop(0)
            self.mutex.unlock()

            img = QImage(path)
            if not img.isNull():
                scaled = img.scaled(w, h, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                canvas = QImage(w, h, QImage.Format.Format_ARGB32_Premultiplied)
                canvas.fill(QColor("#ffffff"))
                p = QPainter(canvas)
                x = (w - scaled.width()) // 2
                y = (h - scaled.height()) // 2
                p.drawImage(x, y, scaled)
                p.end()
                self.loaded.emit(idx, canvas, path, w, h)
            else:
                self.loaded.emit(idx, QImage(), path, w, h)

    def queue_image(self, idx, path, w, h):
        self.mutex.lock()
        self.queue.append((idx, path, w, h))
        self.cond.wakeOne()
        self.mutex.unlock()

    def clear(self):
        self.mutex.lock()
        self.queue.clear()
        self.mutex.unlock()

    def stop(self):
        self.running = False
        self.mutex.lock()
        self.cond.wakeAll()
        self.mutex.unlock()
        self.wait()

# ── Main Window ───────────────────────────────────────────────────────────────
class ShoeDatabase(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("KITE UHN Shoe Database")
        self.resize(1200,760); self.setMinimumSize(900,600); self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self._cards=[]; self._all_shoes=[]; self._selected_shoe=None
        self._compare=None
        self._img_loader = AsyncImageLoader()
        self._img_loader.loaded.connect(self._on_image_loaded)
        self._img_loader.start()
        self._search_timer=QTimer(); self._search_timer.setSingleShot(True); self._search_timer.timeout.connect(self._do_refresh)
        init_db(); self._build_ui(); self.apply_theme(); self._do_refresh()

    def closeEvent(self, event):
        self._img_loader.stop()
        super().closeEvent(event)

    def _build_ui(self):
        c=get_c()
        root=QWidget(); self.setCentralWidget(root)
        self.main_layout=QVBoxLayout(root); self.main_layout.setContentsMargins(0,0,0,0); self.main_layout.setSpacing(0)
        self.app_stack=QStackedWidget(); self.main_layout.addWidget(self.app_stack)
        self.home_page=HomeWidget(self._enter_app); self.app_stack.addWidget(self.home_page)

        self.db_page=QWidget(); self.db_page.setObjectName("db_page")
        db_lay=QVBoxLayout(self.db_page); db_lay.setContentsMargins(0,0,0,0); db_lay.setSpacing(0)
        acc=QFrame(); acc.setFixedHeight(3); acc.setStyleSheet(f"background:{c['ACCENT']};border:none;"); db_lay.addWidget(acc)

        self.topbar=QFrame(); self.topbar.setFixedHeight(56); self.topbar.setObjectName("topbar")
        tl=QHBoxLayout(self.topbar); tl.setContentsMargins(20,0,20,0); tl.setSpacing(10)
        self.top_logo=QLabel()
        if os.path.exists(LOGO_PATH):
            pix=QPixmap(LOGO_PATH).scaledToHeight(24,Qt.TransformationMode.SmoothTransformation)
            self.top_logo.setPixmap(pix); tl.addWidget(self.top_logo)
        title_lbl=QLabel("Shoe Database"); title_lbl.setObjectName("topbar_title"); tl.addWidget(title_lbl); tl.addStretch()

        self.search_box=QLineEdit(); self.search_box.setPlaceholderText("Search brand, model, iDAPT…"); self.search_box.setFixedWidth(200)
        self.search_box.textChanged.connect(lambda: self._search_timer.start(150)); tl.addWidget(self.search_box)

        add_btn=mk_btn("+ Add Shoe","primary_btn"); add_btn.setToolTip("Add New Shoe"); add_btn.clicked.connect(self._add_shoe); tl.addWidget(add_btn)

        self.more_btn=MoreOptionsButton(self); self.more_btn.setProperty("class", "action_btn"); self.more_btn.setToolTip("More Options")
        self.more_menu=QMenu(self)
        self.more_menu.addAction("Toggle Theme", self._toggle_theme)
        self.more_menu.addAction("Export to CSV", self._export_excel)
        self.more_menu.addAction("Sync Data", self._run_sync)
        self.more_menu.addAction("Import Folders", self._import_folder)
        self.more_btn.clicked.connect(lambda: self.more_menu.exec(self.more_btn.mapToGlobal(QPoint(0, self.more_btn.height() + 2))))
        tl.addWidget(self.more_btn)

        db_lay.addWidget(self.topbar)

        self.splitter=QSplitter(Qt.Orientation.Horizontal); db_lay.addWidget(self.splitter,1)
        self.grid_scroll=DynamicScrollArea(); self.grid_scroll.setWidgetResizable(True); self.grid_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.grid_widget=QWidget(); self.grid_widget.setObjectName("grid_widget")
        self.flow=FlowLayout(margin=10,spacing=10); self.grid_widget.setLayout(self.flow); self.grid_scroll.setWidget(self.grid_widget)
        self.splitter.addWidget(self.grid_scroll)
        self.sidebar_stack=QStackedWidget(); self.sidebar_stack.setMinimumWidth(260)
        self._sb_empty=SidebarEmpty(); self._sb_detail=SidebarDetail()
        self._sb_detail.edit_requested.connect(self._edit_shoe); self._sb_detail.delete_requested.connect(self._delete_shoe)
        self._compare=ComparePanel()
        self._compare.set_report_callback(self._open_shoe_report)
        self._compare.unpin_all_requested.connect(self._clear_all_pins)
        self.sidebar_stack.addWidget(self._sb_empty)   # index 0
        self.sidebar_stack.addWidget(self._sb_detail)  # index 1
        self.sidebar_stack.addWidget(self._compare)    # index 2
        self.sidebar_stack.setCurrentIndex(0)
        self.splitter.addWidget(self.sidebar_stack); self.splitter.setSizes([700,400])
        self.app_stack.addWidget(self.db_page)

    def _enter_app(self): self.app_stack.setCurrentIndex(1); self.setFocus()

    def _toggle_pin(self, shoe):
        if self._compare.is_pinned(shoe):
            self._compare.unpin(shoe)
        else:
            self._compare.pin(shoe)
        self._update_sidebar_mode()
        self._refresh_card_pin_states()

    def _update_sidebar_mode(self):
        if self._compare.count() > 0:
            self.sidebar_stack.setCurrentIndex(2)
            self.splitter.setSizes([600, 500])
        else:
            # revert to detail or empty
            if self._selected_shoe:
                self.sidebar_stack.setCurrentIndex(1)
            else:
                self.sidebar_stack.setCurrentIndex(0)
            self.splitter.setSizes([700, 400])

    def _refresh_card_pin_states(self):
        for card in self._cards:
            card.set_pinned(self._compare.is_pinned(card.shoe))

    def _clear_all_pins(self):
        self._compare.clear_all()
        self._refresh_card_pin_states()
        self._update_sidebar_mode()

    def _open_shoe_report(self, shoe):
        reports = []
        try: reports = [p for p in json.loads(shoe.get("all_reports") or "[]") if os.path.isfile(p)]
        except: pass
        if not reports and shoe.get("report_path") and os.path.isfile(shoe["report_path"]):
            reports = [shoe["report_path"]]
        if not reports: QMessageBox.warning(self, "No Report", f"No report found for {shoe.get('idapt_id','this shoe')}."); return
        open_pdf(reports[0])

    def _toggle_theme(self):
        global CURRENT_MODE; CURRENT_MODE="light" if CURRENT_MODE=="dark" else "dark"
        self.apply_theme()

    def apply_theme(self):
        self.setStyleSheet(generate_css()); self._sb_empty.refresh_theme(); self._sb_detail.refresh_theme()
        if self._compare: self._compare.refresh_theme()

    def _run_sync(self):
        dlg=SyncDialog(self); dlg.exec()
        _PIXMAP_CACHE.clear(); self._do_refresh()

    def _export_excel(self):
        path,_=QFileDialog.getSaveFileName(self,"Export Database","UHN_KITE_Database.csv","CSV Files (*.csv)")
        if not path: return
        raw_shoes=get_all_shoes()
        if not raw_shoes: QMessageBox.warning(self,"Export Failed","Database is empty!"); return
        core_keys=["id","idapt_id","brand","model","year","size","sole_type","tread","maa_mean","maa_sd","notes"]
        export_data=[]; all_keys=set(core_keys)
        for shoe in raw_shoes:
            row_data={k: shoe[k] for k in core_keys if k in shoe}
            try:
                lab_data=json.loads(shoe.get("lab_data") or "{}")
                lab_data.pop("image_path",None); lab_data.pop("all_images",None)
                row_data.update(lab_data); all_keys.update(lab_data.keys())
            except: pass
            export_data.append(row_data)
        header=[k for k in core_keys if k in all_keys]
        header.extend(sorted(list(all_keys-set(header))))
        try:
            with open(path,'w',newline='',encoding='utf-8') as f:
                writer=csv.DictWriter(f,fieldnames=header); writer.writeheader(); writer.writerows(export_data)
            QMessageBox.information(self,"Export Successful",f"Exported to:\n{path}")
        except Exception as e: QMessageBox.critical(self,"Error",f"Could not write file:\n{e}")

    def keyPressEvent(self,event):
        if self.search_box.hasFocus() or self.app_stack.currentIndex()==0: super().keyPressEvent(event); return
        if not self._all_shoes or not self._cards: return
        current_idx=next((i for i,s in enumerate(self._all_shoes) if self._selected_shoe and s["id"]==self._selected_shoe["id"]),0)
        items_per_row=max(1,self.grid_scroll.viewport().width()//(CARD_W+10))
        if event.key()==Qt.Key.Key_Right: new_idx=min(len(self._all_shoes)-1,current_idx+1)
        elif event.key()==Qt.Key.Key_Left: new_idx=max(0,current_idx-1)
        elif event.key()==Qt.Key.Key_Down: new_idx=min(len(self._all_shoes)-1,current_idx+items_per_row)
        elif event.key()==Qt.Key.Key_Up: new_idx=max(0,current_idx-items_per_row)
        else: super().keyPressEvent(event); return
        if new_idx!=current_idx:
            self._on_card_click(self._all_shoes[new_idx])
            self.grid_scroll.ensureWidgetVisible(self._cards[new_idx],50,50)

    def _do_refresh(self):
        q=self.search_box.text().strip()
        self._all_shoes=get_all_shoes(q, "all")
        self._rebuild_grid()

    def _rebuild_grid(self):
        self._img_loader.clear()
        c=get_c()
        self.grid_widget.hide()
        
        while self.flow.count():
            item = self.flow.takeAt(0)
            w = item.widget()
            if w:
                w.hide()
                if not isinstance(w, ShoeCard):
                    w.deleteLater()
                    
        self._cards.clear()
        if not hasattr(self, '_card_pool'): self._card_pool = []
        
        if not self._all_shoes:
            emp=QLabel("No shoes found.\nClick '+ Add Shoe' to get started.")
            emp.setAlignment(Qt.AlignmentFlag.AlignCenter); emp.setStyleSheet(f"color:{c['DIM']};font-size:14px;")
            self.flow.addWidget(emp)
            emp.show()
            self.grid_widget.show()
            return
            
        sel_id=self._selected_shoe["id"] if self._selected_shoe else None
        
        for i, shoe in enumerate(self._all_shoes):
            if i < len(self._card_pool):
                card = self._card_pool[i]
                card.update_shoe(shoe)
            else:
                card=ShoeCard(shoe)
                card.clicked.connect(self._on_card_click)
                card.pin_requested.connect(self._toggle_pin)
                self._card_pool.append(card)
                
            card.set_selected(shoe["id"]==sel_id)
            card.set_pinned(self._compare.is_pinned(shoe) if self._compare else False)
            self.flow.addWidget(card)
            self._cards.append(card)
            card.show()
            
            card.img_lbl.setPixmap(_placeholder(CARD_W, THUMB_H))
            
            path = card.shoe.get("image_path", "")
            if path:
                key = (path, CARD_W, THUMB_H)
                if key in _PIXMAP_CACHE:
                    card.img_lbl.setPixmap(_PIXMAP_CACHE[key])
                else:
                    self._img_loader.queue_image(i, path, CARD_W, THUMB_H)
        self.grid_widget.show()

    def _on_image_loaded(self, idx, qimg, path, w, h):
        if idx >= len(self._cards): return
        card = self._cards[idx]
        if card.shoe.get("image_path") != path: return
        
        if not qimg.isNull():
            px = QPixmap.fromImage(qimg)
        else:
            px = _placeholder(w, h)
            
        _PIXMAP_CACHE[(path, w, h)] = px
        card.img_lbl.setPixmap(px)

    def _on_card_click(self,shoe):
        if self._selected_shoe and self._selected_shoe["id"] == shoe["id"]:
            return
            
        prev_id = self._selected_shoe["id"] if self._selected_shoe else None
        self._selected_shoe=shoe
        
        for card in self._cards:
            if card.shoe["id"] == shoe["id"]:
                card.set_selected(True)
            elif prev_id and card.shoe["id"] == prev_id:
                card.set_selected(False)
                
        self._sb_detail.load(shoe)
        if self._compare.count() == 0:
            self.sidebar_stack.setCurrentIndex(1)
        self.setFocus()

    def _add_shoe(self): ShoeForm(self,on_save=self._handle_save).exec()
    def _edit_shoe(self,shoe): ShoeForm(self,shoe=shoe,on_save=lambda d:self._handle_save(d,shoe["id"])).exec()

    def _handle_save(self,data,shoe_id=None):
        save_shoe(data,shoe_id); self._do_refresh()
        if shoe_id:
            updated=next((s for s in get_all_shoes() if s["id"]==shoe_id),None)
            if updated: self._selected_shoe=updated; self._sb_detail.load(updated); self.sidebar_stack.setCurrentIndex(1); self._compare.update_shoe(updated)

    def _delete_shoe(self,shoe):
        name=(f"{shoe.get('brand','')} {shoe.get('model','')}".strip() or "this shoe")
        r=QMessageBox.question(self,"Delete",f"Delete '{name}'?",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)
        if r==QMessageBox.StandardButton.Yes:
            delete_shoe(shoe["id"]); self._compare.unpin(shoe); self._selected_shoe=None; self._sb_detail.load(None); self._refresh_card_pin_states(); self._update_sidebar_mode(); self._do_refresh()

    def _import_folder(self):
        root=QFileDialog.getExistingDirectory(self,"Select 'Photos and Reports' folder")
        if not root: return
        entries=[e for e in os.listdir(root) if os.path.isdir(os.path.join(root,e)) and re.match(r"iDAPT\d+",e,re.IGNORECASE)]
        if not entries: QMessageBox.information(self,"Import","No iDAPT folders found."); return
        r=QMessageBox.question(self,"Import",f"Found {len(entries)} iDAPT folders.\n\nImport all? (Already-imported skipped.)",QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No)
        if r!=QMessageBox.StandardButton.Yes: return
        dlg=ImportDialog(self,root); dlg.exec(); imp,skip=dlg.result_counts(); self._do_refresh()
        QMessageBox.information(self,"Import Complete",f"Imported: {imp} shoes\n⏭ Skipped (already exist): {skip}")


# ── Entry Point ───────────────────────────────────────────────────────────────
if __name__=="__main__":
    app=QApplication(sys.argv)
    app.setStyle("Fusion")
    splash=SplashScreen(); splash.show(); app.processEvents()

    def _launch():
        splash.set_status("Loading shoe records…"); app.processEvents()
        win=ShoeDatabase()
        splash.set_status("Ready!"); app.processEvents()
        def _finish(): splash.close(); win.show()
        QTimer.singleShot(600,_finish)

    QTimer.singleShot(80,_launch)
    sys.exit(app.exec())
