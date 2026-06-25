#!/usr/bin/env python3

import sys
import os
import re
import multiprocessing
import socket
import platform
import subprocess
import ctypes
from datetime import datetime

# console
if sys.platform == "win32":
    _orig_popen_init = subprocess.Popen.__init__
    def _silent_popen_init(self, *args, **kwargs):
        kwargs.setdefault("creationflags", 0)
        kwargs["creationflags"] |= subprocess.CREATE_NO_WINDOW
        _orig_popen_init(self, *args, **kwargs)
    subprocess.Popen.__init__ = _silent_popen_init

import psutil
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QListWidget, QListWidgetItem, QStackedWidget, QFrame,
    QGridLayout, QPushButton, QProgressBar, QScrollArea, QFileDialog,
    QTableWidget, QTableWidgetItem, QHeaderView, QSizePolicy, QSpacerItem
)
from PyQt6.QtCore import Qt, QTimer, QThread, QObject, pyqtSignal, QPoint
from PyQt6.QtGui import QFont, QIcon, QPixmap, QPainter

try:
    import win32pdh
    HAVE_PYWIN32 = True
except ImportError:
    HAVE_PYWIN32 = False

def _ps_cim(wmi_class, properties):
    """Run Get-CimInstance via PowerShell and return list of dicts."""
    SEP = "~~|~~"
    select_expr = f" + '{SEP}' + ".join(f"[string]$_.{p}" for p in properties)
    cmd = (
        f"Get-CimInstance -ClassName {wmi_class} | "
        f"ForEach-Object {{ {select_expr} }}"
    )
    out = safe_run(["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command", cmd], timeout=8)
    results = []
    for line in out.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = line.split(SEP)
        results.append(dict(zip(properties, [p.strip() for p in parts])))
    return results

try:
    from device_smi import Device as SmiDevice
    HAVE_DEVICE_SMI = True
except ImportError:
    HAVE_DEVICE_SMI = False

try:
    import pyopencl as cl
    HAVE_PYOPENCL = True
except ImportError:
    HAVE_PYOPENCL = False

try:
    import clr
    HAVE_PYTHONNET = True
except ImportError:
    HAVE_PYTHONNET = False
# -----------------------------------------------------------------------

APP_NAME = "AlphaINF"
APP_VERSION = "Beta 0.01"

_USB_DEVICES = set()

def _load_cpu_databases():
    """Load Intel and AMD JSON databases from the app directory.
    Returns (intel_dict, amd_list) — both empty on any failure."""
    base = get_app_dir()
    intel, amd = {}, []
    try:
        import json
        intel_path = os.path.join(base, "intel_cpu_database.json")
        if os.path.exists(intel_path):
            with open(intel_path, encoding="utf-8") as f:
                intel = json.load(f)
    except Exception:
        pass
    try:
        import json
        amd_path = os.path.join(base, "amd_cpu_database.json")
        if os.path.exists(amd_path):
            with open(amd_path, encoding="utf-8") as f:
                amd = json.load(f)
    except Exception:
        pass
    return intel, amd


def _norm(s):
    """Lowercase, strip punctuation/spaces for fuzzy name matching."""
    import re
    return re.sub(r"[\s\-_()\u2122\u00ae]+", "", s).lower()


def lookup_cpu_in_database(cpu_name):
    """Search both CPU databases for an entry matching *cpu_name*.
    Returns a flat dict of spec fields, or {} if nothing found."""
    if not cpu_name:
        return {}
    intel_db, amd_db = _load_cpu_databases()
    name_norm = _norm(cpu_name)

    # --- Intel ---
    if intel_db and ("intel" in cpu_name.lower() or
                     any(x in cpu_name.upper() for x in ("I3", "I5", "I7", "I9", "XEON", "CELERON", "PENTIUM", "ULTRA"))):
        import re
        def _extract_intel_model(s):
            m = re.search(r"[iI][3579][\-\s]?\d{3,5}[A-Za-z]*", s)
            return _norm(m.group(0)) if m else ""
        nm = _extract_intel_model(cpu_name)
        best_key, best_score = None, 0
        for key, entry in intel_db.items():
            entry_name = entry.get("name", "")
            em = _extract_intel_model(entry_name)
            if em and nm and em == nm:
                best_key, best_score = key, 100
                break
            # Fallback: overlap score
            score = sum(1 for tok in _norm(entry_name).split() if tok in name_norm)
            if score > best_score:
                best_score, best_key = score, key
        if best_key and best_score >= 2:
            entry = intel_db[best_key]
            perf = entry.get("Performance", {})
            mem  = entry.get("Memory Specifications", {})
            pkg  = entry.get("Package Specifications", {})
            adv  = entry.get("Advanced Technologies", {})
            ess  = entry.get("Essentials", {})
            return {
                "Base Frequency":   perf.get("Processor Base Frequency", ""),
                "Max Turbo":        perf.get("Max Turbo Frequency", ""),
                "TDP":              perf.get("TDP", ""),
                "Cache":            perf.get("Cache", ""),
                "Lithography":      ess.get("Lithography", ""),
                "Launch Date":      ess.get("Launch Date", ""),
                "Memory Types":     mem.get("Memory Types", ""),
                "Max Memory":       mem.get("Max Memory Size (dependent on memory type)", ""),
                "Memory Channels":  mem.get("Max # of Memory Channels", ""),
                "Socket":           pkg.get("Sockets Supported", ""),
                "TjMax":            pkg.get("TCASE", ""),
                "Hyper-Threading":  adv.get("Intel Hyper-Threading Technology", ""),
                "Turbo Boost":      adv.get("Intel Turbo Boost Technology", ""),
                "Instruction Set":  adv.get("Instruction Set Extensions", "") or adv.get("Instruction Set", ""),
                "PCIe Version":     entry.get("Expansion Options", {}).get("PCI Express Revision", ""),
            }

    # --- AMD ---
    if amd_db and ("amd" in cpu_name.lower() or
                   any(x in cpu_name.upper() for x in ("RYZEN", "ATHLON", "EPYC", "THREADRIPPER"))):
        import re
        def extract_amd_model(s):
            m = re.search(r"\d{3,4}[A-Za-z]*", s)
            return m.group(0).upper() if m else ""
        nm = extract_amd_model(cpu_name)
        best, best_score = None, 0
        for entry in amd_db:
            model = entry.get("Model", "")
            if not model:
                continue
            if nm and nm in model.upper():
                score = 10 + len(nm)
            else:
                score = sum(1 for tok in _norm(model).split() if tok in name_norm)
            if score > best_score:
                best_score, best = score, entry
        if best and best_score > 0:
            return {
                "Base Frequency":   best.get("Base Clock", ""),
                "Max Turbo":        best.get("Max. Boost Clock \u00b9 \u00b2", ""),
                "TDP":              best.get("Default TDP", ""),
                "L1 Cache":         best.get("L1 Cache", ""),
                "L2 Cache":         best.get("L2 Cache", ""),
                "L3 Cache":         best.get("L3 Cache", ""),
                "Lithography":      best.get("Processor Technology for CPU Cores", ""),
                "Launch Date":      best.get("Launch Date", ""),
                "Memory Types":     best.get("System Memory Type", ""),
                "Memory Channels":  str(best.get("Memory Channels", "")),
                "Socket":           best.get("CPU Socket", ""),
                "TjMax":            best.get("Max. Operating Temperature (Tjmax)", ""),
                "PCIe Version":     best.get("PCI Express\u00ae Version", ""),
                "Instruction Set":  "",
            }
    return {}



def get_app_dir():
    """Directory the running app actually lives in: the .exe's folder
    when frozen (e.g. via PyInstaller), otherwise this script's folder.
    Used to resolve bundled resources like LHM_DLL_PATH relative to
    %APPLOCATION% instead of the current working directory."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(*parts):
    """Resolve a path to a bundled resource.
    In a PyInstaller --onefile build, assets are extracted to sys._MEIPASS.
    In a normal run, they sit next to this script."""
    if hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)

LHM_DLL_PATH = os.path.join(get_app_dir(), "LHM", "LibreHardwareMonitor.dll")

# ---------------------------------------------------------------------------
# Color palette / stylesheet
# ---------------------------------------------------------------------------
BG = "#15161e"
PANEL = "#1d1f2b"
PANEL_ALT = "#242737"
ACCENT = "#6c5ce7"
ACCENT_SOFT = "#3a3358"
TEXT_MAIN = "#e8e9f1"
TEXT_DIM = "#8c8fa3"
GOOD = "#3ddc97"
WARN = "#ffb454"
BAD = "#ff5d6c"

STYLESHEET = f"""
QMainWindow {{
    background-color: {BG};
}}
QWidget {{
    background-color: {BG};
    color: {TEXT_MAIN};
    font-family: "Segoe UI", "Cantarell", sans-serif;
    font-size: 13px;
}}
QLabel {{
    background-color: transparent;
}}
#TitleBar {{
    background-color: {PANEL};
    border-bottom: 1px solid #2a2c3c;
}}
#Sidebar {{
    background-color: {PANEL};
    border-right: 1px solid #2a2c3c;
}}
#SidebarTitle {{
    color: {TEXT_MAIN};
    font-size: 19px;
    font-weight: 700;
    padding: 22px 18px 4px 18px;
}}
#SidebarSubtitle {{
    color: {TEXT_DIM};
    font-size: 11px;
    padding: 0px 18px 18px 18px;
}}
QListWidget {{
    background-color: transparent;
    border: none;
    outline: none;
    padding: 6px;
}}
QListWidget::item {{
    color: {TEXT_DIM};
    padding: 11px 14px;
    margin: 2px 6px;
    border-radius: 8px;
}}
QListWidget::item:selected {{
    background-color: {ACCENT_SOFT};
    color: {TEXT_MAIN};
    font-weight: 600;
}}
QListWidget::item:hover:!selected {{
    background-color: #20223054;
}}
#PageTitle {{
    font-size: 22px;
    font-weight: 700;
    color: {TEXT_MAIN};
}}
#PageSubtitle {{
    color: {TEXT_DIM};
    font-size: 12px;
    margin-bottom: 6px;
}}
.Card {{
    background-color: {PANEL};
    border-radius: 12px;
}}
.CardLabel {{
    color: {TEXT_DIM};
    font-size: 11px;
    font-weight: 600;
    letter-spacing: 0.5px;
}}
.CardValue {{
    color: {TEXT_MAIN};
    font-size: 17px;
    font-weight: 700;
}}
QProgressBar {{
    background-color: {PANEL_ALT};
    border: none;
    border-radius: 6px;
    height: 12px;
    text-align: center;
    color: {TEXT_MAIN};
    font-size: 10px;
}}
QProgressBar::chunk {{
    background-color: {ACCENT};
    border-radius: 6px;
}}
QTableWidget {{
    background-color: {PANEL};
    border: none;
    border-radius: 10px;
    gridline-color: #2a2c3c;
}}
QHeaderView::section {{
    background-color: {PANEL_ALT};
    color: {TEXT_DIM};
    border: none;
    padding: 8px;
    font-weight: 600;
}}
QTableWidget::item {{
    padding: 6px;
}}
QPushButton {{
    background-color: {ACCENT};
    color: white;
    border: none;
    border-radius: 8px;
    padding: 9px 18px;
    font-weight: 600;
}}
QPushButton:hover {{
    background-color: #7d6cf0;
}}
QScrollArea {{
    border: none;
    background: transparent;
}}
"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def format_freq_ghz(mhz):
    """Format a frequency given in MHz as a GHz string for display."""
    return f"{mhz / 1000:.2f} GHz"


def human_bytes(n):
    if n is None:
        return "N/A"
    units = ["B", "KB", "MB", "GB", "TB", "PB"]
    size = float(n)
    for unit in units:
        if size < 1024:
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} {unit}"
        size /= 1024
    return f"{size:.1f} EB"


def human_mb(n):
    """Format a byte count as a MB string (used for RAM/swap display)."""
    if n is None:
        return "N/A"
    mb = n / (1024 ** 2)
    return f"{mb:,.1f} MB"


def format_temp_c(value):
    """Format a Celsius float from a hardware sensor for display."""
    if value is None:
        return "N/A"
    return f"{value:.1f} °C"


def human_uptime(seconds):
    days, rem = divmod(int(seconds), 86400)
    hours, rem = divmod(rem, 3600)
    minutes, _ = divmod(rem, 60)
    parts = []
    if days:
        parts.append(f"{days}d")
    if hours or days:
        parts.append(f"{hours}h")
    parts.append(f"{minutes}m")
    return " ".join(parts)


def safe_run(cmd, timeout=3):
    try:
        out = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout,
                             creationflags=subprocess.CREATE_NO_WINDOW)
        return out.stdout.strip()
    except Exception:
        return ""


def get_cpu_name():
    # Prefer py-cpuinfo when available: it's pure Python, works the same
    # way on every OS, and tends to give a cleaner brand string than
    # parsing the registry / /proc/cpuinfo / sysctl ourselves.
    if HAVE_PY_CPUINFO:
        try:
            info = py_cpuinfo.get_cpu_info()
            brand = info.get("brand_raw")
            if brand:
                return brand
        except Exception:
            pass  # fall through to the manual platform-specific methods

    system = platform.system()
    try:
        if system == "Windows":
            # Registry has the real name e.g. "Intel(R) Core(TM) i5-760 CPU @ 2.80GHz"
            # platform.processor() returns raw CPUID string on some Windows versions
            try:
                import winreg
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                    r"HARDWARE\DESCRIPTION\System\CentralProcessor\0")
                name, _ = winreg.QueryValueEx(key, "ProcessorNameString")
                winreg.CloseKey(key)
                if name and name.strip():
                    return name.strip()
            except Exception:
                pass
            return platform.processor() or platform.uname().processor or "Unknown CPU"
        elif system == "Darwin":
            name = safe_run(["sysctl", "-n", "machdep.cpu.brand_string"])
            return name or platform.processor() or "Unknown CPU"
        else:
            with open("/proc/cpuinfo") as f:
                for line in f:
                    if "model name" in line:
                        return line.split(":", 1)[1].strip()
    except Exception:
        pass
    return platform.processor() or "Unknown CPU"


def detect_cpu_brand(model_string):
    """Guess a friendly brand label from the raw CPU model string.

    - i3 / i5 / i7 / i9 / Ultra  -> Intel Core
    - leading G or N model (e.g. G6900, N100) -> Intel Pentium/Celeron
    - A4/A6/A8/A9/A10/A12, Ryzen, Athlon, FX -> AMD
    """
    if not model_string:
        return None
    n = model_string.upper()

    # Intel Core family: i3, i5, i7, i9, or the newer "Ultra" branding
    if re.search(r"\bI[3579]\b", n) or "ULTRA" in n:
        return "Intel Core"

    # AMD family: Ryzen, Athlon, FX, or A-series (A4/A6/A8/A9/A10/A12)
    if re.search(r"\b(RYZEN|ATHLON|FX)\b", n) or re.search(r"\bA(4|6|8|9|10|12)\b", n):
        return "AMD"

    # Intel Pentium/Celeron: model strings starting with G or N (e.g. G6900, N100)
    if re.search(r"\bG\d{3,4}\b", n) or re.search(r"\bN\d{3,5}\b", n):
        return "Intel Pentium/Celeron"

    return None


def clean_cpu_model(raw):
    """Strip vendor/trademark noise from a raw CPU string so it can be
    recombined with a clean brand prefix."""
    s = raw


    s = re.sub(r"\((R|TM|C)\)", "", s, flags=re.I)

    s = re.sub(r"\b\d+-Core\b", "", s, flags=re.I)

    for word in ["Intel", "AMD", "Genuine", "Core", "Pentium", "Celeron",
                 "Gold", "Silver", "Processor"]:
        s = re.sub(rf"\b{word}\b", "", s, flags=re.I)
 
    s = re.sub(r"@.*$", "", s)

    s = re.sub(r"\bCPU\b", "", s, flags=re.I)

    # Turn model dashes into spaces (i5-12400F -> i5 12400F)
    s = s.replace("-", " ")

    return re.sub(r"\s+", " ", s).strip()


def get_cpu_display_name():
    """CPU name as 'Brand Model', e.g. 'Intel Core i5 12400F' or
    'AMD Ryzen 5 5600X'."""
    name = get_cpu_name()
    brand = detect_cpu_brand(name)
    if not brand:
        return name
    model = clean_cpu_model(name)
    return f"{brand} {model}".strip()


def detect_gpu_brand(model_string):
    """Guess a friendly brand label from a raw GPU model string."""
    if not model_string:
        return None
    n = model_string.upper()

    if "NVIDIA" in n or re.search(r"\bGEFORCE\b", n) or re.search(r"\bRTX\b", n) or re.search(r"\bGTX\b", n) or re.search(r"\bQUADRO\b", n):
        return "NVIDIA"

    if "AMD" in n or re.search(r"\bRADEON\b", n):
        return "AMD"

    if "INTEL" in n or re.search(r"\b(UHD|IRIS|HD GRAPHICS|ARC)\b", n):
        return "Intel"

    if re.search(r"\bAPPLE\b", n) or re.search(r"\bM[1-4]\b", n):
        return "Apple"

    return None


def clean_gpu_model(raw, brand=None):
    """Strip vendor noise from a raw GPU string so it can be recombined
    with a clean brand prefix."""
    s = raw
    s = re.sub(r"\((R|TM|C)\)", "", s, flags=re.I)
    for word in ["NVIDIA", "AMD", "Intel", "Apple", "Corporation", "Graphics", "Series"]:
        s = re.sub(rf"\b{word}\b", "", s, flags=re.I)

    # Collapse duplicated model numbers some drivers report,
    # e.g. "RX550/550" -> "RX550"
    s = re.sub(r"(\d+)\s*/\s*\1\b", r"\1", s)

    # Add a space between a model's letters and digits for AMD/NVIDIA
    # cards, e.g. "RX550" -> "RX 550", "RTX3060" -> "RTX 3060"
    if brand in ("AMD", "NVIDIA"):
        s = re.sub(r"([A-Za-z])(\d)", r"\1 \2", s)

    return re.sub(r"\s+", " ", s).strip()


def get_gpu_display_name(name):
    """GPU name as 'Brand Model', e.g. 'NVIDIA GeForce RTX 3060' or
    'AMD Radeon RX 580'."""
    brand = detect_gpu_brand(name)
    if not brand:
        return name
    model = clean_gpu_model(name, brand)
    return f"{brand} {model}".strip()


def get_cpu_extra_info():
    """Extra static CPU details from py-cpuinfo, if available.
    Returns a dict that may be empty (never raises)."""
    if not HAVE_PY_CPUINFO:
        return {}
    try:
        info = py_cpuinfo.get_cpu_info()
        return {
            "Architecture": info.get("arch_string_raw", ""),
            "Vendor": info.get("vendor_id_raw", ""),
            "L2 Cache": info.get("l2_cache_size"),
            "L3 Cache": info.get("l3_cache_size"),
        }
    except Exception:
        return {}


def get_base_clock_mhz():
    """Nominal/rated clock speed, used as the multiplier base for the
    live frequency estimate. Falls back to wmic if psutil has nothing."""
    freq = None
    try:
        f = psutil.cpu_freq()
        if f:
            freq = f.max or f.current
    except Exception:
        pass
    if not freq and platform.system() == "Windows":
        try:
            rows = _ps_cim("Win32_Processor", ["MaxClockSpeed"])
            for row in rows:
                val = row.get("MaxClockSpeed", "").strip()
                if val.isdigit():
                    freq = float(val)
                    break
        except Exception:
            pass
    return freq


def get_live_cpu_freq_powershell(base_mhz):
    """Fallback *live* CPU frequency via PowerShell and get counter .

    if not base_mhz:
        return None
    try:
        cmd = [
            "powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command",
            "$r = Get-Counter -Counter "
            "'\\Processor Information(_Total)\\% Processor Performance' "
            "-SampleInterval 1 -MaxSamples 2; "
            "$r[-1].CounterSamples.CookedValue"
        ]
        out = subprocess.run(cmd, capture_output=True, text=True, timeout=4,
                             creationflags=subprocess.CREATE_NO_WINDOW)
        lines = [l.strip() for l in out.stdout.strip().splitlines() if l.strip()]
        if not lines:
            return None
        pct = float(lines[-1])
        return base_mhz * (pct / 100.0)
    except Exception:
        return None


class PdhFreqReader:
   
    COUNTER_PATH = r"\Processor Information(_Total)\% Processor Performance"

    def __init__(self):
        self._query = None
        self._counter = None
        self._first_collect_done = False

    def open(self):
        self._query = win32pdh.OpenQuery()
        self._counter = win32pdh.AddCounter(self._query, self.COUNTER_PATH)

    def collect(self, base_mhz):
        """Call once per poll tick. Returns a frequency in MHz, or None
        if this was the first collection (no delta yet) or on error."""
        try:
            win32pdh.CollectQueryData(self._query)
        except Exception:
            return None

        if not self._first_collect_done:
            # First call establishes the baseline only; PDH rate
            # counters need a second call to compute a real delta.
            self._first_collect_done = True
            return None

        try:
            _, value = win32pdh.GetFormattedCounterValue(
                self._counter, win32pdh.PDH_FMT_DOUBLE
            )
            return base_mhz * (value / 100.0)
        except Exception:
            return None

    def close(self):
        if self._query is not None:
            try:
                win32pdh.CloseQuery(self._query)
            except Exception:
                pass
            self._query = None
            self._counter = None


class FreqPoller(QThread):
    """Polls live CPU frequency on a background thread.

    freq_updated = pyqtSignal(float)

    def __init__(self, base_mhz, interval=2.0):
        super().__init__()
        self.base_mhz = base_mhz
        self.interval = interval
        self._running = True
        self._pdh = None

    def run(self):
        use_pdh = HAVE_PYWIN32
        if use_pdh:
            try:
                self._pdh = PdhFreqReader()
                self._pdh.open()
            except Exception:
                use_pdh = False
                self._pdh = None

        while self._running:
            freq = None
            if use_pdh and self._pdh is not None:
                freq = self._pdh.collect(self.base_mhz)
            else:
                freq = get_live_cpu_freq_powershell(self.base_mhz)

            if freq:
                self.freq_updated.emit(freq)

            # Sleep in small chunks so stop() takes effect quickly
            # instead of blocking the app close for up to `interval` seconds.
            slept = 0.0
            while slept < self.interval and self._running:
                self.msleep(100)
                slept += 0.1

        if self._pdh is not None:
            self._pdh.close()

    def stop(self):
        self._running = False


def get_disk_name_map():
  
    global _USB_DEVICES
    mapping = {}
    system = platform.system()
    try:
        if system == "Windows":
            try:
                # Single PS query: join LogicalDisk -> Partition -> DiskDrive
                # Output format: DriveLetter~~|~~DiskIndex~~|~~Model
                ps_cmd = (
                    "Get-CimInstance Win32_LogicalDiskToPartition | ForEach-Object {"
                    "$part = Get-CimInstance -Query \"SELECT DiskIndex FROM Win32_DiskPartition WHERE DeviceID='$($_.Antecedent.DeviceID)'\"; "
                    "$disk = Get-CimInstance -Query \"SELECT Index,Model FROM Win32_DiskDrive WHERE Index=$($part.DiskIndex)\"; "
                    "$letter = $_.Dependent.DeviceID.TrimEnd(':'); "
                    "\"$letter~~|~~$($disk.Index)~~|~~$($disk.Model)\"}"
                )
                out = safe_run(["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command", ps_cmd], timeout=10)
                disk_models = {}
                for line in out.splitlines():
                    line = line.strip()
                    if not line or "~~|~~" not in line:
                        continue
                    parts = line.split("~~|~~")
                    if len(parts) < 3:
                        continue
                    letter, disk_idx, model = parts[0].strip(), parts[1].strip(), parts[2].strip()
                    if letter and model:
                        mapping[f"{letter}:\\"] = model
                        try:
                            disk_models[int(disk_idx)] = model
                        except Exception:
                            pass

                # Fallback: single disk -> all partitions
                if not mapping:
                    all_disks = []
                    for row in _ps_cim("Win32_DiskDrive", ["Index", "Model", "InterfaceType"]):
                        iface = row.get("InterfaceType", "").strip().upper()
                        model = row.get("Model", "").strip()
                        if model and iface != "USB":
                            all_disks.append(model)
                    if len(all_disks) == 1:
                        for part in psutil.disk_partitions(all=False):
                            mapping[part.device] = all_disks[0]

                # Optical drives
                for row in _ps_cim("Win32_CDROMDrive", ["Drive", "Caption"]):
                    drive = row.get("Drive", "").strip()
                    caption = row.get("Caption", "").strip()
                    if drive:
                        mapping[f"{drive}\\"] = caption or "DVD/CD Drive"

                # Track which drive letters are USB
                for row in _ps_cim("Win32_DiskDrive", ["Index", "InterfaceType"]):
                    iface = row.get("InterfaceType", "").strip().upper()
                    if iface == "USB":
                        try:
                            idx = int(row.get("Index", "").strip())
                            for letter, model in list(mapping.items()):
                                # match via disk_models built earlier if available
                                pass
                        except Exception:
                            pass
                # Simpler: mark any mapped drive whose model contains USB keywords
                usb_keywords = ("usb", "flash", "datatraveler", "sandisk cruzer",
                                "kingston dt", "generic flash")
                _USB_DEVICES = {
                    dev for dev, name in mapping.items()
                    if any(k in name.lower() for k in usb_keywords)
                }
            except Exception:
                pass

            # Fallback: PowerShell Get-Partition if CimInstance gave nothing
            if not mapping:
                ps_cmd = (
                    "Get-Partition | Where-Object DriveLetter | ForEach-Object {"
                    "$d = Get-Disk -Number $_.DiskNumber; "
                    "\"$($_.DriveLetter):,$($d.FriendlyName)\"}"
                )
                out = safe_run(["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command", ps_cmd], timeout=8)
                for line in out.splitlines():
                    line = line.strip()
                    if "," in line:
                        letter, name = line.split(",", 1)
                        letter, name = letter.strip(), name.strip()
                        if letter and name:
                            mapping[f"{letter}:\\"] = name


        elif system == "Darwin":
            for part in psutil.disk_partitions(all=False):
                out = safe_run(["diskutil", "info", part.device])
                for line in out.splitlines():
                    if "Device / Media Name" in line:
                        mapping[part.device] = line.split(":", 1)[1].strip()
                        break
        else:
            # Linux: list every block device with its model, then map
            # each partition back to its parent disk's model.
            out = safe_run(["lsblk", "-o", "NAME,MODEL", "-p", "-n"])
            disk_models = {}
            for line in out.splitlines():
                bits = line.strip().split(None, 1)
                if not bits:
                    continue
                dev = bits[0]
                model = bits[1].strip() if len(bits) > 1 else ""
                if model:
                    disk_models[dev] = model
            for part in psutil.disk_partitions(all=False):
                dev = part.device
                base = re.sub(r"p?\d+$", "", dev)  # /dev/sda1 -> /dev/sda
                if base in disk_models:
                    mapping[dev] = disk_models[base]
    except Exception:
        pass
    return mapping


def get_smi_cpu_info():
    
    if not HAVE_DEVICE_SMI:
        return None
    try:
        dev = SmiDevice("cpu")
        return {
            "vendor": (getattr(dev, "vendor", None) or "").upper() or None,
            "model": getattr(dev, "model", None) or None,
            "count": getattr(dev, "count", None),
            "cores": getattr(dev, "cores", None),
            "threads": getattr(dev, "threads", None),
            "features": getattr(dev, "features", None) or [],
        }
    except Exception:
        return None


def get_smi_gpu_info():
   
    if not HAVE_DEVICE_SMI:
        return []
    results = []
    for prefix in ("cuda", "xpu", "gpu", "rocm"):
        found_any = False
        for i in range(8):
            try:
                dev = SmiDevice(f"{prefix}:{i}")
            except Exception:
                break
            found_any = True
            try:
                entry = {
                    "vendor": getattr(dev, "vendor", None),
                    "model": getattr(dev, "model", None),
                    "memory_total": getattr(dev, "memory_total", None),
                }
                gpu_block = getattr(dev, "gpu", None) or {}
                if isinstance(gpu_block, dict):
                    entry["driver"] = gpu_block.get("driver")
                    entry["firmware"] = gpu_block.get("firmware")
                pcie = getattr(dev, "pcie", None) or {}
                if isinstance(pcie, dict):
                    entry["pcie_gen"] = pcie.get("gen")
                    entry["pcie_speed"] = pcie.get("speed")
                    entry["pcie_id"] = pcie.get("id")
                results.append(entry)
            except Exception:
                pass
        if found_any:
            break
    return results


def get_primary_disk_name():
    """Best-effort physical disk model name (e.g. 'Samsung SSD 860 EVO
    250GB'), not the drive letter/mount point. Falls back to the first
    partition's device path if model detection isn't available."""
    system = platform.system()
    try:
        if system == "Windows":
            try:
                # MediaType 3 = Fixed HDD/SSD, 4 = External, 5 = CD-ROM
                # Also skip obvious USB keywords as fallback filter
                for row in _ps_cim("Win32_DiskDrive", ["Model", "MediaType", "InterfaceType"]):
                    model = row.get("Model", "").strip()
                    media = row.get("MediaType", "").strip()
                    iface = row.get("InterfaceType", "").strip().upper()
                    if not model:
                        continue
                    # Skip USB/removable drives
                    if iface == "USB":
                        continue
                    if media and media not in ("Fixed hard disk media", "3", ""):
                        continue
                    return model
            except Exception:
                pass
        elif system == "Darwin":
            out = safe_run(["system_profiler", "SPStorageDataType"])
            for line in out.splitlines():
                if "Device Name" in line or "Media Name" in line:
                    return line.split(":", 1)[1].strip()
        else:
            # Linux: try lsblk for a clean model name first.
            out = safe_run(["lsblk", "-d", "-no", "MODEL"])
            for line in out.splitlines():
                line = line.strip()
                if line:
                    return line
    except Exception:
        pass

    # Fallback: first real partition's device path.
    try:
        parts = psutil.disk_partitions(all=False)
        if parts:
            return parts[0].device
    except Exception:
        pass
    return "Unknown Disk"


_GPU_ARCH_TABLE = [
    # NVIDIA
    (r"RTX\s?50\d{2}", "Blackwell", "GDDR7"),
    (r"RTX\s?40\d{2}", "Ada Lovelace", "GDDR6X"),
    (r"RTX\s?30\d{2}", "Ampere", "GDDR6X"),
    (r"\bA\d{2,3}0?\b.*GB|A100|A6000|A40\b", "Ampere", "HBM2e"),
    (r"H100|H200", "Hopper", "HBM3"),
    (r"L40|L4\b", "Ada Lovelace", "GDDR6"),
    (r"RTX\s?20\d{2}|GTX\s?16\d{2}", "Turing", "GDDR6"),
    (r"GTX\s?10\d{2}\s?Ti", "Pascal", "GDDR5X"),
    (r"GTX\s?10\d{2}", "Pascal", "GDDR5"),
    (r"GTX\s?9\d{2}", "Maxwell", "GDDR5"),
    (r"GTX\s?7\d{2}|GTX\s?6\d{2}", "Kepler", "GDDR5"),
    (r"V100", "Volta", "HBM2"),
    # AMD
    (r"RX\s?7\d{3}", "RDNA 3", "GDDR6"),
    (r"RX\s?6\d{3}", "RDNA 2", "GDDR6"),
    (r"RX\s?5\d{3}", "RDNA", "GDDR6"),
    (r"Vega", "Vega (GCN 5)", "HBM2"),
    (r"RX\s?4\d{2}|RX\s?5\d{2}", "Polaris", "GDDR5"),
    (r"R9\s?Fury", "Fiji (GCN 3)", "HBM"),
    (r"R9|R7", "GCN", "GDDR5"),
    # Intel
    (r"Arc\s?A\d", "Alchemist (Xe-HPG)", "GDDR6"),
    (r"Iris\s?Xe", "Xe-LP (integrated)", "Shared System Memory"),
    (r"UHD Graphics", "Intel Gen (integrated)", "Shared System Memory"),
    # Apple
    (r"Apple M\d", "Apple Silicon GPU", "Unified Memory"),
]


def get_gpu_arch_and_vram_type(model_name):
    """Best-effort GPU architecture codename + VRAM type, inferred from
    the model name via a lookup table (no OS API exposes either of
    these directly). Returns (architecture, vram_type), each "Unknown"
    if no pattern matches."""
    if not model_name:
        return "Unknown", "Unknown"
    for pattern, arch, vram in _GPU_ARCH_TABLE:
        if re.search(pattern, model_name, re.IGNORECASE):
            return arch, vram
    return "Unknown", "Unknown"


def get_nvidia_live_gpu_specs():
    """Live core/memory clock (MHz) and core voltage (mV, when the
    driver exposes it) per NVIDIA GPU index, via nvidia-smi. Returns a
    dict {index: {"core_clock": .., "mem_clock": .., "voltage": ..}}.
    Voltage is only reported by recent drivers on some cards; absent
    otherwise. Empty dict if nvidia-smi isn't available."""
    specs = {}
    out = safe_run(["nvidia-smi",
                     "--query-gpu=index,clocks.gr,clocks.mem",
                     "--format=csv,noheader,nounits"], timeout=5)
    for line in out.splitlines():
        bits = [b.strip() for b in line.split(",")]
        if len(bits) >= 3:
            try:
                idx = int(bits[0])
                specs[idx] = {"core_clock": bits[1], "mem_clock": bits[2], "voltage": None}
            except ValueError:
                continue

    # Voltage isn't in the CSV query API; only some recent driver/card
    # combos expose it via the verbose -q text report.
    vout = safe_run(["nvidia-smi", "-q", "-d", "VOLTAGE"], timeout=5)
    if vout:
        idx = 0
        for line in vout.splitlines():
            if line.strip().startswith("GPU "):
                m = re.match(r"GPU (\d+)", line.strip())
                if m:
                    idx = int(m.group(1))
            m = re.search(r"Voltage\s*:\s*([\d.]+)\s*mV", line)
            if m and idx in specs:
                specs[idx]["voltage"] = f"{m.group(1)} mV"
    return specs


# --- ADL (AMD Display Library) structures and sensor IDs, module-level
# so they're defined once instead of being rebuilt on every reader
# open() call. ---

# AdapterInfo, as defined in AMD's adl_structures.h (Windows build).
class _ADL_AdapterInfo(ctypes.Structure):
    _fields_ = [
        ("iSize",            ctypes.c_int),
        ("iAdapterIndex",    ctypes.c_int),
        ("strUDID",          ctypes.c_char * 256),
        ("iBusNumber",       ctypes.c_int),
        ("iDeviceNumber",    ctypes.c_int),
        ("iFunctionNumber",  ctypes.c_int),
        ("iVendorID",        ctypes.c_int),
        ("strAdapterName",   ctypes.c_char * 256),
        ("strDisplayName",   ctypes.c_char * 256),
        ("iPresent",         ctypes.c_int),
        # Windows-only tail fields (ADL is only used on Windows here)
        ("iExist",           ctypes.c_int),
        ("strDriverPath",    ctypes.c_char * 256),
        ("strDriverPathExt", ctypes.c_char * 256),
        ("strPNPString",     ctypes.c_char * 256),
        ("iOSDisplayIndex",  ctypes.c_int),
    ]

# Sensor IDs used by ADL2_New_QueryPMLogData_Get (adl_defines.h / ADLSensorType)
_ADL_SENSOR_CLKS_CORE   = 100   # Shader/core clock MHz
_ADL_SENSOR_CLKS_MEMORY = 101   # Memory clock MHz
_ADL_SENSOR_VOLTAGE_GFX = 200   # Core voltage mV


class _ADL_SingleSensorData(ctypes.Structure):
    _fields_ = [("supported", ctypes.c_int),
                ("value",     ctypes.c_int)]


class _ADL_PMLogData(ctypes.Structure):
    _fields_ = [("ulVersion",          ctypes.c_uint),
                ("ulActiveSampleRate", ctypes.c_uint),
                ("ulLastUpdated",      ctypes.c_ulonglong),
                ("ulValidMask",        ctypes.c_ulonglong * 4),
                ("sensors",            _ADL_SingleSensorData * 256)]


class _ADL_PMActivity(ctypes.Structure):
    _fields_ = [("iSize",                    ctypes.c_int),
                ("iEngineClock",              ctypes.c_int),  # 10 kHz units
                ("iMemoryClock",              ctypes.c_int),  # 10 kHz units
                ("iVddc",                     ctypes.c_int),  # mV
                ("iActivityPercent",          ctypes.c_int),
                ("iCurrentPerformanceLevel",  ctypes.c_int),
                ("iCurrentBusSpeed",          ctypes.c_int),
                ("iCurrentBusLanes",          ctypes.c_int),
                ("iMaximumBusLanes",          ctypes.c_int),
                ("iReserved",                 ctypes.c_int)]


class AdlGpuReader:

    def __init__(self):
        self._adl = None
        self._adapter_indices = []
        self._have_adl2 = False
        self._adl2_context = ctypes.c_void_p()
        self._malloc_cb = None

    def open(self):
        """Load the DLL, initialize both ADL APIs, and enumerate
        adapters once. Returns True on success, False if AMD's ADL
        isn't usable on this machine (no AMD GPU, DLL missing, etc).
        Safe to call even when there's no AMD GPU -- just returns False."""
        for dll_name in ("atiadlxx.dll", "atiadlxy.dll"):
            try:
                self._adl = ctypes.WinDLL(dll_name)
                break
            except OSError:
                continue
        if self._adl is None:
            return False

        ADL_MAIN_MALLOC_CALLBACK = ctypes.CFUNCTYPE(ctypes.c_void_p, ctypes.c_int)
        self._malloc_cb = ADL_MAIN_MALLOC_CALLBACK(
            lambda size: ctypes.cast(ctypes.create_string_buffer(size),
                                     ctypes.c_void_p).value
        )

        try:
            rc = self._adl.ADL_Main_Control_Create(self._malloc_cb, 1)
        except Exception:
            self._adl = None
            return False
        if rc != 0:
            self._adl = None
            return False

        try:
            rc2 = self._adl.ADL2_Main_Control_Create(
                self._malloc_cb, 1, ctypes.byref(self._adl2_context))
            self._have_adl2 = (rc2 == 0)
        except Exception:
            self._have_adl2 = False

        num_adapters = ctypes.c_int(0)
        try:
            self._adl.ADL_Adapter_NumberOfAdapters_Get(ctypes.byref(num_adapters))
        except Exception:
            self.close()
            return False

        n = num_adapters.value
        if n <= 0:
            self.close()
            return False

        arr_type = _ADL_AdapterInfo * n
        adapter_buf = arr_type()
        adapter_buf[0].iSize = ctypes.sizeof(_ADL_AdapterInfo)
        try:
            self._adl.ADL_Adapter_AdapterInfo_Get(
                ctypes.cast(adapter_buf, ctypes.c_void_p), ctypes.sizeof(arr_type))
        except Exception:
            self.close()
            return False

        seen_udids = set()
        self._adapter_indices = []
        for i in range(n):
            info = adapter_buf[i]
            udid = info.strUDID
            if udid in seen_udids:
                continue  # ADL lists each adapter once per display output
            seen_udids.add(udid)
            self._adapter_indices.append(info.iAdapterIndex)

        return len(self._adapter_indices) > 0

    def poll(self):
        """Read current clocks/voltage for every enumerated adapter.
        Returns a list of dicts [{core_clock, mem_clock, voltage}, ...]
        in adapter-index order. Safe to call repeatedly -- no DLL
        reload or context recreation happens here, just the sensor
        reads themselves."""
        if self._adl is None:
            return []

        results = []
        for idx in self._adapter_indices:
            entry = {"core_clock": None, "mem_clock": None, "voltage": None}

            if self._have_adl2:
                try:
                    pmlog = _ADL_PMLogData()
                    rc = self._adl.ADL2_New_QueryPMLogData_Get(
                        self._adl2_context, idx, ctypes.byref(pmlog))
                    if rc == 0:
                        def _sensor(sid):
                            s = pmlog.sensors[sid]
                            return s.value if s.supported else None
                        entry["core_clock"] = _sensor(_ADL_SENSOR_CLKS_CORE)
                        entry["mem_clock"]  = _sensor(_ADL_SENSOR_CLKS_MEMORY)
                        v = _sensor(_ADL_SENSOR_VOLTAGE_GFX)
                        if v is not None:
                            entry["voltage"] = f"{v} mV"
                        results.append(entry)
                        continue
                except Exception:
                    pass

            # Legacy OD5 fallback
            try:
                act = _ADL_PMActivity()
                act.iSize = ctypes.sizeof(_ADL_PMActivity)
                rc = self._adl.ADL_Overdrive5_CurrentActivity_Get(idx, ctypes.byref(act))
                if rc == 0:
                    if act.iEngineClock > 0:
                        entry["core_clock"] = act.iEngineClock // 100  # 10kHz -> MHz
                    if act.iMemoryClock > 0:
                        entry["mem_clock"] = act.iMemoryClock // 100
                    if act.iVddc > 0:
                        entry["voltage"] = f"{act.iVddc} mV"
            except Exception:
                pass

            results.append(entry)

        return results

    def close(self):
        if self._adl is None:
            return
        try:
            self._adl.ADL_Main_Control_Destroy()
        except Exception:
            pass
        if self._have_adl2:
            try:
                self._adl.ADL2_Main_Control_Destroy(self._adl2_context)
            except Exception:
                pass
        self._adl = None
        self._adapter_indices = []
        self._have_adl2 = False


def _get_amd_live_specs_adl_windows():
 
    reader = AdlGpuReader()
    try:
        if not reader.open():
            return []
        return reader.poll()
    finally:
        reader.close()


def _get_intel_live_specs_windows():
    
    specs = []
    # PowerShell Get-Counter is the simplest cross-version path
    ps_cmd = (
        "(Get-Counter '\\GPU Engine(*)\\Running Time' -ErrorAction SilentlyContinue"
        " | Select-Object -ExpandProperty CounterSamples"
        " | Where-Object { $_.InstanceName -match 'engtype_3D' }"
        " | Measure-Object CookedValue -Maximum).Maximum"
    )
    try:
        out = safe_run(["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command", ps_cmd], timeout=6)
        # This counter gives % time, not MHz directly — insufficient.
        # Instead use the dedicated frequency counter available in Win10 2004+:
        pass
    except Exception:
        pass

    ps_freq = (
        "Get-Counter '\\GPU Adapter Memory(*)\\Shared Usage' -ErrorAction SilentlyContinue"
        " | Out-Null; "
        "(Get-Counter '\\GPU Engine(engtype_3D)\\Running Time'"
        " -ErrorAction SilentlyContinue"
        " | Select-Object -ExpandProperty CounterSamples)[0].CookedValue"
    )
   
    wmi_out = safe_run(
        ["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-Command",
         "(Get-WmiObject -Query \"SELECT * FROM Win32_VideoController"
         " WHERE Name LIKE '%Intel%'\").CurrentRefreshRate"],
        timeout=5)

    return specs


def _get_amd_live_specs_rocmsmi():
    """Live core clock, memory clock, and voltage for AMD GPUs via rocm-smi.

    rocm-smi --showclocks outputs blocks like:
        GPU[0] : GPU clock level: 2 (1500Mhz)
        GPU[0] : Memory clock level: 2 (875Mhz)
    rocm-smi --showvoltage:
        GPU[0] : Voltage (mV): 1050

    Returns a dict  {gpu_index: {"core_clock": int, "mem_clock": int, "voltage": str}}.
    """
    specs = {}
    # --- clocks ---
    clk_out = safe_run(["rocm-smi", "--showclocks"], timeout=5)
    for line in clk_out.splitlines():
        m_idx  = re.match(r"GPU\[(\d+)\]", line.strip())
        m_core = re.search(r"GPU clock.*?(\d+)\s*[Mm]hz", line, re.IGNORECASE)
        m_mem  = re.search(r"[Mm]emory clock.*?(\d+)\s*[Mm]hz", line, re.IGNORECASE)
        if m_idx and m_core:
            idx = int(m_idx.group(1))
            specs.setdefault(idx, {})["core_clock"] = int(m_core.group(1))
        if m_idx and m_mem:
            idx = int(m_idx.group(1))
            specs.setdefault(idx, {})["mem_clock"] = int(m_mem.group(1))
    # JSON output is more reliable on newer rocm-smi versions
    if not specs:
        clk_json = safe_run(["rocm-smi", "--showclocks", "--json"], timeout=5)
        try:
            import json
            data = json.loads(clk_json)
            for key, val in data.items():
                m = re.match(r"card(\d+)", key)
                if not m:
                    continue
                idx = int(m.group(1))
                specs.setdefault(idx, {})
                for k, v in val.items():
                    if re.search(r"gpu.*(clock|freq)", k, re.IGNORECASE):
                        m2 = re.search(r"(\d+)", str(v))
                        if m2:
                            specs[idx]["core_clock"] = int(m2.group(1))
                    if re.search(r"mem.*(clock|freq)", k, re.IGNORECASE):
                        m2 = re.search(r"(\d+)", str(v))
                        if m2:
                            specs[idx]["mem_clock"] = int(m2.group(1))
        except Exception:
            pass
    # --- voltage ---
    volt_out = safe_run(["rocm-smi", "--showvoltage"], timeout=5)
    for line in volt_out.splitlines():
        m_idx  = re.match(r"GPU\[(\d+)\]", line.strip())
        m_volt = re.search(r"Voltage.*?:\s*(\d+)", line, re.IGNORECASE)
        if m_idx and m_volt:
            idx = int(m_idx.group(1))
            specs.setdefault(idx, {})["voltage"] = f"{m_volt.group(1)} mV"
    return specs


def _get_amd_live_specs_sysfs():
    """Live AMD GPU clocks + voltage from the AMDGPU sysfs interface.

    The AMDGPU kernel driver exposes per-card power/clock data under
    /sys/class/drm/cardN/device/. These files are world-readable (no root
    required) on most distros:

      pp_dpm_sclk  — list of shader (core) clock states; active one is marked *
      pp_dpm_mclk  — list of memory clock states; active one is marked *
      hwmon/hwmon*/in0_input — core voltage in millivolts (hwmon subsystem)

    Returns a list of dicts [{name, core_clock, mem_clock, voltage}, ...],
    one per AMDGPU card found, in card-number order.
    """
    import glob
    results = []
    for card_path in sorted(glob.glob("/sys/class/drm/card[0-9]*/device")):
        # Only consider AMDGPU cards (skip Intel i915, nouveau, etc.)
        driver_link = os.path.join(card_path, "driver")
        try:
            driver = os.path.basename(os.readlink(driver_link))
        except Exception:
            driver = ""
        if "amdgpu" not in driver.lower():
            continue

        entry = {"name": None, "core_clock": None, "mem_clock": None, "voltage": None}

        # Card name from uevent
        try:
            with open(os.path.join(card_path, "uevent")) as f:
                for line in f:
                    if line.startswith("DRIVER="):
                        break
        except Exception:
            pass

        def _read_active_clock(dpm_file):
            """Parse pp_dpm_sclk / pp_dpm_mclk; return MHz int of the active (*) state."""
            try:
                with open(dpm_file) as f:
                    for line in f:
                        if "*" in line:
                            m = re.search(r"(\d+)\s*[Mm]hz", line)
                            if m:
                                return int(m.group(1))
            except Exception:
                pass
            return None

        entry["core_clock"] = _read_active_clock(
            os.path.join(card_path, "pp_dpm_sclk"))
        entry["mem_clock"]  = _read_active_clock(
            os.path.join(card_path, "pp_dpm_mclk"))

        # Voltage via hwmon (in0_input = GPU voltage in mV)
        for hwmon in sorted(glob.glob(os.path.join(card_path, "hwmon", "hwmon*"))):
            try:
                with open(os.path.join(hwmon, "in0_input")) as f:
                    mv = int(f.read().strip())
                    if mv > 0:
                        entry["voltage"] = f"{mv} mV"
                        break
            except Exception:
                pass

        results.append(entry)
    return results


def _get_intel_live_specs_sysfs():
    """Live Intel GPU (i915) clocks from sysfs.

    The i915 kernel driver exposes actual/max frequency under:
      /sys/class/drm/cardN/gt/gt0/  (modern kernels, 5.18+)
      /sys/class/drm/cardN/device/  (older kernels)

    Files: rps_cur_freq_mhz (current), gt_cur_freq_mhz (older name).
    Memory clock is not exposed by i915 — Intel iGPUs share system RAM
    and the memory controller frequency isn't reported per-GPU.
    Voltage is also unavailable without root (requires MSR access).

    Returns a list of dicts [{name, core_clock, mem_clock, voltage}, ...].
    """
    import glob
    results = []
    for card_path in sorted(glob.glob("/sys/class/drm/card[0-9]*/device")):
        driver_link = os.path.join(card_path, "driver")
        try:
            driver = os.path.basename(os.readlink(driver_link))
        except Exception:
            driver = ""
        if "i915" not in driver.lower() and "xe" not in driver.lower():
            continue

        entry = {"name": "Intel GPU", "core_clock": None, "mem_clock": None, "voltage": None}

        # Modern path (kernel 5.18+): gt/gt0/rps_cur_freq_mhz
        for freq_path in [
            os.path.join(card_path, "..", "gt", "gt0", "rps_cur_freq_mhz"),
            os.path.join(card_path, "gt_cur_freq_mhz"),
        ]:
            try:
                with open(freq_path) as f:
                    val = int(f.read().strip())
                    if val > 0:
                        entry["core_clock"] = val
                        break
            except Exception:
                pass

        results.append(entry)
    return results


def get_opencl_live_gpu_specs():
    """Live core clock, memory clock, and voltage for Intel and AMD GPUs.

    Source priority by OS:

    Windows:
      AMD   → ADL (atiadlxx.dll) via ADL2_New_QueryPMLogData_Get (Polaris+)
              or ADL_Overdrive5_CurrentActivity_Get (legacy GCN fallback).
              The DLL ships with every AMD driver since ~2010 — no extra install.
      Intel → OpenCL CL_DEVICE_MAX_CLOCK_FREQUENCY (max rated clock; live
              frequency isn't exposed without DXGI/D3DKMT kernel calls).

    Linux:
      AMD   → rocm-smi (if ROCm is installed), else AMDGPU sysfs pp_dpm_sclk /
              pp_dpm_mclk / hwmon in0_input — world-readable, no root needed.
      Intel → i915/xe sysfs rps_cur_freq_mhz — live actual frequency.

    Returns a list of dicts, one per GPU device found:
        {
            "name":       str,
            "vendor":     str,
            "core_clock": int | None,    # MHz
            "mem_clock":  int | None,    # MHz  (AMD only)
            "voltage":    str | None,    # e.g. "1050 mV"  (AMD only)
        }
    """
    is_windows = platform.system() == "Windows"

    # --- Step 1: enumerate GPUs via OpenCL (name + vendor + max clock) ---
    ocl_devices = []
    if HAVE_PYOPENCL:
        try:
            for plt in cl.get_platforms():
                try:
                    for device in plt.get_devices(device_type=cl.device_type.GPU):
                        vendor    = device.vendor.strip()
                        vendor_up = vendor.upper()
                        ocl_devices.append({
                            "name":       device.name.strip(),
                            "vendor":     vendor,
                            "core_clock": device.max_clock_frequency or None,
                            "mem_clock":  None,
                            "voltage":    None,
                            "_is_amd":    "AMD" in vendor_up or "ADVANCED MICRO" in vendor_up,
                            "_is_intel":  "INTEL" in vendor_up,
                        })
                except Exception:
                    continue
        except Exception:
            pass

    has_amd   = any(d["_is_amd"]   for d in ocl_devices) if ocl_devices else True
    has_intel = any(d["_is_intel"] for d in ocl_devices) if ocl_devices else True

    # --- Step 2: get live data from the right source per OS ---
    if is_windows:
        # AMD: ADL DLL (ships with every AMD driver)
        amd_live_list = _get_amd_live_specs_adl_windows() if has_amd else []
        # Intel: no reliable live-clock source on Windows without vendor SDK;
        # the OpenCL max-clock value captured above is the best we have.
        amd_sysfs = []
        amd_rocm  = {}
        intel_sysfs = []
    else:
        amd_live_list = []
        amd_rocm  = _get_amd_live_specs_rocmsmi()    if has_amd   else {}
        amd_sysfs = _get_amd_live_specs_sysfs()      if (has_amd and not amd_rocm) else []
        intel_sysfs = _get_intel_live_specs_sysfs()  if has_intel else []

    # --- Step 3: merge live data into the OpenCL device list ---
    amd_live_idx    = 0
    amd_sysfs_idx   = 0
    intel_sysfs_idx = 0

    for d in ocl_devices:
        if d["_is_amd"]:
            if is_windows:
                # Match ADL entries positionally (same enumeration order as OCL)
                if amd_live_idx < len(amd_live_list):
                    live = amd_live_list[amd_live_idx]
                    amd_live_idx += 1
                    if live.get("core_clock"):
                        d["core_clock"] = live["core_clock"]
                    if live.get("mem_clock"):
                        d["mem_clock"]  = live["mem_clock"]
                    if live.get("voltage"):
                        d["voltage"]    = live["voltage"]
            else:
                amd_idx = sum(1 for x in ocl_devices[:ocl_devices.index(d)] if x["_is_amd"])
                live = amd_rocm.get(amd_idx, {})
                if live:
                    if live.get("core_clock"): d["core_clock"] = live["core_clock"]
                    if live.get("mem_clock"):  d["mem_clock"]  = live["mem_clock"]
                    if live.get("voltage"):    d["voltage"]    = live["voltage"]
                elif amd_sysfs_idx < len(amd_sysfs):
                    s = amd_sysfs[amd_sysfs_idx]; amd_sysfs_idx += 1
                    if s.get("core_clock"): d["core_clock"] = s["core_clock"]
                    if s.get("mem_clock"):  d["mem_clock"]  = s["mem_clock"]
                    if s.get("voltage"):    d["voltage"]    = s["voltage"]

        elif d["_is_intel"] and not is_windows and intel_sysfs_idx < len(intel_sysfs):
            s = intel_sysfs[intel_sysfs_idx]; intel_sysfs_idx += 1
            if s.get("core_clock"):
                d["core_clock"] = s["core_clock"]

    # Clean up internal flags
    for d in ocl_devices:
        d.pop("_is_amd", None)
        d.pop("_is_intel", None)

    # Sysfs-only fallback when PyOpenCL isn't installed (Linux)
    if not ocl_devices and not is_windows:
        for s in (amd_sysfs or []):
            ocl_devices.append({"name": s.get("name") or "AMD GPU",   "vendor": "AMD",
                                 "core_clock": s.get("core_clock"), "mem_clock": s.get("mem_clock"),
                                 "voltage": s.get("voltage")})
        for s in (intel_sysfs or []):
            ocl_devices.append({"name": "Intel GPU", "vendor": "Intel",
                                 "core_clock": s.get("core_clock"), "mem_clock": None, "voltage": None})

    leftover_amd = amd_live_list[amd_live_idx:] if is_windows else []
    if leftover_amd:
        try:
            amd_basic_names = [
                g["name"] for g in get_gpu_info()
                if "amd" in g.get("name", "").lower()
                or "radeon" in g.get("name", "").lower()
            ]
        except Exception:
            amd_basic_names = []
        for j, live in enumerate(leftover_amd):
            name = (amd_basic_names[j] if j < len(amd_basic_names)
                    else f"AMD GPU {j}")
            ocl_devices.append({
                "name": name,
                "vendor": "AMD",
                "core_clock": live.get("core_clock"),
                "mem_clock":  live.get("mem_clock"),
                "voltage":    live.get("voltage"),
            })

    return ocl_devices


def get_ram_module_specs():
    """Best-effort DDR type + speed (MHz) of installed RAM modules.
    Returns a dict {"type": .., "speed": ..}, values "N/A" if unknown.
    If multiple modules disagree, shows the most common value."""
    system = platform.system()
    types, speeds = [], []
    try:
        if system == "Windows":
            try:
                ddr_map = {20: "DDR", 21: "DDR2", 24: "DDR3", 26: "DDR4", 34: "DDR5"}
                for row in _ps_cim("Win32_PhysicalMemory", ["Speed", "SMBIOSMemoryType"]):
                    sp = row.get("Speed", "").strip()
                    tp = row.get("SMBIOSMemoryType", "").strip()
                    if sp.isdigit():
                        speeds.append(f"{sp} MHz")
                    if tp.isdigit():
                        types.append(ddr_map.get(int(tp), f"Type {tp}"))
            except Exception:
                pass
        elif system == "Darwin":
            out = safe_run(["system_profiler", "SPMemoryDataType"], timeout=8)
            for line in out.splitlines():
                line = line.strip()
                if line.startswith("Type:"):
                    types.append(line.split(":", 1)[1].strip())
                elif line.startswith("Speed:"):
                    speeds.append(line.split(":", 1)[1].strip())
        else:
            # Linux: dmidecode needs root, so this only succeeds if run
            # with sudo or the binary has the right capabilities set.
            out = safe_run(["dmidecode", "-t", "17"], timeout=8)
            for line in out.splitlines():
                line = line.strip()
                if line.startswith("Type:") and "Unknown" not in line:
                    types.append(line.split(":", 1)[1].strip())
                elif line.startswith("Speed:") and "Unknown" not in line:
                    speeds.append(line.split(":", 1)[1].strip())
    except Exception:
        pass

    def most_common(items):
        if not items:
            return "N/A"
        return max(set(items), key=items.count)

    return {"type": most_common(types), "speed": most_common(speeds)}


def get_motherboard_info():
    """Best-effort motherboard manufacturer + model. Returns a dict
    {"brand": .., "model": ..}, "Unidentified" if unavailable."""
    system = platform.system()
    brand, model = "Unidentified", "Unidentified"
    try:
        if system == "Windows":
            try:
                for row in _ps_cim("Win32_BaseBoard", ["Manufacturer", "Product"]):
                    b = row.get("Manufacturer", "").strip()
                    p = row.get("Product", "").strip()
                    if b or p:
                        brand, model = b or "Unidentified", p or "Unidentified"
                        break
            except Exception:
                pass
        elif system == "Darwin":
            out = safe_run(["system_profiler", "SPHardwareDataType"], timeout=8)
            for line in out.splitlines():
                line = line.strip()
                if line.startswith("Model Identifier:"):
                    brand, model = "Apple", line.split(":", 1)[1].strip()
                    break
        else:
            # /sys/class/dmi/id is usually world-readable, unlike
            # dmidecode which needs root - try it first.
            try:
                with open("/sys/class/dmi/id/board_vendor") as f:
                    brand = f.read().strip() or "Unidentified"
            except Exception:
                pass
            try:
                with open("/sys/class/dmi/id/board_name") as f:
                    model = f.read().strip() or "Unidentified"
            except Exception:
                pass
            if brand == "Unidentified" or model == "Unidentified":
                out = safe_run(["dmidecode", "-t", "baseboard"], timeout=8)
                for line in out.splitlines():
                    line = line.strip()
                    if line.startswith("Manufacturer:"):
                        brand = line.split(":", 1)[1].strip() or brand
                    elif line.startswith("Product Name:"):
                        model = line.split(":", 1)[1].strip() or model
    except Exception:
        pass
    return {"brand": brand or "Unidentified", "model": model or "Unidentified"}


def normalize_socket_designation(raw):
    """Keep an OS-reported socket string only if it actually looks like
    a real socket name (LGA/AM/PGA/BGA/...). BIOS/firmware often reports
    junk placeholders like 'Socket 0', 'Other', or 'U3E1' instead."""
    if not raw:
        return None
    val = raw.strip()
    if not val or val.upper() in ("N/A", "UNKNOWN", "OTHER", "TO BE FILLED BY O.E.M."):
        return None
    if re.fullmatch(r"SOCKET\s*\d+", val, re.I):
        return None  # generic BIOS placeholder, not a real socket name
    if re.search(r"\b(LGA\d*|AM[2-5]\+?|PGA\d*|BGA\d*|FCBGA|FCLGA\d*|FM[12]\+?|STRX\d*|TR\d|SP\d)\b", val, re.I):
        return val
    return None


def infer_cpu_socket_from_name(cpu_name):
    """Best-effort socket guess (LGA/AM/PGA/BGA family) from the CPU's
    model string, used when the OS doesn't report a usable socket name.
    Approximate -- covers common consumer Intel/AMD chips only."""
    if not cpu_name:
        return None
    n = cpu_name.upper()

    # AMD: Ryzen, Athlon, FX, A-series APUs
    if re.search(r"\b(RYZEN|FX|ATHLON)\b", n) or re.search(r"\bA(4|6|8|9|10|12)\b", n):
        if "THREADRIPPER" in n:
            return "sTRX4 / TRX40"
        if "FX-" in n or "FX " in n or re.search(r"\bFX\d", n):
            return "AM3+"
        if "ATHLON" in n and not re.search(r"\bRYZEN\b", n):
            return "AM4"
        m = re.search(r"(?<!\d)(\d)\d{3}", n)
        gen = m.group(1) if m else None
        if gen in ("7", "8", "9"):
            return "AM5"
        if gen in ("1", "2", "3", "4", "5"):
            return "AM4"
        return "AM4"

    # Intel Core / Core Ultra
    if "ULTRA" in n:
        return "LGA1851"
    m = re.search(r"\bI[3579]-?\s?(\d{3,5})[A-Z]*\b", n)
    if m:
        digits = m.group(1)
        if len(digits) == 3:
            return "LGA1156"  # 1st-gen Core (Lynnfield/Clarkdale era), e.g. i5-760
        gen = int(digits[:2]) if len(digits) >= 5 else int(digits[0])
        if gen >= 12:
            return "LGA1700"
        if gen >= 10:
            return "LGA1200"
        if gen >= 6:
            return "LGA1151"
        if gen >= 4:
            return "LGA1150"
        if gen >= 2:
            return "LGA1155"
        return None

    # Intel Pentium/Celeron desktop (G-series) or mobile (N-series)
    m = re.search(r"\bG(\d)\d{3}\b", n)
    if m:
        gen = int(m.group(1))
        if gen >= 7:
            return "LGA1700"
        if gen == 6:
            return "LGA1200"
        return "LGA1151"
    if re.search(r"\bN\d{3,5}\b", n):
        return "BGA (Soldered)"

    return None


def get_cpu_socket():
   
    system = platform.system()
    raw = None
    try:
        if system == "Windows":
            try:
                for row in _ps_cim("Win32_Processor", ["SocketDesignation"]):
                    val = row.get("SocketDesignation", "").strip()
                    if val:
                        raw = val
                        break
            except Exception:
                pass
        elif system == "Darwin":
            # macOS doesn't expose a socket designation; Apple Silicon
            # and modern Intel Macs are soldered (no socket).
            raw = None
        else:
            out = safe_run(["dmidecode", "-t", "processor"], timeout=8)
            for line in out.splitlines():
                line = line.strip()
                if line.startswith("Socket Designation:"):
                    val = line.split(":", 1)[1].strip()
                    if val and "Unknown" not in val:
                        raw = val
                    break
    except Exception:
        pass

    clean = normalize_socket_designation(raw)
    if clean:
        return clean

    inferred = infer_cpu_socket_from_name(get_cpu_display_name())
    if inferred:
        return inferred

    return "N/A"


def get_gpu_info():
    """Best-effort GPU detection without extra third-party deps."""
    gpus = []
    # Try NVIDIA via nvidia-smi
    nv = safe_run(["nvidia-smi", "--query-gpu=name,memory.total,driver_version",
                   "--format=csv,noheader"])
    if nv:
        for line in nv.splitlines():
            bits = [b.strip() for b in line.split(",")]
            if len(bits) >= 3:
                # nvidia-smi reports memory as e.g. "4096 MiB" -- normalize
                # the suffix to "MB" for consistent labeling with the
                # other GPU memory sources below, which also use MB.
                mem = bits[1].replace("MiB", "MB")
                gpus.append({"name": bits[0], "memory": mem, "driver": bits[2]})
    if gpus:
        return gpus

    system = platform.system()
    if system == "Windows":
        try:
            seen_names = set()
            for row in _ps_cim("Win32_VideoController", ["Name", "AdapterRAM", "DriverVersion"]):
                name = row.get("Name", "").strip()
                if not name or name in seen_names:
                    continue
                seen_names.add(name)
                try:
                    ram = human_mb(int(row.get("AdapterRAM", "0").strip() or "0"))
                except Exception:
                    ram = "Unidentified"
                driver = row.get("DriverVersion", "N/A").strip() or "N/A"
                gpus.append({"name": name, "memory": ram, "driver": driver})
        except Exception:
            pass
    elif system == "Darwin":
        out = safe_run(["system_profiler", "SPDisplaysDataType"])
        for line in out.splitlines():
            if "Chipset Model" in line:
                gpus.append({"name": line.split(":", 1)[1].strip(),
                             "memory": "N/A", "driver": "N/A"})
    else:
        out = safe_run(["lspci"])
        for line in out.splitlines():
            if "VGA" in line or "3D controller" in line:
                gpus.append({"name": line.split(":")[-1].strip(),
                             "memory": "Unidentified", "driver": "Unidentified"})
    return gpus


# ---------------------------------------------------------------------------
# Reusable UI bits
# ---------------------------------------------------------------------------
class Card(QFrame):
    def __init__(self, label, value, accent=False):
        super().__init__()
        self.setProperty("class", "Card")
        self.setStyleSheet("")  # picked up by .Card selector
        self.setObjectName("CardFrame")
        self.setMinimumHeight(80)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(6)

        lbl = QLabel(label.upper())
        lbl.setProperty("class", "CardLabel")
        val = QLabel(str(value))
        val.setProperty("class", "CardValue")
        val.setWordWrap(True)
        if accent:
            val.setStyleSheet(f"color: {ACCENT};")

        layout.addWidget(lbl)
        layout.addWidget(val)
        self.value_label = val


def section_header(title, subtitle=""):
    box = QVBoxLayout()
    box.setSpacing(2)
    t = QLabel(title)
    t.setObjectName("PageTitle")
    box.addWidget(t)
    if subtitle:
        s = QLabel(subtitle)
        s.setObjectName("PageSubtitle")
        box.addWidget(s)
    return box


def make_scroll(inner_widget):
    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setWidget(inner_widget)
    return scroll


def styled_table(headers):
    table = QTableWidget()
    table.setColumnCount(len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.verticalHeader().setVisible(False)
    table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
    table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
    table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
    table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    table.setShowGrid(False)
    table.setAlternatingRowColors(False)
    return table


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------
class OverviewPage(QWidget):
    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(32, 28, 32, 28)
        outer.setSpacing(18)

        header_row = QHBoxLayout()
        header_row.addLayout(section_header(
            "Overview", "A snapshot of this machine"))
        header_row.addStretch()
        export_btn = QPushButton("Export Report")
        export_btn.clicked.connect(self.export_report)
        header_row.addWidget(export_btn, alignment=Qt.AlignmentFlag.AlignTop)
        outer.addLayout(header_row)

        grid = QGridLayout()
        grid.setSpacing(14)
        outer.addLayout(grid)

        uname = platform.uname()
        mem = psutil.virtual_memory()

        self.cards_info = [
            ("OS", f"{uname.system} {uname.release}"),
            ("CPU", "—"),
            ("GPU", "—"),
            ("RAM", human_mb(mem.total)),
            ("Internal Disk Name", "—"),
        ]

        self.cards = []
        for i, (label, value) in enumerate(self.cards_info):
            card = Card(label, value)
            self.cards.append(card)
            grid.addWidget(card, i // 3, i % 3)

        outer.addSpacing(6)
        temp_label = QLabel("Temperatures")
        temp_label.setStyleSheet(f"color: {TEXT_DIM}; font-size: 11px; "
                                  f"font-weight: 600; letter-spacing: 0.5px;")
        outer.addWidget(temp_label)
        temp_grid = QGridLayout()
        temp_grid.setSpacing(14)
        self.cpu_temp_card = Card("CPU Temperature", "—")
        self.gpu_temp_card = Card("GPU Temperature", "—")
        self.mobo_temp_card = Card("Motherboard Temperature", "—")
        temp_grid.addWidget(self.cpu_temp_card, 0, 0)
        temp_grid.addWidget(self.gpu_temp_card, 0, 1)
        temp_grid.addWidget(self.mobo_temp_card, 0, 2)
        outer.addLayout(temp_grid)

        self.temp_note = QLabel("")
        self.temp_note.setStyleSheet(f"color: {TEXT_DIM}; font-size: 11px;")
        outer.addWidget(self.temp_note)

        outer.addStretch()

    def update_temperatures(self, temps):
        self.cpu_temp_card.value_label.setText(format_temp_c(best_cpu_temp(temps)))
        gpu_vals = gpu_temps_in_order(temps)
        self.gpu_temp_card.value_label.setText(
            format_temp_c(gpu_vals[0]) if gpu_vals else "Unidentified")
        self.mobo_temp_card.value_label.setText(format_temp_c(best_motherboard_temp(temps)))

    def set_temp_status(self, message):
        # Only surface the note when sensors are unavailable; once
        # live data is flowing the note clears itself.
        self.temp_note.setText("" if message == "Live (LibreHardwareMonitor)" else message)

    def export_report(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save AlphaInfo Report", "AlphaInfo_report.txt", "Text Files (*.txt)")
        if not path:
            return
        lines = [f"{APP_NAME} v{APP_VERSION} — System Report",
                 f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ""]
        for label, value in self.cards_info:
            lines.append(f"{label}: {value}")
        lines.append("")
        lines.append("CPU usage per core (%): " +
                      ", ".join(f"{p:.0f}" for p in psutil.cpu_percent(percpu=True, interval=0.3)))
        mem = psutil.virtual_memory()
        lines.append(f"Memory used: {human_mb(mem.used)} / {human_mb(mem.total)} "
                      f"({mem.percent}%)")
        lines.append("")
        lines.append("Disk partitions:")
        for part in psutil.disk_partitions(all=False):
            try:
                u = psutil.disk_usage(part.mountpoint)
                lines.append(f"  {part.device} ({part.mountpoint}) — "
                              f"{human_bytes(u.used)} / {human_bytes(u.total)} ({u.percent}%)")
            except Exception:
                continue
        lines.append("")
        lines.append("Network interfaces:")
        for name, addrs in psutil.net_if_addrs().items():
            for a in addrs:
                if a.family == socket.AF_INET:
                    lines.append(f"  {name}: {a.address}")
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))


class CPUPage(QWidget):
    def __init__(self):
        super().__init__()
        # Wrap everything in a scroll area
        _root = QVBoxLayout(self)
        _root.setContentsMargins(0, 0, 0, 0)
        _scroll = QScrollArea()
        _scroll.setWidgetResizable(True)
        _scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        _inner = QWidget()
        _scroll.setWidget(_inner)
        _root.addWidget(_scroll)
        outer = QVBoxLayout(_inner)
        outer.setContentsMargins(32, 28, 32, 28)
        outer.setSpacing(16)
        outer.addLayout(section_header("Processor", get_cpu_name()))

        self.overall_bar = QProgressBar()
        self.overall_bar.setFormat("Overall CPU Usage: %p%")
        outer.addWidget(self.overall_bar)

        freq_row = QHBoxLayout()
        self.freq_card = Card("Current Frequency", "—")
        self.cores_card = Card("Physical Cores", str(psutil.cpu_count(logical=False) or "N/A"))
        self.threads_card = Card("Threads", str(psutil.cpu_count(logical=True)))
        self.socket_card = Card("Socket Type", "—")
        self.temp_card = Card("CPU Temperature", "—")
        freq_row.addWidget(self.freq_card)
        freq_row.addWidget(self.cores_card)
        freq_row.addWidget(self.threads_card)
        freq_row.addWidget(self.socket_card)
        freq_row.addWidget(self.temp_card)
        outer.addLayout(freq_row)



        # Static fallback value (what psutil/registry reports).
        self.base_mhz = get_base_clock_mhz()
        if self.base_mhz:
            self.freq_card.value_label.setText(format_freq_ghz(self.base_mhz))

        # On Windows, switch to a live estimate from performance counters
        # once the background poller returns its first reading.
        self.freq_poller = None
        self._live_active = False
        if platform.system() == "Windows" and self.base_mhz:
            # PDH polling is cheap (in-process, no subprocess spawn) so
            # it can run more often than the PowerShell fallback, which
            # pays ~0.5-1s of powershell.exe startup cost per sample.
            interval = 1.0 if HAVE_PYWIN32 else 2.0
            self.freq_poller = FreqPoller(self.base_mhz, interval=interval)
            self.freq_poller.freq_updated.connect(self._on_live_freq)
            self.freq_poller.start()

        self.core_bars = []
        core_grid = QGridLayout()
        core_grid.setSpacing(10)
        n = psutil.cpu_count(logical=True) or 1
        cols = 4
        for i in range(n):
            box = QVBoxLayout()
            lbl = QLabel(f"Core {i}")
            lbl.setStyleSheet(f"color: {TEXT_DIM}; font-size: 11px;")
            bar = QProgressBar()
            box.addWidget(lbl)
            box.addWidget(bar)
            wrap = QWidget()
            wrap.setLayout(box)
            core_grid.addWidget(wrap, i // cols, i % cols)
            self.core_bars.append(bar)
        outer.addLayout(core_grid)
        self._socket_loaded = False

        # Database spec section — populated later by apply_static_data
        outer.addSpacing(6)
        self._db_spec_label = QLabel("CPU Specifications")
        self._db_spec_label.setStyleSheet(
            f"color: {TEXT_DIM}; font-size: 11px; font-weight: 600; letter-spacing: 0.5px;")
        self._db_spec_label.setVisible(False)
        outer.addWidget(self._db_spec_label)
        self._db_grid_widget = QWidget()
        self._db_grid_layout = QGridLayout(self._db_grid_widget)
        self._db_grid_layout.setSpacing(14)
        self._db_grid_widget.setVisible(False)
        outer.addWidget(self._db_grid_widget)

        outer.addStretch()

    def _on_live_freq(self, freq_mhz):
        self._live_active = True
        self.freq_card.value_label.setText(format_freq_ghz(freq_mhz))

    def apply_static_data(self, data):
        self.socket_card.value_label.setText(data.get("cpu_socket", "N/A"))

        specs = data.get("cpu_db_specs") or {}
        # Filter out empty values
        specs = {k: v for k, v in specs.items() if v and str(v).strip()}
        if not specs:
            return

        # Key display order
        order = [
            "Base Frequency", "Max Turbo", "TDP", "Cache",
            "L1 Cache", "L2 Cache", "L3 Cache",
            "Lithography", "Launch Date",
            "Memory Types", "Max Memory", "Memory Channels",
            "Socket", "TjMax", "PCIe Version",
            "Hyper-Threading", "Turbo Boost", "Instruction Set",
        ]
        ordered = [(k, specs[k]) for k in order if k in specs]
        # Append any remaining keys not in order
        ordered += [(k, v) for k, v in specs.items() if k not in dict(ordered)]

        cols = 3
        for i, (label, value) in enumerate(ordered):
            self._db_grid_layout.addWidget(Card(label, value), i // cols, i % cols)

        self._db_spec_label.setVisible(True)
        self._db_grid_widget.setVisible(True)

    def refresh(self):
        per_cpu = psutil.cpu_percent(percpu=True)
        overall = psutil.cpu_percent()
        self.overall_bar.setValue(int(overall))
        for bar, pct in zip(self.core_bars, per_cpu):
            bar.setValue(int(pct))
        # Only fall back to psutil's static reading if the live Windows
        # poller hasn't produced a value yet (or isn't running at all).
        if not getattr(self, "_live_active", False):
            try:
                freq = psutil.cpu_freq()
                if freq:
                    self.freq_card.value_label.setText(format_freq_ghz(freq.current))
            except Exception:
                pass

    def stop_polling(self):
        if self.freq_poller:
            self.freq_poller.stop()
            self.freq_poller.wait(3000)

    def update_temperatures(self, temps):
        self.temp_card.value_label.setText(format_temp_c(best_cpu_temp(temps)))


class MemoryPage(QWidget):
    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(32, 28, 32, 28)
        outer.setSpacing(16)
        outer.addLayout(section_header("Memory", "RAM and swap usage"))

        self.ram_bar = QProgressBar()
        outer.addWidget(QLabel("Physical Memory (RAM)"))
        outer.addWidget(self.ram_bar)

        grid = QGridLayout()
        grid.setSpacing(14)
        self.total_card = Card("Total", "—")
        self.used_card = Card("Used", "—")
        self.avail_card = Card("Available", "—")
        for i, c in enumerate([self.total_card, self.used_card, self.avail_card]):
            grid.addWidget(c, 0, i)
        outer.addLayout(grid)

        outer.addSpacing(10)
        self.swap_bar = QProgressBar()
        outer.addWidget(QLabel("Swap"))
        outer.addWidget(self.swap_bar)

        swap_grid = QGridLayout()
        swap_grid.setSpacing(14)
        self.swap_total_card = Card("Swap Total", "—")
        self.swap_used_card = Card("Swap Used", "—")
        swap_grid.addWidget(self.swap_total_card, 0, 0)
        swap_grid.addWidget(self.swap_used_card, 0, 1)
        outer.addLayout(swap_grid)

        outer.addSpacing(10)
        ram_spec_label = QLabel("RAM Specifications")
        ram_spec_label.setStyleSheet(f"color: {TEXT_DIM}; font-size: 11px; "
                                      f"font-weight: 600; letter-spacing: 0.5px;")
        outer.addWidget(ram_spec_label)
        spec_grid = QGridLayout()
        spec_grid.setSpacing(14)
        self.ddr_type_card = Card("DDR Type", "—")
        self.ram_speed_card = Card("Frequency / Speed", "—")
        spec_grid.addWidget(self.ddr_type_card, 0, 0)
        spec_grid.addWidget(self.ram_speed_card, 0, 1)
        outer.addLayout(spec_grid)

        outer.addStretch()

    def apply_static_data(self, data):
        ram_specs = data.get("ram_specs", {})
        self.ddr_type_card.value_label.setText(ram_specs.get("type", "N/A"))
        self.ram_speed_card.value_label.setText(ram_specs.get("speed", "N/A"))

    def refresh(self):
        mem = psutil.virtual_memory()
        self.ram_bar.setValue(int(mem.percent))
        self.total_card.value_label.setText(human_mb(mem.total))
        self.used_card.value_label.setText(human_mb(mem.used))
        self.avail_card.value_label.setText(human_mb(mem.available))

        swap = psutil.swap_memory()
        self.swap_bar.setValue(int(swap.percent))
        self.swap_total_card.value_label.setText(human_mb(swap.total))
        self.swap_used_card.value_label.setText(human_mb(swap.used))


class MainboardPage(QWidget):
    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(32, 28, 32, 28)
        outer.setSpacing(16)
        outer.addLayout(section_header("Mainboard", "Motherboard Informations"))

        outer.addWidget(QLabel("Board Info"))
        hw_grid = QGridLayout()
        hw_grid.setSpacing(14)
        self.mobo_brand_card = Card("Motherboard Brand", "—")
        self.mobo_model_card = Card("Motherboard Model", "—")
        self.socket_card = Card("CPU Socket Type", "—")
        self.mobo_temp_card = Card("Motherboard Temperature", "—")
        hw_grid.addWidget(self.mobo_brand_card, 0, 0)
        hw_grid.addWidget(self.mobo_model_card, 0, 1)
        hw_grid.addWidget(self.socket_card, 1, 0)
        hw_grid.addWidget(self.mobo_temp_card, 1, 1)
        outer.addLayout(hw_grid)
        self._hw_loaded = False

        outer.addStretch()

    def apply_static_data(self, data):
        mobo = data.get("mobo", {})
        self.mobo_brand_card.value_label.setText(mobo.get("brand", "N/A"))
        self.mobo_model_card.value_label.setText(mobo.get("model", "N/A"))
        self.socket_card.value_label.setText(data.get("cpu_socket", "N/A"))

    def refresh(self):
        pass  # all data loaded via apply_static_data from background thread

    def update_temperatures(self, temps):
        self.mobo_temp_card.value_label.setText(format_temp_c(best_motherboard_temp(temps)))


class DiskPage(QWidget):
    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(32, 28, 32, 28)
        outer.setSpacing(16)
        outer.addLayout(section_header("Storage", "Mounted disks"))

        self.table = styled_table(["Disk Name", "Mount Point", "File System",
                                    "Used", "Total", "Usage", "Temperature"])
        self.table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.Stretch)
        outer.addWidget(self.table)

        # Disk model lookup involves spawning a subprocess (PowerShell/
        # lsblk/diskutil), so cache it instead of redoing it on every
        # timer-driven refresh tick.
        self._name_map = None
        # Latest LHM storage temperature readings, updated separately
        # by the shared TempPoller and just re-read here on each
        # refresh() tick rather than rebuilding the table from a signal.
        self._storage_temps = {}

    def apply_static_data(self, data):
        self._name_map = data.get("disk_name_map", {})

    def refresh(self):
        parts = psutil.disk_partitions(all=False)
        name_map = self._name_map or {}
        self.table.setRowCount(len(parts))
        for row, part in enumerate(parts):
            try:
                usage = psutil.disk_usage(part.mountpoint)
                used = human_bytes(usage.used)
                total = human_bytes(usage.total)
                pct = f"{usage.percent}%"
            except Exception:
                used, total, pct = "Unidentified", "Unidentified", "Unidentified"
            disk_name = name_map.get(part.device) or part.device
            if disk_name == part.device and "cdrom" in (part.opts or ""):
                disk_name = "DVD/CD Drive"
            temp = match_storage_temp(disk_name, self._storage_temps)
            is_usb = part.device in _USB_DEVICES or any(
                k in disk_name.lower()
                for k in ("usb", "flash", "datatraveler", "sandisk cruzer",
                          "kingston dt", "generic flash")
            )
            temp_str = "—" if is_usb else format_temp_c(temp)
            for col, text in enumerate([disk_name, part.mountpoint,
                                         part.fstype, used, total, pct,
                                         temp_str]):
                self.table.setItem(row, col, QTableWidgetItem(text))

    def update_temperatures(self, temps):
        self._storage_temps = storage_temps_by_hardware(temps)


class NetworkPage(QWidget):
    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(32, 28, 32, 28)
        outer.setSpacing(16)
        outer.addLayout(section_header("Network", "Interfaces and traffic"))

        grid = QGridLayout()
        grid.setSpacing(14)
        self.sent_card = Card("Total Sent", "—")
        self.recv_card = Card("Total Received", "—")
        self.host_card = Card("Hostname", socket.gethostname())
        self.temp_card = Card("Adapter Temperature", "—")
        grid.addWidget(self.sent_card, 0, 0)
        grid.addWidget(self.recv_card, 0, 1)
        grid.addWidget(self.host_card, 0, 2)
        grid.addWidget(self.temp_card, 0, 3)
        outer.addLayout(grid)

        self.table = styled_table(["Interface", "IPv4 Address", "MAC Address", "Status"])
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch)
        outer.addWidget(self.table)

    def refresh(self):
        io = psutil.net_io_counters()
        self.sent_card.value_label.setText(human_bytes(io.bytes_sent))
        self.recv_card.value_label.setText(human_bytes(io.bytes_recv))

        addrs = psutil.net_if_addrs()
        stats = psutil.net_if_stats()
        self.table.setRowCount(len(addrs))
        for row, (name, addr_list) in enumerate(addrs.items()):
            ipv4 = next((a.address for a in addr_list if a.family == socket.AF_INET), "—")
            mac = next((a.address for a in addr_list
                        if a.family.name == "AF_PACKET" or a.family.name == "AF_LINK"), "—")
            up = stats.get(name)
            status = "Up" if up and up.isup else "Down"
            for col, text in enumerate([name, ipv4, mac, status]):
                self.table.setItem(row, col, QTableWidgetItem(text))

    def update_temperatures(self, temps):
        self.temp_card.value_label.setText(format_temp_c(best_network_temp(temps)))


def get_live_gpu_clock_data(gpus):
    """Live core clock, memory clock, and core voltage for each GPU in
    `gpus` (the list returned by get_gpu_info()).

    Returns {gpu_index: {"core_clock": str, "mem_clock": str, "voltage": str}}
    with "N/A" for any value that isn't available. NVIDIA GPUs are
    matched by index via nvidia-smi; AMD/Intel GPUs are matched by name
    via ADL/OpenCL (get_opencl_live_gpu_specs), since their enumeration
    order isn't guaranteed to line up with the basic wmic-based list.

    This is called both once (to populate the GPU page initially) and
    repeatedly on a background thread (GpuLivePoller) to keep the
    Core Clock / Memory Clock / Core Voltage cards updating in real
    time, the same way the CPU page's frequency card does.
    """
    nv_live = get_nvidia_live_gpu_specs()
    ocl_devices = get_opencl_live_gpu_specs()

    def _ocl_match(gpu_name):
        gpu_name_l = (gpu_name or "").lower()
        for ocl in ocl_devices:
            ocl_name_l = ocl["name"].lower()
            if ocl_name_l in gpu_name_l or gpu_name_l in ocl_name_l:
                return ocl
        return {}

    result = {}
    for i, gpu in enumerate(gpus):
        name = gpu.get("name", "Unknown")
        nv_live_entry = nv_live.get(i, {})
        ocl_entry = _ocl_match(name)

        if nv_live_entry.get("core_clock"):
            core_clock = f"{nv_live_entry['core_clock']} MHz"
        elif ocl_entry.get("core_clock"):
            core_clock = f"{ocl_entry['core_clock']} MHz"
        else:
            core_clock = "N/A"

        if nv_live_entry.get("mem_clock"):
            mem_clock = f"{nv_live_entry['mem_clock']} MHz"
        elif ocl_entry.get("mem_clock"):
            mem_clock = f"{ocl_entry['mem_clock']} MHz"
        else:
            mem_clock = "N/A"

        voltage = (nv_live_entry.get("voltage")
                   or ocl_entry.get("voltage")
                   or "N/A")

        result[i] = {"core_clock": core_clock, "mem_clock": mem_clock,
                      "voltage": voltage}
    return result


class GpuLivePoller(QThread):
    """Polls live GPU core/memory clock + voltage on a background
    thread, the same way FreqPoller does for CPU frequency: chunked
    sleep so stop() takes effect quickly instead of blocking app close,
    and all the actual hardware queries (nvidia-smi / ADL / OpenCL)
    happen off the UI thread since they involve subprocess/DLL calls
    that can take noticeable time.
    """
    data_updated = pyqtSignal(dict)

    def __init__(self, gpus, interval=2.0):
        super().__init__()
        self.gpus = gpus
        self.interval = interval
        self._running = True

    def run(self):
        while self._running:
            try:
                data = get_live_gpu_clock_data(self.gpus)
                if data:
                    self.data_updated.emit(data)
            except Exception:
                pass

            slept = 0.0
            while slept < self.interval and self._running:
                self.msleep(100)
                slept += 0.1

    def stop(self):
        self._running = False


class GPUPage(QWidget):
    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setContentsMargins(32, 28, 32, 28)
        outer.setSpacing(16)
        outer.addLayout(section_header("Graphics", "Detected GPU"))

        self.container = QVBoxLayout()
        outer.addLayout(self.container)
        outer.addStretch()
        self.loaded = False
        self.live_cards = {}  # {gpu_index: {"core_clock": Card, "mem_clock": Card, "voltage": Card}}
        self.poller = None

    def apply_static_data(self, data):
        if self.loaded:
            return
        self.loaded = True
        gpus = data.get("gpus") or []
        self._preloaded_gpus = gpus
        if not gpus:
            note = QLabel("No dedicated GPU detected, or vendor tools "
                           "(nvidia-smi / wmic / system_profiler) are unavailable.")
            note.setStyleSheet(f"color: {TEXT_DIM};")
            note.setWordWrap(True)
            self.container.addWidget(note)
            return

        smi_gpus = data.get("gpu_smi") or []
        live_data = data.get("gpu_live") or {}

        grid = QGridLayout()
        grid.setSpacing(14)
        row = 0
        for i, gpu in enumerate(gpus):
            name = gpu.get("name", "Unknown")
            arch, vram_type = get_gpu_arch_and_vram_type(name)
            grid.addWidget(Card(f"GPU {i} Name", name), row, 0)
            grid.addWidget(Card("Memory", gpu.get("memory", "N/A")), row, 1)
            grid.addWidget(Card("Driver Version", gpu.get("driver", "N/A")), row, 2)
            row += 1

            grid.addWidget(Card("GPU Architecture", arch), row, 0)
            grid.addWidget(Card("VRAM Type (estimated)", vram_type), row, 1)

            live = live_data.get(i, {"core_clock": "N/A", "mem_clock": "N/A", "voltage": "N/A"})
            core_clock_card = Card("Core Clock", live["core_clock"])
            grid.addWidget(core_clock_card, row, 2)
            row += 1

            mem_clock_card = Card("Memory Clock", live["mem_clock"])
            voltage_card = Card("Core Voltage", live["voltage"])
            temp_card = Card("Temperature", "—")
            grid.addWidget(mem_clock_card, row, 0)
            grid.addWidget(voltage_card, row, 1)
            grid.addWidget(temp_card, row, 2)

            self.live_cards[i] = {
                "core_clock": core_clock_card,
                "mem_clock": mem_clock_card,
                "voltage": voltage_card,
                "temperature": temp_card,
            }

            if i < len(smi_gpus):
                smi = smi_gpus[i]
                vram = smi.get("memory_total")
                if vram:
                    grid.addWidget(Card("Exact VRAM (device-smi)",
                                         human_mb(vram)), row, 2)
                    row += 1
            else:
                row += 1

            if i < len(smi_gpus):
                smi = smi_gpus[i]
                pcie_gen = smi.get("pcie_gen")
                pcie_speed = smi.get("pcie_speed")
                firmware = smi.get("firmware")
                if pcie_gen or firmware:
                    pcie_text = (f"Gen {pcie_gen} x{pcie_speed}"
                                 if pcie_gen and pcie_speed else "N/A")
                    grid.addWidget(Card("PCIe Link", pcie_text), row, 0)
                    grid.addWidget(Card("Firmware", firmware or "N/A"), row, 1)
                    row += 1
        self.container.addLayout(grid)

        self.poller = GpuLivePoller(gpus)
        self.poller.data_updated.connect(self._on_gpu_live_data)
        self.poller.start()

    def refresh(self):
        pass  # UI is built in apply_static_data from the background loader

    def _on_gpu_live_data(self, data):
        for i, vals in data.items():
            cards = self.live_cards.get(i)
            if not cards:
                continue
            cards["core_clock"].value_label.setText(vals.get("core_clock", "N/A"))
            cards["mem_clock"].value_label.setText(vals.get("mem_clock", "N/A"))
            cards["voltage"].value_label.setText(vals.get("voltage", "N/A"))

    def stop_polling(self):
        if self.poller:
            self.poller.stop()
            self.poller.wait(3000)

    def update_temperatures(self, temps):
        gpu_vals = gpu_temps_in_order(temps)
        for i, cards in self.live_cards.items():
            value = gpu_vals[i] if i < len(gpu_vals) else None
            cards["temperature"].value_label.setText(format_temp_c(value))


# ---------------------------------------------------------------------------
# Live temperatures (LibreHardwareMonitor)
# ---------------------------------------------------------------------------
class LibreHardwareMonitorBridge:
    """Wraps a local copy of LibreHardwareMonitorLib (via pythonnet) to
    read live temperature sensors for CPU, GPU, Motherboard/SuperIO,
    Storage and Network hardware.

    Mirrors how the other vendor-specific readers in this file behave:
    every failure mode (pythonnet not installed, DLL missing,
    non-Windows, not running elevated, driver unavailable) is caught
    and simply leaves the bridge unavailable rather than raising, so
    the UI just keeps showing its placeholder temperature values.
    """

    def __init__(self, dll_path=LHM_DLL_PATH):
        self.dll_path = dll_path
        self.error = None
        self._computer = None
        self._HardwareType = None
        self._SensorType = None

    def open(self):
        if not HAVE_PYTHONNET:
            self.error = "pythonnet not installed"
            return False
        if platform.system() != "Windows":
            self.error = "Windows only"
            return False

        resolved = os.path.abspath(self.dll_path)
        if not os.path.isfile(resolved):
            self.error = f"DLL not found: {resolved}"
            return False

        try:
            dll_dir = os.path.dirname(resolved)
            if dll_dir not in sys.path:
                sys.path.append(dll_dir)
            assembly_name = os.path.splitext(os.path.basename(resolved))[0]
            clr.AddReference(assembly_name)

            from LibreHardwareMonitor.Hardware import Computer, HardwareType, SensorType
            self._HardwareType = HardwareType
            self._SensorType = SensorType

            computer = Computer()
            computer.IsCpuEnabled = True
            computer.IsGpuEnabled = True
            computer.IsMotherboardEnabled = True
            computer.IsStorageEnabled = True
            computer.IsNetworkEnabled = True
            computer.IsMemoryEnabled = False  # Memory tab has no temp card
            computer.Open()
            self._computer = computer
            return True
        except Exception as e:
            self.error = str(e)
            self._computer = None
            return False

    def read_temperatures(self):
        """Returns {"cpu": [...], "gpu": [...], "motherboard": [...],
        "storage": [...], "network": [...]}, each a list of
        {"hardware": str, "sensor": str, "value": float} readings."""
        result = {"cpu": [], "gpu": [], "motherboard": [], "storage": [], "network": []}
        if self._computer is None:
            return result
        HardwareType = self._HardwareType
        SensorType = self._SensorType
        try:
            for hw in self._computer.Hardware:
                hw.Update()
                self._collect(hw, result, HardwareType, SensorType)
                for sub in hw.SubHardware:
                    sub.Update()
                    self._collect(sub, result, HardwareType, SensorType)
        except Exception:
            pass
        return result

    @staticmethod
    def _collect(hw, result, HardwareType, SensorType):
        ht = hw.HardwareType
        if ht == HardwareType.Cpu:
            key = "cpu"
        elif ht in (HardwareType.GpuNvidia, HardwareType.GpuAmd, HardwareType.GpuIntel):
            key = "gpu"
        elif ht in (HardwareType.Motherboard, HardwareType.SuperIO):
            key = "motherboard"
        elif ht == HardwareType.Storage:
            key = "storage"
        elif ht == HardwareType.Network:
            key = "network"
        else:
            return
        for sensor in hw.Sensors:
            try:
                if sensor.SensorType == SensorType.Temperature and sensor.Value is not None:
                    result[key].append({
                        "hardware": str(hw.Name),
                        "sensor": str(sensor.Name),
                        "value": float(sensor.Value),
                    })
            except Exception:
                continue

    def close(self):
        if self._computer is not None:
            try:
                self._computer.Close()
            except Exception:
                pass
            self._computer = None


def best_cpu_temp(temps):
    """Pick a single representative CPU temperature: a package/die/
    Tctl-style aggregate sensor if present, otherwise the hottest core."""
    readings = temps.get("cpu", [])
    if not readings:
        return None
    for r in readings:
        name = r["sensor"].lower()
        if "package" in name or "tctl" in name or "tdie" in name or "core max" in name:
            return r["value"]
    return max(r["value"] for r in readings)


def gpu_temps_in_order(temps):
    """GPU temperatures grouped by hardware device, in the order LHM
    reported them - matched positionally against get_gpu_info()'s list
    elsewhere, the same best-effort approach already used for
    device-smi GPU matching."""
    grouped = {}
    order = []
    for r in temps.get("gpu", []):
        hw = r["hardware"]
        if hw not in grouped:
            grouped[hw] = []
            order.append(hw)
        grouped[hw].append(r)
    out = []
    for hw in order:
        items = grouped[hw]
        val = None
        for it in items:
            sensor_name = it["sensor"].lower()
            if "core" in sensor_name or "hot spot" in sensor_name:
                val = it["value"]
                break
        out.append(val if val is not None else max(i["value"] for i in items))
    return out


def best_motherboard_temp(temps):
    readings = temps.get("motherboard", [])
    if not readings:
        return None
    for r in readings:
        name = r["sensor"].lower()
        if "motherboard" in name or "system" in name:
            return r["value"]
    return readings[0]["value"]


def storage_temps_by_hardware(temps):
    """{hardware_name: temperature} for each storage device LHM sees."""
    out = {}
    for r in temps.get("storage", []):
        out.setdefault(r["hardware"], r["value"])
    return out


def match_storage_temp(disk_name, storage_temps):
    """Best-effort fuzzy match between a psutil/wmic disk display name
    and an LHM storage hardware name, since the two tools don't share
    a common identifier."""
    if not disk_name or not storage_temps:
        return None
    needle = disk_name.lower()
    for hw_name, value in storage_temps.items():
        hay = hw_name.lower()
        if needle in hay or hay in needle:
            return value
    return None


def best_network_temp(temps):
    readings = temps.get("network", [])
    if not readings:
        return None
    return readings[0]["value"]


class TempPoller(QThread):
    """Polls live temperature sensors from LibreHardwareMonitor on a
    background thread, the same pattern as FreqPoller/GpuLivePoller.

    Opening the underlying Computer object talks to a kernel driver
    and can take a moment, plus typically needs Administrator
    privileges on Windows, so this happens off the UI thread. If
    opening fails for any reason, the thread emits a status message
    once and exits without ever emitting temperature data - pages
    simply keep showing their placeholder values.
    """
    temps_updated = pyqtSignal(dict)
    status_changed = pyqtSignal(str)

    def __init__(self, interval=2.0):
        super().__init__()
        self.interval = interval
        self._running = True
        self._bridge = None

    def run(self):
        self._bridge = LibreHardwareMonitorBridge()
        if not self._bridge.open():
            self.status_changed.emit(self._bridge.error or "Temperature sensors unavailable")
            return
        self.status_changed.emit("Live (LibreHardwareMonitor)")

        while self._running:
            try:
                temps = self._bridge.read_temperatures()
                self.temps_updated.emit(temps)
            except Exception:
                pass
            slept = 0.0
            while slept < self.interval and self._running:
                self.msleep(100)
                slept += 0.1

        self._bridge.close()

    def stop(self):
        self._running = False


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------
class StaticDataLoader(QThread):
    """Loads all slow one-time data (PS/subprocess calls) in background
    at startup so tab switches are instant."""
    finished = pyqtSignal(dict)

    def run(self):
        data = {}
        try:
            data["ram_specs"] = get_ram_module_specs()
        except Exception:
            data["ram_specs"] = {"type": "N/A", "speed": "N/A"}
        try:
            data["mobo"] = get_motherboard_info()
        except Exception:
            data["mobo"] = {"brand": "N/A", "model": "N/A"}
        try:
            data["cpu_socket"] = get_cpu_socket()
        except Exception:
            data["cpu_socket"] = "N/A"
        try:
            data["disk_name_map"] = get_disk_name_map()
        except Exception:
            data["disk_name_map"] = {}
        try:
            data["primary_disk"] = get_primary_disk_name()
        except Exception:
            data["primary_disk"] = "Unknown"
        try:
            data["gpus"] = get_gpu_info()
        except Exception:
            data["gpus"] = []
        try:
            data["gpu_smi"] = get_smi_gpu_info()
        except Exception:
            data["gpu_smi"] = []
        try:
            data["gpu_live"] = get_live_gpu_clock_data(data.get("gpus") or [])
        except Exception:
            data["gpu_live"] = {}
        try:
            data["cpu_name"] = get_cpu_display_name()
        except Exception:
            data["cpu_name"] = platform.processor() or "Unknown CPU"
        try:
            data["cpu_db_specs"] = lookup_cpu_in_database(data.get("cpu_name", ""))
        except Exception:
            data["cpu_db_specs"] = {}
        self.finished.emit(data)



class TitleBar(QWidget):
    """Custom frameless titlebar matching the sidebar color."""
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self._drag_pos = None
        self.setFixedHeight(38)
        self.setObjectName("TitleBar")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 0, 8, 0)
        layout.setSpacing(0)

        icon_lbl = QLabel("AlphaCorp™")
        icon_lbl.setStyleSheet("font-size: 14px; background: transparent; padding-right: 6px;")
        layout.addWidget(icon_lbl)

        title = QLabel(APP_NAME)
        title.setStyleSheet(f"color: {TEXT_MAIN}; font-size: 13px; font-weight: 600; background: transparent;")
        layout.addWidget(title)
        layout.addStretch()

        btn_style = (
            "QPushButton { background: transparent; border: none; color: %s;"
            " font-size: 15px; padding: 0 10px; min-width: 38px; min-height: 38px; }"
            " QPushButton:hover { background: %s; border-radius: 0px; }"
        )

        self.min_btn = QPushButton("─")
        self.min_btn.setStyleSheet(btn_style % (TEXT_DIM, "#2a2c3c"))
        self.min_btn.clicked.connect(parent.showMinimized)

        self.max_btn = QPushButton("□")
        self.max_btn.setStyleSheet(btn_style % (TEXT_DIM, "#2a2c3c"))
        self.max_btn.clicked.connect(self._toggle_max)

        self.close_btn = QPushButton("✕")
        self.close_btn.setStyleSheet(btn_style % (TEXT_DIM, "#c0392b"))
        self.close_btn.clicked.connect(parent.close)

        for btn in (self.min_btn, self.max_btn, self.close_btn):
            layout.addWidget(btn)

    def _toggle_max(self):
        if self.parent.isMaximized():
            self.parent.showNormal()
        else:
            self.parent.showMaximized()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.parent.frameGeometry().topLeft()

    def mouseMoveEvent(self, event):
        if self._drag_pos and event.buttons() == Qt.MouseButton.LeftButton:
            self.parent.move(event.globalPosition().toPoint() - self._drag_pos)

    def mouseReleaseEvent(self, event):
        self._drag_pos = None

    def mouseDoubleClickEvent(self, event):
        self._toggle_max()


class AlphaInfoWindow(QMainWindow):
    PAGES = ["Overview", "CPU", "Memory", "Mainboard", "Storage", "Network", "Graphics"]

    def __init__(self, preloaded_data=None):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}")
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.resize(980, 640)
        self.setMinimumSize(820, 560)

        central = QWidget()
        self.setCentralWidget(central)
        root_v = QVBoxLayout(central)
        root_v.setContentsMargins(0, 0, 0, 0)
        root_v.setSpacing(0)

        # Custom titlebar
        self.titlebar = TitleBar(self)
        root_v.addWidget(self.titlebar)

        # Main content row
        content = QWidget()
        root = QHBoxLayout(content)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root_v.addWidget(content)

        # Sidebar
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(220)
        side_layout = QVBoxLayout(sidebar)
        side_layout.setContentsMargins(0, 0, 0, 0)
        side_layout.setSpacing(0)

        title = QLabel(APP_NAME)
        title.setObjectName("SidebarTitle")
        subtitle = QLabel(f"v{APP_VERSION} · Informations")
        subtitle.setObjectName("SidebarSubtitle")
        side_layout.addWidget(title)
        side_layout.addWidget(subtitle)

        self.nav = QListWidget()
        self.nav.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.nav.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        for name in self.PAGES:
            QListWidgetItem(name, self.nav)
        self.nav.setCurrentRow(0)
        self.nav.currentRowChanged.connect(self.change_page)
        side_layout.addWidget(self.nav)

        root.addWidget(sidebar)

        # Pages
        self.stack = QStackedWidget()
        self.overview_page = OverviewPage()
        self.cpu_page = CPUPage()
        self.memory_page = MemoryPage()
        self.mainboard_page = MainboardPage()
        self.disk_page = DiskPage()
        self.network_page = NetworkPage()
        self.gpu_page = GPUPage()

        for page in [self.overview_page, self.cpu_page, self.memory_page,
                     self.mainboard_page, self.disk_page, self.network_page, self.gpu_page]:
            self.stack.addWidget(page)

        root.addWidget(self.stack, 1)

        # Live temperature sensors (LibreHardwareMonitor), shared
        # across every page that shows one. Opened once here rather
        # than per-page, since opening the underlying Computer object
        # more than once is wasteful (and can conflict) with the
        # hardware driver it talks to.
        self.temp_poller = TempPoller()
        self.temp_poller.temps_updated.connect(self._on_temps_updated)
        self.temp_poller.status_changed.connect(self._on_temp_status)
        self.temp_poller.start()

        if preloaded_data is not None:
            self._on_static_data(preloaded_data)
        else:
            self._static_loader = StaticDataLoader()
            self._static_loader.finished.connect(self._on_static_data)
            self._static_loader.start()

        # Live refresh timer for dynamic pages
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.refresh_current_page)
        self.timer.start(1500)
        self.refresh_current_page()

    def _on_static_data(self, data):
        """Received from background loader - push static data to all pages."""
        cpu_name = data.get("cpu_name", "Unknown CPU")
        gpus = data.get("gpus") or []
        gpu_name = get_gpu_display_name(gpus[0]["name"]) if gpus else "No dedicated GPU detected"
        primary_disk = data.get("primary_disk", "Unknown")
        for card, (label, _) in zip(self.overview_page.cards, self.overview_page.cards_info):
            if label == "CPU":
                card.value_label.setText(cpu_name)
            elif label == "GPU":
                card.value_label.setText(gpu_name)
            elif label == "Internal Disk Name":
                card.value_label.setText(primary_disk)
        for page in (self.cpu_page, self.memory_page, self.mainboard_page, self.disk_page, self.gpu_page):
            if hasattr(page, "apply_static_data"):
                try:
                    page.apply_static_data(data)
                except Exception:
                    pass

    def change_page(self, index):
        self.stack.setCurrentIndex(index)
        self.refresh_current_page()

    def refresh_current_page(self):
        idx = self.stack.currentIndex()
        widget = self.stack.widget(idx)
        if hasattr(widget, "refresh"):
            widget.refresh()

    def _on_temps_updated(self, temps):
        # Every page except Memory gets a temperature reading.
        for page in (self.overview_page, self.cpu_page, self.mainboard_page,
                     self.disk_page, self.network_page, self.gpu_page):
            if hasattr(page, "update_temperatures"):
                try:
                    page.update_temperatures(temps)
                except Exception:
                    pass

    def _on_temp_status(self, message):
        if hasattr(self.overview_page, "set_temp_status"):
            self.overview_page.set_temp_status(message)

    def closeEvent(self, event):
        if hasattr(self.cpu_page, "stop_polling"):
            self.cpu_page.stop_polling()
        if hasattr(self.gpu_page, "stop_polling"):
            self.gpu_page.stop_polling()
        if hasattr(self, "temp_poller"):
            self.temp_poller.stop()
            self.temp_poller.wait(3000)
        event.accept()


def _is_admin():
    """Return True if the current process has Administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def _relaunch_as_admin():
    exe = sys.executable.replace("python.exe", "pythonw.exe")
    if not os.path.exists(exe):
        exe = sys.executable
    params = '"' + os.path.abspath(sys.argv[0]) + '"'
    ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, params, None, 0)


class _PreloadWorker(QObject):
    finished = pyqtSignal(dict)

    def run(self):
        loader = StaticDataLoader()
        result = {}
        loader.finished.connect(lambda d: result.update(d))
        loader.run()
        self.finished.emit(result)


class LoadingWindow(QWidget):
    """512x512 frameless window showing IMG/load.tga while data loads."""
    ready = pyqtSignal(dict)

    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        self.setFixedSize(512, 512)

        img_path = get_resource_path("IMG", "load.png")
        self._pix = QPixmap(img_path)

        # Use a QLabel to display the image — simplest and most reliable
        lbl = QLabel(self)
        lbl.setGeometry(0, 0, 512, 512)
        if not self._pix.isNull():
            lbl.setPixmap(self._pix.scaled(512, 512, Qt.AspectRatioMode.KeepAspectRatio,
                                           Qt.TransformationMode.SmoothTransformation))
        else:
            # Fallback if TGA fails to load
            lbl.setStyleSheet("background-color: #1a1a2e;")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setText("Loading…")

        screen = QApplication.primaryScreen().availableGeometry()
        self.move((screen.width() - 512) // 2, (screen.height() - 512) // 2)

        self._thread = QThread()
        self._worker = _PreloadWorker()
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_done)
        self._thread.start()

    def _on_done(self, data):
        self._thread.quit()
        self._thread.wait()
        self.ready.emit(data)


def main():
    if platform.system() == "Windows" and not _is_admin():
        _relaunch_as_admin()
        os._exit(0)

    # Single-instance guard via named mutex
    if platform.system() == "Windows":
        _mutex = ctypes.windll.kernel32.CreateMutexW(None, True, "AlphaINF_SingleInstance")
        if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            hwnd = ctypes.windll.user32.FindWindowW(None, APP_NAME)
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 9)   # SW_RESTORE
                ctypes.windll.user32.SetForegroundWindow(hwnd)
            os._exit(0)

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLESHEET)
    _icon_path = get_resource_path("icon.ico")
    if os.path.exists(_icon_path):
        app.setWindowIcon(QIcon(_icon_path))

    loader_win = LoadingWindow()

    def _launch(data):
        loader_win.close()
        window = AlphaInfoWindow(preloaded_data=data)
        app._main_window = window
        window.show()

    loader_win.ready.connect(_launch)
    loader_win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()