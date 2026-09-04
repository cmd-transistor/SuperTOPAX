"""
Topax E-Test Data Extractor
============================
GUI for extracting Topax .tgz data files from Intel E-Test servers
(EOL and Inline) and saving them as CSVs.

Requires : pandas, PyUber (Intel internal)
Standard  : tkinter (bundled with Python)
"""

import fnmatch
import io
import os
import glob
import queue
import shlex
import subprocess
import sys
import tarfile
import re
import threading
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ── Optional PyUber import ────────────────────────────────────────────────────
try:
    import PyUber
    _PYUBER_OK = True
except ImportError:
    _PYUBER_OK = False

# ── Constants ─────────────────────────────────────────────────────────────────
_EOL_SERVER    = r'\\rf2p-sram-nas2.rf2prod.mfg.intel.com\pasram'
_INLINE_SERVER = r'\\rf3p-nas-patest.rf3prod.mfg.intel.com\patest'
_PYUBER_DSN    = 'D1D_PROD_XARIES'
_ZIP_PREFIX    = 'v93cleanup_'
_DATA_TYPES    = ['SML', 'VBD', 'FLICKER_NOISE','CAP_SWEEP_VATCTARGET_NOSTRAY', 'TPX']

# Sibling scripts launched independently by the "Plots" buttons (each keeps its
# own CLI defaults/config; this GUI does not pass any extraction settings to them).
_IDVG_PLOTTER_SCRIPT     = Path(__file__).with_name('topax_idvg_plotter.py')
_WAFER_SPARKLINE_SCRIPT  = Path(__file__).with_name('topax_wafer_sparkline.py')

# Default output base folder: honor TOPAX_OUTPUT_DIR if set, else the user's
# Documents folder.  Never defaults to a network share — see _is_network_path().
_ENV_OUTPUT_DIR_VAR  = 'TOPAX_OUTPUT_DIR'
_DEFAULT_OUTPUT_BASE = Path(os.environ[_ENV_OUTPUT_DIR_VAR]) if os.environ.get(_ENV_OUTPUT_DIR_VAR, '').strip() \
    else Path.home() / 'Documents' / 'TopaxData'


def _is_network_path(path: str) -> bool:
    """True if `path` is a UNC path or resolves to a mapped network drive."""
    p = path.strip()
    if not p:
        return False
    if p.startswith('\\\\') or p.startswith('//'):
        return True
    drive, _ = os.path.splitdrive(p)
    if drive and os.name == 'nt':
        try:
            import ctypes
            DRIVE_REMOTE = 4
            if ctypes.windll.kernel32.GetDriveTypeW(f"{drive}\\") == DRIVE_REMOTE:
                return True
        except Exception:
            pass
    return False


def _list_subdirs(path: str) -> list:
    """Sorted immediate subfolder names of `path`, or [] if inaccessible."""
    try:
        return sorted(
            name for name in os.listdir(path)
            if os.path.isdir(os.path.join(path, name))
        )
    except OSError:
        return []


# ══════════════════════════════════════════════════════════════════════════════
# Data-access helpers
# ══════════════════════════════════════════════════════════════════════════════

def _sanitize_lot(pattern: str) -> str:
    """Only allow characters valid in a lot / SQL LIKE pattern."""
    if not re.fullmatch(r'[A-Za-z0-9_%*?-]+', pattern):
        raise ValueError(f"Invalid lot pattern: {pattern!r}")
    return pattern


def _parse_lot_patterns(raw: str) -> list:
    """
    Split a comma-separated lot input into individual SQL LIKE patterns.

    Exact lot IDs that contain no wildcard character automatically receive a
    trailing '%' so that child lots are included too.  For example, 'D608EM10'
    becomes 'D608EM10%' which matches both 'D608EM10' and 'D608EM10H'.
    Patterns that already contain '*' or '?' are left as the user typed them.
    """
    patterns = []
    for tok in raw.split(','):
        tok = tok.strip()
        if tok:
            has_wildcard = '*' in tok or '?' in tok
            sql_pat = _sanitize_lot(tok).replace('*', '%').replace('?', '_')
            if not has_wildcard:
                sql_pat += '%'   # auto-include child lots
            patterns.append(sql_pat)
    if not patterns:
        raise ValueError("At least one Lot ID is required.")
    return patterns


def _sanitize_ops(ops: list) -> list:
    """Validate operation codes — numeric (7118) or alphanumeric (ENGG)."""
    result = []
    for op in ops:
        op = op.strip()
        if not re.fullmatch(r'[A-Za-z0-9_]+', op):
            raise ValueError(f"Invalid operation code: {op!r}")
        result.append(op)
    return result


def query_lot_info(lot_patterns: list, operations: list) -> dict:
    """
    Query PyUber and return {full_lot_id: (etest_process, product)}.

    Parameters
    ----------
    lot_patterns : list  e.g. ["D547FWV0", "D551F%"]  (SQL LIKE patterns)
    operations   : list  e.g. ["7118", "7106"]
    """
    if not _PYUBER_OK:
        raise RuntimeError(
            "PyUber is not installed.  Install it and ensure you are on "
            "the Intel network."
        )

    ops       = _sanitize_ops(operations)
    op_clause = ', '.join(f"'{op}'" for op in ops)
    lot_clause = '\n           OR '.join(f"ets.lot LIKE '{p}'" for p in lot_patterns)

    sql = f"""
        SELECT DISTINCT
              ets.lot         AS LOT
            , ets.process     AS ETEST_PROCESS
            , ets.devrevstep  AS PRODUCT
        FROM  A_ETest_Testing_Session ets
        WHERE ets.operation IN ({op_clause})
          AND ({lot_clause})
    """

    conn = PyUber.connect(_PYUBER_DSN)
    df   = pd.read_sql(sql, conn)

    return {
        row['LOT']: (row['ETEST_PROCESS'], row['PRODUCT'])
        for _, row in df.iterrows()
    }


# ── File discovery ────────────────────────────────────────────────────────────

def find_live_tgz(base_path: str, lot_id: str, struct_patterns: list) -> list:
    """
    Locate .tgz files under  base_path / lot_id / ** / <pattern>.

    Returns list of (abs_path, wafer, die_xy, tgz_filename).
    """
    found = []
    for pat in struct_patterns:
        search = os.path.join(base_path, lot_id, '**', pat)
        for fpath in glob.glob(search, recursive=True):
            parts = fpath.replace('/', os.sep).split(os.sep)
            if len(parts) >= 3:
                wafer    = parts[-3]
                die_xy   = parts[-2]
                tgz_name = parts[-1]
            else:
                wafer = die_xy = 'unknown'
                tgz_name = os.path.basename(fpath)
            found.append((fpath, wafer, die_xy, tgz_name))
    return found


def find_zip_archives(base_path: str, lot_id: str) -> list:
    """Return all v93cleanup_{lot_id}_*.zip archives under base_path."""
    return glob.glob(
        os.path.join(base_path, f'{_ZIP_PREFIX}{lot_id}_*.zip')
    )


def find_tgz_in_zip(zip_path: str, struct_patterns: list) -> list:
    """
    Find matching .tgz entries inside a zip archive.

    Wafer/die are parsed from the last 3 path components inside the zip
    (expected layout: .../<wafer>/<die_xy>/<file>.tgz).  If the internal
    structure differs, wafer/die fall back to 'unknown'.

    Returns list of (zip_path, entry_name, wafer, die_xy, tgz_filename).
    """
    found = []
    with zipfile.ZipFile(zip_path, 'r') as zf:
        for entry in zf.namelist():
            fname = entry.replace('\\', '/').split('/')[-1]
            for pat in struct_patterns:
                if fnmatch.fnmatch(fname, pat):
                    parts = entry.replace('\\', '/').rstrip('/').split('/')
                    if len(parts) >= 3:
                        wafer    = parts[-3]
                        die_xy   = parts[-2]
                        tgz_name = parts[-1]
                    elif len(parts) == 2:
                        wafer    = parts[-2]
                        die_xy   = 'unknown'
                        tgz_name = parts[-1]
                    else:
                        wafer = die_xy = 'unknown'
                        tgz_name = fname
                    found.append((zip_path, entry, wafer, die_xy, tgz_name))
                    break   # don't double-add if multiple patterns match
    return found


# ── TGZ parsing ───────────────────────────────────────────────────────────────

def _parse_topax(source) -> pd.DataFrame:
    """
    Parse a Topax .tgz file (tar.gz archive containing a single CSV).

    All Topax tgz files are proper tar.gz archives.  The inner CSV always
    carries its own column-header row.  SML / VBD / FLICKER_NOISE files
    additionally have a '# Advantest SML Library' comment as the very first
    line, which pandas skips via comment='#'.

    `source` may be a file-path string or a BytesIO object.
    """
    if isinstance(source, io.BytesIO):
        source.seek(0)
        tf = tarfile.open(fileobj=source, mode='r:gz')
    else:
        tf = tarfile.open(source, mode='r:gz')

    with tf:
        csv_members = [m for m in tf.getmembers()
                       if m.name.lower().endswith('.csv') and m.isfile()]
        fobj = tf.extractfile(csv_members[0])
        return pd.read_csv(fobj, comment='#', on_bad_lines='skip')


def read_tgz_file(path: str) -> pd.DataFrame:
    return _parse_topax(path)


def read_tgz_from_zip(zip_path: str, entry_name: str) -> pd.DataFrame:
    with zipfile.ZipFile(zip_path, 'r') as zf:
        data = zf.read(entry_name)
    return _parse_topax(io.BytesIO(data))


# ══════════════════════════════════════════════════════════════════════════════
# GUI
# ══════════════════════════════════════════════════════════════════════════════

class _App:

    _PAD = dict(padx=8, pady=3)

    def __init__(self, root: tk.Tk):
        self.root    = root
        self.root.title("Topax E-Test Data Extractor")
        self.root.minsize(820, 660)

        self._log_q    = queue.Queue()
        self._stop     = threading.Event()
        self._log_file = None          # set while a run is active
        self._out_auto_value = ""      # last value we auto-computed for the Output Folder

        self._build_widgets()
        self._poll_log()

    # ── Widget construction ───────────────────────────────────────────────────

    def _build_widgets(self):
        p = self._PAD

        # ── Input panel ──────────────────────────────────────────────────────
        frm = ttk.LabelFrame(self.root, text="Parameters", padding=6)
        frm.pack(fill=tk.X, padx=10, pady=(8, 4))

        def lbl(text, row):
            ttk.Label(frm, text=text).grid(row=row, column=0, sticky=tk.E, **p)

        def hint(text, row, col=2, span=1):
            ttk.Label(frm, text=text, foreground='grey').grid(
                row=row, column=col, columnspan=span, sticky=tk.W, **p)

        # ── Lot ID
        lbl("Lot ID *", 0)
        self._v_lot = tk.StringVar()
        ttk.Entry(frm, textvariable=self._v_lot, width=32).grid(
            row=0, column=1, sticky=tk.W, **p)
        hint("e.g. D547FWV0  or  D547F*  or  D547F*, D551FYE*  — child lots (e.g. D608EM10H) auto-included", 0)

        # ── Operation(s)
        lbl("Operation(s) *", 1)
        self._v_op = tk.StringVar()
        ttk.Entry(frm, textvariable=self._v_op, width=32).grid(
            row=1, column=1, sticky=tk.W, **p)
        hint("e.g. 7118  or  7118, 7106  or  ENGG", 1)

        # ── Test type (EOL / Inline)
        lbl("Test Type *", 2)
        self._v_mode = tk.StringVar(value='EOL')
        cmb_mode = ttk.Combobox(frm, textvariable=self._v_mode,
                     values=['EOL', 'Inline'], width=10,
                     state='readonly')
        cmb_mode.grid(row=2, column=1, sticky=tk.W, **p)
        cmb_mode.bind('<<ComboboxSelected>>', self._on_test_type_changed)

        # ── Process Node (scanned from the server root; optional override)
        lbl("Process Node", 3)
        self._v_proc_node = tk.StringVar()
        self._cmb_proc_node = ttk.Combobox(frm, textvariable=self._v_proc_node, width=16)
        self._cmb_proc_node.grid(row=3, column=1, sticky=tk.W, **p)
        self._cmb_proc_node.bind('<<ComboboxSelected>>', self._on_process_node_changed)
        self._cmb_proc_node.bind('<FocusOut>', self._on_process_node_changed)
        ttk.Button(frm, text="Auto-detect from Lot ID", command=self._on_lookup_from_lot).grid(
            row=3, column=2, sticky=tk.W, **p)

        # ── Test Chip (scanned after Process Node is entered/selected)
        lbl("Test Chip", 4)
        self._v_test_chip = tk.StringVar()
        self._cmb_test_chip = ttk.Combobox(frm, textvariable=self._v_test_chip, width=16, state='disabled')
        self._cmb_test_chip.grid(row=4, column=1, sticky=tk.W, **p)
        self._cmb_test_chip.bind('<<ComboboxSelected>>', lambda e: self._refresh_auto_output_dir())
        hint("Populated from the selected Process Node's folder. Optional.", 4)

        # ── Data pull type
        lbl("Data Type *", 5)
        self._v_dtype = tk.StringVar(value='SML')
        ttk.Combobox(frm, textvariable=self._v_dtype,
                     values=_DATA_TYPES, width=16,
                     state='readonly').grid(row=5, column=1, sticky=tk.W, **p)

        # ── Structure pattern(s)
        lbl("Structure(s) *", 6)
        self._v_struct = tk.StringVar(value='*.tgz')
        ttk.Entry(frm, textvariable=self._v_struct, width=56).grid(
            row=6, column=1, columnspan=2, sticky=tk.EW, **p)
        hint("Comma-separated for multiple.  Default: *.tgz\n"
             "e.g.  *_PRE_SORT_ROSN-144-50*M3WAC*tgz, *_PRE_SORT_ROSP-144-50*M3WAC*tgz\n"
             "TPX: filename is the same for every die (e.g. _CMOS_UNI_ACDOE27_TPX_IDVG_VD.tgz) — "
             "leave as *.tgz; structures are split by the 'Device' column inside the CSV, not the filename.",
             7, col=1, span=2)

        # ── Output folder
        lbl("Output Folder *", 8)
        self._v_out = tk.StringVar()
        self._v_lot.trace_add('write', lambda *a: self._refresh_auto_output_dir())
        frm_out = ttk.Frame(frm)
        frm_out.grid(row=8, column=1, columnspan=2, sticky=tk.W, **p)
        ttk.Entry(frm_out, textvariable=self._v_out, width=72).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(frm_out, text="Browse…", command=self._browse).pack(
            side=tk.LEFT, padx=4)

        # ── Sub-folder checkbox
        self._v_subdirs = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frm,
            text="Auto-create sub-folders:  {Process}\\{Lot7}\\Topax\\",
            variable=self._v_subdirs,
        ).grid(row=9, column=1, columnspan=2, sticky=tk.W, **p)

        # ── Skip-existing checkbox
        self._v_skip_existing = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frm,
            text="Skip files that already exist in the output folder (don't overwrite)",
            variable=self._v_skip_existing,
        ).grid(row=10, column=1, columnspan=2, sticky=tk.W, **p)

        frm.columnconfigure(1, weight=1)

        self._refresh_auto_output_dir()

        # ── Buttons ───────────────────────────────────────────────────────────
        frm_btns = ttk.Frame(self.root)
        frm_btns.pack(fill=tk.X, padx=10, pady=(2, 4))

        self._btn_run  = ttk.Button(frm_btns, text="▶  Run",  command=self._on_run)
        self._btn_stop = ttk.Button(frm_btns, text="■  Stop", command=self._on_stop,
                                    state=tk.DISABLED)
        self._btn_clr  = ttk.Button(frm_btns, text="Clear Log", command=self._clear_log)

        self._btn_run.pack(side=tk.LEFT, padx=(0, 4))
        self._btn_stop.pack(side=tk.LEFT, padx=4)
        self._btn_clr.pack(side=tk.LEFT, padx=12)

        # ── Plots (independent of the extraction settings above) ─────────────
        frm_plots_outer = ttk.LabelFrame(self.root, text="Plots", padding=6)
        frm_plots_outer.pack(fill=tk.X, padx=10, pady=(0, 4))

        nb = ttk.Notebook(frm_plots_outer)
        nb.pack(fill=tk.X)

        tab_idvg = ttk.Frame(nb, padding=6)
        tab_spark = ttk.Frame(nb, padding=6)
        nb.add(tab_idvg, text="Plot IDVG")
        nb.add(tab_spark, text="Wafer Sparkline")

        self._build_idvg_tab(tab_idvg)
        self._build_sparkline_tab(tab_spark)

        # Auto-fill the Plot tabs' Topax Dir / Output Dir from the Extractor's
        # Output Folder, unless the user has since typed something else there.
        self._plot_dir_auto_value = ""
        self._v_out.trace_add('write', self._sync_plot_dirs_from_extractor)
        self._sync_plot_dirs_from_extractor()

        # ── Status bar ────────────────────────────────────────────────────────
        frm_status = ttk.Frame(self.root)
        frm_status.pack(fill=tk.X, padx=10, pady=(0, 2))

        self._v_status = tk.StringVar(value="Ready")
        ttk.Label(frm_status, textvariable=self._v_status).pack(side=tk.LEFT)

        self._v_pct = tk.DoubleVar(value=0.0)
        ttk.Progressbar(
            frm_status, variable=self._v_pct,
            maximum=100, length=300, mode='determinate',
        ).pack(side=tk.RIGHT, padx=4)

        # ── Run log ───────────────────────────────────────────────────────────
        frm_log = ttk.LabelFrame(self.root, text="Run Log", padding=4)
        frm_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(2, 8))

        self._log_txt = scrolledtext.ScrolledText(
            frm_log, height=18, wrap=tk.WORD, font=("Consolas", 9),
        )
        self._log_txt.pack(fill=tk.BOTH, expand=True)
        # Prevent accidental edits while keeping copy/select
        self._log_txt.bind("<Key>", lambda e: "break")

    # ── Plot tab construction ─────────────────────────────────────────────────

    def _browse_dir_into(self, var: tk.StringVar):
        d = filedialog.askdirectory(title="Select Folder")
        if d:
            var.set(d)

    def _add_path_row(self, parent, label, row, var):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.E, padx=6, pady=2)
        frm = ttk.Frame(parent)
        frm.grid(row=row, column=1, columnspan=3, sticky=tk.EW, padx=6, pady=2)
        ttk.Entry(frm, textvariable=var, width=52).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(frm, text="Browse…", command=lambda: self._browse_dir_into(var)).pack(side=tk.LEFT, padx=4)

    def _add_entry_row(self, parent, label, row, col, var, width=20, hint=None):
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky=tk.E, padx=6, pady=2)
        ttk.Entry(parent, textvariable=var, width=width).grid(row=row, column=col + 1, sticky=tk.W, padx=6, pady=2)
        if hint:
            ttk.Label(parent, text=hint, foreground='grey').grid(row=row, column=col + 2, sticky=tk.W, padx=6, pady=2)

    def _build_idvg_tab(self, parent):
        # Leave every field blank by default so topax_idvg_plotter.py's own
        # module-level defaults are used unless the user overrides them here.
        self._idvg_topax_dir   = tk.StringVar()
        self._idvg_output_dir  = tk.StringVar()
        self._idvg_lot_prefixes = tk.StringVar()
        self._idvg_wafers      = tk.StringVar()
        self._idvg_structures  = tk.StringVar()
        self._idvg_vds         = tk.StringVar()
        self._idvg_max_files   = tk.StringVar()
        self._idvg_plot_source = tk.BooleanVar(value=True)
        self._idvg_plot_gate   = tk.BooleanVar(value=True)
        self._idvg_plot_body   = tk.BooleanVar(value=True)
        self._idvg_skip_ppt    = tk.BooleanVar(value=False)
        self._idvg_extra_args  = tk.StringVar()

        self._add_path_row(parent, "Topax Dir", 0, self._idvg_topax_dir)
        self._add_path_row(parent, "Output Dir", 1, self._idvg_output_dir)
        self._add_entry_row(parent, "Lot Prefixes", 2, 0, self._idvg_lot_prefixes, hint="blank = all lots found")
        self._add_entry_row(parent, "Wafers", 3, 0, self._idvg_wafers, hint="e.g. 250,353  (blank = script default)")
        self._add_entry_row(parent, "Structures", 4, 0, self._idvg_structures, width=40, hint="exact structure name(s), comma-separated")
        self._add_entry_row(parent, "|VDS| magnitudes", 5, 0, self._idvg_vds, hint="e.g. 0.65,0.05")
        self._add_entry_row(parent, "Max files", 6, 0, self._idvg_max_files, width=8, hint="blank = no cap (debug aid)")

        chk_frame = ttk.Frame(parent)
        chk_frame.grid(row=7, column=0, columnspan=4, sticky=tk.W, padx=6, pady=(4, 2))
        ttk.Checkbutton(chk_frame, text="Plot Source (Is)", variable=self._idvg_plot_source).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(chk_frame, text="Plot Gate (Ig)", variable=self._idvg_plot_gate).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(chk_frame, text="Plot Body (Ib)", variable=self._idvg_plot_body).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(chk_frame, text="Skip PPT (PNGs + summary only)", variable=self._idvg_skip_ppt).pack(side=tk.LEFT, padx=4)

        self._add_entry_row(parent, "Extra CLI args", 8, 0, self._idvg_extra_args, width=60,
                             hint="any other --flag from topax_idvg_plotter.py --help")

        self._btn_plot_idvg = ttk.Button(
            parent, text="▶  Run Plot IDVG",
            command=lambda: self._on_run_plot_script(_IDVG_PLOTTER_SCRIPT, self._btn_plot_idvg, self._build_idvg_args),
        )
        self._btn_plot_idvg.grid(row=9, column=0, columnspan=2, sticky=tk.W, padx=6, pady=(6, 2))

        parent.columnconfigure(1, weight=1)

    def _build_sparkline_tab(self, parent):
        self._spark_topax_dir   = tk.StringVar()
        self._spark_output_dir  = tk.StringVar()
        self._spark_lot_prefixes = tk.StringVar()
        self._spark_wafers       = tk.StringVar()
        self._spark_structure_substring = tk.StringVar()
        self._spark_device_type  = tk.StringVar(value="")
        self._spark_vds          = tk.StringVar()
        self._spark_overlay_vt   = tk.BooleanVar(value=True)
        self._spark_plot_source  = tk.BooleanVar(value=True)
        self._spark_plot_gate    = tk.BooleanVar(value=True)
        self._spark_extra_args   = tk.StringVar()

        self._add_path_row(parent, "Topax Dir", 0, self._spark_topax_dir)
        self._add_path_row(parent, "Output Dir", 1, self._spark_output_dir)
        self._add_entry_row(parent, "Lot Prefixes", 2, 0, self._spark_lot_prefixes, hint="blank = script default")
        self._add_entry_row(parent, "Wafers", 3, 0, self._spark_wafers, hint="comma-separated, or 'all'")
        self._add_entry_row(parent, "Structure substring", 4, 0, self._spark_structure_substring, width=40)
        ttk.Label(parent, text="Device type").grid(row=5, column=0, sticky=tk.E, padx=6, pady=2)
        ttk.Combobox(parent, textvariable=self._spark_device_type, values=["", "NS", "PS"], width=6,
                     state='readonly').grid(row=5, column=1, sticky=tk.W, padx=6, pady=2)
        self._add_entry_row(parent, "|VDS| magnitude", 6, 0, self._spark_vds, hint="single value, e.g. 0.65")

        chk_frame = ttk.Frame(parent)
        chk_frame.grid(row=7, column=0, columnspan=4, sticky=tk.W, padx=6, pady=(4, 2))
        ttk.Checkbutton(chk_frame, text="Overlay VT flavors", variable=self._spark_overlay_vt).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(chk_frame, text="Plot Source (Is)", variable=self._spark_plot_source).pack(side=tk.LEFT, padx=4)
        ttk.Checkbutton(chk_frame, text="Plot Gate (Ig)", variable=self._spark_plot_gate).pack(side=tk.LEFT, padx=4)

        self._add_entry_row(parent, "Extra CLI args", 8, 0, self._spark_extra_args, width=60,
                             hint="any other --flag from topax_wafer_sparkline.py --help")

        self._btn_plot_sparkline = ttk.Button(
            parent, text="▶  Run Wafer Sparkline",
            command=lambda: self._on_run_plot_script(_WAFER_SPARKLINE_SCRIPT, self._btn_plot_sparkline, self._build_sparkline_args),
        )
        self._btn_plot_sparkline.grid(row=9, column=0, columnspan=2, sticky=tk.W, padx=6, pady=(6, 2))

        parent.columnconfigure(1, weight=1)

    def _build_idvg_args(self) -> list:
        args = []
        if self._idvg_topax_dir.get().strip():
            args += ["--topax-dir", self._idvg_topax_dir.get().strip()]
        if self._idvg_output_dir.get().strip():
            args += ["--output-dir", self._idvg_output_dir.get().strip()]
        if self._idvg_lot_prefixes.get().strip():
            args += ["--lot-prefixes", self._idvg_lot_prefixes.get().strip()]
        if self._idvg_wafers.get().strip():
            args += ["--wafers", self._idvg_wafers.get().strip()]
        if self._idvg_structures.get().strip():
            args += ["--structures", self._idvg_structures.get().strip()]
        if self._idvg_vds.get().strip():
            args += ["--include-vds-magnitudes", self._idvg_vds.get().strip()]
        if self._idvg_max_files.get().strip():
            args += ["--max-files", self._idvg_max_files.get().strip()]
        if not self._idvg_plot_source.get():
            args.append("--no-plot-both-currents")
        if not self._idvg_plot_gate.get():
            args.append("--no-plot-gate-current")
        if not self._idvg_plot_body.get():
            args.append("--no-plot-body-current")
        if self._idvg_skip_ppt.get():
            args.append("--skip-ppt")
        args += shlex.split(self._idvg_extra_args.get().strip())
        return args

    def _build_sparkline_args(self) -> list:
        args = []
        if self._spark_topax_dir.get().strip():
            args += ["--topax-dir", self._spark_topax_dir.get().strip()]
        if self._spark_output_dir.get().strip():
            args += ["--output-dir", self._spark_output_dir.get().strip()]
        if self._spark_lot_prefixes.get().strip():
            args += ["--lot-prefixes", self._spark_lot_prefixes.get().strip()]
        if self._spark_wafers.get().strip():
            args += ["--wafers", self._spark_wafers.get().strip()]
        if self._spark_structure_substring.get().strip():
            args += ["--structure-substring", self._spark_structure_substring.get().strip()]
        if self._spark_device_type.get().strip():
            args += ["--device-type", self._spark_device_type.get().strip()]
        if self._spark_vds.get().strip():
            args += ["--vds-magnitude", self._spark_vds.get().strip()]
        if not self._spark_overlay_vt.get():
            args.append("--no-overlay-vt-flavors")
        if not self._spark_plot_source.get():
            args.append("--no-plot-source-current")
        if not self._spark_plot_gate.get():
            args.append("--no-plot-gate-current")
        args += shlex.split(self._spark_extra_args.get().strip())
        return args

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _browse(self):
        d = filedialog.askdirectory(title="Select Output Folder")
        if d:
            self._v_out.set(d)

    def _clear_log(self):
        self._log_txt.delete('1.0', tk.END)

    def _current_server(self) -> str:
        return _EOL_SERVER if self._v_mode.get() == 'EOL' else _INLINE_SERVER

    def _on_test_type_changed(self, _event=None):
        """Server changed (EOL/Inline): reset Process Node/Test Chip choices."""
        self._cmb_proc_node['values'] = _list_subdirs(self._current_server())
        self._v_proc_node.set('')
        self._v_test_chip.set('')
        self._cmb_test_chip['values'] = []
        self._cmb_test_chip.configure(state='disabled')

    def _on_process_node_changed(self, _event=None):
        """Process Node entered/selected: scan it for Test Chip subfolders."""
        proc = self._v_proc_node.get().strip()
        if not proc:
            self._cmb_test_chip.configure(state='disabled')
            self._cmb_test_chip['values'] = []
            self._v_test_chip.set('')
            self._refresh_auto_output_dir()
            return

        base = os.path.join(self._current_server(), proc)
        chips = _list_subdirs(base)
        self._cmb_test_chip['values'] = chips
        self._cmb_test_chip.configure(state='readonly' if chips else 'disabled')
        if self._v_test_chip.get() not in chips:
            self._v_test_chip.set('')
        self._refresh_auto_output_dir()

    def _compute_auto_output_dir(self) -> str:
        lot = self._v_lot.get().strip()
        proc = self._v_proc_node.get().strip()
        chip = self._v_test_chip.get().strip()
        lot_token = re.sub(r'[^A-Za-z0-9_-]+', '_', lot) or 'UnspecifiedLot'
        parts = [str(_DEFAULT_OUTPUT_BASE)]
        if proc:
            parts.append(re.sub(r'[^A-Za-z0-9_-]+', '_', proc))
        if chip:
            parts.append(re.sub(r'[^A-Za-z0-9_-]+', '_', chip))
        parts.append(lot_token)
        return os.path.join(*parts)

    def _refresh_auto_output_dir(self):
        """Recompute the default Output Folder, but only overwrite it if the
        user hasn't since typed something else in that field themselves."""
        auto_val = self._compute_auto_output_dir()
        current = self._v_out.get().strip()
        if not current or current == self._out_auto_value:
            self._v_out.set(auto_val)
        self._out_auto_value = auto_val

    def _on_lookup_from_lot(self):
        """Query PyUber for the current Lot ID(s)/Operation(s) and auto-fill
        Process Node / Test Chip from the result."""
        lot_input = self._v_lot.get().strip()
        ops_raw = self._v_op.get().replace(';', ',')
        operations = [o.strip() for o in ops_raw.split(',') if o.strip()]
        if not lot_input or not operations:
            messagebox.showinfo("Auto-detect", "Enter Lot ID and Operation(s) first.", parent=self.root)
            return
        if not _PYUBER_OK:
            messagebox.showerror("PyUber Not Found", "PyUber is not installed; cannot auto-detect.", parent=self.root)
            return

        def _work():
            try:
                lot_patterns = _parse_lot_patterns(lot_input)
                lot_map = query_lot_info(lot_patterns, operations)
            except Exception as exc:
                self._log(f"Auto-detect failed: {exc}")
                self.root.after(0, lambda: messagebox.showerror("Auto-detect Failed", str(exc), parent=self.root))
                return
            if not lot_map:
                self._log("Auto-detect: no lots found for that Lot ID / Operation(s).")
                self.root.after(0, lambda: messagebox.showwarning(
                    "Auto-detect", "No lots found in the database for that Lot ID / Operation(s).", parent=self.root))
                return
            self.root.after(0, lambda: self._apply_lookup_result(lot_map))

        threading.Thread(target=_work, daemon=True).start()

    def _apply_lookup_result(self, lot_map: dict):
        combos = sorted(set(lot_map.values()))
        for lot, (proc, chip) in sorted(lot_map.items()):
            self._log(f"  Auto-detected: {lot} -> {proc} / {chip}")
        if len(combos) == 1:
            proc, chip = combos[0]
            self._cmb_proc_node['values'] = [proc]
            self._v_proc_node.set(proc)
            self._cmb_test_chip['values'] = [chip]
            self._cmb_test_chip.configure(state='readonly')
            self._v_test_chip.set(chip)
            self._refresh_auto_output_dir()
        else:
            procs = sorted({c[0] for c in combos})
            self._cmb_proc_node['values'] = procs
            messagebox.showinfo(
                "Multiple Matches",
                f"Found {len(combos)} different Process Node / Test Chip combinations across the "
                "matched lots (see Run Log). Pick manually from the Process Node dropdown.",
                parent=self.root,
            )

    def _sync_plot_dirs_from_extractor(self, *_args):
        """Propagate the Extractor's Output Folder into the Plot tabs' Topax
        Dir / Output Dir fields, unless the user has typed something else there."""
        out_val = self._v_out.get().strip()
        if not out_val:
            return
        prev_auto = self._plot_dir_auto_value
        for var in (self._idvg_topax_dir, self._idvg_output_dir,
                    self._spark_topax_dir, self._spark_output_dir):
            cur = var.get().strip()
            if not cur or cur == prev_auto:
                var.set(out_val)
        self._plot_dir_auto_value = out_val

    def _poll_log(self):
        """Drain the log queue into the text widget (≤50 lines per 100 ms tick)."""
        for _ in range(50):
            try:
                line = self._log_q.get_nowait()
            except queue.Empty:
                break
            self._log_txt.insert(tk.END, line + '\n')
            self._log_txt.see(tk.END)
        self.root.after(100, self._poll_log)

    def _log(self, msg: str):
        self._log_q.put(msg)
        lf = self._log_file
        if lf:
            try:
                lf.write(msg + '\n')
                lf.flush()
            except Exception:
                pass

    def _set_ui(self, pct=None, status=None):
        """Thread-safe: schedule a GUI update on the main thread."""
        _pct    = pct
        _status = status

        def _apply():
            if _pct is not None:
                self._v_pct.set(_pct)
            if _status is not None:
                self._v_status.set(_status)

        self.root.after(0, _apply)

    # ── Button handlers ───────────────────────────────────────────────────────

    def _on_run(self):
        errs = []
        if not self._v_lot.get().strip():
            errs.append("• Lot ID is required.")
        if not self._v_op.get().strip():
            errs.append("• Operation(s) is required.")
        if not self._v_out.get().strip():
            errs.append("• Output Folder is required.")
        if errs:
            messagebox.showerror("Missing Input", "\n".join(errs), parent=self.root)
            return

        out_val = self._v_out.get().strip()
        if _is_network_path(out_val):
            corrected = self._compute_auto_output_dir()
            messagebox.showwarning(
                "Local Storage Required",
                f"Output folder is on a network location:\n{out_val}\n\n"
                f"Data must be saved locally. Using instead:\n{corrected}",
                parent=self.root,
            )
            self._v_out.set(corrected)
            self._out_auto_value = corrected

        self._btn_run.configure(state=tk.DISABLED)
        self._btn_stop.configure(state=tk.NORMAL)
        self._stop.clear()
        self._v_pct.set(0.0)
        self._v_status.set("Running…")

        threading.Thread(target=self._worker, daemon=True).start()

    def _on_stop(self):
        self._stop.set()
        self._v_status.set("Stopping…")

    def _on_done(self):
        self._btn_run.configure(state=tk.NORMAL)
        self._btn_stop.configure(state=tk.DISABLED)

    def _on_run_plot_script(self, script_path: Path, button: ttk.Button, build_args):
        """Launch a sibling plotting script as an independent subprocess with
        GUI-supplied CLI args, streaming its stdout/stderr into the Run Log
        and driving the shared progress bar from its 'PROGRESS n/total' lines."""
        if not script_path.exists():
            messagebox.showerror("Script Not Found", f"Could not find:\n{script_path}", parent=self.root)
            return

        try:
            extra_args = build_args()
        except ValueError as exc:
            messagebox.showerror("Invalid Extra CLI Args", str(exc), parent=self.root)
            return

        button.configure(state=tk.DISABLED)
        cmd = [sys.executable, str(script_path)] + extra_args
        self._log(f"\n{'─' * 58}\nRunning: {' '.join(cmd)}\n{'─' * 58}")
        self._set_ui(pct=0.0, status=f"Running {script_path.name}…")

        def _stream():
            try:
                proc = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, bufsize=1, cwd=str(script_path.parent),
                )
                for line in proc.stdout:
                    line = line.rstrip('\n')
                    progress_match = re.match(r'^PROGRESS (\d+)/(\d+)$', line.strip())
                    if progress_match:
                        done, total = int(progress_match.group(1)), int(progress_match.group(2))
                        pct = (done / total * 100) if total else 0
                        self._set_ui(pct=pct, status=f"{script_path.name}: {done} / {total}…")
                        continue
                    self._log(line)
                proc.wait()
                self._log(f"{script_path.name} finished (exit code {proc.returncode}).")
                self._set_ui(pct=100.0, status=f"{script_path.name} finished (exit code {proc.returncode})")
            except Exception as exc:
                self._log(f"ERROR launching {script_path.name}: {exc}")
            finally:
                self.root.after(0, lambda: button.configure(state=tk.NORMAL))

        threading.Thread(target=_stream, daemon=True).start()

    # ── Background worker ─────────────────────────────────────────────────────

    def _worker(self):
        # Open a timestamped log file in the chosen output folder
        out = self._v_out.get().strip()
        ts  = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        log_path = os.path.join(out, f'Topax_Extractor_{ts}.txt') if out else None
        if log_path:
            try:
                os.makedirs(out, exist_ok=True)
                self._log_file = open(log_path, 'w', encoding='utf-8')
            except Exception:
                log_path = None   # don't let a bad path block the run

        try:
            if log_path:
                self._log(f"Log file : {log_path}")
            self._run_extraction()
        except Exception as exc:
            self._log(f"\nFATAL ERROR: {exc}")
            self._set_ui(status="Fatal error — see log")
        finally:
            if self._log_file:
                try:
                    self._log_file.close()
                except Exception:
                    pass
                self._log_file = None
            self.root.after(0, self._on_done)

    def _run_extraction(self):
        lot_input   = self._v_lot.get().strip()
        lot_patterns = _parse_lot_patterns(lot_input)   # list of SQL LIKE patterns
        ops_raw    = self._v_op.get().replace(';', ',')
        operations = [o.strip() for o in ops_raw.split(',') if o.strip()]
        mode       = self._v_mode.get()           # 'EOL' | 'Inline'
        dtype      = self._v_dtype.get()          # 'SML' | 'VBD' | …
        struct_raw = self._v_struct.get().strip()
        struct_pats = [s.strip() for s in re.split(r'[,\n]', struct_raw) if s.strip()]
        if not struct_pats:
            struct_pats = ['*.tgz']

        out_base     = self._v_out.get().strip()
        create_subs  = self._v_subdirs.get()
        server       = _EOL_SERVER if mode == 'EOL' else _INLINE_SERVER
        proc_node    = self._v_proc_node.get().strip()
        test_chip    = self._v_test_chip.get().strip()
        manual_override = bool(proc_node) and bool(test_chip)

        sep = '─' * 58

        self._log(sep)
        self._log(f"Lot pattern(s): {lot_patterns}")
        self._log(f"Operations   : {', '.join(operations)}")
        self._log(f"Mode         : {mode}   Data type: {dtype}")
        self._log(f"Structure(s) : {struct_pats}")
        self._log(f"Output base  : {out_base}")
        if manual_override:
            self._log(f"Process Node / Test Chip (manual): {proc_node} / {test_chip}")
        self._log(sep)

        # ── 1. Resolve lot -> (process, test chip) ───────────────────────────
        if manual_override:
            self._set_ui(pct=5, status="Using manual Process Node / Test Chip…")
            self._log("Step 1 — Manual Process Node / Test Chip override (skipping PyUber lookup)…")
            raw_lots = [t.strip() for t in lot_input.split(',') if t.strip() and '*' not in t and '?' not in t]
            if not raw_lots:
                self._log("  No exact Lot ID(s) given — manual override needs literal lot IDs (no wildcards).")
                self._set_ui(pct=0, status="No exact lot IDs for manual override")
                return
            lot_map = {lot: (proc_node, test_chip) for lot in raw_lots}
        else:
            self._set_ui(pct=5, status="Querying database…")
            self._log("Step 1 — Querying PyUber for lot / process info…")

            try:
                lot_map = query_lot_info(lot_patterns, operations)
            except Exception as exc:
                self._log(f"  ERROR: {exc}")
                self._set_ui(pct=0, status="Database query failed")
                return

            if not lot_map:
                self._log(
                    "  No lots found.  Check Lot ID, Operation, VPN, and "
                    "database access."
                )
                self._set_ui(pct=0, status="No lots found")
                return

        self._log(f"  {len(lot_map)} lot(s) found: {sorted(lot_map.keys())}")

        # ── 1b. Filesystem scan for child lots (e.g. D608EM10H) ──────────────
        # PyUber only returns 8-char lot IDs; child lots like D608EM10H are
        # discovered by looking for sibling directories on the network share.
        self._set_ui(pct=8, status="Checking filesystem for child lots…")
        self._log("  Checking filesystem for child lots…")
        child_lots: dict = {}
        for lot_id, (etest_proc, product) in list(lot_map.items()):
            base = os.path.join(server, etest_proc, product, dtype)
            if not os.path.isdir(base):
                continue
            for candidate in glob.glob(os.path.join(base, lot_id + '?*')):
                if os.path.isdir(candidate):
                    child = os.path.basename(candidate)
                    if child not in lot_map and child not in child_lots:
                        child_lots[child] = (etest_proc, product)
                        self._log(f"    Child lot found on filesystem: {child}")
        if child_lots:
            lot_map.update(child_lots)
            self._log(f"  Added {len(child_lots)} child lot(s).")
        else:
            self._log("  No additional child lots found.")

        # ── 2. Discover all files ─────────────────────────────────────────────
        self._set_ui(pct=10, status="Scanning network share…")
        self._log(f"\nStep 2 — Scanning for files…")

        tasks = []

        for lot_id, (etest_proc, product) in sorted(lot_map.items()):
            if self._stop.is_set():
                break

            base = os.path.join(server, etest_proc, product, dtype)
            self._log(f"\n  Lot {lot_id}  →  ...\\{etest_proc}\\{product}\\{dtype}")
            self._log(f"  Full path: {base}")

            if not os.path.isdir(base):
                self._log(f"    ! Path not accessible — skipping")
                continue

            # Live files
            live = find_live_tgz(base, lot_id, struct_pats)
            self._log(f"    Live files  : {len(live)}")
            for item in live:
                tasks.append(dict(lot_id=lot_id, etest_proc=etest_proc,
                                  src='live', item=item))

            # Archived (zip) files — keyed by exact lot_id from PyUber
            zips = find_zip_archives(base, lot_id)
            self._log(f"    Archives    : {len(zips)}")
            for zp in zips:
                zname = os.path.basename(zp)
                try:
                    zipped = find_tgz_in_zip(zp, struct_pats)
                except Exception as exc:
                    self._log(f"      ! Could not read {zname}: {exc}")
                    continue
                self._log(f"      {zname}: {len(zipped)} matching file(s)")
                for item in zipped:
                    tasks.append(dict(lot_id=lot_id, etest_proc=etest_proc,
                                      src='zip', item=item))

        total = len(tasks)
        if total == 0:
            self._log(
                "\nNo files found.  Check structure pattern, data type, "
                "lot ID, and network access."
            )
            self._set_ui(pct=0, status="No files found")
            return

        self._log(f"\n  Total: {total} file(s) to process.")

        # ── 3. Process / convert files ────────────────────────────────────────
        self._set_ui(pct=15, status=f"Processing 0 / {total}…")
        self._log(f"\nStep 3 — Converting files…")

        ok = err = skipped = 0
        skip_existing = self._v_skip_existing.get()

        for idx, task in enumerate(tasks):
            if self._stop.is_set():
                self._log("  Stopped by user.")
                break

            lot_id     = task['lot_id']
            etest_proc = task['etest_proc']
            lot7       = lot_id[:7]

            dest = (
                os.path.join(out_base, etest_proc, lot7, 'Topax')
                if create_subs else out_base
            )
            os.makedirs(dest, exist_ok=True)

            try:
                if task['src'] == 'live':
                    fpath, wafer, die_xy, tgz_name = task['item']
                    stem     = Path(tgz_name).stem
                    csv_name = f"{lot7}_w{wafer}_Die{die_xy}{stem}.csv"
                    self._log(f"  [{idx+1}/{total}] {tgz_name}")
                    df = read_tgz_file(fpath)

                else:  # zip
                    zip_path, entry, wafer, die_xy, tgz_name = task['item']
                    stem     = Path(tgz_name).stem
                    csv_name = f"{lot7}_w{wafer}_Die{die_xy}{stem}.csv"
                    self._log(f"  [{idx+1}/{total}] {tgz_name}  (archived)")
                    df = read_tgz_from_zip(zip_path, entry)

                out_path = os.path.join(dest, csv_name)
                if skip_existing and os.path.exists(out_path):
                    self._log(f"    skipped (already exists)")
                    skipped += 1
                else:
                    df.to_csv(out_path)
                    ok += 1

            except Exception as exc:
                self._log(f"    ERROR: {exc}")
                err += 1

            pct = 15 + (idx + 1) / total * 85
            self._set_ui(pct=pct, status=f"Processing {idx+1} / {total}…")

        # ── Summary ───────────────────────────────────────────────────────────
        self._log(f"\n{sep}")
        if ok == 0 and err == 0 and skipped == 0:
            self._log("No files were processed (stopped or nothing matched).")
        else:
            skip_note = f",  {skipped} skipped" if skipped else ""
            self._log(f"Done: {ok} file(s) saved,  {err} error(s){skip_note}.")
            self._log(f"Output: {out_base}")

        skip_note    = f", {skipped} skipped" if skipped else ""
        final_status = (
            f"Stopped — {ok} saved, {err} error(s){skip_note}"
            if self._stop.is_set()
            else f"Done — {ok} saved, {err} error(s){skip_note}"
        )
        self._set_ui(pct=100, status=final_status)


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════

def main():
    root = tk.Tk()
    app  = _App(root)

    if not _PYUBER_OK:
        messagebox.showwarning(
            "PyUber Not Found",
            "PyUber could not be imported.\n\n"
            "The GUI will open, but database queries will fail.\n"
            "Ensure you are on the Intel network and PyUber is installed.",
            parent=root,
        )

    root.mainloop()


if __name__ == '__main__':
    main()
