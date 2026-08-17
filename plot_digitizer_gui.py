"""
Plot Digitizer - extract data points from a figure image into a table.

Import a screenshot/scan of a plot from a paper, calibrate the axes by
clicking two known points on each axis, then click the data points (or let
the colour auto-detector find them).  The extracted (x, y) values appear in
a live table at the bottom that you can copy straight into Excel or save as
.csv / .xlsx.

Workflow
    1. "Load image..."  -> pick a PNG/JPG/etc. of the plot.
    2. Calibration (right panel): click "Pick X1", then click a point on the
       x-axis whose value you know, and type that value in the box. Repeat for
       X2, Y1, Y2. Tick "log" if an axis is logarithmic.
    3. Series + points (right panel): pick "Add points" mode and click each
       data point.  Use several named series for multi-curve plots.
       Or "Sample colour" on a marker, then "Auto points" / "Auto trace".
    4. The table (bottom) fills in live.  "Copy table" puts it on the
       clipboard (tab-separated -> paste into Excel); or Export CSV / XLSX.

Layout mirrors vplot_gui.py: control bar on top, a settings row, then the
image canvas above a resizable table pane.

Run:  py plot_digitizer_gui.py     (Python 3.11 - has PIL/numpy/pandas/scipy)
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

import numpy as np
import pandas as pd
from PIL import Image

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk,
)
from matplotlib.colors import to_hex

try:                                   # auto-detect is optional
    from scipy import ndimage
    _HAVE_SCIPY = True
except Exception:                      # pragma: no cover
    _HAVE_SCIPY = False

# colour-blind-safe cycle for new series (Okabe-Ito, matching vplot_common)
SERIES_COLORS = [
    "#0072B2", "#D55E00", "#009E73", "#CC79A7",
    "#E69F00", "#56B4E9", "#F0E442", "#000000",
]

CALIB_KEYS = ("x1", "x2", "y1", "y2")


class DigitizerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Plot digitizer - image to data table")
        self.geometry("1180x1000")

        # --- image state -------------------------------------------------
        self.img_rgb = None            # (H, W, 3) uint8 array, or None
        self.img_path = None

        # --- calibration state ------------------------------------------
        # pixel coords of the four reference clicks (px, py) or None
        self.calib = {k: None for k in CALIB_KEYS}
        self.calib_val = {k: tk.StringVar() for k in CALIB_KEYS}
        for v in self.calib_val.values():
            v.trace_add("write", lambda *_: self._refresh_table())
        self.logx = tk.BooleanVar(value=False)
        self.logy = tk.BooleanVar(value=False)
        self.logx.trace_add("write", lambda *_: self._refresh_table())
        self.logy.trace_add("write", lambda *_: self._refresh_table())

        # --- series (data point) state ----------------------------------
        self.series_order = []                 # list of series names
        self.series_data = {}                  # name -> list of [px, py]
        self.series_color = {}                 # name -> hex colour
        self.active_series = tk.StringVar()

        # --- interaction ------------------------------------------------
        self.mode = None               # x1/x2/y1/y2/add/delete/pick_color
        self.target_color = None       # (r, g, b) for auto-detect
        self._overlay = []             # matplotlib overlay artists

        # --- table options ----------------------------------------------
        self.layout_var = tk.StringVar(value="Wide (columns per series)")
        self.layout_var.trace_add("write", lambda *_: self._refresh_table())

        self._build_topbar()
        self._build_settings_row()
        self._build_canvas_and_table()

        self._add_series()             # start with one series
        self._refresh_table()

    # ====================================================================
    #  UI construction
    # ====================================================================
    def _build_topbar(self):
        bar = ttk.Frame(self, padding=8)
        bar.pack(side=tk.TOP, fill=tk.X)

        ttk.Button(bar, text="Load image...", command=self.load_image).pack(side=tk.LEFT)
        self.img_label = ttk.Label(bar, text="(no image loaded)")
        self.img_label.pack(side=tk.LEFT, padx=(8, 16))

        ttk.Button(bar, text="Copy table", command=self.copy_table).pack(side=tk.LEFT)
        ttk.Button(bar, text="Export CSV...", command=self.export_csv).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(bar, text="Export XLSX...", command=self.export_xlsx).pack(side=tk.LEFT, padx=(6, 0))

        ttk.Label(bar, text="Table layout:").pack(side=tk.LEFT, padx=(16, 4))
        ttk.Combobox(bar, textvariable=self.layout_var, state="readonly", width=26,
                     values=["Wide (columns per series)", "Long (Series/X/Y rows)"]
                     ).pack(side=tk.LEFT)

        self.status = tk.StringVar(value="Load an image to begin.")
        ttk.Label(self, textvariable=self.status, relief=tk.SUNKEN,
                  anchor=tk.W).pack(side=tk.BOTTOM, fill=tk.X)

    def _build_settings_row(self):
        row = ttk.Frame(self, padding=(8, 0))
        row.pack(side=tk.TOP, fill=tk.X)

        left = ttk.Frame(row)
        left.pack(side=tk.LEFT, anchor="n")
        right = ttk.Frame(row)
        right.pack(side=tk.LEFT, anchor="n", fill=tk.X, expand=True, padx=(8, 0))

        self._build_calib_panel(left)
        self._build_series_panel(right)

    def _build_calib_panel(self, parent):
        box = ttk.LabelFrame(parent, text="Axis calibration", padding=8)
        box.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        headers = ("Reference", "Pick", "Known value", "Pixel")
        for c, txt in enumerate(headers):
            ttk.Label(box, text=txt).grid(row=0, column=c, padx=6, sticky="w")

        labels = {"x1": "X point 1", "x2": "X point 2",
                  "y1": "Y point 1", "y2": "Y point 2"}
        self.calib_pix_lbl = {}
        for i, key in enumerate(CALIB_KEYS, start=1):
            ttk.Label(box, text=labels[key]).grid(row=i, column=0, padx=6, pady=2, sticky="w")
            ttk.Button(box, text="Pick", width=6,
                       command=lambda k=key: self._set_mode(k)
                       ).grid(row=i, column=1, padx=6, pady=2)
            ttk.Entry(box, textvariable=self.calib_val[key], width=12
                      ).grid(row=i, column=2, padx=6, pady=2)
            lbl = ttk.Label(box, text="-", width=14)
            lbl.grid(row=i, column=3, padx=6, pady=2, sticky="w")
            self.calib_pix_lbl[key] = lbl

        logrow = ttk.Frame(box)
        logrow.grid(row=5, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Checkbutton(logrow, text="log X axis", variable=self.logx).pack(side=tk.LEFT)
        ttk.Checkbutton(logrow, text="log Y axis", variable=self.logy).pack(side=tk.LEFT, padx=(12, 0))

    def _build_series_panel(self, parent):
        box = ttk.LabelFrame(parent, text="Data series + point picking", padding=8)
        box.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        # series row
        srow = ttk.Frame(box)
        srow.grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 6))
        ttk.Label(srow, text="Active series:").pack(side=tk.LEFT)
        self.series_cb = ttk.Combobox(srow, textvariable=self.active_series,
                                      state="readonly", width=18)
        self.series_cb.pack(side=tk.LEFT, padx=6)
        self.series_cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_overlay())
        ttk.Button(srow, text="Add", width=5, command=self._add_series).pack(side=tk.LEFT)
        ttk.Button(srow, text="Rename", width=7, command=self._rename_series).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(srow, text="Colour", width=7, command=self._pick_series_color).pack(side=tk.LEFT, padx=(4, 0))
        self.series_sw = tk.Button(srow, width=3, relief="raised",
                                   command=self._pick_series_color)
        self.series_sw.pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(srow, text="Delete", width=7, command=self._delete_series).pack(side=tk.LEFT, padx=(8, 0))

        # mode row (what a click does)
        mrow = ttk.Frame(box)
        mrow.grid(row=1, column=0, columnspan=6, sticky="w")
        ttk.Label(mrow, text="Click mode:").pack(side=tk.LEFT)
        self.mode_lbl = ttk.Label(mrow, text="(none - pick a tool)", foreground="#666")
        self.mode_lbl.pack(side=tk.LEFT, padx=6)
        ttk.Button(mrow, text="Add points", command=lambda: self._set_mode("add")
                   ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(mrow, text="Delete points", command=lambda: self._set_mode("delete")
                   ).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(mrow, text="Clear series", command=self._clear_series
                   ).pack(side=tk.LEFT, padx=(12, 0))

        # auto-detect row
        arow = ttk.Frame(box)
        arow.grid(row=2, column=0, columnspan=6, sticky="w", pady=(6, 0))
        ttk.Button(arow, text="Sample colour", command=lambda: self._set_mode("pick_color")
                   ).pack(side=tk.LEFT)
        self.color_sw = tk.Button(arow, width=3, relief="sunken", state="disabled")
        self.color_sw.pack(side=tk.LEFT, padx=(4, 8))
        ttk.Label(arow, text="tol").pack(side=tk.LEFT)
        self.tol_var = tk.StringVar(value="45")
        ttk.Entry(arow, textvariable=self.tol_var, width=5).pack(side=tk.LEFT, padx=(2, 8))
        ttk.Label(arow, text="min size").pack(side=tk.LEFT)
        self.minsize_var = tk.StringVar(value="6")
        ttk.Entry(arow, textvariable=self.minsize_var, width=5).pack(side=tk.LEFT, padx=(2, 8))
        state = "normal" if _HAVE_SCIPY else "disabled"
        ttk.Button(arow, text="Auto points", command=self._auto_points, state=state
                   ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(arow, text="Auto trace", command=self._auto_trace
                   ).pack(side=tk.LEFT, padx=(4, 0))
        if not _HAVE_SCIPY:
            ttk.Label(arow, text="(scipy missing: 'Auto points' off)",
                      foreground="#a00").pack(side=tk.LEFT, padx=(8, 0))

    def _build_canvas_and_table(self):
        pane = ttk.PanedWindow(self, orient=tk.VERTICAL)
        pane.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # --- image canvas -----------------------------------------------
        self.canvas_frame = ttk.Frame(pane)
        pane.add(self.canvas_frame, weight=3)

        self.fig = Figure(figsize=(7, 5), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_axis_off()
        self.ax.text(0.5, 0.5, "Load an image of a plot to begin",
                     ha="center", va="center", transform=self.ax.transAxes,
                     color="#888")
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.canvas_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.canvas_frame)
        self.toolbar.update()
        self.canvas.mpl_connect("button_press_event", self._on_click)

        # --- table ------------------------------------------------------
        tframe = ttk.Frame(pane)
        pane.add(tframe, weight=2)
        self.tree = ttk.Treeview(tframe, show="headings", height=8)
        vsb = ttk.Scrollbar(tframe, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tframe, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tframe.rowconfigure(0, weight=1)
        tframe.columnconfigure(0, weight=1)

    # ====================================================================
    #  image loading
    # ====================================================================
    def load_image(self):
        path = filedialog.askopenfilename(
            title="Select a plot image",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp"),
                       ("All files", "*.*")],
        )
        if not path:
            return
        try:
            img = Image.open(path).convert("RGB")
        except Exception as exc:
            messagebox.showerror("Load failed", str(exc))
            return
        self.img_rgb = np.asarray(img)
        self.img_path = path
        self.img_label.config(text=os.path.basename(path))

        self.ax.clear()
        self.ax.imshow(self.img_rgb)
        self.ax.set_axis_off()
        self._overlay = []
        self._refresh_overlay()
        self.status.set("Image loaded. Calibrate the axes, then pick data points.")

    # ====================================================================
    #  mode + click dispatch
    # ====================================================================
    def _set_mode(self, mode):
        if self.img_rgb is None:
            messagebox.showinfo("No image", "Load an image first.")
            return
        if self.toolbar.mode:                 # zoom/pan active -> turn it off
            messagebox.showinfo(
                "Toolbar active",
                "Turn off the zoom/pan tool (in the toolbar) before picking points.")
            return
        self.mode = mode
        pretty = {"x1": "click X reference point 1",
                  "x2": "click X reference point 2",
                  "y1": "click Y reference point 1",
                  "y2": "click Y reference point 2",
                  "add": "click to ADD data points",
                  "delete": "click near a point to DELETE it",
                  "pick_color": "click a marker to sample its colour"}
        self.mode_lbl.config(text=pretty.get(mode, mode))
        self.status.set(f"Mode: {pretty.get(mode, mode)}")

    def _on_click(self, event):
        if self.img_rgb is None or event.inaxes != self.ax:
            return
        if event.button != 1 or event.xdata is None or self.toolbar.mode:
            return
        px, py = float(event.xdata), float(event.ydata)

        if self.mode in CALIB_KEYS:
            self.calib[self.mode] = (px, py)
            self.calib_pix_lbl[self.mode].config(text=f"({px:.0f}, {py:.0f})")
            self.status.set(f"{self.mode.upper()} set. Enter its known value if not done.")
            self._refresh_overlay()
            self._refresh_table()
        elif self.mode == "add":
            name = self.active_series.get()
            if not name:
                return
            self.series_data[name].append([px, py])
            self._refresh_overlay()
            self._refresh_table()
        elif self.mode == "delete":
            self._delete_nearest(px, py)
        elif self.mode == "pick_color":
            r, c = int(round(py)), int(round(px))
            h, w = self.img_rgb.shape[:2]
            if 0 <= r < h and 0 <= c < w:
                self.target_color = tuple(int(v) for v in self.img_rgb[r, c])
                hexc = to_hex([v / 255 for v in self.target_color])
                self.color_sw.config(bg=hexc, activebackground=hexc, state="normal")
                self.status.set(f"Sampled colour {hexc}. Now use Auto points / Auto trace.")

    def _delete_nearest(self, px, py, radius=25):
        best = None
        for name in self.series_order:
            for i, (qx, qy) in enumerate(self.series_data[name]):
                d = (qx - px) ** 2 + (qy - py) ** 2
                if best is None or d < best[0]:
                    best = (d, name, i)
        if best and best[0] <= radius ** 2:
            _, name, i = best
            self.series_data[name].pop(i)
            self._refresh_overlay()
            self._refresh_table()
            self.status.set(f"Deleted a point from '{name}'.")

    # ====================================================================
    #  series management
    # ====================================================================
    def _add_series(self):
        n = len(self.series_order) + 1
        name = f"Series {n}"
        while name in self.series_data:
            n += 1
            name = f"Series {n}"
        self.series_order.append(name)
        self.series_data[name] = []
        self.series_color[name] = SERIES_COLORS[(n - 1) % len(SERIES_COLORS)]
        self.active_series.set(name)
        self._sync_series_cb()
        self._update_series_swatch()
        self._refresh_table()

    def _rename_series(self):
        old = self.active_series.get()
        if not old:
            return
        new = _ask_string(self, "Rename series", "New name:", old)
        if not new or new == old:
            return
        if new in self.series_data:
            messagebox.showerror("Rename", f"'{new}' already exists.")
            return
        idx = self.series_order.index(old)
        self.series_order[idx] = new
        self.series_data[new] = self.series_data.pop(old)
        self.series_color[new] = self.series_color.pop(old)
        self.active_series.set(new)
        self._sync_series_cb()
        self._refresh_overlay()
        self._refresh_table()

    def _delete_series(self):
        name = self.active_series.get()
        if not name:
            return
        if len(self.series_order) == 1:
            messagebox.showinfo("Delete series", "Keep at least one series.")
            return
        if self.series_data[name] and not messagebox.askyesno(
                "Delete series", f"Delete '{name}' and its {len(self.series_data[name])} points?"):
            return
        self.series_order.remove(name)
        self.series_data.pop(name)
        self.series_color.pop(name)
        self.active_series.set(self.series_order[0])
        self._sync_series_cb()
        self._update_series_swatch()
        self._refresh_overlay()
        self._refresh_table()

    def _clear_series(self):
        name = self.active_series.get()
        if name and self.series_data[name]:
            if messagebox.askyesno("Clear series",
                                   f"Remove all {len(self.series_data[name])} points from '{name}'?"):
                self.series_data[name].clear()
                self._refresh_overlay()
                self._refresh_table()

    def _pick_series_color(self):
        name = self.active_series.get()
        if not name:
            return
        _, hx = colorchooser.askcolor(color=self.series_color[name],
                                      title=f"Colour for {name}")
        if hx:
            self.series_color[name] = hx
            self._update_series_swatch()
            self._refresh_overlay()

    def _sync_series_cb(self):
        self.series_cb["values"] = self.series_order

    def _update_series_swatch(self):
        name = self.active_series.get()
        if name:
            hx = self.series_color[name]
            self.series_sw.config(bg=hx, activebackground=hx)

    # ====================================================================
    #  auto-detect
    # ====================================================================
    def _color_mask(self):
        if self.target_color is None:
            messagebox.showinfo("No colour", "Use 'Sample colour' on a marker first.")
            return None
        try:
            tol = float(self.tol_var.get())
        except ValueError:
            tol = 45.0
        diff = self.img_rgb.astype(np.int16) - np.array(self.target_color, dtype=np.int16)
        dist = np.sqrt((diff ** 2).sum(axis=2))
        return dist <= tol

    def _auto_points(self):
        if not _HAVE_SCIPY:
            return
        mask = self._color_mask()
        if mask is None:
            return
        try:
            min_size = float(self.minsize_var.get())
        except ValueError:
            min_size = 6.0
        structure = np.ones((3, 3), dtype=bool)     # 8-connectivity
        lbl, n = ndimage.label(mask, structure=structure)
        if n == 0:
            messagebox.showinfo("Auto points", "No pixels matched that colour/tolerance.")
            return
        idx = range(1, n + 1)
        sizes = ndimage.sum(mask, lbl, idx)
        coms = ndimage.center_of_mass(mask, lbl, idx)
        name = self.active_series.get()
        added = 0
        for size, (py, px) in zip(np.atleast_1d(sizes), np.atleast_2d(coms)):
            if size >= min_size:
                self.series_data[name].append([float(px), float(py)])
                added += 1
        self._refresh_overlay()
        self._refresh_table()
        self.status.set(f"Auto points: added {added} blob centroids to '{name}'.")

    def _auto_trace(self):
        """One point per image column that contains the target colour."""
        mask = self._color_mask()
        if mask is None:
            return
        name = self.active_series.get()
        cols = np.where(mask.any(axis=0))[0]
        if cols.size == 0:
            messagebox.showinfo("Auto trace", "No pixels matched that colour/tolerance.")
            return
        step = max(1, int(cols.size / 300))          # cap at ~300 points
        added = 0
        for px in cols[::step]:
            rows = np.where(mask[:, px])[0]
            py = float(rows.mean())
            self.series_data[name].append([float(px), py])
            added += 1
        self._refresh_overlay()
        self._refresh_table()
        self.status.set(f"Auto trace: added {added} points to '{name}'.")

    # ====================================================================
    #  coordinate mapping
    # ====================================================================
    def _calib_ready(self, axis):
        """axis is 'x' or 'y'. True if both refs picked and values valid."""
        keys = ("x1", "x2") if axis == "x" else ("y1", "y2")
        for k in keys:
            if self.calib[k] is None:
                return False
            try:
                float(self.calib_val[k].get())
            except (ValueError, tk.TclError):
                return False
        log = self.logx.get() if axis == "x" else self.logy.get()
        if log and any(float(self.calib_val[k].get()) <= 0 for k in keys):
            return False
        # need distinct pixels on the mapped component
        comp = 0 if axis == "x" else 1
        p1 = self.calib[keys[0]][comp]
        p2 = self.calib[keys[1]][comp]
        return p1 != p2

    def _map(self, pix, axis):
        keys = ("x1", "x2") if axis == "x" else ("y1", "y2")
        comp = 0 if axis == "x" else 1
        log = self.logx.get() if axis == "x" else self.logy.get()
        p1 = self.calib[keys[0]][comp]
        p2 = self.calib[keys[1]][comp]
        v1 = float(self.calib_val[keys[0]].get())
        v2 = float(self.calib_val[keys[1]].get())
        if log:
            v1, v2 = np.log10(v1), np.log10(v2)
        frac = (pix - p1) / (p2 - p1)
        val = v1 + frac * (v2 - v1)
        return float(10 ** val) if log else float(val)

    def _series_xy(self, name):
        """Return list of (x, y) in data units (or pixel units if uncalibrated)."""
        xok, yok = self._calib_ready("x"), self._calib_ready("y")
        out = []
        for px, py in self.series_data[name]:
            x = self._map(px, "x") if xok else px
            y = self._map(py, "y") if yok else py
            out.append((x, y))
        return out

    # ====================================================================
    #  overlay drawing
    # ====================================================================
    def _refresh_overlay(self):
        if self.img_rgb is None:
            return
        for art in self._overlay:
            try:
                art.remove()
            except Exception:
                pass
        self._overlay = []

        # calibration markers
        cal_style = {"x1": ("X1", "#d62728"), "x2": ("X2", "#d62728"),
                     "y1": ("Y1", "#1f77b4"), "y2": ("Y2", "#1f77b4")}
        for key, (tag, col) in cal_style.items():
            if self.calib[key] is not None:
                px, py = self.calib[key]
                m, = self.ax.plot(px, py, marker="+", ms=14, mew=2, color=col)
                t = self.ax.annotate(tag, (px, py), textcoords="offset points",
                                     xytext=(6, 6), color=col, fontsize=9, fontweight="bold")
                self._overlay += [m, t]

        # data points, active series drawn last/larger
        active = self.active_series.get()
        for name in self.series_order:
            pts = self.series_data[name]
            if not pts:
                continue
            arr = np.array(pts)
            is_active = (name == active)
            sc = self.ax.scatter(arr[:, 0], arr[:, 1],
                                 s=45 if is_active else 25,
                                 facecolors="none" if not is_active else self.series_color[name],
                                 edgecolors=self.series_color[name],
                                 linewidths=1.6 if is_active else 1.0, zorder=5)
            self._overlay.append(sc)

        self.canvas.draw_idle()

    # ====================================================================
    #  table
    # ====================================================================
    def _build_dataframe(self):
        wide = self.layout_var.get().startswith("Wide")
        xok, yok = self._calib_ready("x"), self._calib_ready("y")
        xs = "X" if xok else "X(px)"
        ys = "Y" if yok else "Y(px)"

        if wide:
            data = {}
            maxlen = 0
            for name in self.series_order:
                xy = self._series_xy(name)
                data[f"{name} {xs}"] = [round(x, 6) for x, _ in xy]
                data[f"{name} {ys}"] = [round(y, 6) for _, y in xy]
                maxlen = max(maxlen, len(xy))
            for k in data:                       # pad ragged columns
                data[k] += [""] * (maxlen - len(data[k]))
            return pd.DataFrame(data)
        else:
            rows = []
            for name in self.series_order:
                for i, (x, y) in enumerate(self._series_xy(name), start=1):
                    rows.append([name, i, round(x, 6), round(y, 6)])
            return pd.DataFrame(rows, columns=["Series", "Point", xs, ys])

    def _refresh_table(self):
        df = self._build_dataframe()
        cols = [f"c{i}" for i in range(len(df.columns))]
        self.tree["columns"] = cols
        for cid, name in zip(cols, df.columns):
            self.tree.heading(cid, text=str(name))
            self.tree.column(cid, width=120, anchor="center", stretch=True)
        self.tree.delete(*self.tree.get_children())
        for _, r in df.iterrows():
            self.tree.insert("", "end", values=list(r))
        self._df = df                            # cache for copy/export

    # ====================================================================
    #  copy / export
    # ====================================================================
    def _current_df(self):
        df = getattr(self, "_df", None)
        if df is None or df.empty:
            messagebox.showinfo("No data", "No data points to export yet.")
            return None
        return df

    def copy_table(self):
        df = self._current_df()
        if df is None:
            return
        self.clipboard_clear()
        self.clipboard_append(df.to_csv(sep="\t", index=False))
        self.status.set("Table copied to clipboard - paste into Excel (Ctrl+V).")

    def export_csv(self):
        df = self._current_df()
        if df is None:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV", "*.csv")],
            initialfile=self._default_name("csv"))
        if path:
            df.to_csv(path, index=False)
            self.status.set(f"Saved -> {os.path.basename(path)}")

    def export_xlsx(self):
        df = self._current_df()
        if df is None:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialfile=self._default_name("xlsx"))
        if path:
            df.to_excel(path, index=False)
            self.status.set(f"Saved -> {os.path.basename(path)}")

    def _default_name(self, ext):
        base = os.path.splitext(os.path.basename(self.img_path))[0] if self.img_path else "digitized"
        return f"{base}_data.{ext}"


def _ask_string(parent, title, prompt, initial=""):
    """Small modal text prompt (avoids importing simpledialog styling issues)."""
    from tkinter import simpledialog
    return simpledialog.askstring(title, prompt, initialvalue=initial, parent=parent)


if __name__ == "__main__":
    DigitizerApp().mainloop()
