"""
Tkinter GUI wrapping the vanadium figures (Cij tri-plot, moduli dual-plot).

    - pick which figure to build
    - choose the source .xlsx (defaults to VCIJplotdata.xlsx)
    - set the legend name, colour AND marker shape of each series (CK / FS / poly):
        * an editable legend-name box per series
        * a preset colour-blind-safe palette dropdown, or
        * a per-series colour wheel + hex box (accepts hex or matplotlib names)
        * a per-series marker-shape dropdown
    - render it in an embedded matplotlib canvas with the navigation toolbar
    - save the current figure to PNG / PDF / SVG

Series colours/markers are pushed into vplot_common (vc.COL / vc.MARK) right
before each render, so the plot modules pick them up with no other changes.

Run:  python vplot_gui.py     (Python 3.11 - has pandas/matplotlib)
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk,
)
from matplotlib.colors import is_color_like, to_hex

import vplot_common as vc
import plot_cij_tri
import plot_moduli_dual
import plot_single
import plot_vv0

FIGURES = {
    "Cij tri-plot (C11 / C12 / C44) - CK vs FS": plot_cij_tri,
    "Moduli dual (K / GH) - CK vs FS + poly":    plot_moduli_dual,
    "Single element (pick one) - CK vs FS + poly": plot_single,
    "V/V0 vs P - data + Ding et al. 2007 (BM3)":   plot_vv0,
}

# core series key -> friendly label shown in the style panel.  Comparison
# sources (Katahara, Antonangeli, ...) are auto-discovered from the workbook
# and appended at runtime by _series_list(), so no source is hard-coded here.
CORE_SERIES = [
    ("CK",   "CK (Cook)"),
    ("FS",   "FS (finite strain)"),
    ("poly", "poly (Kpoly / Gpoly)"),
    ("vv0",  "V/V0 data (this study)"),
    ("ding", "Ding 2007 (BM3 curve)"),
]

# friendly marker name -> matplotlib marker code
MARKERS = {
    "Circle": "o", "Square": "s", "Triangle up": "^", "Triangle down": "v",
    "Diamond": "D", "Pentagon": "p", "Star": "*", "Plus": "P",
    "X": "X", "Hexagon": "h",
}

# friendly line name -> matplotlib line-style code.  Choosing one of these
# draws the series as a connecting line (no point markers) instead of scattered
# points - e.g. render the Katahara (Kat) curve as a solid or dashed line.
LINE_STYLES = {
    "Solid line": "-", "Dashed line": "--",
    "Dotted line": ":", "Dash-dot line": "-.",
}

# combined choices shown in the per-series style dropdown (markers first, then
# lines).  STYLE_BY_CODE maps a stored code back to its friendly name.
STYLES = {**MARKERS, **LINE_STYLES}
MARKER_BY_CODE = {v: k for k, v in STYLES.items()}

# how a line-drawn series shows its uncertainty (global; pushed into
# vc.LINE_UNC before each render).  friendly name -> stored code.
LINE_UNC = {
    "None": "none",
    "Error bars": "bars",
    "Shaded band": "band",
}
LINE_UNC_BY_CODE = {v: k for k, v in LINE_UNC.items()}


class VPlotApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Vanadium elastic-property plotter")
        self.geometry("1480x1000")
        self.current_fig = None
        self.canvas = None
        self.toolbar = None

        # Route the window's X through the same hard shutdown as the Force Quit
        # button, so closing the window can never leave a half-dead process
        # (open figures / a live event loop) lingering in the background and
        # blocking the next launch.
        self.protocol("WM_DELETE_WINDOW", self._force_quit)

        # Discover comparison sources from the default workbook up front, so
        # their style rows are present before the first render.  Safe if the
        # file is missing - the rows just appear after the first successful plot.
        self._discover(vc.DEFAULT_XLSX)

        # live per-series style, seeded from the module defaults (now including
        # any auto-discovered sources)
        self.colors = dict(vc.COL)
        self.markers = dict(vc.MARK)
        self.names = dict(vc.LABEL)   # editable legend tags (CK / FS / PC)
        # per-series draw order (higher = on top); seeded from the module
        # default and kept across style-row rebuilds, like colours/markers.
        self.zorders = dict(vc.ZORDER)
        # error-bar visibility toggles (all-vertical / all-horizontal); these
        # gate the MARKER series only.  Line series use line_unc_var below.
        self.show_yerr = tk.BooleanVar(value=vc.SHOW_YERR)
        self.show_xerr = tk.BooleanVar(value=vc.SHOW_XERR)
        # how a line-drawn series shows its uncertainty (None / Error bars /
        # Shaded band), independent of the marker error-bar toggles.
        self.line_unc_var = tk.StringVar(
            value=LINE_UNC_BY_CODE.get(vc.LINE_UNC, "Shaded band"))
        # output resolution + figure size (cm); size blank => module default
        self.dpi_var = tk.StringVar(value=str(vc.DPI))
        self.figw_var = tk.StringVar(value="")
        self.figh_var = tk.StringVar(value="")
        # on-screen display zoom: scales ONLY how large the figure is drawn in
        # the right-hand panel (via the canvas screen dpi), never the figure's
        # true size in inches or the exported file.  2x so the small PRB-sized
        # figures are legible out of the box.
        self.zoom_var = tk.StringVar(value="2.0")
        # font-size multipliers (1.0 = each plot module's default sizes)
        self.legend_fs_var = tk.StringVar(value="1.0")
        self.axis_fs_var = tk.StringVar(value="1.0")
        # PRB size preset selector (fills the Width/Height cm fields on change)
        self.preset_var = tk.StringVar(value="(custom)")
        # legend placement: X/Y anchor (axes fraction, blank => auto/movable)
        # + column count ("auto" => each plot's own default).  Drag the legend
        # on the canvas to fill X/Y, or type them directly.
        self.leg_x_var = tk.StringVar(value="")
        self.leg_y_var = tk.StringVar(value="")
        self.leg_ncol_var = tk.StringVar(value="auto")
        self._dragging_legend = False   # True while a legend drag is in flight
        # widget handles filled in by _build_style_panel
        self.hexvars = {}
        self.swatches = {}
        self.markvars = {}
        self.zordervars = {}     # per-series "Layer" (z-order) entry vars
        self.namevars = {}
        self.showvars = {}       # per-series "plot this dataset?" checkboxes
        self.show_state = {}     # remembered show/hide across style-row rebuilds
        self._built_series_keys = []   # series keys the style rows were built for
        # axis-bounds state (see _build_bounds_panel)
        self.bound_vars = {}     # "x" / panel key -> (lo StringVar, hi StringVar)
        self.bound_store = {}    # key -> [lo str, hi str], kept across figure switches
        # editable per-panel y-axis label state
        self.ylabel_vars = {}    # panel key -> label StringVar
        self.ylabel_store = {}   # key -> label str, kept across figure switches

        self._build_controls()
        self._build_main_split()

    # ------------------------------------------------------------------
    # main split: all control panels stacked on the LEFT half, the figure
    # centred on the RIGHT half (grey space above/below is intentional).
    # The two always-on panels (series style, axis bounds) stay expanded;
    # the other three (error bars, legend, figure size/DPI) live behind
    # collapsible "dropdown" headers to keep the left column short.
    # ------------------------------------------------------------------
    def _build_main_split(self):
        main = ttk.Frame(self)
        main.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # left half: the controls, stacked vertically (natural width)
        left = ttk.Frame(main, padding=(8, 4))
        left.pack(side=tk.LEFT, fill=tk.Y, anchor="n")

        # right half: the figure canvas, centred with grey margins
        right = ttk.Frame(main)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas_frame = right

        # always-visible panels
        self._build_style_panel(left)
        self._build_bounds_panel(left)
        # collapsible "dropdown" panels (start collapsed)
        self._build_display_panel(
            self._make_collapsible(left, "Uncertainty (error bars & bands)"))
        self._build_legend_panel(self._make_collapsible(left, "Legend"))
        self._build_fonts_panel(
            self._make_collapsible(left, "Fonts (legend & axis text)"))
        self._build_output_panel(
            self._make_collapsible(left, "Figure size, DPI & screen zoom"))

    # ------------------------------------------------------------------
    # collapsible section: a full-width header button that shows/hides its
    # body.  Returns the body frame for the caller to fill.
    # ------------------------------------------------------------------
    def _make_collapsible(self, parent, title, expanded=False):
        outer = ttk.Frame(parent)
        outer.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        state = {"open": expanded}
        hdr_var = tk.StringVar()
        body = ttk.Frame(outer)

        def refresh():
            hdr_var.set(("▼  " if state["open"] else "▶  ") + title)
            if state["open"]:
                body.pack(side=tk.TOP, fill=tk.X)
            else:
                body.forget()

        def toggle():
            state["open"] = not state["open"]
            refresh()

        ttk.Button(outer, textvariable=hdr_var, command=toggle,
                   style="Toolbutton").pack(side=tk.TOP, fill=tk.X)
        refresh()
        return body

    # ------------------------------------------------------------------
    def _build_controls(self):
        bar = ttk.Frame(self, padding=8)
        bar.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(bar, text="Figure:").pack(side=tk.LEFT)
        self.fig_choice = tk.StringVar(value=list(FIGURES)[0])
        fig_cb = ttk.Combobox(
            bar, textvariable=self.fig_choice, values=list(FIGURES),
            state="readonly", width=42,
        )
        fig_cb.pack(side=tk.LEFT, padx=(4, 12))
        fig_cb.bind("<<ComboboxSelected>>", self._on_figure_change)

        # single-element selector (only affects the "Single element" figure)
        ttk.Label(bar, text="Element:").pack(side=tk.LEFT)
        self.single_var = tk.StringVar(value=vc.SINGLE)
        el_cb = ttk.Combobox(
            bar, textvariable=self.single_var, values=list(vc.QUANTITIES),
            state="readonly", width=11,
        )
        el_cb.pack(side=tk.LEFT, padx=(4, 12))
        el_cb.bind("<<ComboboxSelected>>", self._on_single_change)

        ttk.Button(bar, text="Plot", command=self.render).pack(side=tk.LEFT)
        ttk.Button(bar, text="Save PNG...", command=self.save).pack(
            side=tk.LEFT, padx=(6, 0))

        # Hard shutdown button - guarantees the whole process dies (matplotlib
        # canvases, the Tk event loop, any in-flight render) instead of the
        # partial teardown the window's X can leave running in the background.
        # Sits at the far right, coloured so it's unmistakable.
        tk.Button(bar, text="Force Quit", command=self._force_quit,
                  bg="#c0392b", fg="white",
                  activebackground="#e74c3c", activeforeground="white").pack(
            side=tk.RIGHT)

        # data-source row
        srcbar = ttk.Frame(self, padding=(8, 0, 8, 8))
        srcbar.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(srcbar, text="Data file:").pack(side=tk.LEFT)
        self.src_var = tk.StringVar(value=vc.DEFAULT_XLSX)
        ttk.Entry(srcbar, textvariable=self.src_var).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        ttk.Button(srcbar, text="Browse...", command=self._browse).pack(side=tk.LEFT)

        self.status = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status, relief=tk.SUNKEN,
                  anchor=tk.W).pack(side=tk.BOTTOM, fill=tk.X)

    @staticmethod
    def _discover(path):
        """Populate vc.LIT_SERIES from a workbook, ignoring any read error."""
        try:
            vc.load_all_data(path)
        except Exception:
            pass

    def _series_list(self):
        """Core series + every auto-discovered comparison source (in order)."""
        rows = list(CORE_SERIES)
        for k in vc.LIT_SERIES:
            rows.append((k, vc.LABEL.get(k, k)))   # friendly label = legend name
        return rows

    def _build_style_panel(self, parent):
        box = ttk.LabelFrame(parent, text="Series style (colour + marker)", padding=8)
        box.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        # preset palette row (static - lives above the re-buildable rows frame)
        prow = ttk.Frame(box)
        prow.pack(side=tk.TOP, anchor="w", pady=(0, 6))
        ttk.Label(prow, text="Palette preset:").pack(side=tk.LEFT)
        self.palette_var = tk.StringVar(value="Okabe-Ito (CB-safe)")
        pcb = ttk.Combobox(prow, textvariable=self.palette_var,
                           values=list(vc.PALETTES), state="readonly", width=24)
        pcb.pack(side=tk.LEFT, padx=6)
        pcb.bind("<<ComboboxSelected>>", self._apply_palette)

        # rows (one per series) live in their own frame so new sources can be
        # added without rebuilding the whole panel.
        self.style_rows_frame = ttk.Frame(box)
        self.style_rows_frame.pack(side=tk.TOP, fill=tk.X)
        self._build_style_rows()

    def _build_style_rows(self):
        """(Re)build the per-series style rows for the current source list.

        Called on start-up and again whenever discovery turns up a new source
        (e.g. after switching to a workbook with an extra tag).  Existing style
        choices are preserved via self.colors / markers / names / show_state.
        """
        frame = self.style_rows_frame

        # remember the current show/hide state before the old widgets die
        for k, v in self.showvars.items():
            self.show_state[k] = v.get()
        for w in frame.winfo_children():
            w.destroy()
        self.hexvars, self.swatches = {}, {}
        self.markvars, self.namevars, self.showvars = {}, {}, {}
        self.zordervars = {}

        # column headers
        for c, txt in enumerate(
                ("Show", "Series", "Legend name", "Colour", "Hex / name",
                 "Marker / line", "Layer")):
            ttk.Label(frame, text=txt).grid(row=0, column=c, padx=6, sticky="w")

        # one row per series (core + auto-discovered sources)
        for i, (key, label) in enumerate(self._series_list(), start=1):
            # seed style state for a source we haven't shown before
            self.colors.setdefault(key, vc.COL.get(key, "#000000"))
            self.markers.setdefault(key, vc.MARK.get(key, "o"))
            self.names.setdefault(key, vc.LABEL.get(key, key))
            self.zorders.setdefault(key, vc.ZORDER.get(key, vc.DEFAULT_ZORDER))

            # visibility checkbox - unchecked drops the whole dataset from the plot
            sv = tk.BooleanVar(value=self.show_state.get(key, vc.visible(key)))
            scb = ttk.Checkbutton(frame, variable=sv, command=self._auto_render)
            scb.grid(row=i, column=0, padx=6, pady=2)
            self.showvars[key] = sv

            ttk.Label(frame, text=label).grid(row=i, column=1, padx=6, pady=2, sticky="w")

            # editable legend tag (the moduli panels prefix it: "K "+tag, etc.)
            nv = tk.StringVar(value=self.names[key])
            nent = ttk.Entry(frame, textvariable=nv, width=12)
            nent.grid(row=i, column=2, padx=6, pady=2, sticky="w")
            nent.bind("<Return>",   lambda e, k=key: self._apply_name(k))
            nent.bind("<FocusOut>", lambda e, k=key: self._apply_name(k))
            self.namevars[key] = nv

            sw = tk.Button(frame, width=3, relief="raised",
                           command=lambda k=key: self._pick_color(k))
            sw.grid(row=i, column=3, padx=6, pady=2)
            self.swatches[key] = sw

            hv = tk.StringVar(value=self.colors[key])
            ent = ttk.Entry(frame, textvariable=hv, width=12)
            ent.grid(row=i, column=4, padx=6, pady=2, sticky="w")
            ent.bind("<Return>",   lambda e, k=key: self._apply_hex(k))
            ent.bind("<FocusOut>", lambda e, k=key: self._apply_hex(k))
            self.hexvars[key] = hv

            mv = tk.StringVar(value=MARKER_BY_CODE.get(self.markers[key], "Circle"))
            mcb = ttk.Combobox(frame, textvariable=mv, values=list(STYLES),
                               state="readonly", width=14)
            mcb.grid(row=i, column=5, padx=6, pady=2, sticky="w")
            mcb.bind("<<ComboboxSelected>>", lambda e, k=key: self._on_marker(k))
            self.markvars[key] = mv

            # Layer (z-order): higher draws on top.  Type a value or use the
            # arrows; larger numbers lift the series above the rest.
            zv = tk.StringVar(value=str(self.zorders[key]))
            zsb = ttk.Spinbox(frame, from_=0, to=999, width=5, textvariable=zv,
                              command=lambda k=key: self._apply_zorder(k))
            zsb.grid(row=i, column=6, padx=6, pady=2, sticky="w")
            zsb.bind("<Return>",   lambda e, k=key: self._apply_zorder(k))
            zsb.bind("<FocusOut>", lambda e, k=key: self._apply_zorder(k))
            self.zordervars[key] = zv

            self._set_color(key, self.colors[key])   # paint the swatch

        self._built_series_keys = [k for k, _ in self._series_list()]

    def _sync_style_rows(self):
        """Rebuild the style rows if discovery has turned up new sources."""
        if [k for k, _ in self._series_list()] != self._built_series_keys:
            self._build_style_rows()

    # ------------------------------------------------------------------
    # display-options panel (error-bar visibility)
    # ------------------------------------------------------------------
    def _build_display_panel(self, parent):
        box = ttk.Frame(parent, padding=8)
        box.pack(side=tk.TOP, fill=tk.X)

        # marker (scatter) series error bars
        mrow = ttk.Frame(box)
        mrow.pack(side=tk.TOP, fill=tk.X)
        ttk.Checkbutton(mrow, text="Show vertical (y) uncertainties",
                        variable=self.show_yerr,
                        command=self._auto_render).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Checkbutton(mrow, text="Show horizontal (x) uncertainties",
                        variable=self.show_xerr,
                        command=self._auto_render).pack(side=tk.LEFT)
        ttk.Label(box, text="(the two boxes above apply to point/marker series)"
                  ).pack(side=tk.TOP, anchor="w", pady=(2, 0))

        # line-drawn series (e.g. a finite-strain fit) uncertainty style
        lrow = ttk.Frame(box)
        lrow.pack(side=tk.TOP, fill=tk.X, pady=(6, 0))
        ttk.Label(lrow, text="Line series uncertainty:").pack(side=tk.LEFT)
        lcb = ttk.Combobox(lrow, textvariable=self.line_unc_var,
                           values=list(LINE_UNC), state="readonly", width=14)
        lcb.pack(side=tk.LEFT, padx=6)
        lcb.bind("<<ComboboxSelected>>", lambda e: self._auto_render())

    # ------------------------------------------------------------------
    # legend panel: position (drag or type) + column count
    # ------------------------------------------------------------------
    def _build_legend_panel(self, parent):
        box = ttk.Frame(parent, padding=8)
        box.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(box, text="X (0-1):").grid(row=0, column=0, padx=6, pady=2, sticky="w")
        xe = ttk.Entry(box, textvariable=self.leg_x_var, width=8)
        xe.grid(row=0, column=1, padx=6, pady=2, sticky="w")

        ttk.Label(box, text="Y (0-1):").grid(row=0, column=2, padx=6, pady=2, sticky="w")
        ye = ttk.Entry(box, textvariable=self.leg_y_var, width=8)
        ye.grid(row=0, column=3, padx=6, pady=2, sticky="w")

        ttk.Label(box, text="Columns:").grid(row=0, column=4, padx=6, pady=2, sticky="w")
        ncb = ttk.Combobox(box, textvariable=self.leg_ncol_var, width=6,
                           state="readonly",
                           values=["auto", "1", "2", "3", "4", "5", "6"])
        ncb.grid(row=0, column=5, padx=6, pady=2, sticky="w")
        ncb.bind("<<ComboboxSelected>>", lambda e: self._auto_render())

        ttk.Button(box, text="Reset position (auto)",
                   command=self._reset_legend).grid(
                       row=1, column=0, columnspan=2, padx=6, pady=(4, 0), sticky="w")
        ttk.Label(box, text="Blank X/Y = auto; or drag the legend on the plot"
                  ).grid(row=1, column=2, columnspan=4, padx=6, pady=(4, 0), sticky="w")

        for ent in (xe, ye):
            ent.bind("<Return>",   lambda e: self._auto_render())
            ent.bind("<FocusOut>", lambda e: self._auto_render())

    def _reset_legend(self):
        """Clear the pinned X/Y so the legend goes back to auto placement."""
        self.leg_x_var.set("")
        self.leg_y_var.set("")
        vc.LEGEND_XY = None
        self._auto_render()

    def _collect_legend(self):
        """Push the legend position + column count into vplot_common."""
        x = self._to_float(self.leg_x_var.get())
        y = self._to_float(self.leg_y_var.get())
        vc.LEGEND_XY = (x, y) if (x is not None and y is not None) else None
        val = self.leg_ncol_var.get().strip().lower()
        if val in ("", "auto"):
            vc.LEGEND_NCOL = None
        else:
            try:
                vc.LEGEND_NCOL = max(1, int(float(val)))
            except ValueError:
                vc.LEGEND_NCOL = None

    # ------------------------------------------------------------------
    # legend drag: locate the on-canvas legend and read its position back
    # into the X/Y boxes when the user drops it.
    # ------------------------------------------------------------------
    def _find_legend(self):
        """Return (legend, axes) for the current figure's legend, or None."""
        if self.current_fig is None:
            return None
        for ax in self.current_fig.axes:
            leg = ax.get_legend()
            if leg is not None:
                return leg, ax
        return None

    def _legend_bbox(self, leg):
        """Display-coordinate bounding box of the legend, or None."""
        try:
            return leg.get_window_extent(self.canvas.get_renderer())
        except Exception:
            try:
                return leg.get_window_extent()
            except Exception:
                return None

    def _on_canvas_press(self, event):
        """Flag a drag only when the press lands on the legend box."""
        self._dragging_legend = False
        found = self._find_legend()
        if found is None or event.x is None:
            return
        bbox = self._legend_bbox(found[0])
        if bbox is not None and bbox.contains(event.x, event.y):
            self._dragging_legend = True

    def _on_canvas_release(self, _event):
        """After a legend drag, write its upper-left corner into X/Y (axes
        fraction) and persist it so the next render keeps the spot."""
        if not self._dragging_legend:
            return
        self._dragging_legend = False
        found = self._find_legend()
        if found is None:
            return
        leg, ax = found
        bbox = self._legend_bbox(leg)
        if bbox is None:
            return
        x0, y1 = ax.transAxes.inverted().transform((bbox.x0, bbox.y1))
        self.leg_x_var.set(f"{x0:.3f}")
        self.leg_y_var.set(f"{y1:.3f}")
        vc.LEGEND_XY = (float(x0), float(y1))   # keep it on the next render

    # PRB figure-size presets: label -> fn(gui) -> (width_in, height_in) or
    # None ("(custom)" clears the size fields so module defaults apply).  The
    # "(auto)" entries size themselves from the current figure's panel count.
    _PRESETS = {
        "(custom)":                        lambda g: None,
        "PRB single-col — 1 panel":   lambda g: vc.prb_single(),
        "PRB double-col — 1 panel":   lambda g: vc.prb_double(),
        "PRB double-col — panels (auto)":
            lambda g: vc.prb_stacked(g._num_panels()),
        "PRB single-col — panels (auto)":
            lambda g: vc.prb_stacked(g._num_panels(), double=False),
    }

    # ------------------------------------------------------------------
    # fonts panel: legend + axis text size multipliers (view + export)
    # ------------------------------------------------------------------
    def _build_fonts_panel(self, parent):
        box = ttk.Frame(parent, padding=8)
        box.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(box, text="Legend font (x):").grid(
            row=0, column=0, padx=6, pady=2, sticky="w")
        lsb = ttk.Spinbox(box, from_=0.4, to=4.0, increment=0.1, width=6,
                          textvariable=self.legend_fs_var,
                          command=self._auto_render)
        lsb.grid(row=0, column=1, padx=6, pady=2, sticky="w")

        ttk.Label(box, text="Axis font (x):").grid(
            row=0, column=2, padx=6, pady=2, sticky="w")
        asb = ttk.Spinbox(box, from_=0.4, to=4.0, increment=0.1, width=6,
                          textvariable=self.axis_fs_var,
                          command=self._auto_render)
        asb.grid(row=0, column=3, padx=6, pady=2, sticky="w")

        ttk.Label(box, text="1.0 = default; axis = labels + tick numbers"
                  ).grid(row=1, column=0, columnspan=4, padx=6, pady=(2, 0),
                         sticky="w")

        for sb in (lsb, asb):
            sb.bind("<Return>",   lambda e: self._auto_render())
            sb.bind("<FocusOut>", lambda e: self._auto_render())

    # ------------------------------------------------------------------
    # output panel: figure size (cm) + export DPI + on-screen zoom
    # ------------------------------------------------------------------
    def _build_output_panel(self, parent):
        box = ttk.Frame(parent, padding=8)
        box.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(box, text="Width (cm):").grid(row=0, column=0, padx=6, pady=2, sticky="w")
        we = ttk.Entry(box, textvariable=self.figw_var, width=8)
        we.grid(row=0, column=1, padx=6, pady=2, sticky="w")

        ttk.Label(box, text="Height (cm):").grid(row=0, column=2, padx=6, pady=2, sticky="w")
        he = ttk.Entry(box, textvariable=self.figh_var, width=8)
        he.grid(row=0, column=3, padx=6, pady=2, sticky="w")

        ttk.Label(box, text="(blank = default)").grid(
            row=0, column=4, padx=6, sticky="w")

        ttk.Label(box, text="DPI:").grid(row=1, column=0, padx=6, pady=2, sticky="w")
        de = ttk.Entry(box, textvariable=self.dpi_var, width=8)
        de.grid(row=1, column=1, padx=6, pady=2, sticky="w")

        # on-screen zoom: view-only magnification of the right-hand canvas.
        # Does NOT change the figure's true size in cm or the exported file.
        ttk.Label(box, text="Screen zoom (x):").grid(
            row=1, column=2, padx=6, pady=2, sticky="w")
        zsb = ttk.Spinbox(box, from_=0.5, to=5.0, increment=0.25, width=6,
                          textvariable=self.zoom_var,
                          command=self._auto_render)
        zsb.grid(row=1, column=3, padx=6, pady=2, sticky="w")

        ttk.Label(box, text="PRB preset:").grid(
            row=2, column=0, padx=6, pady=2, sticky="w")
        pc = ttk.Combobox(box, textvariable=self.preset_var, state="readonly",
                          width=26, values=list(self._PRESETS))
        pc.grid(row=2, column=1, columnspan=3, padx=6, pady=2, sticky="w")
        pc.bind("<<ComboboxSelected>>", self._apply_preset)

        ttk.Label(box, text="(screen zoom is view-only; true size = cm above)"
                  ).grid(row=3, column=0, columnspan=4, padx=6, pady=(2, 0),
                         sticky="w")

        for ent in (we, he, de, zsb):
            ent.bind("<Return>",   lambda e: self._auto_render())
            ent.bind("<FocusOut>", lambda e: self._auto_render())

    # ------------------------------------------------------------------
    # axis-bounds panel (rebuilt whenever the chosen figure changes)
    # ------------------------------------------------------------------
    def _build_bounds_panel(self, parent):
        self.bounds_box = ttk.LabelFrame(
            parent, text="Axis bounds (blank = auto) & labels", padding=8)
        self.bounds_box.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))
        self._populate_bounds()

    def _panel_rows(self):
        """(key, row-label, default y-axis label) rows for the current figure.

        First row is the shared pressure x-axis (no editable y-label -> None);
        then one row per panel carrying that panel's default y-axis label text
        (PANELS[i][1]) so the GUI can seed the editable label box with what the
        plot currently shows.
        """
        module = FIGURES[self.fig_choice.get()]
        rows = [("x", "Pressure  (x, shared)", None)]
        # the single-element figure has one panel driven by the Element dropdown,
        # not a fixed module.PANELS list, so build its row from the selection.
        if module is plot_single:
            key = self.single_var.get()
            ylab = vc.QUANTITIES.get(key, (key, None, None))[0]
            rows.append((key, f"{key}  (y)", ylab))
            return rows
        for p in getattr(module, "PANELS", []):
            key = p[0]
            default_ylab = p[1] if len(p) > 1 else key
            rows.append((key, f"{key}  (y)", default_ylab))
        return rows

    def _populate_bounds(self):
        """Clear and rebuild the bounds + label rows for the current figure."""
        for w in self.bounds_box.winfo_children():
            w.destroy()
        self.bound_vars = {}
        self.ylabel_vars = {}

        for c, txt in enumerate(("Axis / panel", "Lower", "Upper", "Axis label")):
            ttk.Label(self.bounds_box, text=txt).grid(
                row=0, column=c, padx=6, pady=(0, 4), sticky="w")

        for r, (key, label, default_ylab) in enumerate(self._panel_rows(), start=1):
            ttk.Label(self.bounds_box, text=label).grid(
                row=r, column=0, padx=6, pady=2, sticky="w")
            saved = self.bound_store.get(key, ["", ""])
            lo = tk.StringVar(value=saved[0])
            hi = tk.StringVar(value=saved[1])
            for c, var in ((1, lo), (2, hi)):
                ent = ttk.Entry(self.bounds_box, textvariable=var, width=10)
                ent.grid(row=r, column=c, padx=6, pady=2, sticky="w")
                ent.bind("<Return>",   self._on_bound_edit)
                ent.bind("<FocusOut>", self._on_bound_edit)
            self.bound_vars[key] = (lo, hi)

            # editable y-axis label (y panels only; the shared x row has none).
            # Seed with the user's saved text, else the plot's current default
            # so they can see and trim it, e.g. "Shear modulus $G_H$ (GPa)" ->
            # "GH".  Blank falls back to the module default at render time.
            if default_ylab is not None:
                lv = tk.StringVar(
                    value=self.ylabel_store.get(key, default_ylab))
                lent = ttk.Entry(self.bounds_box, textvariable=lv, width=24)
                lent.grid(row=r, column=3, padx=6, pady=2, sticky="w")
                lent.bind("<Return>",   self._on_bound_edit)
                lent.bind("<FocusOut>", self._on_bound_edit)
                self.ylabel_vars[key] = lv

        # "Auto (clear all)" convenience button (bounds only; labels persist)
        ttk.Button(self.bounds_box, text="Auto (clear all)",
                   command=self._clear_bounds).grid(
                       row=len(self.bound_vars) + 1, column=0,
                       columnspan=4, sticky="w", padx=6, pady=(6, 0))

    def _stash_bounds(self):
        """Remember current entries so they survive a figure switch."""
        for key, (lo, hi) in self.bound_vars.items():
            self.bound_store[key] = [lo.get(), hi.get()]
        for key, lv in self.ylabel_vars.items():
            self.ylabel_store[key] = lv.get()

    def _on_figure_change(self, _event=None):
        self._stash_bounds()
        self._populate_bounds()

    def _on_single_change(self, _event=None):
        """Element dropdown changed: rebuild bounds for the new quantity and
        redraw if the single-element figure is the one on screen."""
        self._stash_bounds()
        vc.SINGLE = self.single_var.get()
        self._populate_bounds()
        if FIGURES[self.fig_choice.get()] is plot_single:
            self._auto_render()

    def _on_bound_edit(self, _event=None):
        self._stash_bounds()
        self._auto_render()

    def _clear_bounds(self):
        for lo, hi in self.bound_vars.values():
            lo.set("")
            hi.set("")
        self._stash_bounds()
        self._auto_render()

    @staticmethod
    def _to_float(text):
        """Parse an entry to float; blank or bad text -> None (autoscale)."""
        text = text.strip()
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None

    def _parse_dpi(self):
        """Parse the DPI entry; blank or bad text -> the module default (600)."""
        try:
            d = int(float(self.dpi_var.get().strip()))
            if d > 0:
                return d
        except ValueError:
            pass
        return vc.DPI if vc.DPI else 600

    def _parse_figsize(self):
        """Return (width_in, height_in) from the cm entries, or None (default).

        Both width and height must parse to positive numbers, else the plot
        module's own default size is used.
        """
        w = self._to_float(self.figw_var.get())
        h = self._to_float(self.figh_var.get())
        if w and h and w > 0 and h > 0:
            return (w / vc.CM_PER_IN, h / vc.CM_PER_IN)
        return None

    def _parse_zoom(self):
        """On-screen zoom factor from the entry; blank/bad/<=0 -> 1.0."""
        z = self._to_float(self.zoom_var.get())
        return z if (z and z > 0) else 1.0

    def _parse_scale(self, var, default=1.0):
        """A font-size multiplier StringVar -> float; blank/bad/<=0 -> default."""
        v = self._to_float(var.get())
        return v if (v and v > 0) else default

    def _num_panels(self):
        """Number of stacked y-panels in the current figure (>= 1)."""
        return max(1, sum(1 for _k, _lbl, ylab in self._panel_rows()
                          if ylab is not None))

    def _apply_preset(self, *_):
        """Fill the Width/Height (cm) fields from the chosen PRB preset."""
        fn = self._PRESETS.get(self.preset_var.get())
        if fn is None:
            return
        size_in = fn(self)
        if size_in is None:                 # "(custom)" -> module default
            self.figw_var.set("")
            self.figh_var.set("")
        else:
            w_in, h_in = size_in
            self.figw_var.set(f"{w_in * vc.CM_PER_IN:.2f}")
            self.figh_var.set(f"{h_in * vc.CM_PER_IN:.2f}")
        self._auto_render()

    def _collect_bounds(self):
        """Push the current bound entries into vc.XLIM / vc.YLIM."""
        vc.XLIM[0], vc.XLIM[1] = None, None
        vc.YLIM.clear()
        for key, (lo, hi) in self.bound_vars.items():
            pair = [self._to_float(lo.get()), self._to_float(hi.get())]
            if key == "x":
                vc.XLIM[0], vc.XLIM[1] = pair
            else:
                vc.YLIM[key] = pair

    def _collect_ylabels(self):
        """Push the current per-panel label entries into vc.YLABEL.

        Blank entries are left out, so those panels keep their module default.
        """
        vc.YLABEL.clear()
        for key, lv in self.ylabel_vars.items():
            text = lv.get().strip()
            if text:
                vc.YLABEL[key] = text

    # ------------------------------------------------------------------
    # style callbacks
    # ------------------------------------------------------------------
    def _set_color(self, key, value):
        """Validate a colour (hex or matplotlib name), store it, paint swatch."""
        if not is_color_like(value):
            messagebox.showerror("Bad colour", f"'{value}' is not a valid colour.")
            self.hexvars[key].set(self.colors[key])   # revert entry
            return False
        hx = to_hex(value)
        self.colors[key] = hx
        self.hexvars[key].set(hx)
        self.swatches[key].configure(bg=hx, activebackground=hx)
        return True

    def _pick_color(self, key):
        rgb, hx = colorchooser.askcolor(color=self.colors[key],
                                        title=f"Colour for {key}")
        if hx and self._set_color(key, hx):
            self._auto_render()

    def _apply_hex(self, key):
        val = self.hexvars[key].get().strip()
        if val and val != self.colors[key]:
            if self._set_color(key, val):
                self._auto_render()

    def _apply_palette(self, _event=None):
        pal = vc.PALETTES.get(self.palette_var.get())
        if not pal:
            return
        for key, hx in pal.items():
            self._set_color(key, hx)
        self._auto_render()

    def _on_marker(self, key):
        self.markers[key] = STYLES[self.markvars[key].get()]
        self._auto_render()

    def _apply_zorder(self, key):
        """Store an edited layer (z-order); bad/blank text reverts to current."""
        try:
            z = int(float(self.zordervars[key].get().strip()))
        except (ValueError, AttributeError):
            self.zordervars[key].set(str(self.zorders[key]))   # revert
            return
        if z != self.zorders[key]:
            self.zorders[key] = z
            self.zordervars[key].set(str(z))                   # normalise text
            self._auto_render()

    def _apply_name(self, key):
        """Store an edited legend tag; blank reverts to the current value."""
        val = self.namevars[key].get().strip()
        if not val:
            self.namevars[key].set(self.names[key])   # don't allow empty
            return
        if val != self.names[key]:
            self.names[key] = val
            self._auto_render()

    def _auto_render(self):
        """Re-draw immediately if a figure is already on screen."""
        if self.current_fig is not None:
            self.render()

    # ------------------------------------------------------------------
    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select data workbook",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.src_var.set(path)
            # surface any comparison sources in the new file right away
            self._discover(path)
            self._sync_style_rows()

    def _force_quit(self):
        """Really shut everything down.

        The window's X and a plain ``mainloop`` exit can leave the process
        alive in the background - open matplotlib figures, a canvas timer, or
        a stray render still hold references, so ``python.exe`` keeps running,
        stays out of the taskbar, and blocks the next launch.  This tears the
        GUI down and then calls ``os._exit``, which terminates the interpreter
        (and every thread it owns) immediately, skipping the atexit/cleanup
        handlers that can otherwise hang.  Nothing survives it.
        """
        # best-effort graceful teardown first (ignored if already gone)
        try:
            import matplotlib.pyplot as plt
            plt.close("all")           # drop every open figure
        except Exception:
            pass
        try:
            self._clear_canvas()       # destroy the embedded canvas + toolbar
        except Exception:
            pass
        try:
            self.quit()                # break out of mainloop
            self.destroy()             # tear down all widgets
        except Exception:
            pass
        # hard stop: guarantees the process (and any non-daemon thread) exits
        # now, so nothing is left running in the background.
        os._exit(0)

    def _clear_canvas(self):
        if self.toolbar is not None:
            self.toolbar.destroy()
            self.toolbar = None
        if self.canvas is not None:
            self.canvas.get_tk_widget().destroy()
            self.canvas = None

    def render(self):
        module = FIGURES[self.fig_choice.get()]
        path = self.src_var.get().strip() or None

        # push the current style choices into the shared module
        vc.COL.update(self.colors)
        vc.MARK.update(self.markers)
        vc.LABEL.update(self.names)   # editable legend tags
        vc.ZORDER.update(self.zorders)   # per-series draw order (higher = on top)
        vc.SHOW.update({k: v.get() for k, v in self.showvars.items()})  # per-series on/off
        vc.SHOW_YERR = self.show_yerr.get()   # marker error-bar toggles
        vc.SHOW_XERR = self.show_xerr.get()
        vc.LINE_UNC = LINE_UNC.get(self.line_unc_var.get(), "none")  # line series
        vc.DPI = self._parse_dpi()            # export resolution
        vc.FIGSIZE = self._parse_figsize()    # figure size (in), None => default
        vc.LEGEND_FONT_SCALE = self._parse_scale(self.legend_fs_var)  # legend text
        vc.AXIS_FONT_SCALE = self._parse_scale(self.axis_fs_var)      # axis text
        vc.SINGLE = self.single_var.get()     # which quantity the single plot draws
        self._collect_bounds()          # push axis bounds into vc.XLIM / vc.YLIM
        self._collect_ylabels()         # push per-panel y-axis labels into vc.YLABEL
        self._collect_legend()          # push legend position + columns into vc

        try:
            fig = module.main(path=path, show=False)
        except Exception as exc:                      # surface any data error
            messagebox.showerror("Plot failed", str(exc))
            self.status.set(f"Error: {exc}")
            return

        # a newly-loaded workbook may have introduced comparison sources -
        # give each one its own style row (preserving existing choices).
        self._sync_style_rows()

        # on-screen zoom: bump ONLY the figure's screen dpi so the canvas is
        # drawn larger in the right-hand panel.  The figure's true size in
        # inches (get_size_inches) is untouched, and "Save" passes its own
        # explicit dpi, so neither the reported dimensions nor the exported
        # file are affected - this is purely a magnifying glass for viewing.
        zoom = self._parse_zoom()
        if zoom != 1.0:
            fig.set_dpi(fig.get_dpi() * zoom)

        self._clear_canvas()
        self.current_fig = fig
        self.canvas = FigureCanvasTkAgg(fig, master=self.canvas_frame)
        self.canvas.draw()
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.canvas_frame)
        self.toolbar.update()
        # WYSIWYG: show the figure at its true size (width_in x height_in x the
        # figure's display dpi) instead of stretching it to fill the window, so
        # the on-screen plot visibly scales with the entered dimensions.  No
        # fill => the packer keeps the canvas at its requested pixel size;
        # expand=True donates the leftover right-half space to the canvas's
        # cavity so anchor="center" leaves grey margins above and below it.
        self.canvas.get_tk_widget().pack(anchor="center", expand=True, pady=6)
        # legend drag: report the dropped position back into the X/Y boxes
        self.canvas.mpl_connect("button_press_event", self._on_canvas_press)
        self.canvas.mpl_connect("button_release_event", self._on_canvas_release)
        w_cm = fig.get_size_inches()[0] * vc.CM_PER_IN
        h_cm = fig.get_size_inches()[1] * vc.CM_PER_IN
        self.status.set(
            f"Rendered: {self.fig_choice.get()}  "
            f"({w_cm:.1f} x {h_cm:.1f} cm @ {vc.DPI} dpi"
            + (f", shown at {zoom:g}x" if zoom != 1.0 else "") + ")")

    def save(self):
        if self.current_fig is None:
            messagebox.showinfo("Nothing to save", "Plot a figure first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG", "*.png"), ("PDF", "*.pdf"), ("SVG", "*.svg")],
        )
        if path:
            self.current_fig.savefig(path, dpi=self._parse_dpi(),
                                     bbox_inches="tight")
            self.status.set(f"Saved -> {os.path.basename(path)}")


if __name__ == "__main__":
    VPlotApp().mainloop()
