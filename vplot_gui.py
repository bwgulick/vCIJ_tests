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

FIGURES = {
    "Cij tri-plot (C11 / C12 / C44) - CK vs FS": plot_cij_tri,
    "Moduli dual (K / GH) - CK vs FS + poly":    plot_moduli_dual,
}

# series key -> friendly label shown in the style panel
SERIES = [
    ("CK",   "CK (Cook)"),
    ("FS",   "FS (finite strain)"),
    ("poly", "poly (Kpoly / Gpoly)"),
]

# friendly marker name -> matplotlib marker code
MARKERS = {
    "Circle": "o", "Square": "s", "Triangle up": "^", "Triangle down": "v",
    "Diamond": "D", "Pentagon": "p", "Star": "*", "Plus": "P",
    "X": "X", "Hexagon": "h",
}
MARKER_BY_CODE = {v: k for k, v in MARKERS.items()}


class VPlotApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Vanadium elastic-property plotter")
        self.geometry("1040x960")
        self.current_fig = None
        self.canvas = None
        self.toolbar = None

        # live per-series style, seeded from the module defaults
        self.colors = dict(vc.COL)
        self.markers = dict(vc.MARK)
        self.names = dict(vc.LABEL)   # editable legend tags (CK / FS / PC)
        # error-bar visibility toggles (all-vertical / all-horizontal)
        self.show_yerr = tk.BooleanVar(value=vc.SHOW_YERR)
        self.show_xerr = tk.BooleanVar(value=vc.SHOW_XERR)
        # widget handles filled in by _build_style_panel
        self.hexvars = {}
        self.swatches = {}
        self.markvars = {}
        self.namevars = {}
        # axis-bounds state (see _build_bounds_panel)
        self.bound_vars = {}     # "x" / panel key -> (lo StringVar, hi StringVar)
        self.bound_store = {}    # key -> [lo str, hi str], kept across figure switches

        self._build_controls()
        self._build_settings_row()
        self._build_canvas_area()

    # ------------------------------------------------------------------
    # settings row: series style (left) beside axis bounds + error bars
    # (right), so the controls stay short and the figure gets the height.
    # ------------------------------------------------------------------
    def _build_settings_row(self):
        row = ttk.Frame(self, padding=(8, 0))
        row.pack(side=tk.TOP, fill=tk.X)

        left = ttk.Frame(row)
        left.pack(side=tk.LEFT, anchor="n")
        right = ttk.Frame(row)
        right.pack(side=tk.LEFT, anchor="n", fill=tk.X, expand=True, padx=(8, 0))

        self._build_style_panel(left)
        self._build_bounds_panel(right)
        self._build_display_panel(right)

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

        ttk.Button(bar, text="Plot", command=self.render).pack(side=tk.LEFT)
        ttk.Button(bar, text="Save PNG...", command=self.save).pack(
            side=tk.LEFT, padx=(6, 0))

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

    def _build_style_panel(self, parent):
        box = ttk.LabelFrame(parent, text="Series style (colour + marker)", padding=8)
        box.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        # preset palette row
        prow = ttk.Frame(box)
        prow.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 6))
        ttk.Label(prow, text="Palette preset:").pack(side=tk.LEFT)
        self.palette_var = tk.StringVar(value="Okabe-Ito (CB-safe)")
        pcb = ttk.Combobox(prow, textvariable=self.palette_var,
                           values=list(vc.PALETTES), state="readonly", width=24)
        pcb.pack(side=tk.LEFT, padx=6)
        pcb.bind("<<ComboboxSelected>>", self._apply_palette)

        # column headers
        for c, txt in enumerate(
                ("Series", "Legend name", "Colour", "Hex / name", "Marker")):
            ttk.Label(box, text=txt).grid(row=1, column=c, padx=6, sticky="w")

        # one row per series
        for i, (key, label) in enumerate(SERIES, start=2):
            ttk.Label(box, text=label).grid(row=i, column=0, padx=6, pady=2, sticky="w")

            # editable legend tag (the moduli panels prefix it: "K "+tag, etc.)
            nv = tk.StringVar(value=self.names[key])
            nent = ttk.Entry(box, textvariable=nv, width=12)
            nent.grid(row=i, column=1, padx=6, pady=2, sticky="w")
            nent.bind("<Return>",   lambda e, k=key: self._apply_name(k))
            nent.bind("<FocusOut>", lambda e, k=key: self._apply_name(k))
            self.namevars[key] = nv

            sw = tk.Button(box, width=3, relief="raised",
                           command=lambda k=key: self._pick_color(k))
            sw.grid(row=i, column=2, padx=6, pady=2)
            self.swatches[key] = sw

            hv = tk.StringVar(value=self.colors[key])
            ent = ttk.Entry(box, textvariable=hv, width=12)
            ent.grid(row=i, column=3, padx=6, pady=2, sticky="w")
            ent.bind("<Return>",   lambda e, k=key: self._apply_hex(k))
            ent.bind("<FocusOut>", lambda e, k=key: self._apply_hex(k))
            self.hexvars[key] = hv

            mv = tk.StringVar(value=MARKER_BY_CODE.get(self.markers[key], "Circle"))
            mcb = ttk.Combobox(box, textvariable=mv, values=list(MARKERS),
                               state="readonly", width=14)
            mcb.grid(row=i, column=4, padx=6, pady=2, sticky="w")
            mcb.bind("<<ComboboxSelected>>", lambda e, k=key: self._on_marker(k))
            self.markvars[key] = mv

            self._set_color(key, self.colors[key])   # paint the swatch

    # ------------------------------------------------------------------
    # display-options panel (error-bar visibility)
    # ------------------------------------------------------------------
    def _build_display_panel(self, parent):
        box = ttk.LabelFrame(parent, text="Error bars", padding=8)
        box.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))
        ttk.Checkbutton(box, text="Show vertical (y) uncertainties",
                        variable=self.show_yerr,
                        command=self._auto_render).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Checkbutton(box, text="Show horizontal (x) uncertainties",
                        variable=self.show_xerr,
                        command=self._auto_render).pack(side=tk.LEFT)

    # ------------------------------------------------------------------
    # axis-bounds panel (rebuilt whenever the chosen figure changes)
    # ------------------------------------------------------------------
    def _build_bounds_panel(self, parent):
        self.bounds_box = ttk.LabelFrame(
            parent, text="Axis bounds (blank = auto)", padding=8)
        self.bounds_box.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))
        self._populate_bounds()

    def _panel_rows(self):
        """(key, label) rows for the current figure: shared x + one per panel."""
        module = FIGURES[self.fig_choice.get()]
        rows = [("x", "Pressure  (x, shared)")]
        for p in getattr(module, "PANELS", []):
            key = p[0]
            rows.append((key, f"{key}  (y)"))
        return rows

    def _populate_bounds(self):
        """Clear and rebuild the bounds rows for the current figure."""
        for w in self.bounds_box.winfo_children():
            w.destroy()
        self.bound_vars = {}

        for c, txt in enumerate(("Axis / panel", "Lower", "Upper")):
            ttk.Label(self.bounds_box, text=txt).grid(
                row=0, column=c, padx=6, pady=(0, 4), sticky="w")

        for r, (key, label) in enumerate(self._panel_rows(), start=1):
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

        # "Auto (clear all)" convenience button
        ttk.Button(self.bounds_box, text="Auto (clear all)",
                   command=self._clear_bounds).grid(
                       row=len(self.bound_vars) + 1, column=0,
                       columnspan=3, sticky="w", padx=6, pady=(6, 0))

    def _stash_bounds(self):
        """Remember current entries so they survive a figure switch."""
        for key, (lo, hi) in self.bound_vars.items():
            self.bound_store[key] = [lo.get(), hi.get()]

    def _on_figure_change(self, _event=None):
        self._stash_bounds()
        self._populate_bounds()

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

    def _build_canvas_area(self):
        self.canvas_frame = ttk.Frame(self)
        self.canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

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
        self.markers[key] = MARKERS[self.markvars[key].get()]
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
        vc.SHOW_YERR = self.show_yerr.get()   # error-bar visibility toggles
        vc.SHOW_XERR = self.show_xerr.get()
        self._collect_bounds()          # push axis bounds into vc.XLIM / vc.YLIM

        try:
            fig = module.main(path=path, show=False)
        except Exception as exc:                      # surface any data error
            messagebox.showerror("Plot failed", str(exc))
            self.status.set(f"Error: {exc}")
            return

        self._clear_canvas()
        self.current_fig = fig
        self.canvas = FigureCanvasTkAgg(fig, master=self.canvas_frame)
        self.canvas.draw()
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.canvas_frame)
        self.toolbar.update()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        self.status.set(f"Rendered: {self.fig_choice.get()}")

    def save(self):
        if self.current_fig is None:
            messagebox.showinfo("Nothing to save", "Plot a figure first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG", "*.png"), ("PDF", "*.pdf"), ("SVG", "*.svg")],
        )
        if path:
            self.current_fig.savefig(path, dpi=200, bbox_inches="tight")
            self.status.set(f"Saved -> {os.path.basename(path)}")


if __name__ == "__main__":
    VPlotApp().mainloop()
