"""
WaveformPlotter - make publication-style waveform figures from exported views.

Reads the Excel (or CSV) files written by Baosheng_PEO.py's "Export view"
button - a `waveform` sheet with a time column (us) and one or more amplitude
columns - and draws a single-panel figure in the style of an ultrasonic
echo-train plot (amplitude vs time).

Features
    * Load .xlsx (waveform sheet) or .csv.
    * Pick which amplitude column to plot; pick the trace colour.
    * Three interface labels - "Anvil/Buffer-rod", "Buffer-rod/Sample" and
      "Sample/Backing" - are always shown and can be dragged with the mouse.
    * Optional panel letter "(a)", "(b)", ... (also draggable) for multi-panel
      figures.
    * Y axis autoscales symmetrically to the max amplitude; X axis (in us) is
      editable.
    * Save the figure to PNG / PDF / SVG at 300 dpi.

Run:  py WaveformPlotter.py
"""
import os
import numpy as np
import pandas as pd

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg, NavigationToolbar2Tk,
)

# default interface labels (always present) and their start positions
# in axes-fraction coordinates (0-1), roughly matching an echo train.
DEFAULT_LABELS = [
    ("Anvil/Buffer-rod",  0.14, 0.16),
    ("Buffer-rod/Sample", 0.52, 0.70),
    ("Sample/Backing",    0.80, 0.24),
]
TRACE_COLORS = ["tab:purple", "tab:red", "tab:blue", "black",
                "tab:green", "tab:orange"]


class DraggableTexts:
    """Make a set of matplotlib Text artists draggable with the mouse."""

    def __init__(self, canvas, ax, toolbar=None):
        self.canvas = canvas
        self.ax = ax
        self.toolbar = toolbar
        self.texts = []
        self._drag = None
        self.on_move = None      # optional callback(artist) fired while dragging
        canvas.mpl_connect("button_press_event", self.on_press)
        canvas.mpl_connect("motion_notify_event", self.on_motion)
        canvas.mpl_connect("button_release_event", self.on_release)

    def add(self, text_artist):
        self.texts.append(text_artist)

    def _toolbar_active(self):
        # don't hijack clicks while pan/zoom is engaged
        return bool(self.toolbar and getattr(self.toolbar, "mode", ""))

    def on_press(self, event):
        if event.inaxes != self.ax or self._toolbar_active():
            return
        # topmost first so overlapping labels grab in draw order
        for t in reversed(self.texts):
            contains, _ = t.contains(event)
            if contains:
                self._drag = t
                break

    def on_motion(self, event):
        if self._drag is None or event.x is None or event.y is None:
            return
        fx, fy = self.ax.transAxes.inverted().transform((event.x, event.y))
        self._drag.set_position((fx, fy))
        if self.on_move is not None:
            self.on_move(self._drag)
        self.canvas.draw_idle()

    def on_release(self, event):
        self._drag = None


class WaveformPlotter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WaveformPlotter - echo-train figure maker")
        self.geometry("1100x800")

        # --- data / state ---------------------------------------------------
        self.df = None
        self.time_col = None
        self.filename = "(none)"

        self.amp_col = tk.StringVar()
        self.color = tk.StringVar(value=TRACE_COLORS[0])
        self.xlabel = tk.StringVar(value="Time (μs)")   # us
        self.ylabel = tk.StringVar(value="Amplitude (mV)")
        self.xmin = tk.DoubleVar(value=0.0)
        self.xmax = tk.DoubleVar(value=1.0)
        self.ypad = tk.DoubleVar(value=1.10)                 # max-abs padding
        self.panel = tk.StringVar(value="(a)")
        self.show_panel = tk.BooleanVar(value=True)
        self.label_size = tk.IntVar(value=10)

        self.dpi = tk.IntVar(value=600)         # save resolution (raster)
        self.fig_w = tk.DoubleVar(value=7.5)    # figure width  (inches)
        self.fig_h = tk.DoubleVar(value=3.2)    # figure height (inches)

        self.label_vars = []     # editable text StringVars for the 3 labels
        self.label_artists = []  # the Text artists on the plot
        self.panel_artist = None

        # live x/y (axes fraction 0-1) for the 3 labels + panel letter
        self.coord_x = []        # list of DoubleVar
        self.coord_y = []        # list of DoubleVar
        self.coord_map = {}      # artist -> index into coord_x/coord_y
        self._all_positions = [(fx, fy) for _, fx, fy in DEFAULT_LABELS] \
            + [(0.94, 0.88)]     # + panel default

        self._build_controls()
        self._build_plot()

    # ==================================================================
    # UI construction
    # ==================================================================
    def _build_controls(self):
        bar = ttk.Frame(self, padding=6)
        bar.pack(side=tk.TOP, fill=tk.X)

        # --- file ---
        f_file = ttk.LabelFrame(bar, text="File", padding=4)
        f_file.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Button(f_file, text="Load data", command=self.load).pack(fill=tk.X)
        self.lbl_file = ttk.Label(f_file, text=self.filename, width=20)
        self.lbl_file.pack()

        # --- trace ---
        f_tr = ttk.LabelFrame(bar, text="Trace", padding=4)
        f_tr.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Label(f_tr, text="amplitude").grid(row=0, column=0, sticky="e")
        self.cmb_amp = ttk.Combobox(f_tr, textvariable=self.amp_col,
                                    width=18, state="readonly")
        self.cmb_amp.grid(row=0, column=1, padx=2)
        self.cmb_amp.bind("<<ComboboxSelected>>", lambda e: self.redraw())
        ttk.Label(f_tr, text="colour").grid(row=1, column=0, sticky="e")
        cmb_col = ttk.Combobox(f_tr, textvariable=self.color, width=18,
                               values=TRACE_COLORS, state="readonly")
        cmb_col.grid(row=1, column=1, padx=2)
        cmb_col.bind("<<ComboboxSelected>>", lambda e: self.redraw())

        # --- axes ---
        f_ax = ttk.LabelFrame(bar, text="Axes", padding=4)
        f_ax.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Label(f_ax, text="x min").grid(row=0, column=0, sticky="e")
        e0 = ttk.Entry(f_ax, textvariable=self.xmin, width=8)
        e0.grid(row=0, column=1)
        ttk.Label(f_ax, text="x max").grid(row=0, column=2, sticky="e")
        e1 = ttk.Entry(f_ax, textvariable=self.xmax, width=8)
        e1.grid(row=0, column=3)
        for e in (e0, e1):
            e.bind("<Return>", lambda ev: self.redraw())
        ttk.Button(f_ax, text="Apply x", command=self.redraw).grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(2, 0))
        ttk.Button(f_ax, text="Reset x", command=self.reset_x).grid(
            row=1, column=2, columnspan=2, sticky="ew", pady=(2, 0))
        ttk.Label(f_ax, text="x label").grid(row=2, column=0, sticky="e")
        ex = ttk.Entry(f_ax, textvariable=self.xlabel, width=14)
        ex.grid(row=2, column=1, columnspan=3, sticky="ew")
        ex.bind("<Return>", lambda ev: self.redraw())
        ttk.Label(f_ax, text="y label").grid(row=3, column=0, sticky="e")
        ey = ttk.Entry(f_ax, textvariable=self.ylabel, width=14)
        ey.grid(row=3, column=1, columnspan=3, sticky="ew")
        ey.bind("<Return>", lambda ev: self.redraw())

        # --- labels ---
        f_lb = ttk.LabelFrame(bar, text="Interface labels (drag on plot)",
                              padding=4)
        f_lb.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        for i, (txt, _, _) in enumerate(DEFAULT_LABELS):
            var = tk.StringVar(value=txt)
            self.label_vars.append(var)
            ent = ttk.Entry(f_lb, textvariable=var, width=20)
            ent.grid(row=i, column=0, pady=1)
            ent.bind("<Return>", lambda ev: self.redraw())
        row = len(DEFAULT_LABELS)
        ttk.Label(f_lb, text="font").grid(row=row, column=0, sticky="w")
        sp = ttk.Spinbox(f_lb, from_=6, to=20, textvariable=self.label_size,
                         width=5, command=self.redraw)
        sp.grid(row=row, column=0, sticky="e")
        ttk.Button(f_lb, text="Reset label positions",
                   command=self.reset_label_positions).grid(
            row=row + 1, column=0, sticky="ew", pady=(2, 0))

        # --- label positions (axes fraction 0-1) ---
        f_xy = ttk.LabelFrame(bar, text="Label positions (axes 0-1)", padding=4)
        f_xy.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Label(f_xy, text="x", width=7, anchor="center").grid(row=0, column=1)
        ttk.Label(f_xy, text="y", width=7, anchor="center").grid(row=0, column=2)
        names = [t.split("/")[0] for t, _, _ in DEFAULT_LABELS] + ["Panel"]
        for i, (name, (dx, dy)) in enumerate(zip(names, self._all_positions)):
            ttk.Label(f_xy, text=name, width=7).grid(row=i + 1, column=0,
                                                     sticky="w")
            vx = tk.DoubleVar(value=round(dx, 3))
            vy = tk.DoubleVar(value=round(dy, 3))
            self.coord_x.append(vx)
            self.coord_y.append(vy)
            ex = ttk.Entry(f_xy, textvariable=vx, width=7)
            ex.grid(row=i + 1, column=1)
            ey = ttk.Entry(f_xy, textvariable=vy, width=7)
            ey.grid(row=i + 1, column=2)
            ex.bind("<Return>", lambda ev: self.apply_positions())
            ey.bind("<Return>", lambda ev: self.apply_positions())
        nrow = len(names) + 1
        ttk.Button(f_xy, text="Apply", command=self.apply_positions).grid(
            row=nrow, column=0, columnspan=3, sticky="ew", pady=(2, 0))
        ttk.Button(f_xy, text="Align Y (labels)",
                   command=self.align_labels_y).grid(
            row=nrow + 1, column=0, columnspan=3, sticky="ew")

        # --- panel + save ---
        f_pn = ttk.LabelFrame(bar, text="Panel / save", padding=4)
        f_pn.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Checkbutton(f_pn, text="show panel letter",
                        variable=self.show_panel,
                        command=self.redraw).grid(row=0, column=0,
                                                  columnspan=2, sticky="w")
        ttk.Label(f_pn, text="text").grid(row=1, column=0, sticky="e")
        ep = ttk.Entry(f_pn, textvariable=self.panel, width=6)
        ep.grid(row=1, column=1, sticky="w")
        ep.bind("<Return>", lambda ev: self.redraw())

        ttk.Label(f_pn, text="size (in) w").grid(row=2, column=0, sticky="e")
        ew = ttk.Entry(f_pn, textvariable=self.fig_w, width=6)
        ew.grid(row=2, column=1, sticky="w")
        ttk.Label(f_pn, text="h").grid(row=3, column=0, sticky="e")
        eh = ttk.Entry(f_pn, textvariable=self.fig_h, width=6)
        eh.grid(row=3, column=1, sticky="w")
        ttk.Label(f_pn, text="dpi").grid(row=4, column=0, sticky="e")
        ed = ttk.Entry(f_pn, textvariable=self.dpi, width=6)
        ed.grid(row=4, column=1, sticky="w")
        for e in (ew, eh, ed):
            e.bind("<Return>", lambda ev: self.apply_figsize())
        ttk.Button(f_pn, text="Apply size",
                   command=self.apply_figsize).grid(row=5, column=0,
                                                    columnspan=2, sticky="ew",
                                                    pady=(2, 0))
        ttk.Button(f_pn, text="Save figure...",
                   command=self.save_figure).grid(row=6, column=0,
                                                  columnspan=2,
                                                  sticky="ew", pady=(4, 0))

    def _build_plot(self):
        # Display at a fixed dpi so on-screen pixels == inches * dpi. This is
        # what makes the width/height controls actually change what you see
        # (WYSIWYG) instead of the figure just filling the window.
        self.fig = Figure(figsize=(self.fig_w.get(), self.fig_h.get()),
                          dpi=100, tight_layout=True)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_xlabel(self.xlabel.get())
        self.ax.set_ylabel(self.ylabel.get())
        (self.line,) = self.ax.plot([], [], color=self.color.get(), lw=0.8)
        self.ax.grid(False)

        frame = ttk.Frame(self)
        frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.canvas = FigureCanvasTkAgg(self.fig, master=frame)
        # Do NOT fill/stretch the canvas. Let it sit at the figure's true size
        # (centred in the frame) so what you see is exactly what gets saved.
        # If we stretched it, matplotlib would resize the figure to the widget
        # on every window resize and the size controls would have no effect.
        self.canvas.get_tk_widget().pack(side=tk.TOP, expand=True)
        self.toolbar = NavigationToolbar2Tk(self.canvas, frame)

        # create the always-present labels + panel letter, then make draggable
        self.dragger = DraggableTexts(self.canvas, self.ax, self.toolbar)
        self.dragger.on_move = self._sync_coord_vars
        for i, ((txt, fx, fy), var) in enumerate(zip(DEFAULT_LABELS,
                                                     self.label_vars)):
            art = self.ax.text(fx, fy, var.get(), transform=self.ax.transAxes,
                               fontsize=self.label_size.get(), ha="center",
                               va="center")
            self.label_artists.append(art)
            self.dragger.add(art)
            self.coord_map[art] = i
        self.panel_artist = self.ax.text(
            0.94, 0.88, self.panel.get(), transform=self.ax.transAxes,
            fontsize=14, ha="center", va="center")
        self.dragger.add(self.panel_artist)
        self.coord_map[self.panel_artist] = len(self.label_artists)

    # ==================================================================
    # data loading
    # ==================================================================
    def load(self):
        path = filedialog.askopenfilename(
            title="Load waveform view",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"),
                       ("Excel workbook", "*.xlsx *.xls"),
                       ("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            df = self._read_any(path)
        except Exception as exc:
            messagebox.showerror("Load error", str(exc))
            return

        # identify the time column (name contains "time"), else first column
        time_col = next((c for c in df.columns
                         if "time" in str(c).lower()), df.columns[0])
        amp_cols = [c for c in df.columns if c != time_col
                    and pd.api.types.is_numeric_dtype(df[c])]
        if not amp_cols:
            messagebox.showerror("Load error",
                                 "No numeric amplitude column found.")
            return

        self.df = df
        self.time_col = time_col
        self.filename = os.path.basename(path)
        self.lbl_file.config(text=self.filename)

        self.cmb_amp["values"] = amp_cols
        # prefer a filtered column if the export provided one
        default = next((c for c in amp_cols
                        if "filter" in str(c).lower()), amp_cols[0])
        self.amp_col.set(default)

        self.reset_x()   # sets x range from data + redraw

    @staticmethod
    def _read_any(path):
        ext = os.path.splitext(path)[1].lower()
        if ext in (".xlsx", ".xls"):
            xls = pd.ExcelFile(path)
            sheet = "waveform" if "waveform" in xls.sheet_names \
                else xls.sheet_names[0]
            return xls.parse(sheet)
        return pd.read_csv(path)

    # ==================================================================
    # drawing
    # ==================================================================
    def reset_x(self):
        if self.df is None:
            return
        t = self.df[self.time_col].to_numpy(dtype=float)
        self.xmin.set(round(float(np.nanmin(t)), 4))
        self.xmax.set(round(float(np.nanmax(t)), 4))
        self.redraw()

    def apply_figsize(self):
        """Resize the on-screen figure to the typed width/height (inches)."""
        try:
            w, h = float(self.fig_w.get()), float(self.fig_h.get())
        except (tk.TclError, ValueError):
            return
        if w <= 0 or h <= 0:
            return
        # set the figure size in inches, then resize the *Tk widget* to the
        # matching pixel size. On an embedded canvas set_size_inches alone does
        # NOT resize the widget, which left the old (larger) render painted
        # next to the new one - the "overlay". forward=False avoids poking the
        # window manager; the explicit widget resize + full draw() is what
        # actually updates and clears the canvas.
        self.fig.set_size_inches(w, h, forward=False)
        dpi = self.fig.get_dpi()
        self.canvas.get_tk_widget().configure(width=int(round(w * dpi)),
                                              height=int(round(h * dpi)))
        self.canvas.draw()

    def reset_label_positions(self):
        for art, (_, fx, fy) in zip(self.label_artists, DEFAULT_LABELS):
            art.set_position((fx, fy))
        if self.panel_artist is not None:
            self.panel_artist.set_position((0.94, 0.88))
        # push defaults back into the coordinate boxes
        for i, (dx, dy) in enumerate(self._all_positions):
            self.coord_x[i].set(round(dx, 3))
            self.coord_y[i].set(round(dy, 3))
        self.canvas.draw_idle()

    def _sync_coord_vars(self, artist):
        """Update the x/y boxes live while a label is being dragged."""
        i = self.coord_map.get(artist)
        if i is None:
            return
        x, y = artist.get_position()
        self.coord_x[i].set(round(float(x), 3))
        self.coord_y[i].set(round(float(y), 3))

    def apply_positions(self):
        """Move every label to the exact x/y typed in the boxes."""
        for art, i in self.coord_map.items():
            try:
                art.set_position((float(self.coord_x[i].get()),
                                  float(self.coord_y[i].get())))
            except (tk.TclError, ValueError):
                continue
        self.canvas.draw_idle()

    def align_labels_y(self):
        """Snap the three interface labels to a common y (the first one's)."""
        if not self.coord_y:
            return
        y0 = self.coord_y[0].get()
        for i in range(len(self.label_artists)):   # exclude the panel letter
            self.coord_y[i].set(round(float(y0), 3))
        self.apply_positions()

    def redraw(self):
        # keep labels/axis text in sync even before any data is loaded
        for art, var in zip(self.label_artists, self.label_vars):
            art.set_text(var.get())
            art.set_fontsize(self.label_size.get())
        self.panel_artist.set_text(self.panel.get())
        self.panel_artist.set_visible(self.show_panel.get())
        self.ax.set_xlabel(self.xlabel.get())
        self.ax.set_ylabel(self.ylabel.get())

        if self.df is not None:
            t = self.df[self.time_col].to_numpy(dtype=float)
            y = self.df[self.amp_col.get()].to_numpy(dtype=float)
            self.line.set_data(t, y)
            self.line.set_color(self.color.get())

            try:
                x0, x1 = float(self.xmin.get()), float(self.xmax.get())
            except (tk.TclError, ValueError):
                x0, x1 = float(np.nanmin(t)), float(np.nanmax(t))
            if x1 <= x0:
                x0, x1 = float(np.nanmin(t)), float(np.nanmax(t))
            self.ax.set_xlim(x0, x1)

            # y autoscale symmetrically to max amplitude within the x window
            win = (t >= x0) & (t <= x1)
            yy = y[win] if np.any(win) else y
            ymax = np.nanmax(np.abs(yy)) if yy.size else 1.0
            ymax = ymax * float(self.ypad.get()) if ymax > 0 else 1.0
            self.ax.set_ylim(-ymax, ymax)

        self.canvas.draw_idle()

    # ==================================================================
    # save
    # ==================================================================
    def save_figure(self):
        if self.df is None:
            messagebox.showinfo("Save figure", "Load data first.")
            return
        path = filedialog.asksaveasfilename(
            title="Save figure",
            defaultextension=".png",
            filetypes=[("PNG image", "*.png"), ("PDF", "*.pdf"),
                       ("SVG", "*.svg")],
            initialfile=os.path.splitext(self.filename)[0] + "_figure.png",
        )
        if not path:
            return
        self.apply_figsize()   # honour any just-typed size before saving
        try:
            dpi = int(self.dpi.get())
        except (tk.TclError, ValueError):
            dpi = 600
        try:
            # No bbox_inches="tight": it recrops to a different bounding box
            # than the one on screen (shifting the axes under the labels),
            # which is why saved figures didn't match the view. Save exactly
            # the figure as displayed.
            self.fig.savefig(path, dpi=dpi)
        except Exception as exc:
            messagebox.showerror("Save figure", str(exc))
            return
        messagebox.showinfo("Save figure",
                            f"Saved to\n{path}\n({dpi} dpi)")


if __name__ == "__main__":
    WaveformPlotter().mainloop()
