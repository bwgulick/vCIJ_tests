"""
Baosheng_PEO - Pulse-Echo-Overlap (PEO) travel-time measurement GUI.

Recreates the manual pulse-echo-overlap method of Baosheng Li et al. (2002,
J. Phys.: Condens. Matter 14 11337, "Sound velocity measurement using transfer
function method", figure 4) for measuring ultrasonic *travel times*.

Workflow
    1. Load a broadband waveform CSV (default: K1246-600bar.csv).
    2. Optionally band-pass filter it at a chosen centre frequency
       (30 MHz for S waves, 50 MHz for P waves).
    3. A movable *copy* of the waveform is overlaid on the original in the top
       panel.  Slide the copy in time (Fine = 0.0002 us, Jump = 0.05 us) until
       the buffer-rod echo of the copy sits on top of the sample echo of the
       original.  The shift = the travel time (typically ~0.3050 us).
    4. The bottom panel shows the interference (Sum or Difference) of the
       original and the shifted copy - constructive / destructive.
    5. Record travel times at each frequency and export to CSV.

No velocity is computed - this tool measures travel times only.

Run:  py Baosheng_PEO.py     (Python 3.11 - has numpy/pandas/scipy/matplotlib)
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
from matplotlib.widgets import RectangleSelector

from scipy.signal import butter, filtfilt

DEFAULT_CSV = "K1246-600bar.csv"

# ----- default step / filter parameters -------------------------------------
FINE_STEP = 0.0002   # us  (one sample interval)
JUMP_STEP = 0.05     # us  (coarse jump)
Y_FACTOR  = 1.05     # copy amplitude multiply / divide per Y Up / Y Down
DEF_FC    = 30.0     # MHz centre frequency (30 = S wave, 50 = P wave)
DEF_BW    = 20.0     # MHz band-pass bandwidth
DEF_ORDER = 4        # Butterworth order


def load_waveform(path):
    """Return (t_us, amp) arrays from a Tektronix-style CSV.

    Columns 3 (time, s) and 4 (amplitude) hold the trace; the first rows also
    carry metadata in columns 0-2 which parse to NaN and are dropped.
    """
    df = pd.read_csv(path, header=None, low_memory=False)
    if df.shape[1] < 5:
        raise ValueError("CSV needs at least 5 columns (time in col 4, amp in col 5).")
    t = pd.to_numeric(df[3], errors="coerce").to_numpy()
    a = pd.to_numeric(df[4], errors="coerce").to_numpy()
    good = ~np.isnan(t) & ~np.isnan(a)
    t, a = t[good], a[good]
    if t.size == 0:
        raise ValueError("No numeric time/amplitude data found in columns 4/5.")
    return t * 1e6, a          # time -> microseconds


class PEOApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Baosheng_PEO - pulse-echo-overlap travel-time tool")
        self.geometry("1200x900")

        # --- data / state ---------------------------------------------------
        self.t_us = None         # time axis (us)
        self.raw = None          # raw amplitude
        self.signal = None       # working signal (raw or filtered)
        self.dt_us = FINE_STEP   # sample interval (us)
        self.filename = "(none)"

        self.shift = tk.DoubleVar(value=0.3050)   # travel time (us)
        self.yscale = tk.DoubleVar(value=1.0)     # copy amplitude scale
        self.yscale_orig = tk.DoubleVar(value=1.0)  # original amplitude scale
        self.fine_step = tk.DoubleVar(value=FINE_STEP)
        self.jump_step = tk.DoubleVar(value=JUMP_STEP)
        self.fc = tk.DoubleVar(value=DEF_FC)
        self.bw = tk.DoubleVar(value=DEF_BW)
        self.order = tk.IntVar(value=DEF_ORDER)
        self.filter_on = tk.BooleanVar(value=True)
        self.interf_mode = tk.StringVar(value="sum")   # "sum" or "diff"
        self.zoom_on = tk.BooleanVar(value=False)      # box-zoom mode armed?

        self.records = []        # list of (freq_MHz, travel_time_us)

        # --- build UI -------------------------------------------------------
        self._build_controls()
        self._build_plots()

        # keyboard shortcuts for fast overlap alignment
        self.bind("<Left>",     lambda e: self.nudge(-self.fine_step.get()))
        self.bind("<Right>",    lambda e: self.nudge(+self.fine_step.get()))
        self.bind("<Prior>",    lambda e: self.nudge(+self.jump_step.get()))   # PageUp
        self.bind("<Next>",     lambda e: self.nudge(-self.jump_step.get()))   # PageDown
        self.bind("<Up>",       lambda e: self.scale_amp(Y_FACTOR))
        self.bind("<Down>",     lambda e: self.scale_amp(1.0 / Y_FACTOR))

        # auto-load default file if present
        if os.path.exists(DEFAULT_CSV):
            self._do_load(DEFAULT_CSV)

    # ==================================================================
    # UI construction
    # ==================================================================
    def _build_controls(self):
        bar = ttk.Frame(self, padding=6)
        bar.pack(side=tk.TOP, fill=tk.X)

        # --- file ---
        f_file = ttk.LabelFrame(bar, text="File", padding=4)
        f_file.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Button(f_file, text="Load CSV", command=self.load_csv).pack(fill=tk.X)
        self.lbl_file = ttk.Label(f_file, text=self.filename, width=22)
        self.lbl_file.pack()

        # --- filter ---
        f_filt = ttk.LabelFrame(bar, text="Band-pass filter (MHz)", padding=4)
        f_filt.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Checkbutton(f_filt, text="enabled", variable=self.filter_on,
                        command=self.apply_filter).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(f_filt, text="fc").grid(row=1, column=0, sticky="e")
        ttk.Entry(f_filt, textvariable=self.fc, width=6).grid(row=1, column=1)
        ttk.Button(f_filt, text="30 S", width=4,
                   command=lambda: self._set_fc(30)).grid(row=1, column=2, padx=1)
        ttk.Button(f_filt, text="50 P", width=4,
                   command=lambda: self._set_fc(50)).grid(row=1, column=3, padx=1)
        ttk.Label(f_filt, text="bw").grid(row=2, column=0, sticky="e")
        ttk.Entry(f_filt, textvariable=self.bw, width=6).grid(row=2, column=1)
        ttk.Label(f_filt, text="order").grid(row=2, column=2, sticky="e")
        ttk.Entry(f_filt, textvariable=self.order, width=4).grid(row=2, column=3)
        ttk.Button(f_filt, text="Apply", command=self.apply_filter).grid(
            row=3, column=0, columnspan=4, sticky="ew", pady=(2, 0))

        # --- shift / travel time ---
        f_sh = ttk.LabelFrame(bar, text="Shift = travel time (us)", padding=4)
        f_sh.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Button(f_sh, text="◀◀ Jump", width=8,
                   command=lambda: self.nudge(-self.jump_step.get())).grid(row=0, column=0)
        ttk.Button(f_sh, text="◀ Fine", width=8,
                   command=lambda: self.nudge(-self.fine_step.get())).grid(row=0, column=1)
        ttk.Button(f_sh, text="Fine ▶", width=8,
                   command=lambda: self.nudge(+self.fine_step.get())).grid(row=0, column=2)
        ttk.Button(f_sh, text="Jump ▶▶", width=8,
                   command=lambda: self.nudge(+self.jump_step.get())).grid(row=0, column=3)
        ttk.Label(f_sh, text="shift").grid(row=1, column=0, sticky="e")
        e_sh = ttk.Entry(f_sh, textvariable=self.shift, width=10)
        e_sh.grid(row=1, column=1)
        e_sh.bind("<Return>", lambda e: self.set_shift(self.shift.get()))
        ttk.Label(f_sh, text="fine").grid(row=1, column=2, sticky="e")
        ttk.Entry(f_sh, textvariable=self.fine_step, width=8).grid(row=1, column=3)
        ttk.Label(f_sh, text="jump").grid(row=2, column=2, sticky="e")
        ttk.Entry(f_sh, textvariable=self.jump_step, width=8).grid(row=2, column=3)

        # --- amplitude ---
        f_amp = ttk.LabelFrame(bar, text="Amplitude scale", padding=4)
        f_amp.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        # original (blue)
        ttk.Label(f_amp, text="original").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Button(f_amp, text="Y Up", width=5,
                   command=lambda: self.scale_amp(Y_FACTOR, "orig")).grid(row=1, column=0)
        ttk.Button(f_amp, text="Y Dn", width=5,
                   command=lambda: self.scale_amp(1.0 / Y_FACTOR, "orig")).grid(row=1, column=1)
        e_yo = ttk.Entry(f_amp, textvariable=self.yscale_orig, width=7)
        e_yo.grid(row=1, column=2)
        e_yo.bind("<Return>", lambda e: self.redraw())
        # copy (red)
        ttk.Label(f_amp, text="copy").grid(row=2, column=0, columnspan=3, sticky="w")
        ttk.Button(f_amp, text="Y Up", width=5,
                   command=lambda: self.scale_amp(Y_FACTOR, "copy")).grid(row=3, column=0)
        ttk.Button(f_amp, text="Y Dn", width=5,
                   command=lambda: self.scale_amp(1.0 / Y_FACTOR, "copy")).grid(row=3, column=1)
        e_y = ttk.Entry(f_amp, textvariable=self.yscale, width=7)
        e_y.grid(row=3, column=2)
        e_y.bind("<Return>", lambda e: self.redraw())

        # --- interference ---
        f_int = ttk.LabelFrame(bar, text="Interference", padding=4)
        f_int.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Radiobutton(f_int, text="Sum (orig+copy)", value="sum",
                        variable=self.interf_mode, command=self.redraw).pack(anchor="w")
        ttk.Radiobutton(f_int, text="Difference (orig-copy)", value="diff",
                        variable=self.interf_mode, command=self.redraw).pack(anchor="w")

        # --- view / zoom ---
        f_view = ttk.LabelFrame(bar, text="View / zoom", padding=4)
        f_view.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Checkbutton(f_view, text="Box zoom (drag)", variable=self.zoom_on,
                        command=self.toggle_zoom).grid(row=0, column=0, columnspan=2,
                                                       sticky="w")
        ttk.Label(f_view, text="Y axis").grid(row=1, column=0, columnspan=2, sticky="w")
        ttk.Button(f_view, text="Y out", width=6,
                   command=lambda: self.zoom_y(1.25)).grid(row=2, column=0)
        ttk.Button(f_view, text="Y in", width=6,
                   command=lambda: self.zoom_y(0.80)).grid(row=2, column=1)
        ttk.Button(f_view, text="Reset view",
                   command=self.reset_view).grid(row=3, column=0, columnspan=2,
                                                 sticky="ew", pady=(2, 0))

        # --- record ---
        f_rec = ttk.LabelFrame(bar, text="Travel-time records", padding=4)
        f_rec.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Button(f_rec, text="Record", command=self.record).pack(fill=tk.X)
        ttk.Button(f_rec, text="Export...", command=self.export).pack(fill=tk.X)
        self.rec_list = tk.Listbox(f_rec, height=4, width=20)
        self.rec_list.pack()

        # --- figure export ---
        f_fig = ttk.LabelFrame(bar, text="Figure export", padding=4)
        f_fig.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Button(f_fig, text="Export view (xlsx)",
                   command=self.export_view).pack(fill=tk.X)
        ttk.Label(f_fig, text="writes the visible\ntime window to Excel",
                  justify="left").pack(anchor="w")

        # --- readout ---
        self.readout = ttk.Label(self, text="Travel time: -- us",
                                 font=("Segoe UI", 16, "bold"), anchor="w")
        self.readout.pack(side=tk.TOP, fill=tk.X, padx=8, pady=2)

    def _build_plots(self):
        # Two separate figures inside a vertical PanedWindow: drag the sash
        # between them up/down to change the height split (e.g. 80% overlap /
        # 20% sum).  The x-axes are linked manually (see _on_top_xlim).
        self.paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        self.paned.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # --- top figure: overlay ------------------------------------------
        top_frame = ttk.Frame(self.paned)
        self.fig_top = Figure(figsize=(11, 4.4), tight_layout=True)
        self.ax_top = self.fig_top.add_subplot(111)
        self.ax_top.set_ylabel("amplitude")
        self.ax_top.set_title("PEO overlay - original (blue) + movable copy (red)")
        (self.ln_orig,) = self.ax_top.plot([], [], color="tab:blue", lw=0.8,
                                           label="original")
        (self.ln_copy,) = self.ax_top.plot([], [], color="tab:red", lw=0.8,
                                           alpha=0.8, label="copy (shifted)")
        self.ax_top.legend(loc="upper right", fontsize=8)
        self.ax_top.grid(True, alpha=0.3)
        self.canvas_top = FigureCanvasTkAgg(self.fig_top, master=top_frame)
        self.canvas_top.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(self.canvas_top, top_frame)  # pan / home / save
        self.paned.add(top_frame, weight=4)               # ~80% of extra space

        # --- bottom figure: interference ----------------------------------
        bot_frame = ttk.Frame(self.paned)
        self.fig_bot = Figure(figsize=(11, 2.0), tight_layout=True)
        self.ax_bot = self.fig_bot.add_subplot(111)
        self.ax_bot.set_ylabel("interference")
        self.ax_bot.set_xlabel("time (us)")
        (self.ln_int,) = self.ax_bot.plot([], [], color="tab:purple", lw=0.8)
        self.ax_bot.grid(True, alpha=0.3)
        self.canvas_bot = FigureCanvasTkAgg(self.fig_bot, master=bot_frame)
        self.canvas_bot.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.paned.add(bot_frame, weight=1)               # ~20%

        # keep the two x-axes in sync (separate figures -> link manually)
        self._syncing_x = False
        self.ax_top.callbacks.connect("xlim_changed", self._on_top_xlim)

        # box-zoom on the top axis; armed/disarmed by the "Box zoom" toggle
        self.selector = RectangleSelector(
            self.ax_top, self.on_box_zoom, useblit=True, button=[1],
            interactive=False, spancoords="data",
            props=dict(facecolor="tab:orange", edgecolor="tab:orange",
                       fill=True, alpha=0.15, linestyle="--", linewidth=1.0),
        )
        self.selector.set_active(False)

        # set the initial 80/20 split once the panes have a real size
        self.after(200, self._init_sash)

    def _init_sash(self):
        try:
            h = self.paned.winfo_height()
            if h > 1:
                self.paned.sashpos(0, int(h * 0.75))
        except Exception:
            pass

    def _draw_all(self):
        self.canvas_top.draw_idle()
        self.canvas_bot.draw_idle()

    def _on_top_xlim(self, ax):
        """Mirror the top panel's x-range onto the bottom panel."""
        if self._syncing_x:
            return
        self._syncing_x = True
        try:
            self.ax_bot.set_xlim(ax.get_xlim())
            self.canvas_bot.draw_idle()
        finally:
            self._syncing_x = False

    # ==================================================================
    # view / zoom controls
    # ==================================================================
    def toggle_zoom(self):
        """Arm/disarm the drag-box zoom on the overlay panel."""
        self.selector.set_active(self.zoom_on.get())

    def zoom_y(self, factor):
        """Zoom the overlay's amplitude (y) axis about its centre.

        factor > 1 zooms *out* (wider y-range -> tall signal fits);
        factor < 1 zooms *in* (narrower y-range -> signal grows taller).
        """
        lo, hi = self.ax_top.get_ylim()
        mid = 0.5 * (lo + hi)
        half = 0.5 * (hi - lo) * factor
        if half <= 0:
            return
        self.ax_top.set_ylim(mid - half, mid + half)
        self.canvas_top.draw_idle()

    # ==================================================================
    # data loading / filtering
    # ==================================================================
    def load_csv(self):
        path = filedialog.askopenfilename(
            title="Load waveform CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if path:
            self._do_load(path)

    def _do_load(self, path):
        try:
            t, a = load_waveform(path)
        except Exception as exc:
            messagebox.showerror("Load error", str(exc))
            return
        self.t_us, self.raw = t, a
        self.dt_us = float(np.median(np.diff(t))) if t.size > 1 else FINE_STEP
        self.filename = os.path.basename(path)
        self.lbl_file.config(text=self.filename)
        # default fine step to the true sample interval
        self.fine_step.set(round(self.dt_us, 6))
        self.apply_filter()   # sets self.signal + full redraw + autoscale
        self.reset_view()

    def apply_filter(self):
        if self.raw is None:
            return
        if not self.filter_on.get():
            self.signal = self.raw.copy()
        else:
            try:
                fs = 1.0 / self.dt_us                # sampling freq in MHz
                nyq = fs / 2.0
                fc, bw = float(self.fc.get()), float(self.bw.get())
                f_lo, f_hi = fc - bw / 2.0, fc + bw / 2.0
                if not (0 < f_lo < f_hi < nyq):
                    raise ValueError(
                        f"Band {f_lo:.3g}-{f_hi:.3g} MHz outside (0, {nyq:.3g}) MHz.")
                b, a = butter(int(self.order.get()), [f_lo, f_hi],
                              btype="band", fs=fs)
                self.signal = filtfilt(b, a, self.raw)
            except Exception as exc:
                messagebox.showerror("Filter error", str(exc))
                self.filter_on.set(False)
                self.signal = self.raw.copy()
        self.redraw(autoscale=True)

    def _set_fc(self, value):
        self.fc.set(value)
        self.filter_on.set(True)
        self.apply_filter()

    # ==================================================================
    # shift / amplitude actions
    # ==================================================================
    def nudge(self, delta):
        self.set_shift(self.shift.get() + delta)

    def set_shift(self, value):
        self.shift.set(round(float(value), 6))
        self.redraw()

    def scale_amp(self, factor, which="copy"):
        var = self.yscale_orig if which == "orig" else self.yscale
        var.set(round(var.get() * factor, 4))
        self.redraw()

    # ==================================================================
    # drawing
    # ==================================================================
    def _copy_on_grid(self):
        """Copy resampled onto the original time grid (for interference)."""
        ys = self.yscale.get()
        sh = self.shift.get()
        return ys * np.interp(self.t_us, self.t_us + sh, self.signal,
                              left=0.0, right=0.0)

    def redraw(self, autoscale=False):
        if self.signal is None:
            return
        sh = self.shift.get()
        ys = self.yscale.get()
        yo = self.yscale_orig.get()

        # top panel: original + shifted/scaled copy (cheap x-shift)
        self.ln_orig.set_data(self.t_us, yo * self.signal)
        self.ln_copy.set_data(self.t_us + sh, ys * self.signal)

        # bottom panel: interference
        orig_g = yo * self.signal
        copy_g = self._copy_on_grid()
        if self.interf_mode.get() == "sum":
            interf = orig_g + copy_g
            self.ax_bot.set_ylabel("original + copy")
        else:
            interf = orig_g - copy_g
            self.ax_bot.set_ylabel("original - copy")
        self.ln_int.set_data(self.t_us, interf)

        if autoscale:
            self.ax_top.relim(); self.ax_top.autoscale()
            self.ax_bot.relim(); self.ax_bot.autoscale()

        self.readout.config(
            text=f"Travel time: {sh:.4f} us   |   orig amp x{yo:.3f}"
                 f"   copy amp x{ys:.3f}   |   {self.filename}")
        self._draw_all()

    def reset_view(self):
        if self.t_us is None:
            return
        self.ax_top.set_xlim(self.t_us.min(), self.t_us.max())
        self.ax_top.relim(); self.ax_top.autoscale(axis="y")
        self.ax_bot.relim(); self.ax_bot.autoscale(axis="y")
        self._draw_all()

    # ==================================================================
    # box zoom (click-drag on top panel)
    # ==================================================================
    def on_box_zoom(self, eclick, erelease):
        if None in (eclick.xdata, erelease.xdata, eclick.ydata, erelease.ydata):
            return
        x0, x1 = sorted((eclick.xdata, erelease.xdata))
        y0, y1 = sorted((eclick.ydata, erelease.ydata))
        if x1 - x0 <= 0 or y1 - y0 <= 0:
            return
        self.ax_top.set_xlim(x0, x1)       # xlim_changed -> bottom x follows
        self.ax_top.set_ylim(y0, y1)
        self.ax_bot.relim(); self.ax_bot.autoscale(axis="y")
        self._draw_all()

    # ==================================================================
    # records
    # ==================================================================
    def record(self):
        fc = self.fc.get() if self.filter_on.get() else float("nan")
        tt = self.shift.get()
        self.records.append((fc, tt))
        tag = f"{fc:.0f} MHz" if self.filter_on.get() else "raw"
        self.rec_list.insert(tk.END, f"{tag}:  {tt:.4f} us")

    def export(self):
        if not self.records:
            messagebox.showinfo("Export", "No records to export yet.")
            return
        path = filedialog.asksaveasfilename(
            title="Export travel-time records",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="peo_travel_times.csv",
        )
        if not path:
            return
        df = pd.DataFrame(self.records, columns=["center_freq_MHz", "travel_time_us"])
        df["source_file"] = self.filename
        df.to_csv(path, index=False)
        messagebox.showinfo("Export", f"Wrote {len(df)} record(s) to\n{path}")

    def export_view(self):
        """Write the samples inside the current (zoomed) time window to Excel.

        Use this to grab the focused echo region so it can be re-plotted for a
        paper figure (amplitude vs time, e.g. the P- and S-wave echo trains).
        The visible x-range of the top panel defines the exported window; the
        shared x-axis means the bottom (interference) panel covers the same span.
        """
        if self.signal is None:
            messagebox.showinfo("Export view", "Load a waveform first.")
            return

        x0, x1 = self.ax_top.get_xlim()
        lo, hi = (x0, x1) if x0 <= x1 else (x1, x0)
        mask = (self.t_us >= lo) & (self.t_us <= hi)
        if not np.any(mask):
            messagebox.showinfo("Export view", "No samples in the current view.")
            return

        filt_on = self.filter_on.get()
        fc = float(self.fc.get()) if filt_on else float("nan")
        # label common P/S centre frequencies for a friendly default filename
        if filt_on and abs(fc - 50) < 1e-6:
            wave = "P"
        elif filt_on and abs(fc - 30) < 1e-6:
            wave = "S"
        else:
            wave = "raw"
        stem = os.path.splitext(self.filename)[0] if self.filename != "(none)" else "waveform"

        path = filedialog.asksaveasfilename(
            title="Export visible waveform to Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=f"{stem}_{wave}_view.xlsx",
        )
        if not path:
            return

        t = self.t_us[mask]
        wave_df = pd.DataFrame({
            "time_us": t,
            "raw_amplitude": self.raw[mask],
        })
        if filt_on:
            wave_df[f"filtered_amplitude_{fc:g}MHz"] = self.signal[mask]

        info_df = pd.DataFrame({
            "field": [
                "source_file", "filter_enabled", "center_freq_MHz",
                "bandwidth_MHz", "filter_order", "window_start_us",
                "window_end_us", "n_samples", "sample_interval_us",
            ],
            "value": [
                self.filename, filt_on, fc,
                float(self.bw.get()) if filt_on else float("nan"),
                int(self.order.get()) if filt_on else float("nan"),
                float(t.min()), float(t.max()), int(t.size), self.dt_us,
            ],
        })

        try:
            with pd.ExcelWriter(path, engine="openpyxl") as xl:
                wave_df.to_excel(xl, sheet_name="waveform", index=False)
                info_df.to_excel(xl, sheet_name="info", index=False)
        except Exception as exc:
            messagebox.showerror("Export view", str(exc))
            return

        messagebox.showinfo(
            "Export view",
            f"Wrote {t.size} samples ({lo:.4f}-{hi:.4f} us) to\n{path}")


if __name__ == "__main__":
    PEOApp().mainloop()
