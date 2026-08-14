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
        ttk.Button(f_int, text="Reset view",
                   command=self.reset_view).pack(fill=tk.X, pady=(2, 0))

        # --- record ---
        f_rec = ttk.LabelFrame(bar, text="Travel-time records", padding=4)
        f_rec.pack(side=tk.LEFT, fill=tk.Y, padx=3)
        ttk.Button(f_rec, text="Record", command=self.record).pack(fill=tk.X)
        ttk.Button(f_rec, text="Export...", command=self.export).pack(fill=tk.X)
        self.rec_list = tk.Listbox(f_rec, height=4, width=20)
        self.rec_list.pack()

        # --- readout ---
        self.readout = ttk.Label(self, text="Travel time: -- us",
                                 font=("Segoe UI", 16, "bold"), anchor="w")
        self.readout.pack(side=tk.TOP, fill=tk.X, padx=8, pady=2)

    def _build_plots(self):
        self.fig = Figure(figsize=(11, 6.4), tight_layout=True)
        self.ax_top = self.fig.add_subplot(211)
        self.ax_bot = self.fig.add_subplot(212, sharex=self.ax_top)

        self.ax_top.set_ylabel("amplitude")
        self.ax_top.set_title("PEO overlay - original (blue) + movable copy (red)")
        self.ax_bot.set_ylabel("interference")
        self.ax_bot.set_xlabel("time (us)")

        (self.ln_orig,) = self.ax_top.plot([], [], color="tab:blue", lw=0.8,
                                           label="original")
        (self.ln_copy,) = self.ax_top.plot([], [], color="tab:red", lw=0.8,
                                           alpha=0.8, label="copy (shifted)")
        (self.ln_int,)  = self.ax_bot.plot([], [], color="tab:purple", lw=0.8)
        self.ax_top.legend(loc="upper right", fontsize=8)
        self.ax_top.grid(True, alpha=0.3)
        self.ax_bot.grid(True, alpha=0.3)

        canvas_frame = ttk.Frame(self)
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.canvas = FigureCanvasTkAgg(self.fig, master=canvas_frame)
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        NavigationToolbar2Tk(self.canvas, canvas_frame)  # pan / zoom / home

        # click-drag a box on the top axis to zoom (no visible selection box)
        self.selector = RectangleSelector(
            self.ax_top, self.on_box_zoom, useblit=True, button=[1],
            interactive=False, spancoords="data",
            props=dict(facecolor="none", edgecolor="none", fill=False, alpha=0.0),
        )

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
        self.canvas.draw_idle()

    def reset_view(self):
        if self.t_us is None:
            return
        self.ax_top.set_xlim(self.t_us.min(), self.t_us.max())
        self.ax_top.relim(); self.ax_top.autoscale(axis="y")
        self.ax_bot.relim(); self.ax_bot.autoscale(axis="y")
        self.canvas.draw_idle()

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
        self.ax_top.set_xlim(x0, x1)       # shared x -> bottom panel follows
        self.ax_top.set_ylim(y0, y1)
        self.ax_bot.relim(); self.ax_bot.autoscale(axis="y")
        self.canvas.draw_idle()

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


if __name__ == "__main__":
    PEOApp().mainloop()
