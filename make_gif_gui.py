"""
GUI front-end for building an animated GIF from a sequence of images.

Run with:  py make_gif_gui.py

The window starts blank except for the controls. Load files (or a whole
folder), tweak the parameters, and click "Create GIF". Every parameter that
used to live in the CONFIG block of make_gif.py is exposed here.
"""

import os
import queue
import re
import threading

import tkinter as tk
from tkinter import colorchooser, filedialog, messagebox, ttk

import numpy as np
from PIL import Image, ImageColor, ImageDraw, ImageFont, ImageSequence, ImageTk

# Image types offered in the file/folder dialogs.
IMAGE_EXTS = (".tif", ".tiff", ".png", ".jpg", ".jpeg", ".bmp", ".gif")

# Trailing "_NNN" counter (any digit count), stripped when building a label.
TRAILING_COUNTER = re.compile(r"_(\d+)$")


# ---------------------------------------------------------------------------
# Frame-building logic (parameterised versions of the make_gif.py functions)
# ---------------------------------------------------------------------------
def label_for(filename, replace_underscores):
    label = os.path.splitext(filename)[0]
    label = TRAILING_COUNTER.sub("", label)
    if replace_underscores:
        label = label.replace("_", " ")
    return label


def load_font(size):
    try:
        return ImageFont.truetype("arial.ttf", size)
    except OSError:
        return ImageFont.load_default()


def darkness_centroid_shift(image, threshold_pct):
    """Return (dx, dy) that moves the dark region's centroid to the frame center.

    Pixels in the darkest ``threshold_pct`` percent are treated as the feature;
    their darkness-weighted centroid is compared against the geometric center.
    """
    arr = np.asarray(image.convert("L"), dtype=np.float64)
    h, w = arr.shape
    thr = np.percentile(arr, threshold_pct)
    mask = arr <= thr
    if not mask.any():
        return 0.0, 0.0
    ys, xs = np.nonzero(mask)
    weights = (thr - arr[mask]) + 1e-6  # darker pixels count more
    cx = np.average(xs, weights=weights)
    cy = np.average(ys, weights=weights)
    return (w - 1) / 2.0 - cx, (h - 1) / 2.0 - cy


def median_fill(image, mode):
    """A background fill value (from the frame median) for exposed edges."""
    arr = np.asarray(image)
    if mode == "L":
        return int(np.median(arr))
    return tuple(int(np.median(arr[..., c])) for c in range(3))


def build_frame(path, font, opts):
    mode = "L" if opts["grayscale"] else "RGB"
    image = Image.open(path).convert(mode)

    if opts.get("stabilize"):
        dx, dy = darkness_centroid_shift(image, opts["dark_threshold"])
        image = image.transform(
            image.size, Image.AFFINE, (1, 0, -dx, 0, 1, -dy),
            resample=Image.BILINEAR, fillcolor=median_fill(image, mode))

    if opts["scale"] != 1.0:
        new_size = (round(image.width * opts["scale"]), round(image.height * opts["scale"]))
        image = image.resize(new_size, Image.LANCZOS)

    if not opts["show_title"]:
        return image.convert("P", palette=Image.ADAPTIVE)

    bg = ImageColor.getcolor(opts["bg_color"], mode)
    fg = ImageColor.getcolor(opts["text_color"], mode)

    label_h = opts["label_height"]
    canvas = Image.new(mode, (image.width, image.height + label_h), color=bg)
    canvas.paste(image, (0, label_h))

    draw = ImageDraw.Draw(canvas)
    label = label_for(os.path.basename(path), opts["replace_underscores"])
    bbox = draw.textbbox((0, 0), label, font=font)
    text_x = (canvas.width - (bbox[2] - bbox[0])) // 2
    text_y = (label_h - (bbox[3] - bbox[1])) // 2 - bbox[1]
    draw.text((text_x, text_y), label, font=font, fill=fg)

    return canvas.convert("P", palette=Image.ADAPTIVE)


def natural_key(name):
    """Split a name into text/number chunks so 'x2' sorts before 'x10'."""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r"(\d+)", name)]


def frame_durations(count, opts):
    """Per-frame durations: the default everywhere, longer for the trailing frames."""
    durations = [opts["duration"]] * count
    if opts.get("slow_tail") and opts["slow_tail_count"] > 0:
        for i in range(max(0, count - opts["slow_tail_count"]), count):
            durations[i] = opts["slow_tail_ms"]
    return durations


# ---------------------------------------------------------------------------
# A small canvas-drawn on/off toggle switch (slides left/right).
# ---------------------------------------------------------------------------
class ToggleSwitch(tk.Canvas):
    def __init__(self, master, variable, command=None, width=52, height=26, **kw):
        super().__init__(master, width=width, height=height,
                         highlightthickness=0, bd=0, **kw)
        self.var = variable
        self.command = command
        self._sw, self._sh = width, height
        self.bind("<Button-1>", self._on_click)
        self.var.trace_add("write", lambda *_: self._draw())
        self._draw()

    def _on_click(self, _event):
        if str(self["state"]) == "disabled":
            return
        self.var.set(not self.var.get())
        if self.command:
            self.command()

    def _draw(self):
        self.delete("all")
        on = bool(self.var.get())
        pad = 3
        w, h = self._sw, self._sh
        cap = h - 2 * pad  # diameter of the rounded pill ends
        track = "#3f9c4b" if on else "#b9b9b9"
        # Pill background: two round caps joined by a rectangle.
        self.create_oval(pad, pad, pad + cap, h - pad, fill=track, outline=track)
        self.create_oval(w - pad - cap, pad, w - pad, h - pad, fill=track, outline=track)
        self.create_rectangle(pad + cap / 2, pad, w - pad - cap / 2, h - pad,
                              fill=track, outline=track)
        # Knob.
        knob = cap - 2
        kx = (w - pad - 1 - knob) if on else (pad + 1)
        ky = pad + 1
        self.create_oval(kx, ky, kx + knob, ky + knob, fill="white", outline="white")


# ---------------------------------------------------------------------------
# Main application window
# ---------------------------------------------------------------------------
class GifMakerApp:
    def __init__(self, root):
        self.root = root
        root.title("GIF Maker")
        root.minsize(560, 560)

        self.files = []            # ordered list of source image paths
        self.queue = queue.Queue()  # worker -> UI messages

        # Preview animation state
        self._anim_job = None
        self._preview_frames = []
        self._preview_idx = 0
        self._preview_durations = [500]

        pad = {"padx": 8, "pady": 4}
        row = 0

        container = ttk.Frame(root)
        container.pack(fill="both", expand=True)

        main = ttk.Frame(container, padding=12)
        main.pack(side="left", fill="both", expand=True)
        main.columnconfigure(1, weight=1)

        preview = ttk.LabelFrame(container, text="Preview", padding=8)
        preview.pack(side="right", fill="both", expand=True)

        # Top: the single frame for whichever file is clicked in the list.
        self.frame_caption = ttk.Label(preview, anchor="center",
                                       text="(click a file to view its frame)",
                                       justify="center")
        self.frame_caption.pack(fill="x")
        self.frame_view_label = ttk.Label(preview, anchor="center",
                                          justify="center", width=40)
        self.frame_view_label.pack(fill="both", expand=True, pady=(0, 8))

        ttk.Separator(preview, orient="horizontal").pack(fill="x", pady=(0, 8))

        # Bottom: the animated GIF preview.
        ttk.Label(preview, anchor="center", text="GIF").pack(fill="x")
        self.preview_label = ttk.Label(preview, anchor="center",
                                       text="(preview plays here after you\ncreate a GIF)",
                                       justify="center", width=40)
        self.preview_label.pack(fill="both", expand=True)

        # --- Load buttons ----------------------------------------------------
        loadbar = ttk.Frame(main)
        loadbar.grid(row=row, column=0, columnspan=3, sticky="w", **pad)
        ttk.Button(loadbar, text="Load Files…", command=self.load_files).pack(side="left")
        ttk.Button(loadbar, text="Load Folder…", command=self.load_folder).pack(side="left", padx=6)
        ttk.Label(loadbar, text="Order:").pack(side="left", padx=(16, 4))
        self.sort_var = tk.StringVar(value="Modification time")
        sort_box = ttk.Combobox(loadbar, textvariable=self.sort_var, width=18, state="readonly",
                                values=["Modification time", "Name (natural)", "Name (alphabetical)"])
        sort_box.pack(side="left")
        sort_box.bind("<<ComboboxSelected>>", lambda _e: self.apply_sort())
        row += 1

        # --- File list -------------------------------------------------------
        listframe = ttk.LabelFrame(main, text="Frames (in play order)")
        listframe.grid(row=row, column=0, columnspan=3, sticky="nsew", **pad)
        main.rowconfigure(row, weight=1)
        listframe.columnconfigure(0, weight=1)
        listframe.rowconfigure(0, weight=1)
        self.listbox = tk.Listbox(listframe, height=8, activestyle="none",
                                  selectmode="extended")
        self.listbox.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(listframe, orient="vertical", command=self.listbox.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.bind("<Delete>", lambda _e: self.remove_selected())
        self.listbox.bind("<<ListboxSelect>>", self.show_selected_frame)

        listbtns = ttk.Frame(listframe)
        listbtns.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Button(listbtns, text="Remove Selected", command=self.remove_selected).pack(side="left")
        ttk.Button(listbtns, text="Clear All", command=self.clear_all).pack(side="left", padx=6)
        ttk.Label(listbtns, text="(Ctrl/Shift-click to multi-select, or press Delete)").pack(side="left", padx=6)
        row += 1

        # --- Playback parameters --------------------------------------------
        params = ttk.LabelFrame(main, text="Playback")
        params.grid(row=row, column=0, columnspan=3, sticky="ew", **pad)
        params.columnconfigure(1, weight=1)
        params.columnconfigure(3, weight=1)

        self.duration_var = tk.StringVar(value="500")
        self.loop_var = tk.StringVar(value="0")
        self.scale_var = tk.StringVar(value="0.5")

        ttk.Label(params, text="Milliseconds per frame:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(params, textvariable=self.duration_var, width=10).grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(params, text="Loop count (0 = forever):").grid(row=0, column=2, sticky="w", **pad)
        ttk.Entry(params, textvariable=self.loop_var, width=10).grid(row=0, column=3, sticky="w", **pad)

        ttk.Label(params, text="Scale (1.0 = full size):").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(params, textvariable=self.scale_var, width=10).grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(params, text="Grayscale:").grid(row=1, column=2, sticky="w", **pad)
        self.grayscale_var = tk.BooleanVar(value=True)
        ToggleSwitch(params, self.grayscale_var).grid(row=1, column=3, sticky="w", **pad)

        ttk.Label(params, text="Stabilize (center dark spot):").grid(row=2, column=0, sticky="w", **pad)
        self.stabilize_var = tk.BooleanVar(value=False)
        ToggleSwitch(params, self.stabilize_var).grid(row=2, column=1, sticky="w", **pad)
        ttk.Label(params, text="Dark threshold (%):").grid(row=2, column=2, sticky="w", **pad)
        self.dark_thr_var = tk.StringVar(value="20")
        ttk.Entry(params, textvariable=self.dark_thr_var, width=10).grid(row=2, column=3, sticky="w", **pad)

        ttk.Label(params, text="Slow down last frames:").grid(row=3, column=0, sticky="w", **pad)
        self.slow_tail_var = tk.BooleanVar(value=False)
        ToggleSwitch(params, self.slow_tail_var,
                     command=self._sync_slow_tail_state).grid(row=3, column=1, sticky="w", **pad)

        self.tail_count_var = tk.StringVar(value="100")
        self.tail_ms_var = tk.StringVar(value="400")
        self.tail_count_label = ttk.Label(params, text="How many trailing frames:")
        self.tail_count_label.grid(row=4, column=0, sticky="w", **pad)
        self.tail_count_entry = ttk.Entry(params, textvariable=self.tail_count_var, width=10)
        self.tail_count_entry.grid(row=4, column=1, sticky="w", **pad)
        self.tail_ms_label = ttk.Label(params, text="Milliseconds for those:")
        self.tail_ms_label.grid(row=4, column=2, sticky="w", **pad)
        self.tail_ms_entry = ttk.Entry(params, textvariable=self.tail_ms_var, width=10)
        self.tail_ms_entry.grid(row=4, column=3, sticky="w", **pad)
        row += 1

        # --- Title / label parameters ---------------------------------------
        self.title_frame = ttk.LabelFrame(main, text="Title from filename")
        self.title_frame.grid(row=row, column=0, columnspan=3, sticky="ew", **pad)
        self.title_frame.columnconfigure(1, weight=1)
        self.title_frame.columnconfigure(3, weight=1)

        ttk.Label(self.title_frame, text="Show title:").grid(row=0, column=0, sticky="w", **pad)
        self.show_title_var = tk.BooleanVar(value=True)
        ToggleSwitch(self.title_frame, self.show_title_var,
                     command=self._sync_title_state).grid(row=0, column=1, sticky="w", **pad)

        ttk.Label(self.title_frame, text="Replace underscores:").grid(row=0, column=2, sticky="w", **pad)
        self.replace_var = tk.BooleanVar(value=True)
        ToggleSwitch(self.title_frame, self.replace_var).grid(row=0, column=3, sticky="w", **pad)

        self.label_height_var = tk.StringVar(value="60")
        self.font_size_var = tk.StringVar(value="36")
        self.bg_var = tk.StringVar(value="white")
        self.fg_var = tk.StringVar(value="black")

        ttk.Label(self.title_frame, text="Label height (px):").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(self.title_frame, textvariable=self.label_height_var, width=10).grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(self.title_frame, text="Font size:").grid(row=1, column=2, sticky="w", **pad)
        ttk.Entry(self.title_frame, textvariable=self.font_size_var, width=10).grid(row=1, column=3, sticky="w", **pad)

        ttk.Label(self.title_frame, text="Background color:").grid(row=2, column=0, sticky="w", **pad)
        self._color_row(self.title_frame, 2, 1, self.bg_var)
        ttk.Label(self.title_frame, text="Text color:").grid(row=2, column=2, sticky="w", **pad)
        self._color_row(self.title_frame, 2, 3, self.fg_var)
        row += 1

        # --- Output ----------------------------------------------------------
        outbar = ttk.Frame(main)
        outbar.grid(row=row, column=0, columnspan=3, sticky="ew", **pad)
        outbar.columnconfigure(1, weight=1)
        ttk.Label(outbar, text="Output:").grid(row=0, column=0, sticky="w")
        self.output_var = tk.StringVar(value="")
        ttk.Entry(outbar, textvariable=self.output_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(outbar, text="Save As…", command=self.choose_output).grid(row=0, column=2)
        row += 1

        # --- Action + status -------------------------------------------------
        actionbar = ttk.Frame(main)
        actionbar.grid(row=row, column=0, columnspan=3, sticky="ew", **pad)
        actionbar.columnconfigure(1, weight=1)
        self.create_btn = ttk.Button(actionbar, text="Create GIF", command=self.create_gif)
        self.create_btn.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(actionbar, mode="determinate")
        self.progress.grid(row=0, column=1, sticky="ew", padx=8)
        row += 1

        self.status_var = tk.StringVar(value="No files loaded.")
        ttk.Label(main, textvariable=self.status_var, anchor="w").grid(
            row=row, column=0, columnspan=3, sticky="ew", **pad)

        self._sync_title_state()
        self._sync_slow_tail_state()

    # -- small helpers -------------------------------------------------------
    def _color_row(self, parent, r, c, var):
        holder = ttk.Frame(parent)
        holder.grid(row=r, column=c, sticky="w", padx=8, pady=4)
        ttk.Entry(holder, textvariable=var, width=10).pack(side="left")
        ttk.Button(holder, text="…", width=3,
                   command=lambda: self._pick_color(var)).pack(side="left", padx=3)

    def _pick_color(self, var):
        try:
            initial = var.get() or "white"
            rgb, hexval = colorchooser.askcolor(color=initial, parent=self.root)
        except tk.TclError:
            rgb, hexval = colorchooser.askcolor(parent=self.root)
        if hexval:
            var.set(hexval)

    def _sync_slow_tail_state(self):
        state = "normal" if self.slow_tail_var.get() else "disabled"
        for w in (self.tail_count_label, self.tail_count_entry,
                  self.tail_ms_label, self.tail_ms_entry):
            self._set_state(w, state)

    def _sync_title_state(self):
        state = "normal" if self.show_title_var.get() else "disabled"
        for child in self.title_frame.winfo_children():
            # Leave the "Show title:" label and its toggle always active.
            info = child.grid_info()
            if info.get("row") == 0 and info.get("column") in (0, 1):
                continue
            self._set_state(child, state)

    def _set_state(self, widget, state):
        try:
            widget.configure(state=state)
        except tk.TclError:
            pass
        for sub in widget.winfo_children():
            self._set_state(sub, state)

    # -- loading -------------------------------------------------------------
    def load_files(self):
        patterns = " ".join(f"*{e}" for e in IMAGE_EXTS)
        paths = filedialog.askopenfilenames(
            title="Select image files",
            filetypes=[("Images", patterns), ("All files", "*.*")])
        if paths:
            self.files = list(paths)
            self.apply_sort()
            self._default_output(os.path.dirname(self.files[0]))

    def load_folder(self):
        folder = filedialog.askdirectory(title="Select a folder of images")
        if not folder:
            return
        found = [os.path.join(folder, f) for f in os.listdir(folder)
                 if f.lower().endswith(IMAGE_EXTS)]
        if not found:
            messagebox.showwarning("No images", "No supported images found in that folder.")
            return
        self.files = found
        self.apply_sort()
        self._default_output(folder)

    def _default_output(self, folder):
        if not self.output_var.get():
            self.output_var.set(os.path.join(folder, "sequence.gif"))

    def apply_sort(self):
        mode = self.sort_var.get()
        if mode == "Modification time":
            self.files.sort(key=lambda f: (os.path.getmtime(f), os.path.basename(f)))
        elif mode == "Name (natural)":
            self.files.sort(key=lambda f: natural_key(os.path.basename(f)))
        else:
            self.files.sort(key=lambda f: os.path.basename(f).lower())
        self._refresh_list()

    def _refresh_list(self):
        self.listbox.delete(0, "end")
        for i, f in enumerate(self.files, 1):
            self.listbox.insert("end", f"{i:>3}.  {os.path.basename(f)}")
        self.status_var.set(f"{len(self.files)} file(s) loaded.")

    def remove_selected(self):
        selected = self.listbox.curselection()
        if not selected:
            return
        top_fraction = self.listbox.yview()[0]  # remember scroll position
        for i in sorted(selected, reverse=True):
            del self.files[i]
        self._refresh_list()
        self.listbox.yview_moveto(top_fraction)  # keep the view where it was
        if self.files:
            nxt = min(selected[0], len(self.files) - 1)
            self.listbox.selection_set(nxt)
            self.listbox.activate(nxt)

    def clear_all(self):
        self.files = []
        self._refresh_list()
        self.frame_view_label.config(image="", text="")
        self.frame_view_label.image = None
        self.frame_caption.config(text="(click a file to view its frame)")

    def choose_output(self):
        path = filedialog.asksaveasfilename(
            title="Save GIF as", defaultextension=".gif",
            filetypes=[("GIF", "*.gif")],
            initialfile=os.path.basename(self.output_var.get()) or "sequence.gif")
        if path:
            self.output_var.set(path)

    # -- creating ------------------------------------------------------------
    def _gather_opts(self):
        opts = {
            "duration": int(float(self.duration_var.get())),
            "slow_tail": self.slow_tail_var.get(),
            "slow_tail_count": int(float(self.tail_count_var.get())),
            "slow_tail_ms": int(float(self.tail_ms_var.get())),
            "loop": int(float(self.loop_var.get())),
            "scale": float(self.scale_var.get()),
            "grayscale": self.grayscale_var.get(),
            "stabilize": self.stabilize_var.get(),
            "dark_threshold": float(self.dark_thr_var.get()),
            "show_title": self.show_title_var.get(),
            "replace_underscores": self.replace_var.get(),
            "label_height": int(float(self.label_height_var.get())),
            "font_size": int(float(self.font_size_var.get())),
            "bg_color": self.bg_var.get().strip() or "white",
            "text_color": self.fg_var.get().strip() or "black",
            "output": self.output_var.get().strip(),
        }
        if opts["duration"] <= 0:
            raise ValueError("Milliseconds per frame must be greater than 0.")
        if opts["slow_tail"]:
            if opts["slow_tail_count"] <= 0:
                raise ValueError("Number of trailing frames to slow down must be greater than 0.")
            if opts["slow_tail_ms"] <= 0:
                raise ValueError("Milliseconds for the trailing frames must be greater than 0.")
        if opts["scale"] <= 0:
            raise ValueError("Scale must be greater than 0.")
        if opts["stabilize"] and not 0 < opts["dark_threshold"] < 100:
            raise ValueError("Dark threshold must be between 0 and 100.")
        if not opts["output"]:
            raise ValueError("Choose an output path.")
        # Validate colors early so failures surface before the long render.
        mode = "L" if opts["grayscale"] else "RGB"
        if opts["show_title"]:
            ImageColor.getcolor(opts["bg_color"], mode)
            ImageColor.getcolor(opts["text_color"], mode)
        return opts

    def create_gif(self):
        if not self.files:
            messagebox.showinfo("No files", "Load some images first.")
            return
        try:
            opts = self._gather_opts()
        except ValueError as exc:
            messagebox.showerror("Invalid setting", str(exc))
            return

        self.stop_preview()
        self.create_btn.config(state="disabled")
        self.progress.config(maximum=len(self.files), value=0)
        self.status_var.set("Building frames…")

        worker = threading.Thread(target=self._render, args=(list(self.files), opts), daemon=True)
        worker.start()
        self.root.after(100, self._poll_queue)

    def _render(self, files, opts):
        try:
            font = load_font(opts["font_size"])
            frames = []
            for i, path in enumerate(files, 1):
                frames.append(build_frame(path, font, opts))
                self.queue.put(("progress", i))
            durations = frame_durations(len(frames), opts)
            frames[0].save(
                opts["output"], save_all=True, append_images=frames[1:],
                duration=durations, loop=opts["loop"], optimize=True)
            self.queue.put(("done", (opts["output"], len(frames), durations)))
        except Exception as exc:  # surface any PIL/IO error to the UI
            self.queue.put(("error", str(exc)))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "progress":
                    self.progress.config(value=payload)
                    self.status_var.set(f"Building frame {payload}/{len(self.files)}…")
                elif kind == "done":
                    path, n, durations = payload
                    total = sum(durations) / 1000
                    self.progress.config(value=self.progress["maximum"])
                    self.status_var.set(f"Wrote {n} frames to {path}  ({total:.1f}s total)")
                    self.create_btn.config(state="normal")
                    self.start_preview(path, durations)
                    return
                elif kind == "error":
                    self.status_var.set("Failed.")
                    self.create_btn.config(state="normal")
                    messagebox.showerror("Error", payload)
                    return
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # -- single-frame viewer -------------------------------------------------
    def show_selected_frame(self, _event=None, max_dim=460):
        """Render the frame for the file clicked in the list, above the GIF."""
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[-1]  # the frame the user most recently clicked
        if idx >= len(self.files):
            return
        path = self.files[idx]
        try:
            frame = self._render_single_frame(path, max_dim)
        except Exception as exc:
            self.frame_view_label.config(text=f"(could not show frame:\n{exc})", image="")
            self.frame_view_label.image = None
            self.frame_caption.config(text=f"Frame {idx + 1}: {os.path.basename(path)}")
            return
        photo = ImageTk.PhotoImage(frame)
        self.frame_view_label.config(image=photo, text="")
        self.frame_view_label.image = photo  # keep a reference alive
        self.frame_caption.config(text=f"Frame {idx + 1}: {os.path.basename(path)}")

    def _render_single_frame(self, path, max_dim=460):
        """Build the frame the way it will appear in the GIF (best effort).

        Falls back to the raw image if the current settings are incomplete
        (e.g. no output path chosen yet), so the viewer always shows something.
        """
        try:
            opts = self._gather_opts()
            font = load_font(opts["font_size"])
            frame = build_frame(path, font, opts).convert("RGB")
        except Exception:
            frame = Image.open(path).convert("RGB")
        scale = min(max_dim / frame.width, max_dim / frame.height, 1.0)
        if scale < 1.0:
            frame = frame.resize((max(1, round(frame.width * scale)),
                                  max(1, round(frame.height * scale))), Image.LANCZOS)
        return frame

    # -- preview -------------------------------------------------------------
    def start_preview(self, path, durations, max_dim=460):
        """Load the freshly-written GIF and loop it in the preview pane.

        ``durations`` may be a single int or a per-frame list (variable rate).
        """
        self.stop_preview()
        try:
            gif = Image.open(path)
            frames = []
            for frame in ImageSequence.Iterator(gif):
                fr = frame.convert("RGB")
                scale = min(max_dim / fr.width, max_dim / fr.height, 1.0)
                if scale < 1.0:
                    fr = fr.resize((max(1, round(fr.width * scale)),
                                    max(1, round(fr.height * scale))), Image.LANCZOS)
                frames.append(ImageTk.PhotoImage(fr))
            self._preview_frames = frames
        except Exception as exc:
            self.preview_label.config(text=f"(could not preview:\n{exc})", image="")
            return
        if not self._preview_frames:
            return
        if isinstance(durations, (list, tuple)):
            self._preview_durations = [max(20, int(d)) for d in durations]
        else:
            self._preview_durations = [max(20, int(durations))] * len(self._preview_frames)
        self._preview_idx = 0
        self._animate()

    def _animate(self):
        if not self._preview_frames:
            return
        self.preview_label.config(image=self._preview_frames[self._preview_idx], text="")
        delay = self._preview_durations[self._preview_idx % len(self._preview_durations)]
        self._preview_idx = (self._preview_idx + 1) % len(self._preview_frames)
        self._anim_job = self.root.after(delay, self._animate)

    def stop_preview(self):
        if self._anim_job is not None:
            self.root.after_cancel(self._anim_job)
            self._anim_job = None


def main():
    root = tk.Tk()
    GifMakerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
