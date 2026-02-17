"""
PHASE 1: INTELLIGENT SCREENSHOT EXTRACTOR (VIDEO -> OCR OPTIMIZED)
===================================================================
Scrollable left panel - works on any screen size.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import cv2
import os
from pathlib import Path
from PIL import Image, ImageTk
import threading
from datetime import timedelta
import numpy as np


# ── colours ──────────────────────────────────────────────
C_HEADER  = "#2c3e50"
C_BG      = "#f0f2f5"
C_PANEL   = "#ffffff"
C_ACCENT  = "#27ae60"
C_ACCENT2 = "#2980b9"
C_ORANGE  = "#e67e22"
C_TEXT    = "#2c3e50"
C_GRAY    = "#7f8c8d"
C_BORDER  = "#dce1e7"


def make_scrollable_frame(parent):
    """
    Returns (outer_frame, inner_frame).
    Pack outer_frame into the layout.  Put widgets inside inner_frame.
    Mousewheel scrolling works automatically.
    """
    outer = tk.Frame(parent, bg=C_BG)

    canvas = tk.Canvas(outer, bg=C_BG, highlightthickness=0, bd=0)
    scrollbar = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    inner = tk.Frame(canvas, bg=C_BG)
    window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
        # keep inner frame same width as canvas
        canvas.itemconfig(window_id, width=event.width if event.width > 1
                          else canvas.winfo_width())

    inner.bind("<Configure>", lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", _on_configure)

    # Mousewheel
    def _scroll(event):
        if event.delta:
            canvas.yview_scroll(int(-event.delta / 120), "units")
        elif event.num == 4:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")

    canvas.bind_all("<MouseWheel>", _scroll)
    canvas.bind_all("<Button-4>",   _scroll)
    canvas.bind_all("<Button-5>",   _scroll)

    return outer, inner


class VideoScreenshotExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("Phase 1: Video Screenshot Extractor  -  OCR Optimized")
        self.root.minsize(900, 600)

        # Maximise on start (works on Windows + Linux)
        try:
            self.root.state("zoomed")
        except Exception:
            self.root.attributes("-zoomed", True)

        # Variables
        self.video_path            = tk.StringVar()
        self.output_folder         = tk.StringVar()
        self.extraction_mode       = tk.StringVar(value="interval")
        self.interval_seconds      = tk.DoubleVar(value=0.5)
        self.frame_skip            = tk.IntVar(value=15)
        self.apply_ocr_enhancement = tk.BooleanVar(value=True)
        self.upscale_factor        = tk.DoubleVar(value=3.0)
        self.sharpen_amount        = tk.DoubleVar(value=2.5)
        self.clahe_clip            = tk.DoubleVar(value=3.0)
        self.darken_text           = tk.BooleanVar(value=True)

        self.video_info = {}
        self._tk_img    = None

        self._build_ui()

    # ─────────────────────────── UI ───────────────────────────

    def _build_ui(self):
        # ── TOP HEADER ──────────────────────────────────────
        hdr = tk.Frame(self.root, bg=C_HEADER)
        hdr.pack(fill=tk.X, side=tk.TOP)
        tk.Label(hdr, text="PHASE 1  \u2014  VIDEO \u2192 OCR SCREENSHOT EXTRACTOR",
                 font=("Segoe UI", 15, "bold"),
                 bg=C_HEADER, fg="white", pady=12, padx=20).pack(side=tk.LEFT)

        # ── BOTTOM STATUS BAR ────────────────────────────────
        bot = tk.Frame(self.root, bg=C_HEADER, pady=3)
        bot.pack(fill=tk.X, side=tk.BOTTOM)
        self.progress = ttk.Progressbar(bot, orient=tk.HORIZONTAL,
                                         mode="determinate", length=300)
        self.progress.pack(side=tk.LEFT, padx=10, pady=3)
        self.lbl_status = tk.Label(bot, text="Ready",
                                    bg=C_HEADER, fg="white",
                                    font=("Segoe UI", 9), padx=10)
        self.lbl_status.pack(side=tk.LEFT)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("TProgressbar", troughcolor="#3d5166",
                        background=C_ACCENT, thickness=14)

        # ── MAIN BODY ────────────────────────────────────────
        body = tk.Frame(self.root, bg=C_BG)
        body.pack(fill=tk.BOTH, expand=True)

        # ── LEFT SCROLLABLE PANEL ──
        scroll_outer, left = make_scrollable_frame(body)
        scroll_outer.pack(side=tk.LEFT, fill=tk.Y)
        scroll_outer.configure(width=340)
        scroll_outer.pack_propagate(False)

        # ── RIGHT PREVIEW ──
        right = tk.Frame(body, bg="#1a1a2e")
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        tk.Label(right, text="Live Preview  (Enhanced Output)",
                 font=("Segoe UI", 10, "bold"),
                 bg="#1a1a2e", fg=C_ACCENT, pady=6).pack(anchor="w", padx=12)

        self.canvas_preview = tk.Canvas(right, bg="#0d0d1a",
                                         highlightthickness=0)
        self.canvas_preview.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        # ── LEFT CONTENT ─────────────────────────────────────
        pad = {"padx": 12, "pady": 6}

        # Section helper
        def section(title, color=C_ACCENT2):
            f = tk.LabelFrame(left, text=f"  {title}  ",
                               font=("Segoe UI", 9, "bold"),
                               fg=color, bg=C_BG,
                               relief=tk.GROOVE, bd=1,
                               padx=10, pady=8)
            f.pack(fill=tk.X, padx=10, pady=(8, 0))
            return f

        def btn(parent, text, color, cmd):
            return tk.Button(parent, text=text,
                              font=("Segoe UI", 9, "bold"),
                              bg=color, fg="white",
                              activebackground="#1a252f",
                              activeforeground="white",
                              relief=tk.FLAT, bd=0,
                              pady=6, cursor="hand2",
                              command=cmd)

        def entry_browse(parent, var, browse_cmd, lbl_text):
            tk.Label(parent, text=lbl_text, bg=C_BG,
                     fg=C_GRAY, font=("Segoe UI", 8)).pack(anchor="w")
            row = tk.Frame(parent, bg=C_BG)
            row.pack(fill=tk.X, pady=(2, 0))
            tk.Entry(row, textvariable=var,
                     font=("Segoe UI", 9), relief=tk.SOLID, bd=1,
                     bg="white", fg=C_TEXT).pack(side=tk.LEFT,
                                                  fill=tk.X, expand=True)
            btn(row, "Browse", C_ACCENT2, browse_cmd).pack(
                side=tk.RIGHT, padx=(4, 0))

        def slider(parent, label, var, frm, to_, res, note=""):
            lf = tk.Frame(parent, bg=C_BG)
            lf.pack(fill=tk.X, pady=(4, 0))
            tk.Label(lf, text=label, bg=C_BG, fg=C_TEXT,
                     font=("Segoe UI", 8, "bold")).pack(anchor="w")
            if note:
                tk.Label(lf, text=note, bg=C_BG, fg=C_GRAY,
                         font=("Segoe UI", 7)).pack(anchor="w")
            tk.Scale(lf, variable=var, from_=frm, to=to_,
                     resolution=res, orient=tk.HORIZONTAL,
                     bg=C_BG, fg=C_TEXT, troughcolor=C_BORDER,
                     highlightthickness=0,
                     activebackground=C_ACCENT2).pack(fill=tk.X)

        # 1 — Select Video
        s1 = section("1.  Select Video", C_ACCENT2)
        entry_browse(s1, self.video_path, self.browse_video, "Video file path:")
        self.lbl_info = tk.Label(s1, text="No video selected",
                                  bg=C_BG, fg=C_GRAY,
                                  font=("Segoe UI", 8),
                                  justify=tk.LEFT, wraplength=280)
        self.lbl_info.pack(anchor="w", pady=(4, 0))

        # 2 — Output Folder
        s2 = section("2.  Output Folder", C_ACCENT2)
        entry_browse(s2, self.output_folder, self.browse_output, "Save screenshots to:")

        # 3 — Extraction Settings
        s3 = section("3.  Extraction Settings", C_ORANGE)

        tk.Checkbutton(s3, text="Apply OCR Enhancement  (Recommended)",
                       variable=self.apply_ocr_enhancement,
                       font=("Segoe UI", 9, "bold"),
                       fg=C_ORANGE, bg=C_BG,
                       selectcolor="white",
                       activebackground=C_BG).pack(anchor="w", pady=(0, 6))

        mode_f = tk.Frame(s3, bg=C_BG)
        mode_f.pack(fill=tk.X)
        for lbl, val in [("Time Interval (seconds)", "interval"),
                          ("Frame Skip", "frames")]:
            tk.Radiobutton(mode_f, text=lbl, variable=self.extraction_mode,
                           value=val, bg=C_BG, fg=C_TEXT,
                           selectcolor="white",
                           font=("Segoe UI", 9),
                           activebackground=C_BG).pack(anchor="w")

        slider(s3, "Interval (seconds):", self.interval_seconds,
               0.1, 10.0, 0.1, "How often to capture a frame")
        slider(s3, "Frame Skip:", self.frame_skip,
               1, 120, 1, "Capture every N frames")

        # 4 — Enhancement Controls
        s4 = section("4.  Enhancement Controls", "#8e44ad")

        slider(s4, "Upscale Factor:", self.upscale_factor,
               1.5, 5.0, 0.5, "3–5x recommended for small UI fonts")
        slider(s4, "Sharpen Strength:", self.sharpen_amount,
               0.5, 5.0, 0.5, "2–3x makes text bolder and crisper")
        slider(s4, "CLAHE Contrast Clip:", self.clahe_clip,
               1.0, 8.0, 0.5, "3–4 recommended")

        tk.Checkbutton(s4, text="Extra Gamma Darkening  (darker text on light BG)",
                       variable=self.darken_text,
                       font=("Segoe UI", 8), fg="#8e44ad", bg=C_BG,
                       selectcolor="white",
                       activebackground=C_BG).pack(anchor="w", pady=(6, 0))

        # 5 — Start
        s5 = section("5.  Run Extraction", C_ACCENT)
        self.btn_extract = btn(s5, "START EXTRACTION", C_ACCENT,
                                self.start_extraction)
        self.btn_extract.pack(fill=tk.X)

        # spacer at bottom so last section doesn't hug the edge
        tk.Frame(left, bg=C_BG, height=20).pack()

    # ─────────────────────────── VIDEO ───────────────────────────

    def browse_video(self):
        f = filedialog.askopenfilename(
            title="Select video file",
            filetypes=[("Video files", "*.mp4 *.avi *.mkv *.mov *.wmv *.flv")])
        if f:
            self.video_path.set(f)
            self.analyze_video(f)
            p = Path(f)
            if not self.output_folder.get():
                self.output_folder.set(str(p.parent / f"{p.stem}_screenshots"))

    def browse_output(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self.output_folder.set(d)

    def analyze_video(self, path):
        cap   = cv2.VideoCapture(path)
        fps   = cap.get(cv2.CAP_PROP_FPS)
        count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        w     = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        h     = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        dur   = count / fps if fps > 0 else 0
        cap.release()
        self.video_info = {"fps": fps, "count": count}
        self.lbl_info.config(
            text=f"Resolution: {w}\u00d7{h}   |   FPS: {fps:.2f}\n"
                 f"Duration: {timedelta(seconds=int(dur))}   |   Frames: {count}",
            fg=C_TEXT)

    # ─────────────────────────── ENHANCEMENT ───────────────────────────

    def enhance_frame_for_ocr(self, frame):
        scale = self.upscale_factor.get()
        sharp = self.sharpen_amount.get()
        clip  = self.clahe_clip.get()

        h, w = frame.shape[:2]
        frame = cv2.resize(frame, (int(w * scale), int(h * scale)),
                           interpolation=cv2.INTER_LANCZOS4)

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        gray = cv2.fastNlMeansDenoising(gray, h=7,
                                         templateWindowSize=7,
                                         searchWindowSize=21)

        clahe = cv2.createCLAHE(clipLimit=clip, tileGridSize=(8, 8))
        gray  = clahe.apply(gray)

        blurred = cv2.GaussianBlur(gray, (0, 0), sigmaX=2)
        gray    = cv2.addWeighted(gray, 1.0 + sharp, blurred, -sharp, 0)
        gray    = np.clip(gray, 0, 255).astype(np.uint8)

        if self.darken_text.get():
            lut  = np.array([int((i / 255.0) ** 0.65 * 255) for i in range(256)],
                             dtype=np.uint8)
            gray = cv2.LUT(gray, lut)

        kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]], np.float32)
        gray   = cv2.filter2D(gray, -1, kernel)
        gray   = np.clip(gray, 0, 255).astype(np.uint8)
        return gray

    # ─────────────────────────── EXTRACTION ───────────────────────────

    def start_extraction(self):
        if not os.path.exists(self.video_path.get()):
            messagebox.showerror("Error", "Please select a valid video file.")
            return
        out = self.output_folder.get()
        if not out:
            messagebox.showerror("Error", "Please select an output folder.")
            return
        os.makedirs(out, exist_ok=True)
        self.btn_extract.config(state="disabled", text="Extracting…")
        threading.Thread(target=self._process_video, daemon=True).start()

    def _process_video(self):
        cap   = cv2.VideoCapture(self.video_path.get())
        fps   = self.video_info.get("fps", 30)
        total = self.video_info.get("count", 0)

        if self.extraction_mode.get() == "interval":
            step = int(fps * self.interval_seconds.get())
        else:
            step = self.frame_skip.get()
        step = max(1, step)

        frame_id = 0
        saved    = 0

        self._set_status("Extracting frames…")

        while cap.isOpened():
            ret, frame = cap.read()
            if not ret:
                break

            if frame_id % step == 0:
                img = (self.enhance_frame_for_ocr(frame)
                       if self.apply_ocr_enhancement.get() else frame)

                fname = f"screenshot_{saved:04d}_frame_{frame_id:06d}.png"
                cv2.imwrite(os.path.join(self.output_folder.get(), fname),
                            img, [cv2.IMWRITE_PNG_COMPRESSION, 1])

                self._update_preview(img)
                saved += 1

            frame_id += 1

            if frame_id % 50 == 0:
                pct = (frame_id / total * 100) if total else 0
                self.root.after(0, lambda p=pct:
                                self.progress.configure(value=p))
                self.root.after(0, lambda s=saved, f=frame_id:
                                self._set_status(
                                    f"Saved {s} frames  |  {f}/{total}"))

        cap.release()
        self.root.after(0, lambda:
                        self.btn_extract.config(state="normal",
                                                 text="START EXTRACTION"))
        self.root.after(0, lambda: self.progress.configure(value=100))
        self.root.after(0, lambda:
                        messagebox.showinfo("Done",
                                            f"Saved {saved} screenshots.\n\n"
                                            f"Folder:\n{self.output_folder.get()}"))

    def _update_preview(self, cv_img):
        try:
            if len(cv_img.shape) == 2:
                cv_img = cv2.cvtColor(cv_img, cv2.COLOR_GRAY2RGB)
            else:
                cv_img = cv2.cvtColor(cv_img, cv2.COLOR_BGR2RGB)

            cw = max(self.canvas_preview.winfo_width(),  800)
            ch = max(self.canvas_preview.winfo_height(), 500)
            h, w = cv_img.shape[:2]
            scale  = min(cw / w, ch / h)
            nw, nh = int(w * scale), int(h * scale)
            img = Image.fromarray(cv_img).resize((nw, nh), Image.LANCZOS)

            self._tk_img = ImageTk.PhotoImage(img)
            self.canvas_preview.delete("all")
            self.canvas_preview.create_image(cw // 2, ch // 2,
                                              image=self._tk_img,
                                              anchor=tk.CENTER)
        except Exception:
            pass

    def _set_status(self, msg):
        self.root.after(0, lambda: self.lbl_status.config(text=msg))


if __name__ == "__main__":
    root = tk.Tk()
    app  = VideoScreenshotExtractor(root)
    root.mainloop()