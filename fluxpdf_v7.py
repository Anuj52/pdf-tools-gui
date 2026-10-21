#!/usr/bin/env python3
"""
FluxPDF v8.0 (Polished Edition)
Fixed scrolling, improved UI, added keyboard shortcuts and zoom controls.
Designed by Ach (https://github.com/Anuj52)
"""

import os
import sys
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, Canvas
from datetime import datetime

# --- Libraries ---
import customtkinter as ctk
from PIL import Image, ImageTk, ImageOps, ImageEnhance

# --- PDF Engine ---
try:
    import fitz  # PyMuPDF
    ENGINE_READY = True
except ImportError:
    ENGINE_READY = False

# --- Configuration ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")
HISTORY_FILE = "flux_v8_session.json"

COLORS = {
    "bg": "#1e1e1e",
    "toolbar": "#2b2b2b",
    "sidebar": "#252525",
    "accent": "#3B8ED0",
    "text": "#ffffff",
    "sepia_bg": "#f4ecd8",
    "sepia_fg": "#5e4b35"
}

class StateManager:
    """Simple Session Manager"""
    def __init__(self):
        self.data = self._load()
    
    def _load(self):
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, 'r') as f: return json.load(f)
            except: return {}
        return {}

    def save(self, path, page, view, theme, zoom):
        self.data["last_open"] = path
        self.data[path] = {"page": page, "view": view, "theme": theme, "zoom": zoom, "time": str(datetime.now())}
        with open(HISTORY_FILE, 'w') as f: json.dump(self.data, f)

    def get(self, path):
        return self.data.get(path, {"page": 0, "view": "Single", "theme": "Day", "zoom": 1.0})

    def get_recent(self):
        return [k for k in self.data.keys() if k != "last_open"]

class FluxPDF(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("FluxPDF v8.0")
        self.geometry("1300x850")
        self.configure(fg_color=COLORS["bg"])

        if not ENGINE_READY:
            messagebox.showerror("Critical Error", "PyMuPDF is missing.\nPlease run: pip install pymupdf")
            self.destroy()
            return

        self.state = StateManager()
        
        # Runtime Variables
        self.doc = None
        self.curr_path = None
        self.curr_page = 0
        self.curr_zoom = 1.0
        self.curr_view = "Single" # Single, Book, Reflow
        self.curr_theme = "Day"
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._setup_ui()
        self._bind_events()

        # Auto-load last file
        last = self.state.data.get("last_open")
        if last and os.path.exists(last):
            self.load_file(last)

    # =========================================================================
    # 1. MODERN UI SETUP
    # =========================================================================
    def _setup_ui(self):
        # --- Top Toolbar (The Control Center) ---
        self.toolbar = ctk.CTkFrame(self, height=60, fg_color=COLORS["toolbar"], corner_radius=0)
        self.toolbar.grid(row=0, column=0, columnspan=2, sticky="ew")

        # File Section
        ctk.CTkButton(self.toolbar, text="📂 Open Book", fg_color="#444", width=100, command=self.browse_file).pack(side="left", padx=15, pady=10)
        
        # View Controls (Center)
        ctrl_frame = ctk.CTkFrame(self.toolbar, fg_color="transparent")
        ctrl_frame.pack(side="left", expand=True)
        
        self.seg_view = ctk.CTkSegmentedButton(ctrl_frame, values=["Single", "Book", "Reflow"], command=self.set_view)
        self.seg_view.set("Single")
        self.seg_view.pack(side="left", padx=10)

        self.cmb_theme = ctk.CTkOptionMenu(ctrl_frame, values=["Day", "Night", "Sepia"], width=90, command=self.set_theme)
        self.cmb_theme.set("Day")
        self.cmb_theme.pack(side="left", padx=10)

        # Zoom Controls
        ctk.CTkButton(ctrl_frame, text="➖", width=40, fg_color="#333", command=self.zoom_out).pack(side="left", padx=2)
        self.lbl_zoom = ctk.CTkLabel(ctrl_frame, text="100%", width=50)
        self.lbl_zoom.pack(side="left", padx=2)
        ctk.CTkButton(ctrl_frame, text="➕", width=40, fg_color="#333", command=self.zoom_in).pack(side="left", padx=2)

        # Right Section
        ctk.CTkButton(self.toolbar, text="Fullscreen", width=80, fg_color=COLORS["accent"], command=self.toggle_fs).pack(side="right", padx=15)

        # --- Sidebar (Navigation) ---
        self.sidebar = ctk.CTkFrame(self, width=250, fg_color=COLORS["sidebar"], corner_radius=0)
        self.sidebar.grid(row=1, column=0, sticky="nsew")
        
        self.tab_view = ctk.CTkTabview(self.sidebar, fg_color=COLORS["sidebar"])
        self.tab_view.pack(fill="both", expand=True, padx=5, pady=5)
        self.tab_toc = self.tab_view.add("Chapters")
        self.tab_recent = self.tab_view.add("History")
        
        self.scroll_toc = ctk.CTkScrollableFrame(self.tab_toc, fg_color="transparent")
        self.scroll_toc.pack(fill="both", expand=True)
        self.scroll_recent = ctk.CTkScrollableFrame(self.tab_recent, fg_color="transparent")
        self.scroll_recent.pack(fill="both", expand=True)
        self._update_recent_ui()

        # --- Main Reader Area ---
        self.reader_area = ctk.CTkFrame(self, fg_color="#111")
        self.reader_area.grid(row=1, column=1, sticky="nsew")

        # 1. Canvas (For Image Based Views)
        self.canvas = Canvas(self.reader_area, bg="#111", highlightthickness=0)
        
        # 2. Textbox (For Reflow Mode)
        self.text_view = ctk.CTkTextbox(self.reader_area, font=("Georgia", 20), wrap="word", fg_color="#1a1a1a")

        # Default View
        self.canvas.pack(fill="both", expand=True)

        # --- Floating Status Bar ---
        self.status = ctk.CTkFrame(self.reader_area, height=40, fg_color=COLORS["toolbar"], corner_radius=10)
        self.status.place(relx=0.5, rely=0.95, anchor="center")
        
        ctk.CTkButton(self.status, text="◀", width=40, fg_color="transparent", command=self.prev_page).pack(side="left")
        self.lbl_page = ctk.CTkLabel(self.status, text="No File Loaded", font=("Arial", 13, "bold"))
        self.lbl_page.pack(side="left", padx=15)
        ctk.CTkButton(self.status, text="▶", width=40, fg_color="transparent", command=self.next_page).pack(side="left")

    def _bind_events(self):
        # Scroll Wheel
        self.reader_area.bind_all("<MouseWheel>", self._on_mouse_scroll) # Windows
        self.reader_area.bind_all("<Button-4>", self._on_mouse_scroll)   # Linux Up
        self.reader_area.bind_all("<Button-5>", self._on_mouse_scroll)   # Linux Down
        
        # Keyboard
        self.bind("<Left>", lambda e: self.prev_page())
        self.bind("<Right>", lambda e: self.next_page())
        self.bind("<Up>", lambda e: self._scroll_reflow("up"))
        self.bind("<Down>", lambda e: self._scroll_reflow("down"))

    # =========================================================================
    # 2. LOGIC & LOADING
    # =========================================================================
    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if path: self.load_file(path)

    def load_file(self, path):
        try:
            self.doc = fitz.open(path)
            self.curr_path = path
            
            # Load Saved State
            saved = self.state.get(path)
            self.curr_page = saved["page"]
            self.curr_view = saved["view"]
            self.curr_theme = saved["theme"]
            self.curr_zoom = saved["zoom"]

            # Sync UI Controls
            self.seg_view.set(self.curr_view)
            self.cmb_theme.set(self.curr_theme)
            self.lbl_zoom.configure(text=f"{int(self.curr_zoom*100)}%")

            self._load_toc()
            self._update_recent_ui()
            self.render()
        except Exception as e:
            messagebox.showerror("Load Error", str(e))

    def _load_toc(self):
        for w in self.scroll_toc.winfo_children(): w.destroy()
        toc = self.doc.get_toc()
        if not toc:
            ctk.CTkLabel(self.scroll_toc, text="No Chapters").pack()
            return
        for lvl, title, pg in toc:
            ctk.CTkButton(self.scroll_toc, text="  "*(lvl-1) + title, anchor="w", fg_color="transparent", 
                          height=28, text_color="gray90", hover_color="#333",
                          command=lambda p=pg: self.goto_page(p-1)).pack(fill="x")

    def _update_recent_ui(self):
        for w in self.scroll_recent.winfo_children(): w.destroy()
        for path in self.state.get_recent():
            if os.path.exists(path):
                name = os.path.basename(path)
                ctk.CTkButton(self.scroll_recent, text=f"📄 {name}", anchor="w", fg_color="transparent",
                              height=35, command=lambda p=path: self.load_file(p)).pack(fill="x")

    # =========================================================================
    # 3. RENDERING ENGINE (V8)
    # =========================================================================
    def render(self):
        if not self.doc: return
        
        # Boundary Checks
        if self.curr_page < 0: self.curr_page = 0
        if self.curr_page >= len(self.doc): self.curr_page = len(self.doc) - 1

        self.lbl_page.configure(text=f"Page {self.curr_page + 1} / {len(self.doc)}")
        self.state.save(self.curr_path, self.curr_page, self.curr_view, self.curr_theme, self.curr_zoom)

        # Handle Reflow vs Image
        if self.curr_view == "Reflow":
            self._render_reflow()
        else:
            self._render_image()

    def _render_reflow(self):
        self.canvas.pack_forget()
        self.text_view.pack(fill="both", expand=True, padx=100, pady=20)
        
        # Theme Colors
        if self.curr_theme == "Night":
            fg, bg = "#cccccc", "#000000"
        elif self.curr_theme == "Sepia":
            fg, bg = COLORS["sepia_fg"], COLORS["sepia_bg"]
        else:
            fg, bg = "#000000", "#ffffff"
            
        self.text_view.configure(fg_color=bg, text_color=fg)
        text = self.doc[self.curr_page].get_text("text")
        self.text_view.delete("0.0", "end")
        self.text_view.insert("0.0", text)

    def _render_image(self):
        self.text_view.pack_forget()
        self.canvas.pack(fill="both", expand=True)
        
        # Theme Backgrounds
        bg_col = "#000000" if self.curr_theme == "Night" else "#333333"
        self.canvas.configure(bg=bg_col)

        # Get Page Image
        def get_img(pg_idx):
            if pg_idx >= len(self.doc): return None
            page = self.doc[pg_idx]
            mat = fitz.Matrix(self.curr_zoom * 1.5, self.curr_zoom * 1.5) # High DPI
            pix = page.get_pixmap(matrix=mat)
            return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        img = get_img(self.curr_page)

        # Book Mode Logic
        if self.curr_view == "Book" and self.curr_page + 1 < len(self.doc):
            img2 = get_img(self.curr_page + 1)
            w = img.width + img2.width
            h = max(img.height, img2.height)
            merged = Image.new("RGB", (w, h))
            merged.paste(img, (0,0))
            merged.paste(img2, (img.width, 0))
            img = merged

        # Image Processing for Themes
        if self.curr_theme == "Night":
            img = ImageOps.invert(img)
            img = ImageEnhance.Contrast(img).enhance(0.8)
        elif self.curr_theme == "Sepia":
            sepia = Image.new("RGB", img.size, COLORS["sepia_bg"])
            img = Image.blend(img, sepia, 0.3)

        # Display
        self.tk_img = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        
        # Center Image on Canvas
        c_w = self.canvas.winfo_width()
        c_h = self.canvas.winfo_height()
        
        # If canvas isn't ready, default to 0,0
        x_pos = max(0, (c_w - img.width) // 2) if c_w > 1 else 0
        
        self.canvas.create_image(x_pos, 20, image=self.tk_img, anchor="nw")
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    # =========================================================================
    # 4. ACTIONS
    # =========================================================================
    def prev_page(self):
        step = 2 if self.curr_view == "Book" else 1
        self.curr_page -= step
        self.render()

    def next_page(self):
        step = 2 if self.curr_view == "Book" else 1
        self.curr_page += step
        self.render()

    def goto_page(self, p):
        self.curr_page = p
        self.render()

    def set_view(self, view):
        self.curr_view = view
        self.render()

    def set_theme(self, theme):
        self.curr_theme = theme
        self.render()

    def zoom_in(self):
        self.curr_zoom += 0.1
        self.lbl_zoom.configure(text=f"{int(self.curr_zoom*100)}%")
        self.render()

    def zoom_out(self):
        if self.curr_zoom > 0.3:
            self.curr_zoom -= 0.1
            self.lbl_zoom.configure(text=f"{int(self.curr_zoom*100)}%")
            self.render()

    def toggle_fs(self):
        self.attributes("-fullscreen", not self.attributes("-fullscreen"))

    # --- SCROLL LOGIC ---
    def _on_mouse_scroll(self, event):
        """Handle scroll wheel to flip pages or scroll text"""
        if not self.doc: return
        
        # Direction detection
        delta = 0
        if sys.platform == "darwin": delta = event.delta 
        elif event.num == 4: delta = 1
        elif event.num == 5: delta = -1
        else: delta = int(event.delta / 120) # Windows

        if self.curr_view == "Reflow":
            if delta > 0: self.text_view.yview_scroll(-1, "units")
            else: self.text_view.yview_scroll(1, "units")
        else:
            # Page Flip Logic
            if delta > 0: self.prev_page()
            else: self.next_page()

    def _scroll_reflow(self, direction):
        if self.curr_view == "Reflow":
            if direction == "up": self.text_view.yview_scroll(-1, "units")
            else: self.text_view.yview_scroll(1, "units")
        else:
            # Arrow keys for page flip
            if direction == "up": self.prev_page()
            else: self.next_page()

if __name__ == "__main__":
    app = FluxPDF()
    app.mainloop()