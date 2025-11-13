import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from pathlib import Path
from datetime import datetime
import yaml
from playwright.sync_api import sync_playwright
from scrape_and_generate import (
    find_app_frame, get_by_label, click_partners_tab, find_partners_frame,
    get_initial_reporter_name, get_facility_name_and_address, 
    read_external_refs, get_event_date, read_all_products, 
    read_event_description, read_associated_transactions_complete,
    read_analysis_summary_for_txid, read_investigation_summary_for_txid,
    fill_docx, extract_product_code, wait_for_search_with_retries,
    wait_find_in_any_frame, log
)
class GradientFrame(tk.Canvas):
    def __init__(self, parent, color1, color2, **kwargs):
        tk.Canvas.__init__(self, parent, **kwargs)
        self.color1 = color1
        self.color2 = color2
        self.bind("<Configure>", self._draw_gradient)
    def _draw_gradient(self, event=None):
        self.delete("gradient")
        width = self.winfo_width()
        height = self.winfo_height()
        limit = height
        r1, g1, b1 = self.winfo_rgb(self.color1)
        r2, g2, b2 = self.winfo_rgb(self.color2)
        r_ratio = (r2 - r1) / limit
        g_ratio = (g2 - g1) / limit
        b_ratio = (b2 - b1) / limit
        for i in range(limit):
            nr = int(r1 + (r_ratio * i))
            ng = int(g1 + (g_ratio * i))
            nb = int(b1 + (b_ratio * i))
            color = f'#{nr>>8:02x}{ng>>8:02x}{nb>>8:02x}'
            self.create_line(0, i, width, i, tags=("gradient",), fill=color)
        self.lower("gradient")
class ModernButton(tk.Canvas):
    def __init__(self, parent, text, command, style="primary", width=200, height=48):
        super().__init__(parent, width=width, height=height, 
                        highlightthickness=0, cursor="hand2")
        self.command = command
        self.text = text
        self.style = style
        self.disabled = False
        styles = {
            "primary": {"normal": "#6366f1", "hover": "#4f46e5", "text": "white"},
            "secondary": {"normal": "#e5e7eb", "hover": "#d1d5db", "text": "#374151"},
            "success": {"normal": "#10b981", "hover": "#059669", "text": "white"},
            "danger": {"normal": "#ef4444", "hover": "#dc2626", "text": "white"}
        }
        self.colors = styles.get(style, styles["primary"])
        self.configure(bg=parent.cget('bg'))
        self.rect = self.create_rounded_rect(2, 2, width-2, height-2, radius=10, 
                                             fill=self.colors["normal"], outline="")
        self.shadow = self.create_rounded_rect(0, 4, width, height+2, radius=10,
                                               fill="#000000", outline="")
        self.itemconfig(self.shadow, state='hidden')
        self.text_item = self.create_text(width//2, height//2, text=text,
                                         fill=self.colors["text"],
                                         font=("Inter", 11, "bold"))
        self.tag_raise(self.rect)
        self.tag_raise(self.text_item)
        self.bind("<Button-1>", self._on_click)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
    def create_rounded_rect(self, x1, y1, x2, y2, radius=25, **kwargs):
        points = [x1+radius, y1,
                  x1+radius, y1,
                  x2-radius, y1,
                  x2-radius, y1,
                  x2, y1,
                  x2, y1+radius,
                  x2, y1+radius,
                  x2, y2-radius,
                  x2, y2-radius,
                  x2, y2,
                  x2-radius, y2,
                  x2-radius, y2,
                  x1+radius, y2,
                  x1+radius, y2,
                  x1, y2,
                  x1, y2-radius,
                  x1, y2-radius,
                  x1, y1+radius,
                  x1, y1+radius,
                  x1, y1]
        return self.create_polygon(points, smooth=True, **kwargs)
    def _on_enter(self, e):
        if not self.disabled:
            self.itemconfig(self.rect, fill=self.colors["hover"])
            self.itemconfig(self.shadow, state='normal')
            self.configure(cursor="hand2")
    def _on_leave(self, e):
        if not self.disabled:
            self.itemconfig(self.rect, fill=self.colors["normal"])
            self.itemconfig(self.shadow, state='hidden')
    def _on_press(self, e):
        if not self.disabled:
            self.move(self.rect, 0, 2)
            self.move(self.text_item, 0, 2)
    def _on_release(self, e):
        if not self.disabled:
            self.move(self.rect, 0, -2)
            self.move(self.text_item, 0, -2)
    def _on_click(self, e):
        if not self.disabled and self.command:
            self.command()
    def set_state(self, state, text=None):
        self.disabled = (state == "disabled")
        if text:
            self.itemconfig(self.text_item, text=text)
        if self.disabled:
            self.itemconfig(self.rect, fill="#d1d5db")
            self.itemconfig(self.text_item, fill="#9ca3af")
            self.itemconfig(self.shadow, state='hidden')
            self.configure(cursor="arrow")
        else:
            self.itemconfig(self.rect, fill=self.colors["normal"])
            self.itemconfig(self.text_item, fill=self.colors["text"])
            self.configure(cursor="hand2")
class AnimatedProgressBar(tk.Canvas):
    def __init__(self, parent, steps=4, **kwargs):
        super().__init__(parent, height=80, highlightthickness=0, **kwargs)
        self.steps = steps
        self.current_step = 1
        self.configure(bg=parent.cget('bg'))
    def draw_progress(self, current_step):
        self.delete("all")
        self.current_step = current_step
        width = self.winfo_width()
        if width <= 1:
            self.after(100, lambda: self.draw_progress(current_step))
            return
        step_width = width / self.steps
        for i in range(self.steps):
            x = i * step_width + step_width / 2
            if i + 1 < current_step:
                state = "completed"
                color = "#10b981"
                text_color = "white"
            elif i + 1 == current_step:
                state = "active"
                color = "#6366f1"
                text_color = "white"
            else:
                state = "inactive"
                color = "#e5e7eb"
                text_color = "#9ca3af"
            self.create_oval(x-20, 15, x+20, 55, fill=color, outline="", width=0)
            if state == "completed":
                self.create_text(x, 35, text="✓", fill=text_color, 
                               font=("Inter", 16, "bold"))
            else:
                self.create_text(x, 35, text=str(i+1), fill=text_color,
                               font=("Inter", 14, "bold"))
            if i < self.steps - 1:
                line_color = "#10b981" if i + 1 < current_step else "#e5e7eb"
                self.create_line(x+20, 35, x+step_width-20, 35, 
                               fill=line_color, width=3)
            labels = ["Search", "Review Data", "Final Review", "Complete"]
            self.create_text(x, 70, text=labels[i], fill="#6b7280",
                           font=("Inter", 9))
class ComplaintLetterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Complaint Letter Generator")
        self.root.geometry("1100x800")
        self.bg_gradient_start = "#f0f4f8"
        self.bg_gradient_end = "#d9e2ec"
        self.card_bg = "#ffffff"
        self.primary = "#6366f1"
        self.success = "#10b981"
        self.text_primary = "#1f2937"
        self.text_secondary = "#6b7280"
        self.border_color = "#e5e7eb"
        self.complaint_id = tk.StringVar()
        self.current_step = 1
        self.all_data = {}
        self.products = []
        config_path = Path(__file__).parent / "config.yaml"
        if not config_path.exists():
            messagebox.showerror("Error", f"config.yaml not found at {config_path}")
            root.destroy()
            return
        self.config = yaml.safe_load(config_path.read_text())
        self.bg = GradientFrame(root, self.bg_gradient_start, self.bg_gradient_end)
        self.bg.place(x=0, y=0, relwidth=1, relheight=1)
        self.container = tk.Frame(self.bg, bg=self.card_bg)
        self.container.place(relx=0.5, rely=0.5, anchor="center", 
                           width=1000, height=700)
        self.shadow = tk.Frame(self.bg, bg="#b0bec5")
        self.shadow.place(relx=0.5, rely=0.505, anchor="center",
                        width=1000, height=700)
        self.shadow.lower()
        self.create_header()
        self.progress_bar = AnimatedProgressBar(self.container, steps=4, bg=self.card_bg)
        self.progress_bar.pack(fill=tk.X, padx=40, pady=(20, 30))
        self.progress_bar.draw_progress(1)
        self.content_frame = tk.Frame(self.container, bg=self.card_bg)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=(0, 30))
        self.create_step1()
        self.create_step2()
        self.create_step3()
        self.create_step4()
        self.show_step(1)
    def create_header(self):
        header = tk.Frame(self.container, bg=self.card_bg, height=100)
        header.pack(fill=tk.X, padx=40, pady=(30, 0))
        header.pack_propagate(False)
        title = tk.Label(header, text="Complaint Letter Generator",
                        font=("Inter", 28, "bold"), bg=self.card_bg,
                        fg=self.text_primary)
        title.pack(side=tk.LEFT, pady=20)
    def create_modern_label(self, parent, text, size=11, bold=False):
        weight = "bold" if bold else "normal"
        return tk.Label(parent, text=text, font=("Inter", size, weight),
                       bg=self.card_bg, fg=self.text_primary, anchor=tk.W)
    def create_modern_entry(self, parent, textvariable=None, placeholder=""):
        frame = tk.Frame(parent, bg="#f9fafb", highlightthickness=1,
                        highlightbackground=self.border_color, highlightcolor=self.primary)
        entry = tk.Entry(frame, textvariable=textvariable, font=("Inter", 12),
                        bg="#f9fafb", fg=self.text_primary, relief=tk.FLAT,
                        insertbackground=self.primary, border=0)
        entry.pack(fill=tk.BOTH, expand=True, padx=15, pady=12)
        return frame, entry
    def create_modern_text(self, parent, height=6):
        frame = tk.Frame(parent, bg="#f9fafb", highlightthickness=1,
                        highlightbackground=self.border_color, highlightcolor=self.primary)
        
        text = tk.Text(frame, font=("Inter", 10), bg="#f9fafb",
                      fg=self.text_primary, relief=tk.FLAT,
                      insertbackground=self.primary, height=height,
                      wrap=tk.WORD, border=0, padx=12, pady=12)
        scrollbar = tk.Scrollbar(frame, command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        return frame, text
    def create_step1(self):
        self.step1 = tk.Frame(self.content_frame, bg=self.card_bg)
        title = self.create_modern_label(self.step1, "Search for Complaint", 18, bold=True)
        title.pack(anchor=tk.W, pady=(0, 10))
        subtitle = self.create_modern_label(self.step1, 
            "Enter the PE Number to retrieve and process information", 10)
        subtitle.pack(anchor=tk.W, pady=(0, 30))
        label = self.create_modern_label(self.step1, "PE Number", 11, bold=True)
        label.pack(anchor=tk.W, pady=(0, 10))
        entry_frame, self.complaint_entry = self.create_modern_entry(
            self.step1, self.complaint_id)
        entry_frame.pack(fill=tk.X, pady=(0, 25))
        self.complaint_entry.bind('<Return>', lambda e: self.run_scraper())
        self.complaint_entry.focus()
        btn_container = tk.Frame(self.step1, bg=self.card_bg)
        btn_container.pack(pady=(10, 20))
        self.search_btn = ModernButton(btn_container, "Search & Process",
                                      self.run_scraper, style="primary", 
                                      width=220, height=50)
        self.search_btn.pack()
        log_label = self.create_modern_label(self.step1, "Activity Log", 11, bold=True)
        log_label.pack(anchor=tk.W, pady=(20, 10))
        log_frame = tk.Frame(self.step1, bg="#1f2937", highlightthickness=1,
                            highlightbackground="#374151")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.progress_log = scrolledtext.ScrolledText(
            log_frame, height=12, state='disabled', wrap=tk.WORD,
            font=("Consolas", 9), bg="#1f2937", fg="#10b981",
            relief=tk.FLAT, border=0, padx=15, pady=12,
            insertbackground="#10b981")
        self.progress_log.pack(fill=tk.BOTH, expand=True)
    def create_step2(self):
        self.step2 = tk.Frame(self.content_frame, bg=self.card_bg)
        title = self.create_modern_label(self.step2, "Review Initial Reporter Data", 18, bold=True)
        title.pack(anchor=tk.W, pady=(0, 10))
        subtitle = self.create_modern_label(self.step2,
            "Verify and edit the information before proceeding", 10)
        subtitle.pack(anchor=tk.W, pady=(0, 25))
        label = self.create_modern_label(self.step2, "Initial Reporter & Address", 11, bold=True)
        label.pack(anchor=tk.W, pady=(0, 8))
        ir_frame, self.ir_text = self.create_modern_text(self.step2, height=4)
        ir_frame.pack(fill=tk.X, pady=(0, 20))
        label = self.create_modern_label(self.step2, "Event Description", 11, bold=True)
        label.pack(anchor=tk.W, pady=(0, 8))
        event_frame, self.event_text = self.create_modern_text(self.step2, height=5)
        event_frame.pack(fill=tk.X, pady=(0, 20))
        label = self.create_modern_label(self.step2, "Associated Products", 11, bold=True)
        label.pack(anchor=tk.W, pady=(0, 8))
        products_frame = tk.Frame(self.step2, bg="#f9fafb", highlightthickness=1,
                                 highlightbackground=self.border_color)
        products_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        self.products_text = scrolledtext.ScrolledText(
            products_frame, height=4, state='disabled', wrap=tk.WORD,
            font=("Inter", 10), bg="#f9fafb", fg=self.text_primary,
            relief=tk.FLAT, border=0, padx=12, pady=12)
        self.products_text.pack(fill=tk.BOTH, expand=True)
        btn_frame = tk.Frame(self.step2, bg=self.card_bg)
        btn_frame.pack(fill=tk.X)
        back_btn = ModernButton(btn_frame, "Back", lambda: self.show_step(1),
                               style="secondary", width=120, height=48)
        back_btn.pack(side=tk.LEFT, padx=(0, 10))
        next_btn = ModernButton(btn_frame, "Continue", lambda: self.show_step(3),
                               style="primary", width=140, height=48)
        next_btn.pack(side=tk.LEFT)
    def create_step3(self):
        self.step3 = tk.Frame(self.content_frame, bg=self.card_bg)
        title = self.create_modern_label(self.step3, "Review Analysis & Investigation", 18, bold=True)
        title.pack(anchor=tk.W, pady=(0, 10))
        subtitle = self.create_modern_label(self.step3,
            "Final review before generating the document", 10)
        subtitle.pack(anchor=tk.W, pady=(0, 25))
        label = self.create_modern_label(self.step3, "Investigation Summary", 11, bold=True)
        label.pack(anchor=tk.W, pady=(0, 8))
        inv_frame, self.investigation_text_widget = self.create_modern_text(self.step3, height=5)
        inv_frame.pack(fill=tk.X, pady=(0, 20))
        label = self.create_modern_label(self.step3, "Analysis Results", 11, bold=True)
        label.pack(anchor=tk.W, pady=(0, 8))
        analysis_frame, self.analysis_text_widget = self.create_modern_text(self.step3, height=5)
        analysis_frame.pack(fill=tk.X, pady=(0, 20))
        label = self.create_modern_label(self.step3, "Save Location", 11, bold=True)
        label.pack(anchor=tk.W, pady=(0, 8))
        save_container = tk.Frame(self.step3, bg=self.card_bg)
        save_container.pack(fill=tk.X, pady=(0, 25))
        self.save_path = tk.StringVar()
        save_frame, save_entry = self.create_modern_entry(save_container, self.save_path)
        save_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        save_entry.config(state='readonly')
        browse_btn = ModernButton(save_container, "Browse", self.select_save_location,
                                 style="secondary", width=120, height=48)
        browse_btn.pack(side=tk.LEFT)
        btn_frame = tk.Frame(self.step3, bg=self.card_bg)
        btn_frame.pack(fill=tk.X)
        back_btn = ModernButton(btn_frame, "Back", lambda: self.show_step(2),
                               style="secondary", width=120, height=48)
        back_btn.pack(side=tk.LEFT, padx=(0, 10))
        self.generate_btn = ModernButton(btn_frame, "Generate Document",
                                        self.generate_document,
                                        style="success", width=180, height=48)
        self.generate_btn.pack(side=tk.LEFT)
    def create_step4(self):
        self.step4 = tk.Frame(self.content_frame, bg=self.card_bg)
        center = tk.Frame(self.step4, bg=self.card_bg)
        center.place(relx=0.5, rely=0.5, anchor="center")
        icon_canvas = tk.Canvas(center, width=120, height=120, bg=self.card_bg,
                               highlightthickness=0)
        icon_canvas.pack(pady=(0, 30))
        icon_canvas.create_oval(10, 10, 110, 110, fill=self.success, outline="")
        icon_canvas.create_text(60, 60, text="✓", font=("Inter", 50, "bold"),
                              fill="white")
        title = self.create_modern_label(center, "Document Generated Successfully!",
                                        22, bold=True)
        title.pack(pady=(0, 15))
        subtitle = self.create_modern_label(center, "Your complaint letter is ready",
                                           12)
        subtitle.pack(pady=(0, 30))
        path_frame = tk.Frame(center, bg="#f0f9ff", highlightthickness=1,
                             highlightbackground="#bfdbfe")
        path_frame.pack(fill=tk.X, pady=(0, 40), padx=50)
        self.saved_path_label = tk.Label(path_frame, text="", font=("Inter", 10),
                                         bg="#f0f9ff", fg=self.primary,
                                         wraplength=600, pady=20, padx=25)
        self.saved_path_label.pack()
        reset_btn = ModernButton(center, "Process Another Complaint",
                                self.reset_app, style="primary",
                                width=260, height=50)
        reset_btn.pack()
    def show_step(self, step):
        for s in [self.step1, self.step2, self.step3, self.step4]:
            s.pack_forget()
        self.current_step = step
        if step == 1:
            self.step1.pack(fill=tk.BOTH, expand=True)
        elif step == 2:
            self.step2.pack(fill=tk.BOTH, expand=True)
        elif step == 3:
            self.step3.pack(fill=tk.BOTH, expand=True)
        elif step == 4:
            self.step4.pack(fill=tk.BOTH, expand=True)
        self.progress_bar.draw_progress(step)
    def select_save_location(self):
        default_name = f"Customer_Letter_{self.complaint_id.get()}.docx"
        filename = filedialog.asksaveasfilename(
            title="Save Customer Letter",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word Documents", "*.docx"), ("All files", "*.*")]
        )
        if filename:
            self.save_path.set(filename)
    def log_progress(self, message):
        self.progress_log.configure(state='normal')
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.progress_log.insert(tk.END, f"[{timestamp}] {message}\n")
        self.progress_log.see(tk.END)
        self.progress_log.configure(state='disabled')
        self.root.update_idletasks()
    def run_scraper(self):
        if not self.complaint_id.get().strip():
            messagebox.showerror("Error", "Please enter a PE Number")
            return
        self.search_btn.set_state("disabled", "Processing...")
        self.progress_log.configure(state='normal')
        self.progress_log.delete(1.0, tk.END)
        self.progress_log.configure(state='disabled')
        thread = threading.Thread(target=self._scrape_data)
        thread.daemon = True
        thread.start()
    def _scrape_data(self):
        try:
            complaint_id = self.complaint_id.get()
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=self.config.get('headless', False))
                context = browser.new_context()
                page = context.new_page()
                self.log_progress(f"Navigating to CRM...")
                page.goto(self.config['crm_url'], wait_until="load")
                sso_wait = self.config.get('sso_pause_seconds', 0)
                if sso_wait > 0:
                    s = self.config.get('search', {})
                    try:
                        wait_find_in_any_frame(page, [s.get('selector')] + 
                                             (s.get('fallback_selectors', []) or []),
                                             timeout_ms=2500, poll_ms=150)
                        self.log_progress("Authentication complete")
                    except Exception:
                        self.log_progress("Authenticating...")
                        try:
                            wait_for_search_with_retries(
                                page, s,
                                max_attempts=self.config.get('sso_max_attempts', 8),
                                probe_period_ms=self.config.get('sso_probe_period_ms', 2000),
                                reload_between_attempts=self.config.get('sso_reload_between_attempts', True),
                                total_timeout_ms=self.config.get('sso_total_timeout_ms', 240000),
                            )
                        except Exception:
                            page.wait_for_timeout(sso_wait * 1000)
                frame = find_app_frame(
                    page,
                    frame_name_regex=self.config.get('frame_name_regex'),
                    url_regex=self.config.get('frame_url_regex')
                )
                if 'search' in self.config:
                    s = self.config['search']
                    if s.get('selector'):
                        self.log_progress(f"Searching for complaint: {complaint_id}")
                        fallbacks = s.get('fallback_selectors', [])
                        target, target_ctx, used_sel = wait_find_in_any_frame(
                            page, [s['selector']] + fallbacks, timeout_ms=5000
                        )
                        if s.get('clear', True):
                            try:
                                target.fill("")
                            except Exception:
                                pass
                        target.click()
                        frame.wait_for_timeout(100)
                        target.fill(complaint_id)
                        if s.get('submit_selector'):
                            btn = target_ctx.locator(s['submit_selector']).first
                            if btn.count():
                                btn.click()
                        else:
                            target.press(s.get('press_key', 'Enter'))
                        if s.get('wait_for'):
                            target_ctx.wait_for_selector(s['wait_for'], 
                                                        timeout=s.get('wait_timeout', 60000))
                        else:
                            target_ctx.wait_for_load_state("networkidle")
                            target_ctx.wait_for_timeout(s.get('post_wait_ms', 3000))
                values = {}
                values['complaint_id'] = complaint_id
                values["todays_date"] = datetime.now().strftime("%B %d, %Y").replace(" 0", " ")
                for key, conf in self.config.get('field_map', {}).items():
                    if isinstance(conf, str):
                        values[key] = get_by_label(frame, conf)
                    elif isinstance(conf, list):
                        values[key] = get_by_label(frame, conf)
                    elif isinstance(conf, dict):
                        if conf.get('type') == 'label':
                            values[key] = get_by_label(frame, conf.get('labels', []))
                        elif conf.get('type') == 'literal':
                            values[key] = conf.get('value', '')
                self.log_progress("Retrieving partner information...")
                if click_partners_tab(page, frame):
                    pframe = find_partners_frame(page)
                    if pframe:
                        irname = get_initial_reporter_name(pframe)
                        if irname:
                            values['ir_name'] = irname
                        facility_block = get_facility_name_and_address(pframe)
                        if facility_block:
                            values['ir_with_address'] = facility_block
                self.log_progress("Reading external references...")
                ext = read_external_refs(page, frame)
                if ext.get("rb_reference"):
                    values["rb_reference"] = ext["rb_reference"]
                if ext.get("report_number"):
                    values["report_number"] = ext["report_number"]
                self.log_progress("Extracting event date...")
                event_date_text = get_event_date(page)
                if event_date_text:
                    values['event_date'] = event_date_text
                self.log_progress("Loading product line items...")
                products = read_all_products(page, frame)
                self.products = products
                self.log_progress("Reading event description...")
                desc = read_event_description(page, frame)
                if desc:
                    import re
                    desc = re.sub(r'^\s*it\s+was\s+reported(?:\s+that)?[,:-]?\s*', 
                                '', desc, flags=re.I).lstrip()
                    desc = re.sub(r'([.!?])\1+', r'\1', desc)
                    values["event_description"] = desc
                self.log_progress("Processing associated transactions...")
                assoc = read_associated_transactions_complete(page, frame)
                pa_ids = assoc.get("product_analysis", [])
                pa_summary_lines = []
                for txid in pa_ids:
                    self.log_progress(f"Analyzing product analysis: {txid}")
                    summary = read_analysis_summary_for_txid(page, txid)
                    if summary:
                        pa_summary_lines.append(summary)                
                default_pa = ("Information provided to Medtronic indicated that the "
                            "complaint device was not available for evaluation.")
                if pa_summary_lines:
                    values["analysis_results"] = "\n\n".join(pa_summary_lines)
                else:
                    values["analysis_results"] = default_pa                
                inv_ids = assoc.get("investigation", [])
                inv_summary_lines = []
                for txid in inv_ids:
                    self.log_progress(f"Processing investigation: {txid}")
                    summary = read_investigation_summary_for_txid(page, txid)
                    if summary:
                        inv_summary_lines.append(summary)
                if inv_summary_lines:
                    values["investigation_summary"] = "\n\n".join(inv_summary_lines)
                else:
                    values["investigation_summary"] = ""
                values.setdefault("analysis_1", values.get("analysis_results", default_pa))
                values.setdefault("investigation_1", values.get("investigation_summary", ""))
                for idx, p in enumerate(products, start=1):
                    code = (p.get("code") or extract_product_code(p.get("desc",""))).upper()
                    values[f"product_id_{idx}"] = (p.get("id") or code)
                    values[f"product_desc_{idx}"] = p.get("desc", "")
                    sn = (p.get("sn", "") or "").strip()
                    lot = (p.get("lot", "") or "").strip()
                    values[f"product_sn_{idx}"] = sn
                    values[f"product_lot_{idx}"] = lot
                    if sn and lot:
                        first_table_display = f"SN: {sn} / LN: {lot}"
                    elif sn:
                        first_table_display = f"SN: {sn}"
                    elif lot:
                        first_table_display = f"LN: {lot}"
                    else:
                        first_table_display = ""
                    values[f"serial_or_lot_{idx}"] = first_table_display
                values["_product_count"] = len(products)
                if not (values.get('ir_name') or '').strip():
                    values['ir_name'] = 'Customer'
                self.all_data = values
                browser.close()
                self.log_progress("Data collection complete")
                self.root.after(0, self._update_step2_ui)
        except Exception as e:
            self.log_progress(f"Error: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            self.root.after(0, lambda: self.search_btn.set_state("normal", "Search & Process"))
    def _update_step2_ui(self):
        self.ir_text.delete(1.0, tk.END)
        self.ir_text.insert(1.0, self.all_data.get('ir_with_address', 
                                                   self.all_data.get('ir_name', '')))
        self.event_text.delete(1.0, tk.END)
        self.event_text.insert(1.0, self.all_data.get('event_description', ''))
        self.products_text.configure(state='normal')
        self.products_text.delete(1.0, tk.END)
        for idx, p in enumerate(self.products, 1):
            self.products_text.insert(tk.END, f"{idx}. ", "number")
            self.products_text.insert(tk.END, f"{p.get('id', '')} ", "bold")
            self.products_text.insert(tk.END, f"— {p.get('desc', '')}\n")
            if p.get('sn') or p.get('lot'):
                parts = []
                if p.get('sn'):
                    parts.append(f"SN: {p['sn']}")
                if p.get('lot'):
                    parts.append(f"LOT: {p['lot']}")
                self.products_text.insert(tk.END, f"   {' / '.join(parts)}\n")
            self.products_text.insert(tk.END, "\n")
        self.products_text.tag_config("bold", font=("Inter", 10, "bold"))
        self.products_text.tag_config("number", foreground=self.primary, 
                                     font=("Inter", 10, "bold"))
        self.products_text.configure(state='disabled')
        self.investigation_text_widget.delete(1.0, tk.END)
        self.investigation_text_widget.insert(1.0, 
                                             self.all_data.get('investigation_summary', ''))
        self.analysis_text_widget.delete(1.0, tk.END)
        self.analysis_text_widget.insert(1.0, self.all_data.get('analysis_results', ''))
        self.search_btn.set_state("normal", "Search & Process")
        self.show_step(2)
    def generate_document(self):
        if not self.save_path.get():
            messagebox.showerror("Error", "Please select a save location")
            return
        self.generate_btn.set_state("disabled", "Generating...")
        thread = threading.Thread(target=self._generate_document)
        thread.daemon = True
        thread.start()
    def _generate_document(self):
        try:
            self.all_data['ir_with_address'] = self.ir_text.get(1.0, tk.END).strip()
            self.all_data['ir_name'] = self.all_data['ir_with_address'].split('\n')[0]
            self.all_data['event_description'] = self.event_text.get(1.0, tk.END).strip()
            self.all_data['investigation_summary'] = self.investigation_text_widget.get(
                1.0, tk.END).strip()
            self.all_data['investigation_1'] = self.all_data['investigation_summary']
            self.all_data['analysis_results'] = self.analysis_text_widget.get(1.0, tk.END).strip()
            self.all_data['analysis_1'] = self.all_data['analysis_results']
            template_path = Path(self.config['template_path']).expanduser()
            out_path = self.save_path.get()            
            fill_docx(str(template_path), out_path, self.all_data, self.products)
            self.saved_path_label.configure(text=out_path)
            self.root.after(0, lambda: self.show_step(4))
            self.root.after(0, lambda: self.generate_btn.set_state("normal", "Generate Document"))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", 
                f"Failed to generate document: {str(e)}"))
            self.root.after(0, lambda: self.generate_btn.set_state("normal", "Generate Document"))
    def reset_app(self):
        self.complaint_id.set('')
        self.save_path.set('')
        self.all_data = {}
        self.products = []
        self.progress_log.configure(state='normal')
        self.progress_log.delete(1.0, tk.END)
        self.progress_log.configure(state='disabled')
        self.ir_text.delete(1.0, tk.END)
        self.event_text.delete(1.0, tk.END)
        self.investigation_text_widget.delete(1.0, tk.END)
        self.analysis_text_widget.delete(1.0, tk.END)
        self.show_step(1)
def main():
    root = tk.Tk()
    try:
        root.iconbitmap('icon.ico')
    except:
        pass
    app = ComplaintLetterGUI(root)
    root.update_idletasks()
    width = 1100
    height = 800
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    root.resizable(False, False)
    root.mainloop()
if __name__ == "__main__":
    main()