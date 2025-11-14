import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

from scrape_and_generate import scrape_complaint, fill_docx

DEFAULT_CONFIG_PATH = Path("config.yaml")


ACCENT = "#4f46e5"      # Indigo-ish
BG_DARK = "#0f172a"     # Dark background
CARD_BG = "#111827"
TEXT_LIGHT = "#e5e7eb"


class ComplaintWizard(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Complaint Letter Wizard")
        self.minsize(800, 550)
        self.configure(bg=BG_DARK)

        # ---- Theming / Styles ----
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Main.TFrame",
            background=BG_DARK,
        )
        style.configure(
            "Card.TFrame",
            background=CARD_BG,
            relief="ridge",
            borderwidth=1,
        )
        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 18, "bold"),
            foreground=TEXT_LIGHT,
            background=BG_DARK,
        )
        style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 11),
            foreground="#9ca3af",
            background=BG_DARK,
        )
        style.configure(
            "CardTitle.TLabel",
            font=("Segoe UI", 13, "bold"),
            foreground=TEXT_LIGHT,
            background=CARD_BG,
        )
        style.configure(
            "CardText.TLabel",
            font=("Segoe UI", 10),
            foreground="#d1d5db",
            background=CARD_BG,
        )
        style.configure(
            "Accent.TButton",
            font=("Segoe UI", 10, "bold"),
            foreground="white",
            background=ACCENT,
            borderwidth=0,
            padding=6,
        )
        style.map(
            "Accent.TButton",
            background=[("active", "#4338ca")],
        )
        style.configure(
            "Ghost.TButton",
            font=("Segoe UI", 10),
            foreground="#e5e7eb",
            background="#1f2937",
            borderwidth=0,
            padding=6,
        )
        style.map(
            "Ghost.TButton",
            background=[("active", "#374151")],
        )
        style.configure(
            "Step.TLabel",
            font=("Segoe UI", 10, "bold"),
            foreground="#9ca3af",
            background=BG_DARK,
        )

        # Data from backend
        self.values = {}
        self.products = []
        self.cfg = None
        self.template_path = None
        self.out_dir = None

        # For analysis step
        self.current_analysis_idx = 0  # 0-based index into products

        self.complaint_var = tk.StringVar()
        self.status_var = tk.StringVar()

        # Top header (title + step indicator)
        header = ttk.Frame(self, style="Main.TFrame", padding=(20, 15, 20, 5))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        self.title_label = ttk.Label(header, text="Complaint Letter Wizard", style="Title.TLabel")
        self.title_label.grid(row=0, column=0, sticky="w")

        self.step_label = ttk.Label(header, text="Step 1 of 4 · Enter Complaint Number", style="Step.TLabel")
        self.step_label.grid(row=1, column=0, sticky="w", pady=(4, 0))

        # Main area where steps are swapped
        self.main_frame = ttk.Frame(self, style="Main.TFrame", padding=(20, 10, 20, 20))
        self.main_frame.grid(row=1, column=0, sticky="nsew")
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        # Step frames (cards)
        self.step1_frame = ttk.Frame(self.main_frame, style="Card.TFrame", padding=15)
        self.step2_frame = ttk.Frame(self.main_frame, style="Card.TFrame", padding=15)
        self.step3_frame = ttk.Frame(self.main_frame, style="Card.TFrame", padding=15)  # analysis
        self.step4_frame = ttk.Frame(self.main_frame, style="Card.TFrame", padding=15)  # investigation+save

        self._build_step1()
        self._build_step2()
        self._build_step3_analysis()
        self._build_step4_investigation()

        self._show_step(self.step1_frame, "Step 1 of 4 · Enter Complaint Number")

    # ------------------------------------------------------------------
    # Step 1: Complaint input + Go
    # ------------------------------------------------------------------
    def _build_step1(self):
        f = self.step1_frame

        ttk.Label(f, text="Enter Complaint Number", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(
            f,
            text="We’ll pull the IR data, event description, product info, analyses, and investigations from CRM.",
            style="CardText.TLabel",
            wraplength=600,
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(5, 15))

        ttk.Label(f, text="Complaint number:", style="CardText.TLabel").grid(row=2, column=0, sticky="w", pady=(0, 5))

        entry = ttk.Entry(f, textvariable=self.complaint_var, width=30)
        entry.grid(row=2, column=1, sticky="w", padx=(5, 0), pady=(0, 5))

        go_btn = ttk.Button(f, text="Go", style="Accent.TButton", command=self.on_go_clicked)
        go_btn.grid(row=2, column=2, padx=(10, 0), pady=(0, 5))

        status = ttk.Label(f, textvariable=self.status_var, style="CardText.TLabel", wraplength=650)
        status.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))

        f.grid_columnconfigure(1, weight=1)

    def on_go_clicked(self):
        complaint_id = self.complaint_var.get().strip()
        if not complaint_id:
            messagebox.showerror("Missing complaint number", "Please enter a complaint number first.")
            return

        self.status_var.set("Contacting CRM and collecting data…")
        self.update_idletasks()

        try:
            values, products, cfg, template_path, out_dir = scrape_complaint(
                complaint_id,
                str(DEFAULT_CONFIG_PATH),
            )
        except Exception as e:
            self.status_var.set("")
            messagebox.showerror("Error", f"Failed to scrape data:\n{e}")
            return

        self.values = values or {}
        self.products = products or []
        self.cfg = cfg
        self.template_path = template_path
        self.out_dir = out_dir

        # Pre-fill IR with address for next step
        ir_text = self.values.get("ir_with_address", "") or ""
        self.ir_text_widget.delete("1.0", "end")
        self.ir_text_widget.insert("1.0", ir_text)

        self.status_var.set("")
        self._show_step(self.step2_frame, "Step 2 of 4 · Edit IR / Address")

    # ------------------------------------------------------------------
    # Step 2: IR with address edit
    # ------------------------------------------------------------------
    def _build_step2(self):
        f = self.step2_frame

        ttk.Label(f, text="Initial Reporter & Facility Address", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            f,
            text="Review and adjust the initial reporter / facility block before it goes into the template.",
            style="CardText.TLabel",
            wraplength=650,
        ).grid(row=1, column=0, sticky="w", pady=(5, 10))

        self.ir_text_widget = tk.Text(f, width=90, height=10, wrap="word", bg="#020617", fg=TEXT_LIGHT, insertbackground=TEXT_LIGHT)
        self.ir_text_widget.grid(row=2, column=0, sticky="nsew", pady=(5, 10))

        btn_frame = ttk.Frame(f, style="Card.TFrame")
        btn_frame.grid(row=3, column=0, sticky="e")

        back_btn = ttk.Button(btn_frame, text="← Back", style="Ghost.TButton",
                              command=lambda: self._show_step(self.step1_frame, "Step 1 of 4 · Enter Complaint Number"))
        back_btn.grid(row=0, column=0, padx=5)

        next_btn = ttk.Button(btn_frame, text="Next · Analyses →", style="Accent.TButton", command=self.on_ir_next)
        next_btn.grid(row=0, column=1, padx=5)

        f.grid_rowconfigure(2, weight=1)
        f.grid_columnconfigure(0, weight=1)

    def on_ir_next(self):
        # Save updated IR text
        self.values["ir_with_address"] = self.ir_text_widget.get("1.0", "end-1c")

        if not self.products:
            # No products → skip analysis step
            self._show_step(self.step4_frame, "Step 4 of 4 · Investigation & Save")
            return

        # Start at first product's analysis
        self.current_analysis_idx = 0
        self._load_current_analysis()
        self._show_step(self.step3_frame, "Step 3 of 4 · Edit Analyses")

    # ------------------------------------------------------------------
    # Step 3: per-product analysis editing
    # ------------------------------------------------------------------
    def _build_step3_analysis(self):
        f = self.step3_frame

        self.analysis_header_label = ttk.Label(
            f, text="Product Analysis", style="CardTitle.TLabel"
        )
        self.analysis_header_label.grid(row=0, column=0, sticky="w")

        self.analysis_product_label = ttk.Label(
            f,
            text="",
            style="CardText.TLabel",
            wraplength=650,
        )
        self.analysis_product_label.grid(row=1, column=0, sticky="w", pady=(4, 10))

        self.analysis_text_widget = tk.Text(
            f,
            width=90,
            height=12,
            wrap="word",
            bg="#020617",
            fg=TEXT_LIGHT,
            insertbackground=TEXT_LIGHT,
        )
        self.analysis_text_widget.grid(row=2, column=0, sticky="nsew", pady=(5, 10))

        btn_frame = ttk.Frame(f, style="Card.TFrame")
        btn_frame.grid(row=3, column=0, sticky="e")

        self.analysis_back_btn = ttk.Button(
            btn_frame, text="← Back", style="Ghost.TButton", command=self.on_analysis_back
        )
        self.analysis_back_btn.grid(row=0, column=0, padx=5)

        self.analysis_next_btn = ttk.Button(
            btn_frame, text="Next Analysis →", style="Accent.TButton", command=self.on_analysis_next
        )
        self.analysis_next_btn.grid(row=0, column=1, padx=5)

        f.grid_rowconfigure(2, weight=1)
        f.grid_columnconfigure(0, weight=1)

    def _load_current_analysis(self):
        """Load analysis text for the current product index into the text widget."""
        if not self.products:
            self.analysis_product_label.config(text="No products found.")
            self.analysis_text_widget.delete("1.0", "end")
            return

        idx = self.current_analysis_idx  # 0-based
        product_num = idx + 1

        # Get product info
        prod = self.products[idx] if idx < len(self.products) else {}
        pid = (prod.get("id") or self.values.get(f"product_id_{product_num}", "") or "").strip()
        pdesc = (prod.get("desc") or self.values.get(f"product_desc_{product_num}", "") or "").strip()

        header_text = f"Analysis for Product {product_num} of {len(self.products)}"
        self.analysis_header_label.config(text=header_text)

        product_line = f"Product: {pid or '(no ID)'}"
        if pdesc:
            product_line += f" — {pdesc}"
        self.analysis_product_label.config(text=product_line)

        # Fill analysis text
        analysis_key = f"analysis_{product_num}"
        analysis_text = self.values.get(analysis_key, "") or ""
        self.analysis_text_widget.delete("1.0", "end")
        self.analysis_text_widget.insert("1.0", analysis_text)

        # Button label: last one goes to Investigation step
        if product_num == len(self.products):
            self.analysis_next_btn.config(text="Next · Investigation →")
        else:
            self.analysis_next_btn.config(text="Next Analysis →")

    def _save_current_analysis(self):
        if not self.products:
            return
        idx = self.current_analysis_idx
        product_num = idx + 1
        analysis_key = f"analysis_{product_num}"
        text = self.analysis_text_widget.get("1.0", "end-1c")
        self.values[analysis_key] = text

    def on_analysis_next(self):
        if self.products:
            self._save_current_analysis()

        if not self.products or self.current_analysis_idx >= len(self.products) - 1:
            # Done with analyses → Investigation step
            self._show_step(self.step4_frame, "Step 4 of 4 · Investigation & Save")
        else:
            self.current_analysis_idx += 1
            self._load_current_analysis()
            self._show_step(self.step3_frame, "Step 3 of 4 · Edit Analyses")

    def on_analysis_back(self):
        if not self.products:
            self._show_step(self.step2_frame, "Step 2 of 4 · Edit IR / Address")
            return

        # Save current before moving
        self._save_current_analysis()

        if self.current_analysis_idx == 0:
            # Back to IR step
            self._show_step(self.step2_frame, "Step 2 of 4 · Edit IR / Address")
        else:
            self.current_analysis_idx -= 1
            self._load_current_analysis()
            self._show_step(self.step3_frame, "Step 3 of 4 · Edit Analyses")

    # ------------------------------------------------------------------
    # Step 4: Investigation summary + Save
    # ------------------------------------------------------------------
    def _build_step4_investigation(self):
        f = self.step4_frame

        ttk.Label(f, text="Investigation Summary", style="CardTitle.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            f,
            text="Review and edit the investigation summary before generating the final Word document.",
            style="CardText.TLabel",
            wraplength=650,
        ).grid(row=1, column=0, sticky="w", pady=(5, 10))

        self.inv_text_widget = tk.Text(
            f,
            width=90,
            height=10,
            wrap="word",
            bg="#020617",
            fg=TEXT_LIGHT,
            insertbackground=TEXT_LIGHT,
        )
        self.inv_text_widget.grid(row=2, column=0, sticky="nsew", pady=(5, 10))

        btn_frame = ttk.Frame(f, style="Card.TFrame")
        btn_frame.grid(row=3, column=0, sticky="e")

        back_btn = ttk.Button(
            btn_frame,
            text="← Back to Analyses",
            style="Ghost.TButton",
            command=self.on_inv_back,
        )
        back_btn.grid(row=0, column=0, padx=5)

        save_btn = ttk.Button(
            btn_frame,
            text="Save Letter…",
            style="Accent.TButton",
            command=self.on_save_clicked,
        )
        save_btn.grid(row=0, column=1, padx=5)

        f.grid_rowconfigure(2, weight=1)
        f.grid_columnconfigure(0, weight=1)

    def on_inv_back(self):
        # Back to analyses if there are products, else to IR
        if self.products:
            # Load the current analysis again for continuity
            self._load_current_analysis()
            self._show_step(self.step3_frame, "Step 3 of 4 · Edit Analyses")
        else:
            self._show_step(self.step2_frame, "Step 2 of 4 · Edit IR / Address")

    def _update_combined_analysis_results(self):
        """Rebuild values['analysis_results'] from per-product analysis_N fields."""
        if not self.products:
            return
        blocks = []
        for i in range(1, len(self.products) + 1):
            txt = (self.values.get(f"analysis_{i}") or "").strip()
            if txt:
                blocks.append(txt)
        if blocks:
            self.values["analysis_results"] = "\n\n".join(blocks)

    def on_save_clicked(self):
        # Save investigation edits
        self.values["investigation_summary"] = self.inv_text_widget.get("1.0", "end-1c")

        # Ensure combined analysis_results is in sync with per-product edits
        self._update_combined_analysis_results()

        if not self.template_path:
            messagebox.showerror("Error", "Template path is not set.")
            return

        # Build default filename from pattern
        pattern = (self.cfg or {}).get("output_name_pattern", "Customer_Letter_{complaint_id}.docx")
        try:
            default_name = pattern.format(**self.values)
        except Exception:
            default_name = f"Customer_Letter_{self.values.get('complaint_id', 'letter')}.docx"

        initial_dir = str(self.out_dir) if self.out_dir else "."

        save_path = filedialog.asksaveasfilename(
            title="Save generated letter",
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
            initialdir=initial_dir,
            initialfile=default_name,
        )
        if not save_path:
            return

        try:
            fill_docx(str(self.template_path), save_path, self.values, self.products)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate document:\n{e}")
            return

        messagebox.showinfo("Done", f"Letter generated and saved to:\n{save_path}")

    # ------------------------------------------------------------------
    # Shared helper: show step + update step label
    # ------------------------------------------------------------------
    def _show_step(self, frame_to_show: ttk.Frame, step_text: str):
        for f in (self.step1_frame, self.step2_frame, self.step3_frame, self.step4_frame):
            f.grid_forget()
        frame_to_show.grid(row=0, column=0, sticky="nsew")
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.step_label.config(text=step_text)

        # When entering investigation step, pre-fill its text from values
        if frame_to_show is self.step4_frame:
            inv_text = self.values.get("investigation_summary", "") or ""
            self.inv_text_widget.delete("1.0", "end")
            self.inv_text_widget.insert("1.0", inv_text)


if __name__ == "__main__":
    app = ComplaintWizard()
    app.mainloop()
