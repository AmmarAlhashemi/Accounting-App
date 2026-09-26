from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import config
from core.excel_handler import ExcelHandler
from core.pdf_exporter import export_excel_to_pdf
from core.reconciler import perform_reconciliation


class NNN_Accounting_App:

    def __init__(self, root):
        self.root = root
        self.root.title("NNN AC Data Entry")
        self.root.geometry("1220x690+65+0")

        self.excel_handler = ExcelHandler(config.BASE_EXCEL_PATH)
        self.transaction_cache = []
        self.is_new_balance = tk.BooleanVar(value=True)

        # --- Theme variables ---
        self.is_light = tk.BooleanVar(value=False)
        self.ui_elements = {
            "bg_main": [],
            "bg_sticky": [],
            "lbl_sticky": [],
            "lbl_main": [],
            "lbl_dimmed": [],
            "entry": [],
            "labelframe": [],
        }

        self.create_widgets()
        self.load_rate_and_ac_no()

        # Keyboard Bindings
        self.root.bind("<Return>", lambda event: self.add_to_cache())
        self.root.bind("<Delete>", lambda event: self.clear_inputs())
        self.root.bind("<Control-s>", lambda event: self.process_excel())
        self.root.bind("<Control-S>", lambda event: self.process_excel())
        self.root.bind("<Control-p>", lambda event: self.print_to_pdf())
        self.root.bind("<Control-P>", lambda event: self.print_to_pdf())
        self.root.bind_all("<Alt-d>", self.focus_date_day)
        self.root.bind_all("<Alt-D>", self.focus_date_day)
        self.root.bind_all("<Alt-r>", self.focus_rate)
        self.root.bind_all("<Alt-R>", self.focus_rate)
        self.root.bind_all("<Alt-a>", self.focus_ac)
        self.root.bind_all("<Alt-A>", self.focus_ac)

    def focus_date_day(self, event=None):
        self.ent_date.focus_set()
        self.ent_date.selection_range(0, 2)
        self.ent_date.icursor(2)
        return "break"

    def focus_rate(self, event=None):
        self.ent_rate.focus_set()
        self.ent_rate.select_range(0, tk.END)
        self.ent_rate.icursor(tk.END)
        return "break"

    def focus_ac(self, event=None):
        self.ent_ac_no.focus_set()
        self.ent_ac_no.select_range(0, tk.END)
        self.ent_ac_no.icursor(tk.END)
        return "break"

    def create_widgets(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # --- Main Container ---
        main_container = tk.Frame(self.root)
        self.ui_elements["bg_main"].append(main_container)
        main_container.pack(fill="both", expand=True)

        # --- Left Panel: Inputs & Actions ---
        left_panel = tk.Frame(main_container, width=400, padx=20, pady=10)
        self.ui_elements["bg_main"].append(left_panel)
        left_panel.pack(side="left", fill="y")

        # Theme Switch
        self.chk_theme = tk.Checkbutton(
            left_panel,
            text="Light Mode ☀️",
            variable=self.is_light,
            command=self.update_theme,
            font=config.FONT_BOLD,
        )
        self.chk_theme.pack(anchor="w", pady=(0, 10))

        # --- Sticky Data ---
        f_sticky = tk.Frame(
            left_panel, padx=15, pady=15, relief="groove", bd=2
        )
        self.ui_elements["bg_sticky"].append(f_sticky)
        f_sticky.pack(fill="x", pady=(0, 15))

        # Date
        date_container = tk.Frame(f_sticky)
        self.ui_elements["bg_sticky"].append(date_container)
        date_container.pack(side="left", expand=True, fill="x", padx=5)

        lbl_date = tk.Label(
            date_container, text="Date:", font=config.FONT_BOLD
        )
        self.ui_elements["lbl_sticky"].append(lbl_date)
        lbl_date.pack(anchor="w")

        self.ent_date = tk.Entry(
            date_container,
            width=12,
            font=config.FONT_NORM,
            relief="solid",
            bd=1,
        )
        self.ui_elements["entry"].append(self.ent_date)
        self.ent_date.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.ent_date.pack(anchor="w")

        # Rate
        rate_container = tk.Frame(f_sticky)
        self.ui_elements["bg_sticky"].append(rate_container)
        rate_container.pack(side="left", expand=True, fill="x", padx=(20, 5))

        lbl_rate = tk.Label(
            rate_container, text="Last Rate:", font=config.FONT_BOLD
        )
        self.ui_elements["lbl_sticky"].append(lbl_rate)
        lbl_rate.pack(anchor="w")

        self.ent_rate = tk.Entry(
            rate_container,
            width=10,
            font=config.FONT_NORM,
            relief="solid",
            bd=1,
        )
        self.ui_elements["entry"].append(self.ent_rate)
        self.ent_rate.pack(anchor="w")

        # AC No
        ac_container = tk.Frame(f_sticky)
        self.ui_elements["bg_sticky"].append(ac_container)
        ac_container.pack(side="left", expand=True, fill="x", padx=(20, 5))

        lbl_ac = tk.Label(ac_container, text="Last AC:", font=config.FONT_BOLD)
        self.ui_elements["lbl_sticky"].append(lbl_ac)
        lbl_ac.pack(anchor="w")

        self.ent_ac_no = tk.Entry(
            ac_container,
            width=10,
            font=config.FONT_NORM,
            relief="solid",
            bd=1,
        )
        self.ui_elements["entry"].append(self.ent_ac_no)
        self.ent_ac_no.pack(anchor="w")

        # --- Transaction Entry Form ---
        f_entry = tk.LabelFrame(
            left_panel,
            text=" New Transaction ",
            padx=10,
            pady=10,
            font=config.FONT_BOLD,
        )
        self.ui_elements["labelframe"].append(f_entry)
        f_entry.pack(fill="x", pady=5)
        f_entry.columnconfigure(1, weight=1)

        lbl_t = tk.Label(f_entry, text="Type:", font=config.FONT_NORM)
        self.ui_elements["lbl_main"].append(lbl_t)
        lbl_t.grid(row=0, column=0, sticky="w")

        self.combo_type = ttk.Combobox(
            f_entry,
            values=["TRN", "OTHERS", "ORDER"],
            state="readonly",
            font=config.FONT_NORM,
        )
        self.combo_type.current(0)
        self.combo_type.grid(row=0, column=1, pady=5, padx=5, sticky="ew")
        self.combo_type.bind("<<ComboboxSelected>>", self.on_type_change)

        self.lbl_s = tk.Label(f_entry, text="Seq:", font=config.FONT_NORM)
        self.ui_elements["lbl_main"].append(self.lbl_s)
        self.lbl_s.grid(row=1, column=0, sticky="w")

        self.ent_seq = tk.Entry(
            f_entry, font=config.FONT_NORM, relief="solid", bd=1
        )
        self.ui_elements["entry"].append(self.ent_seq)
        self.ent_seq.grid(row=1, column=1, pady=5, padx=5, sticky="ew")

        lbl_d = tk.Label(f_entry, text="Description:", font=config.FONT_NORM)
        self.ui_elements["lbl_main"].append(lbl_d)
        lbl_d.grid(row=2, column=0, sticky="w")

        self.ent_desc = tk.Entry(
            f_entry, width=25, font=config.FONT_NORM, relief="solid", bd=1
        )
        self.ui_elements["entry"].append(self.ent_desc)
        self.ent_desc.insert(0, "BOA DEPOSIT TO A/C NO.")
        self.ent_desc.grid(row=2, column=1, pady=5, padx=5, sticky="ew")

        lbl_a = tk.Label(f_entry, text="Amount:", font=config.FONT_NORM)
        self.ui_elements["lbl_main"].append(lbl_a)
        lbl_a.grid(row=3, column=0, sticky="w")

        self.ent_amount = tk.Entry(
            f_entry, font=config.FONT_NORM, relief="solid", bd=1
        )
        self.ui_elements["entry"].append(self.ent_amount)
        self.ent_amount.grid(row=3, column=1, pady=3, padx=5, sticky="ew")

        # Form Controls
        f_entry_actions = tk.Frame(f_entry)
        self.ui_elements["bg_main"].append(f_entry_actions)
        f_entry_actions.grid(
            row=4,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(10, 2),
            padx=(0, 5),
        )

        lbl_marked = tk.Label(
            f_entry_actions, text="Org Date:", font=config.FONT_NORM
        )
        self.ui_elements["lbl_main"].append(lbl_marked)
        lbl_marked.pack(side="left", padx=(0, 2))

        self.ent_marked_date = tk.Entry(
            f_entry_actions,
            width=7,
            font=config.FONT_NORM,
            relief="solid",
            bd=1,
        )
        self.ui_elements["entry"].append(self.ent_marked_date)
        self.ent_marked_date.pack(side="left", padx=(0, 8))

        self.btn_clear = tk.Button(
            f_entry_actions,
            command=self.clear_inputs,
            text="Clear",
            font=("Segoe UI", 10),
            relief="ridge",
            cursor="hand2",
        )
        self.btn_clear.pack(side="left", fill="x", expand=True, padx=(0, 5))

        self.btn_add = tk.Button(
            f_entry_actions,
            command=self.add_to_cache,
            text="➕ Add",
            font=config.FONT_BOLD,
            relief="ridge",
            cursor="hand2",
        )
        self.btn_add.pack(side="left", fill="x", expand=True)

        excel_frame = tk.Frame(left_panel)
        self.ui_elements["bg_main"].append(excel_frame)
        excel_frame.pack(fill="x", pady=(10, 5))
        excel_frame.columnconfigure(0, weight=1)
        excel_frame.columnconfigure(1, weight=1)

        self.chk_new_balance = tk.Checkbutton(
            excel_frame,
            text="New Opening Balance",
            variable=self.is_new_balance,
            font=config.FONT_NORM,
            anchor="w",
        )
        self.chk_new_balance.grid(
            row=0, column=0, sticky="ew", padx=0, pady=(0, 10)
        )

        self.btn_save_all = tk.Button(
            excel_frame,
            command=self.process_excel,
            text="Save to Excel",
            font=config.FONT_BTN1,
            height=1,
            relief="ridge",
            cursor="hand2",
        )
        self.btn_save_all.grid(
            row=0, column=1, sticky="ew", padx=(3, 0), pady=(0, 10)
        )

        self.btn_print = tk.Button(
            left_panel,
            command=self.print_to_pdf,
            text="Print PDF",
            font=config.FONT_BOLD,
            height=1,
            relief="ridge",
            cursor="hand2",
        )
        self.btn_print.pack(fill="x", pady=(0, 10))

        nnn_frame = tk.LabelFrame(
            left_panel,
            text=" Reconciliation ",
            padx=10,
            pady=3,
            font=config.FONT_BOLD,
        )
        self.ui_elements["labelframe"].append(nnn_frame)
        nnn_frame.pack(fill="x", pady=5)

        self.btn_compare = tk.Button(
            nnn_frame,
            command=self.reconcile,
            text="Compare NNN AC",
            font=config.FONT_BOLD,
            relief="ridge",
            cursor="hand2",
        )
        self.btn_compare.pack(fill="x", pady=(0, 5))

        # --- Shortcuts ---
        f_shortcuts = tk.LabelFrame(
            left_panel,
            text=" Quick Shortcuts ",
            padx=10,
            pady=3,
            font=config.FONT_BOLD,
        )
        self.ui_elements["labelframe"].append(f_shortcuts)
        f_shortcuts.pack(side="bottom", fill="x", pady=3)

        shortcuts = [
            "Alt + D → Date | Alt + R → Rate | Alt + A → AC",
            "Enter → Add to Queue | Del → Clear Inputs",
            "Ctrl+S → Save to Excel | Ctrl + P → Print PDF",
        ]
        for s in shortcuts:
            lbl_sc = tk.Label(f_shortcuts, text=s, font=config.FONT_CODE, anchor="w")
            self.ui_elements["lbl_dimmed"].append(lbl_sc)
            lbl_sc.pack(fill="x")

        # --- Right Panel: Queue Viewer ---
        right_panel = tk.Frame(main_container, padx=15, pady=20)
        self.ui_elements["bg_main"].append(right_panel)
        right_panel.pack(side="right", fill="both", expand=True)

        self.lbl_status = tk.Label(
            right_panel,
            text="Queue: 0 transactions",
            font=("Segoe UI", 11, "italic"),
        )
        self.lbl_status.pack(anchor="w", pady=(0, 5))

        f_view = tk.LabelFrame(
            right_panel,
            text=" Transactions in Queue (Not Saved Yet) ",
            padx=5,
            pady=5,
            font=config.FONT_BOLD,
        )
        self.ui_elements["labelframe"].append(f_view)
        f_view.pack(fill="both", expand=True)

        columns = ("date", "seq", "desc", "rate", "amount")
        self.tree = ttk.Treeview(f_view, columns=columns, show="headings")

        self.tree.heading("date", text="Date")
        self.tree.heading("seq", text="Ord / Seq")
        self.tree.heading("desc", text="Description")
        self.tree.heading("rate", text="Rate")
        self.tree.heading("amount", text="Amount")

        self.tree.column("date", width=100, anchor="center")
        self.tree.column("seq", width=100, anchor="center")
        self.tree.column("desc", width=230, anchor="w")
        self.tree.column("rate", width=90, anchor="center")
        self.tree.column("amount", width=110, anchor="e")

        self.tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            f_view, orient="vertical", command=self.tree.yview
        )
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # --- Right Panel Action Buttons ---
        f_tree_actions = tk.Frame(right_panel)
        self.ui_elements["bg_main"].append(f_tree_actions)
        f_tree_actions.pack(fill="x", pady=(3, 0))
        f_tree_actions.columnconfigure(0, weight=1)
        f_tree_actions.columnconfigure(1, weight=1)

        self.btn_delete_selected = tk.Button(
            f_tree_actions,
            command=self.delete_selected_row,
            text="❌ Delete Selected",
            font=config.FONT_BOLD,
            relief="ridge",
            cursor="hand2",
        )
        self.btn_delete_selected.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        self.btn_delete_all = tk.Button(
            f_tree_actions,
            command=self.delete_all_rows,
            text="🗑️Delete All",
            font=config.FONT_BOLD,
            relief="ridge",
            cursor="hand2",
        )
        self.btn_delete_all.grid(row=0, column=1, padx=(5, 0), sticky="ew")

        self.update_theme()

    def update_theme(self):
        if self.is_light.get():
            self.chk_theme.config(text="Light Mode ☀️")
            BG_MAIN, BG_STICKY = "#F8FAFC", "#E2E8F0"
            FG_TEXT, FG_DIMMED = "#0F172A", "#475569"
            ENTRY_BG, BTN_CLEAR_BG, BTN_CLEAR_ACTIVE = (
                "#FFFFFF",
                "#94A3B8",
                "#64748B",
            )
            STATUS_FG = "#2563EB"
        else:
            self.chk_theme.config(text="Dark Mode 🌙")
            BG_MAIN, BG_STICKY = "#1C2535", "#0F172A"
            FG_TEXT, FG_DIMMED = "#F8FAFC", "#94A3B8"
            ENTRY_BG, BTN_CLEAR_BG, BTN_CLEAR_ACTIVE = (
                "#334155",
                "#475569",
                "#334155",
            )
            STATUS_FG = "#85BCFF"

        BTN_ADD_BG, BTN_ADD_ACTIVE = "#1259CA", "#04349B"
        BTN_SAVE_BG, BTN_SAVE_ACTIVE = "#008A5C", "#005F41"
        TREE_SELECT_BG, TREE_SELECT_FG = "#1C5FCA", "white"
        BTN_PRINT_BG, BTN_PRINT_ACTIVE = "#8B0F4D", "#630031"
        BTN_COMP, BTN_COMP_ACTIVE = "#0C5E94", "#004E81"
        BTN_DEL_BG, BTN_DEL_ACTIVE = "#911515", "#691818"

        self.root.configure(bg=BG_MAIN)
        for w in self.ui_elements["bg_main"]:
            w.configure(bg=BG_MAIN)
        for w in self.ui_elements["bg_sticky"]:
            w.configure(bg=BG_STICKY)
        for w in self.ui_elements["lbl_sticky"]:
            w.configure(bg=BG_STICKY, fg=FG_TEXT)
        for w in self.ui_elements["lbl_main"]:
            w.configure(bg=BG_MAIN, fg=FG_TEXT)
        for w in self.ui_elements["lbl_dimmed"]:
            w.configure(bg=BG_MAIN, fg=FG_DIMMED)
        for w in self.ui_elements["entry"]:
            w.configure(bg=ENTRY_BG, fg=FG_TEXT, insertbackground=FG_TEXT)
        for w in self.ui_elements["labelframe"]:
            w.configure(bg=BG_MAIN, fg=FG_TEXT)

        self.chk_theme.configure(
            bg=BG_MAIN,
            fg=FG_TEXT,
            selectcolor=ENTRY_BG,
            activebackground=BG_MAIN,
            activeforeground=FG_TEXT,
        )
        self.chk_new_balance.configure(
            bg=BG_MAIN,
            fg=FG_TEXT,
            selectcolor=ENTRY_BG,
            activebackground=BG_MAIN,
            activeforeground=FG_TEXT,
        )
        self.btn_add.configure(
            bg=BTN_ADD_BG,
            fg="white",
            activebackground=BTN_ADD_ACTIVE,
            activeforeground="white",
        )
        self.btn_save_all.configure(
            bg=BTN_SAVE_BG,
            fg="white",
            activebackground=BTN_SAVE_ACTIVE,
            activeforeground="white",
        )
        self.btn_print.configure(
            bg=BTN_PRINT_BG,
            fg="white",
            activebackground=BTN_PRINT_ACTIVE,
            activeforeground="white",
        )
        self.btn_clear.configure(
            bg=BTN_CLEAR_BG,
            fg="white",
            activebackground=BTN_CLEAR_ACTIVE,
            activeforeground="white",
        )
        self.btn_compare.configure(
            bg=BTN_COMP,
            fg="white",
            activebackground=BTN_COMP_ACTIVE,
            activeforeground="white",
        )
        self.btn_delete_selected.configure(
            bg=BTN_DEL_BG,
            fg="white",
            activebackground=BTN_DEL_ACTIVE,
            activeforeground="white",
        )
        self.btn_delete_all.configure(
            bg=BTN_DEL_BG,
            fg="white",
            activebackground=BTN_DEL_ACTIVE,
            activeforeground="white",
        )
        self.lbl_status.configure(bg=BG_MAIN, fg=STATUS_FG)

        self.style.configure(
            "Treeview",
            font=config.FONT_NORM,
            rowheight=28,
            background=BG_MAIN,
            fieldbackground=BG_MAIN,
            foreground=FG_TEXT,
            borderwidth=0,
        )
        self.style.map(
            "Treeview",
            background=[("selected", TREE_SELECT_BG)],
            foreground=[("selected", TREE_SELECT_FG)],
        )
        self.style.configure(
            "Treeview.Heading",
            font=config.FONT_BOLD,
            background=BG_STICKY,
            foreground=FG_TEXT,
            borderwidth=1,
        )

        self.style.configure(
            "TCombobox",
            font=config.FONT_NORM,
            fieldbackground=ENTRY_BG,
            background=BG_MAIN,
            foreground=FG_TEXT,
            arrowcolor=FG_TEXT,
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", ENTRY_BG)],
            foreground=[("readonly", FG_TEXT)],
            selectbackground=[("readonly", TREE_SELECT_BG)],
            selectforeground=[("readonly", TREE_SELECT_FG)],
        )

    def load_rate_and_ac_no(self):
        rate, ac_no = self.excel_handler.load_rate_and_ac_no()
        if rate:
            self.ent_rate.delete(0, tk.END)
            self.ent_rate.insert(0, rate)
        if ac_no:
            self.ent_ac_no.delete(0, tk.END)
            self.ent_ac_no.insert(0, ac_no)

    def add_to_cache(self):
        try:
            desc = self.ent_desc.get().upper()
            if self.combo_type.get() == "TRN":
                desc += self.ent_ac_no.get()

            marked_date = self.ent_marked_date.get().strip()
            if marked_date:
                desc += f" {marked_date}"

            data = {
                "date": self.ent_date.get(),
                "type": self.combo_type.get(),
                "seq": (
                    self.ent_seq.get()
                    if self.ent_seq.get().isdigit()
                    else self.ent_seq.get().upper()
                ),
                "desc": desc,
                "amount": float(self.ent_amount.get()),
                "rate": float(self.ent_rate.get()),
            }
            if not data["seq"]:
                raise ValueError

            self.transaction_cache.append(data)
            self.update_queue_display()
            self.clear_inputs()
            self.ent_seq.focus_set()
        except ValueError:
            messagebox.showwarning("Warning", "Check seq and Amount!")

    def update_queue_display(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        last_item = None
        for item in self.transaction_cache:
            last_item = self.tree.insert(
                "",
                "end",
                values=(
                    item["date"],
                    item["seq"],
                    item["desc"],
                    (
                        f"{item['rate']:,.2f}"
                        if item["type"] == "ORDER"
                        else ""
                    ),
                    f"${item['amount']:,.2f}",
                ),
            )

        self.lbl_status.config(
            text=f"Queue: {len(self.transaction_cache)} transactions"
        )
        if last_item:
            self.tree.see(last_item)

    def delete_selected_row(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a row to delete!")
            return
        idx = self.tree.index(selected_item[0])
        del self.transaction_cache[idx]
        self.update_queue_display()

    def delete_all_rows(self):
        if not self.transaction_cache:
            messagebox.showinfo("Info", "Queue is already empty!")
            return
        if messagebox.askyesno(
            "Confirm Delete", "Are you sure you want to clear the entire queue?"
        ):
            self.transaction_cache = []
            self.update_queue_display()

    def process_excel(self):
        if not self.transaction_cache:
            messagebox.showwarning(
                "Empty", "No transactions in queue to save!"
            )
            return
        if not os.path.exists(config.BASE_EXCEL_PATH):
            messagebox.showerror(
                "Error", f"File {config.BASE_EXCEL_PATH} not found!"
            )
            return

        try:
            self.excel_handler.process_excel_batch(
                self.transaction_cache, self.is_new_balance.get()
            )
            messagebox.showinfo(
                "Success",
                f"Successfully processed {len(self.transaction_cache)} entries!",
            )
            self.transaction_cache = []
            self.lbl_status.config(text="Queue: 0 transactions")
            self.is_new_balance.set(False)
            self.update_queue_display()
        except Exception as e:
            messagebox.showerror(
                "Critical Error", f"Failed to save batch: {e}"
            )

    def print_to_pdf(self):
        if not os.path.exists(config.BASE_EXCEL_PATH):
            messagebox.showerror(
                "Error", f"File {config.BASE_EXCEL_PATH} not found!"
            )
            return

        num_blocks = simpledialog.askinteger(
            "Print PDF",
            "How many blocks would you like to print?",
            minvalue=1,
            initialvalue=1,
        )
        if num_blocks is None:
            return

        try:
            pdf_path = export_excel_to_pdf(
                config.BASE_EXCEL_PATH, config.BASE_PDF_DIR, num_blocks
            )
            open_file = messagebox.askyesno(
                "Success",
                f"{os.path.basename(pdf_path)} created successfully\n\nDo you want to open it?",
            )
            if open_file:
                os.startfile(pdf_path)
        except Exception as e:
            messagebox.showerror(
                "Critical Error", f"Failed to create PDF file.\n{e}"
            )

    def reconcile(self):
        messagebox.showinfo(
            "Alert",
            f"Before you select the pdf, mark the range of target data in your file {os.path.basename(config.BASE_EXCEL_PATH)}:\n\n"
            "1. Open the Excel file.\n"
            "2. Locate Column I.\n"
            "3. Type REC in Column I on the first row of your target data.\n"
            "4. Type REC in Column I on the last row of your target data.\n",
        )

        pdf_file = filedialog.askopenfilename(
            filetypes=[("PDF files", "*.pdf")]
        )
        if not pdf_file:
            return

        success = perform_reconciliation(
            config.BASE_EXCEL_PATH, pdf_file, config.RECON_OUTPUT_PATH
        )
        if success:
            open_file = messagebox.askyesno(
                "Success",
                f"The file created successfully at:\n{config.RECON_OUTPUT_PATH}\n\nDo you want to open it?",
            )
            if open_file:
                os.startfile(config.RECON_OUTPUT_PATH)
        else:
            messagebox.showerror(
                "Error", "Reconciliation failed. Check console or file format."
            )

    def on_type_change(self, event=None):
        selection = self.combo_type.get()
        self.ent_desc.delete(0, tk.END)
        self.ent_seq.delete(0, tk.END)

        if selection == "TRN":
            self.lbl_s.config(text="Seq:")
            self.ent_desc.insert(0, "BOA DEPOSIT TO A/C NO.")
        elif selection == "OTHERS":
            self.lbl_s.config(text="Country:")
        elif selection == "ORDER":
            self.lbl_s.config(text="Order No:")

    def clear_inputs(self):
        self.ent_seq.delete(0, tk.END)
        self.ent_amount.delete(0, tk.END)
        self.ent_marked_date.delete(0, tk.END)
        self.on_type_change()