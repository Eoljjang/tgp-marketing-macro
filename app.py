import os
from tkinter import filedialog, messagebox
import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD
import xlwings as xw

from excel_processor import MarketingMacroProcessor

# Set up CustomTkinter appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class MarketingMacroApp(ctk.CTk, TkinterDnD.DnDWrapper):

  def __init__(self):
    super().__init__()

    # Initialize TkinterDnD
    self.TkdndVersion = TkinterDnD._require(self)

    # Window Configuration
    self.title("Work in progress")
    self.geometry("750x1050")
    self.resizable(False, False)

    # File paths storage (1 main file + 7 support files)
    self.marketing_file = None
    self.support_files = [None] * 7

    # Base support file labels
    self.support_file_labels = [
        "AD (Macro # - 2) - Layout",
        "AD (Macro # - 1) - Layout",
        "Made in Canada",
        "SKURank",
        "Mccormick, Kraft, Nestle",
        "HQ",
        "TGP PS & Made in Report",
    ]

    self.create_widgets()

  def create_widgets(self):
    # Main Title Label
    title_label = ctk.CTkLabel(
        self,
        text="Work in Progress",
        font=ctk.CTkFont(size=24, weight="bold"),
    )
    title_label.pack(pady=(20, 10))

    # Scrollable frame for cleaner layout
    main_frame = ctk.CTkScrollableFrame(
        self, width=700, height=860, fg_color="transparent"
    )
    main_frame.pack(fill="both", expand=True, padx=20, pady=10)

    # 0. Marketing Macro # Input Section
    input_frame = ctk.CTkFrame(
        main_frame, fg_color="#2b2b2b", corner_radius=8, height=60
    )
    input_frame.pack(fill="x", pady=6)
    input_frame.pack_propagate(False)

    macro_label = ctk.CTkLabel(
        input_frame,
        text="Marketing Macro #",
        font=ctk.CTkFont(size=14, weight="bold"),
    )
    macro_label.pack(side="left", padx=15)

    self.macro_entry = ctk.CTkEntry(
        input_frame,
        placeholder_text="e.g. 15 or 20",
        width=200,
        height=32,
        font=ctk.CTkFont(size=13),
    )
    self.macro_entry.pack(side="right", padx=15)

    # Separator
    sep1 = ctk.CTkFrame(main_frame, height=2, fg_color="#333333")
    sep1.pack(fill="x", pady=15)

    # 1. Big Upload Area (Marketing Macro)
    self.create_upload_section(
        parent=main_frame,
        title_text="Upload your Marketing Macro file with Extract",
        is_big=True,
        index=None,
    )

    # Separator
    separator = ctk.CTkFrame(main_frame, height=2, fg_color="#333333")
    separator.pack(fill="x", pady=15)

    # Section label for support files
    support_label = ctk.CTkLabel(
        main_frame,
        text="Support Files (7 Required)",
        font=ctk.CTkFont(size=16, weight="bold"),
    )
    support_label.pack(anchor="w", pady=(0, 5))

    # 2. Seven Smaller Upload Areas for specific support files
    for i in range(7):
      self.create_upload_section(
          parent=main_frame,
          title_text=f"Upload your {self.support_file_labels[i]} file",
          is_big=False,
          index=i,
      )

    # 3. Generate Button
    self.generate_btn = ctk.CTkButton(
        main_frame,
        text="Generate marketing macro",
        font=ctk.CTkFont(size=15, weight="bold"),
        height=45,
        fg_color="#1f6aa5",
        hover_color="#144870",
        command=self.on_generate_click,
    )
    self.generate_btn.pack(fill="x", pady=(25, 5))

    # Loading status text and spinner container
    self.loading_label = ctk.CTkLabel(
        main_frame,
        text="",
        font=ctk.CTkFont(size=13, slant="italic", weight="bold"),
        text_color="#3498db",
    )
    self.loading_label.pack(pady=(5, 2))

    self.progress_bar = ctk.CTkProgressBar(
        main_frame, mode="indeterminate", width=300, height=12
    )

  def create_upload_section(self, parent, title_text, is_big, index):
    frame = ctk.CTkFrame(
        parent,
        fg_color="#2b2b2b",
        corner_radius=8,
        height=100 if is_big else 60,
    )
    frame.pack(fill="x", pady=6)
    frame.pack_propagate(False)

    label = ctk.CTkLabel(
        frame, text=title_text, font=ctk.CTkFont(size=13, weight="normal")
    )
    label.pack(side="left", padx=15)

    status_label = ctk.CTkLabel(
        frame,
        text="Drag & drop or browse file",
        font=ctk.CTkFont(size=12, slant="italic"),
        text_color="gray",
    )
    status_label.pack(side="right", padx=15)

    # Click to upload button
    upload_btn = ctk.CTkButton(
        frame,
        text="Browse...",
        width=90,
        height=32,
        command=lambda: self.browse_file(is_big, index, status_label),
    )
    upload_btn.pack(side="right", padx=5)

    # Register drop target for drag-and-drop capability using tkinterdnd2
    for widget in (frame, label, status_label):
      widget.drop_target_register(DND_FILES)
      widget.dnd_bind(
          "<<Drop>>",
          lambda e: self.handle_drop(e.data, is_big, index, status_label),
      )

  def validate_support_filename(self, index, filename):
    fn_lower = filename.lower()
    if index in (0, 1):
      return True
    elif index == 2:
      return "made in canada" in fn_lower
    elif index == 3:
      return "skurank" in fn_lower
    elif index == 4:
      return "mccormick" in fn_lower
    elif index == 5:
      return "hq" in fn_lower
    elif index == 6:
      return "tgp ps" in fn_lower or "made in report" in fn_lower
    return True

  def process_file_path(self, file_path, is_big, index, status_label):
    file_path = file_path.strip('{}""\'')

    if not os.path.exists(file_path):
      messagebox.showerror(
          "Error", f"The dropped path does not exist:\n{file_path}"
      )
      return

    allowed_exts = (
        (".xlsx", ".xls", ".xlsm", ".csv")
        if (not is_big and index == 5)
        else (".xlsx", ".xls", ".xlsm")
    )

    if not file_path.lower().endswith(allowed_exts):
      messagebox.showerror(
          "Invalid File Type",
          "Please provide a valid file matching the required format.",
      )
      return

    file_name = os.path.basename(file_path)

    if not is_big and not self.validate_support_filename(index, file_name):
      expected_patterns = [
          "",
          "",
          "Made in Canada",
          "SKURank",
          "Mccormick, Kraft, Nestle",
          "HQ",
          "TGP PS & Made in Report",
      ]
      messagebox.showerror(
          "Invalid File Name",
          f"The selected file does not match the required naming convention for"
          f" '{self.support_file_labels[index]}'.\n\nFilename must contain:"
          f" '{expected_patterns[index]}'.",
      )
      return

    status_label.configure(text=file_name, text_color="#2ecc71")

    if is_big:
      self.marketing_file = file_path
    else:
      self.support_files[index] = file_path

  def browse_file(self, is_big, index, status_label):
    filetypes = (
        [
            ("Excel & CSV Files", "*.xlsx *.xls *.xlsm *.csv"),
            ("All Files", "*.*"),
        ]
        if (not is_big and index == 5)
        else [("Excel Files", "*.xlsx *.xls *.xlsm"), ("All Files", "*.*")]
    )

    file_path = filedialog.askopenfilename(
        title="Select Support File", filetypes=filetypes
    )
    if file_path:
      self.process_file_path(file_path, is_big, index, status_label)

  def handle_drop(self, raw_data, is_big, index, status_label):
    files = self.tk.splitlist(raw_data)
    if files:
      self.process_file_path(files[0], is_big, index, status_label)

  def on_generate_click(self):
    macro_input = self.macro_entry.get().strip()
    if not macro_input:
      messagebox.showwarning(
          "Missing Input", "Please enter a value for 'Marketing Macro #'."
      )
      return

    try:
      macro_num = int(macro_input)
    except ValueError:
      messagebox.showerror(
          "Invalid Input", "'Marketing Macro #' must be a valid integer number."
      )
      return

    if not self.marketing_file:
      messagebox.showwarning(
          "Missing File", "Please upload the Marketing Macro file with Extract."
      )
      return

    for i, file in enumerate(self.support_files):
      if not file:
        messagebox.showwarning(
            "Missing Files",
            f"Please upload all support files. Missing:"
            f" '{self.support_file_labels[i]}'",
        )
        return

    self.generate_btn.configure(state="disabled", fg_color="gray")
    self.loading_label.configure(text="Generating please wait...")
    self.progress_bar.pack(pady=5)
    self.progress_bar.start()
    self.update()

    self.after(150, lambda: self.execute_generation(macro_num))

  def execute_generation(self, macro_num):
    excel_app = None
    try:
      excel_app = xw.App(visible=False)
      excel_app.screen_updating = False
      excel_app.display_alerts = False

      wb_macro = excel_app.books.open(self.marketing_file)

      # Offload Excel manipulation logic to the separate module
      MarketingMacroProcessor.process_all(
          excel_app, wb_macro, self.support_files, macro_num
      )

      save_path = filedialog.asksaveasfilename(
          defaultextension=".xlsx",
          filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
          title="Save Marketing Macro File As",
      )

      if save_path:
        wb_macro.save(save_path)
        messagebox.showinfo(
            "Success",
            f"Marketing macro generated and saved successfully to:\n{save_path}",
        )
      else:
        messagebox.showwarning(
            "Cancelled", "File save was cancelled. Workbook was not saved."
        )

    except Exception as e:
      messagebox.showerror(
          "Excel Processing Error", f"An error occurred with xlwings: {str(e)}"
      )
    finally:
      if excel_app:
        try:
          excel_app.quit()
        except:
          pass

      self.progress_bar.stop()
      self.progress_bar.pack_forget()
      self.loading_label.configure(text="")
      self.generate_btn.configure(state="normal", fg_color="#1f6aa5")