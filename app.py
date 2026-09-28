import os
from tkinter import filedialog, messagebox
import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD
import xlwings as xw

# Set up CustomTkinter appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


# Inherit from TkinterDnD.Tk to enable native drag-and-drop hooks
class MarketingMacroApp(ctk.CTk, TkinterDnD.DnDWrapper):

  def __init__(self):
    super().__init__()

    # Initialize TkinterDnD
    self.TkdndVersion = TkinterDnD._require(self)

    # Window Configuration
    self.setTitle("Work in progress")
    self.geometry("750x1000")
    self.resizable(False, False)

    # File paths storage (1 main file + 6 support files)
    self.marketing_file = None
    self.support_files = [None] * 6

    # Labels for the 6 specific support files (t-2 first, then t-1)
    self.support_file_labels = [
        "AD (t-2) - Layout",
        "AD (t-1) - Layout",
        "Made in Canada",
        "SKURank",
        "Mccormick, Kraft, Nestle",
        "HQ",
    ]

    self.create_widgets()

  def setTitle(self, title):
    self.title(title)

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
        self, width=700, height=820, fg_color="transparent"
    )
    main_frame.pack(fill="both", expand=True, padx=20, pady=10)

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
        text="Support Files (6 Required)",
        font=ctk.CTkFont(size=16, weight="bold"),
    )
    support_label.pack(anchor="w", pady=(0, 5))

    # 2. Six Smaller Upload Areas for specific support files
    for i in range(6):
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
    frame.drop_target_register(DND_FILES)
    frame.dnd_bind(
        '<<Drop>>',
        lambda e: self.handle_drop(e.data, is_big, index, status_label),
    )

    label.drop_target_register(DND_FILES)
    label.dnd_bind(
        '<<Drop>>',
        lambda e: self.handle_drop(e.data, is_big, index, status_label),
    )

    status_label.drop_target_register(DND_FILES)
    status_label.dnd_bind(
        '<<Drop>>',
        lambda e: self.handle_drop(e.data, is_big, index, status_label),
    )

  def validate_support_filename(self, index, filename):
    fn_lower = filename.lower()

    if index == 0:  # AD (t-2) - Layout -> No naming check
      return True
    elif index == 1:  # AD (t-1) - Layout -> No naming check
      return True
    elif index == 2:  # Made in Canada
      return "made in canada" in fn_lower
    elif index == 3:  # SKURank
      return "skurank" in fn_lower
    elif index == 4:  # Mccormick, Kraft, Nestle
      return "mccormick" in fn_lower
    elif index == 5:  # HQ
      return "hq" in fn_lower
    return True

  def process_file_path(self, file_path, is_big, index, status_label):
    file_path = file_path.strip('{}""\'')

    if not os.path.exists(file_path):
      messagebox.showerror(
          "Error", f"The dropped path does not exist:\n{file_path}"
      )
      return

    if not is_big and index == 5:
      allowed_exts = (".xlsx", ".xls", ".xlsm", ".csv")
    else:
      allowed_exts = (".xlsx", ".xls", ".xlsm")

    if not file_path.lower().endswith(allowed_exts):
      messagebox.showerror(
          "Invalid File Type",
          "Please provide a valid file matching the required format.",
      )
      return

    file_name = os.path.basename(file_path)

    if not is_big:
      if not self.validate_support_filename(index, file_name):
        expected_patterns = [
            "",
            "",
            "Made in Canada",
            "SKURank",
            "Mccormick, Kraft, Nestle",
            "HQ",
        ]
        messagebox.showerror(
            "Invalid File Name",
            f"The selected file does not match the required naming convention for "
            f"'{self.support_file_labels[index]}'.\n\nFilename must contain: "
            f"'{expected_patterns[index]}'.",
        )
        return

    status_label.configure(text=file_name, text_color="#2ecc71")

    if is_big:
      self.marketing_file = file_path
    else:
      self.support_files[index] = file_path

  def browse_file(self, is_big, index, status_label):
    if not is_big and index == 5:
      filetypes = [
          ("Excel & CSV Files", "*.xlsx *.xls *.xlsm *.csv"),
          ("All Files", "*.*"),
      ]
    else:
      filetypes = [("Excel Files", "*.xlsx *.xls *.xlsm"), ("All Files", "*.*")]

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
    # Validation check for main file
    if not self.marketing_file:
      messagebox.showwarning(
          "Missing File",
          "Please upload the Marketing Macro file with Extract.",
      )
      return

    # Validation check for all 6 support files
    for i, file in enumerate(self.support_files):
      if not file:
        messagebox.showwarning(
            "Missing Files",
            f"Please upload all support files. Missing: '{self.support_file_labels[i]}'",
        )
        return

    # Gray out button, display loading text and start spinner
    self.generate_btn.configure(state="disabled", fg_color="gray")
    self.loading_label.configure(text="Generating please wait...")
    self.progress_bar.pack(pady=5)
    self.progress_bar.start()

    # Force UI update to render loading state before heavy Excel operations run
    self.update()

    # Run processing after a brief delay so the spinner animation initiates
    self.after(150, self.execute_generation)

  def execute_generation(self):
    excel_app = None
    try:
      # Initialize Excel in the background (visible=False)
      excel_app = xw.App(visible=False)
      wb_macro = xw.Book(self.marketing_file)

      # Call individual handler methods in order (Index 0 is t-2, Index 1 is t-1)
      self.handle_ad_t2_layout(wb_macro, self.support_files[0])
      self.handle_ad_t1_layout(wb_macro, self.support_files[1])
      self.handle_made_in_canada(wb_macro, self.support_files[2])
      self.handle_sku_rank(wb_macro, self.support_files[3])
      self.handle_mccormick_kraft_nestle(wb_macro, self.support_files[4])
      self.handle_hq(wb_macro, self.support_files[5])

      # Prompt user where to save their new marketing macro file
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
      # Cleanly close background Excel app instance if opened
      if excel_app:
        excel_app.quit()

      # Reset UI state (stop spinner, clear loading label, re-enable button)
      self.progress_bar.stop()
      self.progress_bar.pack_forget()
      self.loading_label.configure(text="")
      self.generate_btn.configure(state="normal", fg_color="#1f6aa5")

  # ==========================================
  # Support File Handler Methods
  # ==========================================

  def handle_ad_t2_layout(self, wb_macro, file_path):
    wb_support = xw.Book(file_path)

    sheet_macro = wb_macro.sheets["Sheet1"]
    sheet_support = wb_support.sheets[0]

    # Set header row in column BQ for t-2
    sheet_macro.range("BQ2").value = "t-2"

    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)
      support_sheet_name = sheet_support.name

      formula = f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$B, 1, 0)"
      sheet_macro.range(f"BQ3:BQ{last_row}").formula = formula

    wb_support.close()

  def handle_ad_t1_layout(self, wb_macro, file_path):
    wb_support = xw.Book(file_path)

    sheet_macro = wb_macro.sheets["Sheet1"]
    sheet_support = wb_support.sheets[0]

    # Set header row in column BR for t-1
    sheet_macro.range("BR2").value = "t-1"

    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)
      support_sheet_name = sheet_support.name

      formula = f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$B, 1, 0)"
      sheet_macro.range(f"BR3:BR{last_row}").formula = formula

    wb_support.close()

  def handle_made_in_canada(self, wb_macro, file_path):
    wb_support = xw.Book(file_path)

    sheet_macro = wb_macro.sheets["Sheet1"]
    sheet_support = wb_support.sheets[0]

    sheet_macro.range("BP2").value = "Canada"

    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)
      support_sheet_name = sheet_support.name

      formula = f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$B, 1, 0)"
      sheet_macro.range(f"BP3:BP{last_row}").formula = formula

    wb_support.close()

  def handle_sku_rank(self, wb_macro, file_path):
    wb_support = xw.Book(file_path)

    sheet_macro = wb_macro.sheets["Sheet1"]
    sheet_support = wb_support.sheets[0]

    # Set headers for BS ("SNS") and BT ("$")
    sheet_macro.range("BS2").value = "SNS"
    sheet_macro.range("BT2").value = "$"

    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)
      support_sheet_name = sheet_support.name

      # Column BS (SNS) -> VLOOKUP returning column index 24
      formula_bs = f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$Z, 24, 0)"
      sheet_macro.range(f"BS3:BS{last_row}").formula = formula_bs

      # Column BT ($) -> VLOOKUP returning column index 15
      formula_bt = f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$Z, 15, 0)"
      bt_range = sheet_macro.range(f"BT3:BT{last_row}")
      bt_range.formula = formula_bt

      # Format BT as currency with 2 decimal places
      bt_range.number_format = '"$"#,##0.00'

    wb_support.close()

  def handle_mccormick_kraft_nestle(self, wb_macro, file_path):
    wb_support = xw.Book(file_path)
    # TODO: Add your xlwings logic for McCormick, Kraft, Nestle here
    wb_support.close()

  def handle_hq(self, wb_macro, file_path):
    wb_support = xw.Book(file_path)
    # TODO: Add your xlwings logic for HQ CSV here
    wb_support.close()


if __name__ == "__main__":
  app = MarketingMacroApp()
  app.mainloop()