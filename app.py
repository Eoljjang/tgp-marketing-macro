import os
from tkinter import filedialog, messagebox
import customtkinter as ctk
import xlwings as xw

# Set up CustomTkinter appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class MarketingMacroApp(ctk.CTk):

  def __init__(self):
    super().__init__()

    # Window Configuration
    self.setTitle("Marketing Macro Generator")
    self.geometry("750x850")
    self.resizable(False, False)

    # File paths storage
    self.marketing_file = None
    self.xyz_files = [None] * 5

    self.create_widgets()

  def setTitle(self, title):
    self.title(title)

  def create_widgets(self):
    # Main Title Label
    title_label = ctk.CTkLabel(
        self,
        text="Marketing Macro Generator",
        font=ctk.CTkFont(size=24, weight="bold"),
    )
    title_label.pack(pady=(20, 10))

    # Scrollable frame for cleaner layout if needed, or simple frame
    main_frame = ctk.CTkScrollableFrame(
        self, width=700, height=700, fg_color="transparent"
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

    # Section label for smaller files
    xyz_label = ctk.CTkLabel(
        main_frame,
        text="XYZ Support Files (5 Required)",
        font=ctk.CTkFont(size=16, weight="bold"),
    )
    xyz_label.pack(anchor="w", pady=(0, 5))

    # 2. Five Smaller Upload Areas (XYZ files)
    for i in range(5):
      self.create_upload_section(
          parent=main_frame,
          title_text=f"Upload your xyz file #{i+1}",
          is_big=False,
          index=i,
      )

    # 3. Generate Button
    generate_btn = ctk.CTkButton(
        main_frame,
        text="Generate Marketing Macro",
        font=ctk.CTkFont(size=15, weight="bold"),
        height=45,
        fg_color="#1f6aa5",
        hover_color="#144870",
        command=self.on_generate_click,
    )
    generate_btn.pack(fill="x", pady=(25, 20))

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
        text="No file selected",
        font=ctk.CTkFont(size=12, slant="italic"),
        text_color="gray",
    )
    status_label.pack(side="right", padx=15)

    # Store references to update status later
    if is_big:
      self.marketing_status = status_label
    else:
      if not hasattr(self, "xyz_statuses"):
        self.xyz_statuses = []
      self.xyz_statuses.append(status_label)

    # Click to upload button
    upload_btn = ctk.CTkButton(
        frame,
        text="Browse...",
        width=90,
        height=32,
        command=lambda: self.browse_file(is_big, index, status_label),
    )
    upload_btn.pack(side="right", padx=5)

  def browse_file(self, is_big, index, status_label):
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls *.xlsm"), ("All Files", "*.*")],
    )

    if file_path:
      # Validate extension
      if not file_path.lower().endswith((".xlsx", ".xls", ".xlsm")):
        messagebox.showerror(
            "Invalid File Type",
            "Please select a valid Excel file (.xlsx, .xls, or .xlsm).",
        )
        return

      file_name = os.path.basename(file_path)
      status_label.configure(text=file_name, text_color="#2ecc71")

      if is_big:
        self.marketing_file = file_path
      else:
        self.xyz_files[index] = file_path

  def on_generate_click(self):
    # Validation check
    if not self.marketing_file:
      messagebox.showwarning(
          "Missing File",
          "Please upload the Marketing Macro file with Extract.",
      )
      return

    for i, file in enumerate(self.xyz_files):
      if not file:
        messagebox.showwarning(
            "Missing Files", f"Please upload all 5 xyz files. Missing file #{i+1}."
        )
        return

    # All files validated successfully!
    messagebox.showinfo("Success", "All files validated successfully!")

    # Template Logic using xlwings
    try:
      # Example xlwings setup (leave implementation details as requested)
      app = xw.App(visible=True)  # Set to False if you want it running in background

      # Open Marketing Macro file
      wb_macro = xw.Book(self.marketing_file)

      # Template logic loop for each uploaded xyz file
      for i, xyz_path in enumerate(self.xyz_files):
        wb_xyz = xw.Book(xyz_path)

        # TODO: Add your custom xlwings data processing / macro logic here
        # Example:
        # data = wb_xyz.sheets[0].range('A1:D10').value
        # wb_macro.sheets['Summary'].range(f'A{i*10 + 1}').value = data

        wb_xyz.close()

      # Save or handle output workbook as needed
      # wb_macro.save("Path_To_Output.xlsx")

    except Exception as e:
      messagebox.showerror(
          "Excel Processing Error", f"An error occurred with xlwings: {str(e)}"
      )


if __name__ == "__main__":
  app = MarketingMacroApp()
  app.mainloop()