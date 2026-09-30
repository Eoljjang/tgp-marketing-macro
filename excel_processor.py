import os


class MarketingMacroProcessor:

  @staticmethod
  def process_all(excel_app, wb_macro, support_files, macro_num):
    # Calculate dynamic header numbers (-2 and -1)
    val_t2 = macro_num - 2
    val_t1 = macro_num - 1

    MarketingMacroProcessor.handle_ad_t2_layout(
        excel_app, wb_macro, support_files[0], val_t2
    )
    MarketingMacroProcessor.handle_ad_t1_layout(
        excel_app, wb_macro, support_files[1], val_t1
    )
    MarketingMacroProcessor.handle_made_in_canada(
        excel_app, wb_macro, support_files[2]
    )
    MarketingMacroProcessor.handle_sku_rank(excel_app, wb_macro, support_files[3])
    MarketingMacroProcessor.handle_mccormick_kraft_nestle(
        excel_app, wb_macro, support_files[4]
    )
    MarketingMacroProcessor.handle_hq(excel_app, wb_macro, support_files[5])
    MarketingMacroProcessor.handle_tgp_ps_report(
        excel_app, wb_macro, support_files[6]
    )

  @staticmethod
  def handle_ad_t2_layout(excel_app, wb_macro, file_path, header_val):
    wb_support = excel_app.books.open(file_path)
    sheet_macro = wb_macro.sheets["Sheet1"]
    sheet_support = wb_support.sheets[0]

    sheet_macro.range("BQ2").value = header_val
    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)
      support_sheet_name = sheet_support.name
      formula = (
          f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$B, 1, 0)"
      )
      sheet_macro.range(f"BQ3:BQ{last_row}").formula = formula

    wb_support.close()

  @staticmethod
  def handle_ad_t1_layout(excel_app, wb_macro, file_path, header_val):
    wb_support = excel_app.books.open(file_path)
    sheet_macro = wb_macro.sheets["Sheet1"]
    sheet_support = wb_support.sheets[0]

    sheet_macro.range("BR2").value = header_val
    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)
      support_sheet_name = sheet_support.name
      formula = (
          f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$B, 1, 0)"
      )
      sheet_macro.range(f"BR3:BR{last_row}").formula = formula

    wb_support.close()

  @staticmethod
  def handle_made_in_canada(excel_app, wb_macro, file_path):
    wb_support = excel_app.books.open(file_path)
    sheet_macro = wb_macro.sheets["Sheet1"]
    sheet_support = wb_support.sheets[0]

    sheet_macro.range("BP2").value = "Canada"
    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)
      support_sheet_name = sheet_support.name
      formula = (
          f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$B, 1, 0)"
      )
      sheet_macro.range(f"BP3:BP{last_row}").formula = formula

    wb_support.close()

  @staticmethod
  def handle_sku_rank(excel_app, wb_macro, file_path):
    wb_support = excel_app.books.open(file_path)
    sheet_macro = wb_macro.sheets["Sheet1"]
    sheet_support = wb_support.sheets[0]

    sheet_macro.range("BS2").value = "SNS"
    sheet_macro.range("BT2").value = "$"
    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)
      support_sheet_name = sheet_support.name

      formula_bs = (
          f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$Z, 24, 0)"
      )
      sheet_macro.range(f"BS3:BS{last_row}").formula = formula_bs

      formula_bt = (
          f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$Z, 15, 0)"
      )
      bt_range = sheet_macro.range(f"BT3:BT{last_row}")
      bt_range.formula = formula_bt
      bt_range.number_format = '"$"#,##0.00'

    wb_support.close()

  @staticmethod
  def handle_mccormick_kraft_nestle(excel_app, wb_macro, file_path):
    wb_support = excel_app.books.open(file_path)
    sheet_macro = wb_macro.sheets["Sheet1"]

    sheet_macro.range("BU2").value = "Kraft"
    sheet_macro.range("BV2").value = "Nestle"
    sheet_macro.range("BW2").value = "McCormick"
    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)

      formula_kraft = f"=VLOOKUP(A3, '[{support_name}]Kraft'!$A:$B, 1, 0)"
      sheet_macro.range(f"BU3:BU{last_row}").formula = formula_kraft

      formula_nestle = f"=VLOOKUP(A3, '[{support_name}]Nestle'!$A:$B, 1, 0)"
      sheet_macro.range(f"BV3:BV{last_row}").formula = formula_nestle

      formula_mccormick = f"=VLOOKUP(A3, '[{support_name}]McCormick'!$A:$B, 1, 0)"
      sheet_macro.range(f"BW3:BW{last_row}").formula = formula_mccormick

    wb_support.close()

  @staticmethod
  def handle_hq(excel_app, wb_macro, file_path):
    wb_support = excel_app.books.open(file_path)
    sheet_macro = wb_macro.sheets["Sheet1"]
    sheet_support = wb_support.sheets[0]

    sheet_macro.range("BY2").value = "HQ"
    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)
      support_sheet_name = sheet_support.name
      formula = (
          f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$B:$I, 6, 0)"
      )
      sheet_macro.range(f"BY3:BY{last_row}").formula = formula

    wb_support.close()

  @staticmethod
  def handle_tgp_ps_report(excel_app, wb_macro, file_path):
    wb_support = excel_app.books.open(file_path)
    sheet_macro = wb_macro.sheets["Sheet1"]
    sheet_support = wb_support.sheets[0]

    sheet_macro.range("BZ2").value = "HO"
    last_row = sheet_macro.range(
        "A" + str(sheet_macro.cells.last_cell.row)
    ).end("up").row

    if last_row >= 3:
      support_name = os.path.basename(file_path)
      support_sheet_name = sheet_support.name
      formula = (
          f"=VLOOKUP(A3, '[{support_name}]{support_sheet_name}'!$A:$Z, 4, 0)"
      )
      sheet_macro.range(f"BZ3:BZ{last_row}").formula = formula

    wb_support.close()