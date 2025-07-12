# -*- coding: utf-8 -*-
"""
@author: JesusMMA
"""

import xlwings as xw
import os
import re
from datetime import date,datetime
from UserInputs import ask_open_file, show_info,show_warning,ask_user_number

from PyQt5.QtWidgets import (QLabel,QPushButton, QVBoxLayout, QApplication,QDialog)
import sys
import Load_SAP_info
# -----------------------------------
# Aux Functions
# -----------------------------------
class RangeSelectorWindow(QDialog):
    """
    PyQt5 dialog for selecting a cell range in Excel to either store or transfer into a template.
    Used in SAP-related workflows to handle manual input of invoice ranges and posting entries.

    Workflow:
    - Launches a modal window prompting the user to select a cell range in Excel
    - Offers two actions:
        - 'Pasar a Template': Copies the selected range into a target template sheet
        - 'Almacenar rango': Saves the selected range to use later in the script
    - Handles cancellations gracefully by notifying the user
    - On confirmation, the selected data is either transferred or stored for downstream logic

    Parameters:
    - wbTemplate (Workbook): The Excel workbook containing the destination sheet
    - destination_start_cell (str): Top-left cell where copied data should begin (e.g. 'D10')

    Attributes:
    - selected_range: Stores the user-selected cell range for external use

    Usage:
    This dialog is typically triggered via `_launch_range_selector()` after a workbook is loaded.
    """
    def __init__(self, wbTemplate, destination_start_cell):
        super().__init__()
        self.wbTemplate = wbTemplate
        self.destination_start_cell = destination_start_cell
        self.selected_range = None 
        self.setWindowTitle("Confirmar Selección")
        self.setFixedSize(300, 150)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        label = QLabel("Selecciona el rango en Excel\nLuego haz clic en OK para copiar.")
        layout.addWidget(label)

        to_template_btn = QPushButton("Pasar a Template")
        to_template_btn.clicked.connect(self.on_to_template_btn)
        layout.addWidget(to_template_btn)
        
        save_range_btn = QPushButton("Almacenar rango")
        save_range_btn.clicked.connect(self.on_save_range_btn)
        layout.addWidget(save_range_btn)

        cancel_btn = QPushButton("Cancelar")
        cancel_btn.clicked.connect(self.on_cancel)
        layout.addWidget(cancel_btn)

        self.setLayout(layout)
        self.show()  # Modeless by default

    def on_to_template_btn(self):
        selected_range = self.wbTemplate.app.selection

        if not selected_range:
            show_info("Cancelar","No hay una selección válida.")
            return
        self.transfer_range(selected_range)
        self.accept()
    def on_save_range_btn(self):
        selected_range = self.wbTemplate.app.selection
        if not selected_range:
            show_info("Cancelar","No hay una selección válida.")
            return
        self.selected_range = selected_range
        self.accept()  # Close the dialog and return control
        
    def on_cancel(self):
        show_info("Cancelar","Selección cancelada por el usuario.")
        self.cancel()

    def transfer_range(self, selected_range):
        wsTemplate = self.wbTemplate.sheets[0]
        start_row = xw.Range(self.destination_start_cell).row
        start_col = xw.Range(self.destination_start_cell).column

        for i, row in enumerate(selected_range.rows):
            for j, cell in enumerate(row):
                wsTemplate.cells(start_row + i, start_col + j).value = cell.value
        self.wbTemplate.app.status_bar = "Listo"

# Launch the PyQt5 window after workbook is open
def launch_range_selector(wbTemplate, destination_start_cell='D10'):
    """
    Opens a PyQt5 dialog for manual cell range selection from an Excel workbook.
    Ensures that the application instance is initialized, then blocks until selection is confirmed or canceled.
    
    Workflow:
    - Checks for an existing QApplication instance; creates one if missing
    - Instantiates and launches RangeSelectorWindow with the active workbook and starting cell
    - Waits for user interaction via dialog box
    - If selection is confirmed, returns the selected range
    - If canceled, returns None
    
    Parameters:
    - wbTemplate (Workbook): The workbook object to enable range selection in
    - destination_start_cell (str, optional): Starting cell for range placement (default is 'D10')
    
    Returns:
    - Range or None: Selected cell range if accepted; otherwise None
    """
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    dialog = RangeSelectorWindow(wbTemplate, destination_start_cell)
    result = dialog.exec_()  # Blocks until dialog is closed
    if result == QDialog.Accepted:
        return dialog.selected_range
    return None

# Aux function to check if a workbook is already opened 
def check_wb_open(path):
    """
    Checks whether the specified Excel workbook is already open.
    If open, returns the existing instance; otherwise, opens it fresh.

    Workflow:
    - Normalize file path for consistent comparison
    - Loop through all open workbooks and compare paths
    - If match is found, return the already open workbook
    - If not, open the workbook from the given path

    Parameters:
    - path (str): Full file path to the workbook

    Returns:
    - Workbook: A reference to the corresponding Excel workbook object
    """
    path = os.path.normcase(os.path.normpath(path))
    # Ensure at least one Excel app instance is running
    if not xw.apps:
        xw.App(visible=True)  
    # Check if workbook is already open
    wb = None
    for book in xw.books:
        book_path = os.path.normcase(os.path.normpath(book.fullname))
        if book_path == path:
            return book
    # If not open, then open it
    if wb is None:
        return xw.Book(path)

# Aux function to convert number to Excel letter column
def letter_from_number(n):
    """
    Converts a 1-indexed Excel column number into its corresponding letter representation.
    
    Workflow:
    - Uses base-26 arithmetic to calculate each letter from the given number
    - Builds the column name from right to left (e.g. 1 → A, 27 → AA, 703 → AAA)
    
    Parameters:
    - n (int): A 1-based index of the Excel column number

    Returns:
    - str: Excel-style column letter (e.g. 'A', 'Z', 'AA', 'AB')
    """
    result = ''
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result 

# Aux function to clean str in order to name a Worksheet or file
def sanitize_sheet_name(name):
    # Remove invalid characters: : \ / ? * [ ]
    name = re.sub(r'[:\\/*?\[\]]', '', name)
    # Strip leading/trailing whitespace and limit to 31 characters
    name = name.strip()[:31]
    # Optional: replace empty name with a default one
    return name if name else "Sheet1"

# Aux function to setup headears in a Worksheet
def setup_headers(ws,report_name):
    """
    Configures the header row with predefined labels and colors, and applies bold formatting and borders.

    Workflow:
    - Deletes columns L and M to remove old headers or data
    - Writes new headers into cells L1 to O1 with specified background colors
    - Applies thick borders to the entire header row
    - Sets the font to bold for header visibility
    
    Parameters:
    - ws: Excel Worksheet
    
    Returns:
    - None: updates Excel file
    """
    # Format Worksheet
    report_detail = Load_SAP_info.config[f"{report_name}_detail"]
    delete_columns = report_detail["delete_columns"]
    insert_columns = report_detail["insert_columns"]
    if len(delete_columns) > 0:
        for col in delete_columns:
            ws.range(col).delete()
    if len(insert_columns) > 0:
        for col in insert_columns:
            ws.range(col).api.Insert()
    headers = report_detail["headers"]
    for item in headers:
        cell = item["cell"]
        text = item["text"]
        color = tuple(item["color"])
        rng = ws.range(cell)
        rng.value = text
        rng.color = color
    last_col = ws.api.Range("1:1").Find("*", LookIn=-4123, LookAt=2, SearchOrder=1, SearchDirection=2).Column + 1
    last_col_letter = letter_from_number(last_col)
    ws.range(f"A1:{last_col_letter}1").api.Borders.Weight = 2
    ws.range(f"A1:{last_col_letter}1").font.bold = True

# Aux function to sett Validation data in a indicated range
def set_data_validation(ws, col_val_index, last_row, validation_source, use_range=False, wb=None):
    """
    Applies dropdown data validation to a specified column in an Excel worksheet.

    Parameters:
    - ws: Worksheet to apply validation to
    - col_val_index: Index of the column to apply the validation
    - last_row: Last used row in the worksheet
    - validation_source: Either a list of values or a range (if use_range=True)
    - use_range: Boolean indicating whether validation_source is a range object
    - wb: Optional Workbook reference (required if validation_source is a range from another sheet)

    Returns:
    - None
    """
    target_range = ws.range((2, col_val_index), (last_row, col_val_index))
    try:
        target_range.api.Validation.Delete()

        if use_range:
            # If it's a range, build the formula from its address
            formula_address = validation_source.get_address(external=True)
            target_range.api.Validation.Add(
                Type=3, AlertStyle=1, Operator=1, Formula1=f"={formula_address}"
            )
        else:
            # If it's a list, join values with commas for inline validation
            list_str = ";".join(validation_source)
            target_range.api.Validation.Add(
                Type=3, AlertStyle=1, Operator=1, Formula1=f'"{list_str}"'
            )
    except Exception as e:
        print(f"[ERROR] Validation failed: {e}")


def get_unique_column_values(ws, col_index, temp_col="ZZ"):
    """
    Extracts and returns a list of unique values from a specified column in an Excel worksheet.

    Workflow:
    - Identify the last non-empty cell in column A (assumed to be the anchor for detecting last row).
    - Define the target range in the specified column from row 1 to the last detected row.
    - Use Excel's AdvancedFilter feature (via COM API) to copy only unique values to a temporary column.
    - Read those unique values into a Python list, replacing None with empty strings for consistency.
    - Clear contents of the temporary column used for filtering.
    - Return the cleaned list of unique values.
    
    Notes:
    - The use of column "ZZ" helps avoid collision with live worksheet data.
    - This function relies on COM API access (via xlwings) for AdvancedFilter.
    Returns:
    - List: of unique values
    """
    # Get the last non-empty row by scanning up from the bottom of column A
    last_row = ws.range("A" + str(ws.cells.last_cell.row)).end("up").row
    # Define source range in the target column (from row 1 to last row)
    source_range = ws.range((1, col_index), (last_row, col_index))
    # Use a distant column (ZZ) to store filtered results temporarily
    temp_cell = ws.range(f"{temp_col}1")
    # Use Excel's AdvancedFilter to get unique values from source_range into temp_col
    source_range.api.AdvancedFilter(
        Action=2,                  # Copy output to another location
        CriteriaRange=None,       # No filter conditions
        CopyToRange=temp_cell.api,
        Unique=True               # Keep only unique values
    )
    # Determine last row in temporary column to read extracted values
    temp_row = ws.range(f"{temp_col}" + str(ws.cells.last_cell.row)).end("up").row
    unique_range = ws.range(f"{temp_col}2:{temp_col}{temp_row}").value
    # Convert range values to list, replacing None with empty string
    unique_list=[]
    if temp_row == 2:
        if unique_range is None:
            unique_list.append("")
        elif isinstance(unique_range, datetime):
            unique_range = unique_range.date()
            unique_list.append(unique_range.strftime("%Y/%m/%d"))
        else:
            unique_list.append(unique_range)
    else:    
        for v in unique_range:
            if v is None:
                unique_list.append("")
            elif isinstance(v, datetime):
                v = v.date()
                unique_list.append(v.strftime("%Y/%m/%d"))
            else:
                unique_list.append(v)
    # Clean up: remove temporary filter results
    ws.range(f"{temp_col}:{temp_col}").clear_contents()
    return unique_list

def split_by_filter(wb,ws,col_index):
    """
    Separates data by 'GESTOR' value, creating one sheet per account manager and copying filtered rows into them.
    
    Workflow:
    - Retrive unique vals from the indicated column 
    - For each value in indicated column :
      - Filters the sheet to show only their rows
      - Creates a new sheet (or reuses existing one) with their name
      - Copies visible rows into the target sheet
    - Resets filter when finished
    
    Parameters:
    - ws_ini: Excel Worksheet containing the full data set
    - wb: Excel Workbook where new account manager-specific sheets will be added
    - col index (int): index of the column to filter by
    
    Returns:
    - None: updates workbook with segmented sheets per account manager
    """
    # Apply AutoFilter to the indicate column
    ws.api.UsedRange.AutoFilter(col_index)
    unique_vals = get_unique_column_values(ws, col_index)
    for val in unique_vals:
        ws.api.Rows(1).AutoFilter(Field=col_index, Criteria1=val)
        if not val:
            val = "SIN DATOS"
        # Check if filtered data exists
        if isinstance(val, str):
            val = sanitize_sheet_name(val)
        try:
            visible_cells = ws.api.UsedRange.SpecialCells(12)  # xlCellTypeVisible = 12
        except Exception:
            continue  # No visible rows after filtering
        # Create or select target sheet
        try:
            target_sheet = wb.sheets[val]
        except Exception:
            target_sheet = wb.sheets.add(after=wb.sheets[-1])
            old_name = ws.name
            if "Hoja" in old_name:
                old_name = ""
            target_sheet.name = f"{old_name}{val}"
        # Copy filtered visible cells
        visible_cells.Copy(Destination=target_sheet.range("A1").api)
    ws.api.AutoFilterMode = False  # Clear filter at end


def merge_sheets(wb, base_sheet, sheet_names):
    """
    Consolidates data from specified account manager sheets into the base sheet.
    
    Workflow:
    - Iterates over sheet names in `sheet_names`
    - Extracts the used data range from each sheet (from A2 to last filled row in column A)
    - Calculates target row in the base sheet to avoid overwriting existing data
    - Copies the data range into the base sheet starting at the computed row
    - Clears the content from the source sheet after merging
    - Logs an error if a sheet cannot be accessed or copied
    
    Parameters:
    - wb: Excel Workbook object containing both source and base sheets
    - base_sheet: Excel Worksheet where data will be consolidated
    - sheet_names (list[str]): List of sheet names to be merged
    
    Returns:
    - None: updates Excel content
    """
    # Merge all sheets if name matches for the passed list 
    for name in sheet_names:
        try:
            sheet = wb.sheets[name]
            sheet_last_row = sheet.range('A' + str(base_sheet.cells.last_cell.row)).end('up').row
            last_col = sheet.api.Range("1:1").Find("*", LookIn=-4123, LookAt=2, SearchOrder=1, SearchDirection=2).Column + 1
            last_col_letter = letter_from_number(last_col)
            used_range = sheet.range(f"A2:{last_col_letter}{sheet_last_row}")
            last_row = base_sheet.range('A' + str(base_sheet.cells.last_cell.row)).end('up').row
            target_row = last_row + 1
            used_range.copy(base_sheet.range(f"A{target_row}"))
            #used_range.clear_contents()
            sheet.api.UsedRange.ClearContents()
        except Exception:
            print(f"[ERROR] La hoja {name} no se ha podido añadir a la Base de datos")
            continue



def detail_handler(client_name,client_aux_name=""):
    """
    From invoices detail, categorizes data and populates a SAP-compatible template,
    and prepares structured data for final processing.
    
    Workflow:
    - Loads client configuration based on provided names
    - Opens invoice payment detail Excel file
    - Validates total invoice amount against user input
    - Extracts and classifies invoice references by type
    - Populates a template Excel file with structured invoice entries
    - Clears template and payment file cells if needed for recalculation
    - Assembles all relevant data into a final dictionary for SAP processing

    Parameters:
    - client_name (str): Primary client identifier
    - client_aux_name (str, optional): Auxiliary client identifier for alternate configurations
    
    Returns:
    - None: writes data to the SAP batch template and invokes downstream processing
    """
    # Final dictionary with all required fields for SAP processing
    payment_dic = {}
    # Load client-specific dictionaries and configurations based on provided names from JSON
    client_name_lower=client_name.lower()
    if client_aux_name:
       client_detail = Load_SAP_info.config[f"{client_aux_name}_detail"]
       if "_" in client_aux_name:
           client_aux_name=client_aux_name.split("_")[0]
       clients_dic=Load_SAP_info.config[f"{client_aux_name}_dic"]
    else:
        clients_dic=Load_SAP_info.config[f"{client_name_lower}_dic"]
        client_detail = Load_SAP_info.config[f"{client_name_lower}_detail"]    
    """
    client_detail structure
    "amount_col": int,
    "inv_ref_col": int or list[int],
    "corp_name": int or str or list[int],
    "total_amount": None or list[int],
    "due_date": list[int],
    "payment_number": int or list[int],
    "doc_type_col": int or None,
    "entry_match": list[str],
    "entry_comment": str,
    "start_row": int,
    "client_category": str,
    "payment_method": str,
    "SGLIndicator": str,
    "ajd_allowed": list[str],
    "invoices_allowed": list[str],
    "debit_allowed": list[str],
    "credit_allowed": list[str],
    "ajd_assignment": str or None
    
    """
    # Prompt user to open invoice detail file
    payment_detail_path = ask_open_file("Abre el archivo con el detalle de Facturas")
    if not payment_detail_path:
        return
    wb = check_wb_open(payment_detail_path)
    ws = wb.sheets[0]
    # Retrive the column number for invoices amount from the detail dictcionary
    amount_col=client_detail["amount_col"]  
    amount_col_letter = letter_from_number(amount_col)
    # Extract starting and ending rows for invoice detail section
    start_row = client_detail["start_row"]
    end_row = ws.range(f"{amount_col_letter}{start_row}").end("down").row
    # Retrive total amount cell (row,col) 
    total_amount_cell = client_detail["total_amount"]
    if total_amount_cell:
        total_amount_row = total_amount_cell[0]
        total_amount_col = total_amount_cell[1]
        # Clear 'total amount' cell if overlapping with invoice amount column
        if amount_col == total_amount_col:
            ws.cells(total_amount_row,total_amount_col).clear() 
    # Define date-related variables
    due_date_cell = client_detail["due_date"]
    if len(due_date_cell) > 1:
        due_date_row = due_date_cell[0]
        due_date_col = due_date_cell[1]
    else:
        due_date_col = due_date_cell[0]
        due_date_row = start_row
    due_date = ws.cells(due_date_row, due_date_col).value.date()
    doc_date_assignment = date.today().strftime("%Y%m%d")
    due_date_assignment = due_date.strftime("%Y%m%d") 
    due_date = due_date.strftime("%d.%m.%Y")
    doc_date = date.today().strftime("%d.%m.%Y")
    # Clear 'due date' cell if overlapping with invoice amount column
    if amount_col == due_date_col:
        ws.cells(due_date_row, due_date_col).clear_contents()
    # Prompt user for total promissory note amount and validate against calculated sum
    user_amount = ask_user_number("Introduce el total del detalle")
    # Get the total amount from the payment details by summing column D
    detail_amount = round(ws.api.Application.WorksheetFunction.Sum(ws.range(f"{amount_col_letter}:{amount_col_letter}").api),2)
    # Exit if amounts don't match (payment details might be incomplete)
    if user_amount != detail_amount:
        show_warning("Cancelación", "El importe introducido no cuadra con el detalle")
        return
    # Clean invoice reference column (remove hyphens)
    inv_ref_col = client_detail["inv_ref_col"]
    if isinstance(inv_ref_col, list):
        inv_ref_col_aux = inv_ref_col[1]
        inv_ref_col=inv_ref_col[0]
    inv_ref_col_letter = letter_from_number(inv_ref_col)
    col = ws.range(f'{inv_ref_col_letter}:{inv_ref_col_letter}')
    col.api.Replace(What="-", Replacement="", LookAt=2, MatchCase=False, SearchFormat=False, ReplaceFormat=False)
    payment_number = client_detail["payment_number"]
    if isinstance(payment_number, int):
        split_by_filter(wb,ws,payment_number)
        ws.delete()
    elif isinstance(payment_number, list):
        payment_number_row = payment_number[0]
        payment_number_col = payment_number[1]
        payment_number = ws.cells(payment_number_row,payment_number_col).api.Text
        ws.name = payment_number
    for sheet in wb.sheets:
        if len(due_date_cell) == 1:
            uniques_due_dates = get_unique_column_values(sheet,due_date_col)
            if len(uniques_due_dates) == 1:
                continue
            else:
                split_by_filter(wb, sheet, due_date_col)
                sheet.delete()
        else:
            break
    # Load SAP batch template and clear previously populated rows
    batch_template_path = Load_SAP_info.config["batch_template_path2"]
    for sheet in wb.sheets:
        # Define date-related variables
        due_date_cell = client_detail["due_date"]
        if len(due_date_cell) > 1:
            due_date_row = due_date_cell[0]
            due_date_col = due_date_cell[1]
        else:
            due_date_col = due_date_cell[0]
            due_date_row = start_row
        due_date = sheet.cells(due_date_row, due_date_col).value.date()
        doc_date_assignment = date.today().strftime("%Y%m%d")
        due_date_assignment = due_date.strftime("%Y%m%d") 
        due_date = due_date.strftime("%d.%m.%Y")
        doc_date = date.today().strftime("%d.%m.%Y")
        # Define end_row
        end_row = sheet.range(f"{amount_col_letter}{start_row}").end("down").row
        wb_template = check_wb_open(batch_template_path)
        ws_template = wb_template.sheets[0]
        # Clear template before insert data
        temp_ini_row = 10
        temp_end_row = ws_template.range(f"D{temp_ini_row}").end("down").row
        if temp_ini_row < temp_end_row:
            ws_template.range(f'D{temp_ini_row}:D{temp_end_row}').clear_contents()
        # Populate template header with metadata
        client_category =client_detail["client_category"]
        ws_template.cells(2,5).value = doc_date
        ws_template.cells(2, 7).value = doc_date
        ws_template.cells(6, 6).value = client_category
        # If there is only one client code add it into template so only Items from that account are selected
        if len(clients_dic) == 1:
            client_code = clients_dic[client_name.upper()]
            ws_template.cells(6,5).value = client_code
        # Initialize containers for invoice, credit, and debit aggregation
        invoices = float(0)
        ajd_amount = float(0)
        credit_amounts ={}
        debit_amounts ={}
        for key_name in clients_dic:
            credit_amounts[key_name]=float(0)
            debit_amounts[key_name]=float(0)
        # Define empty dictionaries for future use
        entries_dic = {}
        invoices_dic = {}
        invoices_allowed = client_detail["invoices_allowed"]
        debit_allowed = client_detail["debit_allowed"]
        credit_allowed = client_detail["credit_allowed"]
        entry_match = client_detail["entry_match"]
        doc_type_col = client_detail["doc_type_col"]
        ajd_allowed = client_detail["ajd_allowed"]
        corp_name_detail = client_detail["corp_name"]
        if isinstance(corp_name_detail, list):
            corp_name_row = corp_name_detail[0]
            corp_name_col = corp_name_detail[1]
            corp_name = (sheet.cells(corp_name_row,corp_name_col).api.Text).strip().upper()
        elif isinstance(corp_name_detail, str):
            corp_name = corp_name_detail
        # Iterate through invoice rows to classify and extract data based on document type
        for i in range(start_row,end_row + 1):
            inv_ref = sheet.cells(i,inv_ref_col).api.Text
            amount = round(float(sheet.cells(i,amount_col).value),2)
            if doc_type_col:
                doc_type = sheet.cells(i,doc_type_col).api.Text.upper()
            else:
                doc_type = inv_ref[0]
            left_ref = inv_ref[0]
            len_ref = len(inv_ref)
            if isinstance(corp_name_detail, int):
                corp_name = (sheet.cells(i,corp_name_detail).api.Text).strip().upper()
            # For known invoice types, copy references to template and update totals
            if all(ref in invoices_allowed for ref in [doc_type, left_ref]):
                if len_ref == 8:
                    ws_template.cells(temp_ini_row,4).value = inv_ref 
                    temp_ini_row += 1
                    invoices = round(invoices + amount,2)
                    invoices_dic[inv_ref] = i
                elif len_ref == 7:
                    inv_ref = "X{inv_ref}"
                    ws_template.cells(temp_ini_row,4).value = inv_ref
                    temp_ini_row += 1
                    invoices_dic[inv_ref] = i
                    inv_ref = "V{inv_ref}"
                    ws_template.cells(temp_ini_row,4).value = inv_ref
                    temp_ini_row += 1
                    invoices_dic[inv_ref] = i
                    invoices = round(invoices + amount,2)
                else:
                    show_warning("Error", f"Tipo de factura no contemplada\nEn la fila {i}")
                    return
            # For credit notes, classify based on matching criteria or corporate name
            elif doc_type in credit_allowed and amount > 0:
                if left_ref in entry_match:
                    entries_dic[i] = i 
                else:
                    # Categorize data under the appropriate client
                    for key in credit_amounts:
                        if key in corp_name:
                            credit_amounts[key] = round(credit_amounts[key] + amount,2)
            # For debit notes, classify based on matching criteria or corporate name
            elif doc_type in debit_allowed and amount < 0:
                if left_ref in entry_match:
                    entries_dic[i] = i # Add data to the entries_dic
                else:
                    # Categorize data under the appropriate client
                    for key in debit_amounts:
                        if key in corp_name:
                            debit_amounts[key] = round(debit_amounts[key] + amount,2)
            elif doc_type in ajd_allowed:
                if amount < 0:
                    amount = abs(amount)
                ajd_amount = amount
            # All non-classified types are individual entries
            else:
                entries_dic[i] = i
        # Save and close populated SAP template
        wb_template.save()
        wb_template.close()
        # Fill final dictionary with structured payment data
        payment_dic["client_name"] = client_name
        payment_dic["due_date"] = due_date
        payment_dic["due_date_assignment"] = due_date_assignment
        payment_dic["doc_date"] = doc_date
        payment_dic["doc_date_assignment"] = doc_date_assignment
        payment_dic["payment_amount"] = user_amount
        payment_dic["invoices_amount"] = invoices
        payment_dic["ajd_amount"] = ajd_amount
        payment_dic["debit_amounts"] = debit_amounts
        payment_dic["credit_amounts"] = credit_amounts
        payment_dic["entries_dic"] = entries_dic
        payment_dic["invoices_dic"] = invoices_dic
        # Invoke template loader for further processing and SAP upload
        from PaymentsModule import payment_batch_template
        payment_batch_template(clients_dic, client_detail, payment_dic, wb, sheet, payment_detail_path)
        
# ---------
# Debug
# ---------   
# Saveguard
if __name__ == "__main__":
    print("Nice")
    detail_handler("Consum")