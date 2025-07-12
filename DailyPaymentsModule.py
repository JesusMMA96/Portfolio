# -*- coding: utf-8 -*-
"""
@author: JesusMMA
"""

import xlwings as xw
import re
from PyQt5.QtWidgets import QMessageBox, QInputDialog

from datetime import date, datetime
from SAPAux import (call_transaction, new_entry, new_entry_add_data,
                    search_items, simulate, enter_position, save_entry, get_entry_number,
                    batch_input)
from UserInputs import (ask_open_file, show_info,show_question,
                        show_warning,ask_user_number,save_confirmation
                        )
from Utilities import launch_range_selector,check_wb_open,set_data_validation,setup_headers
import Load_SAP_info

# Call back in daily_payments() program 
def _pass_row(ws,i,title="Cancelado",msg=None):
    """
    Skips processing for the current row by marking it as 'No Aplicado' and formatting it visibly.
    Displays a warning message to notify the user and applies red font styling to highlight the status.

    Workflow:
    - If no message is passed, generates a default warning message based on row number
    - Shows warning popup with provided title and message
    - Marks the payment row in column 12 as 'No Aplicado'
    - Applies red font color to the full row for visual tracking

    Parameters:
    - ws (Worksheet): Excel worksheet where the payment row exists
    - i (int): Row index to be marked as skipped
    - title (str, optional): Title of the warning popup (defaults to "Cancelado")
    - msg (str, optional): Message content; auto-generated if not provided

    Returns:
    - None: modifies worksheet directly and shows message dialog
    """
    # Default title and msg to _pass_row
    if msg is None:
        msg=f'Fila {i}: se omite y pasa al siguiente.'
    show_warning(title,msg)
    ws.cells(i, 12).value = "No Aplicado"
    ws.range(f'{i}:{i}').api.Font.Color = 255

# Call back in bank_file() program
def _new_concept(description, doc_date):
    """
    Extracts and formats the payment concept from a bank transaction description.
    Matches sender details for transfers or retains original label for specific keywords.

    Workflow:
    - Uses regex to detect standard bank transfer format and extract the sender
    - Formats concept as "Tr [Sender] [Date]" using the document date
    - If keywords "INGRESO" or "PAGO" are present, retains original description
    - If no match is found, returns None

    Parameters:
    - description (str): Raw transaction description from bank file
    - doc_date (date): Associated document date to embed in the concept

    Returns:
    - str or None: Formatted concept string or None if no valid match is found
    """
    match = re.search(r"Transferencia(?: Inmediata)? De\s+(.*?)(,|$)", description)
    if match:
        remitente = match.group(1).strip()
        return f"Tr {remitente} {doc_date}"
    elif "INGRESO" in description or "PAGO" in description:
        return description
    return None

# Call back in daily_payments()
def _load_template(doc_date , client_category:str = "",client_code:str ="" ):
    """
    Loads and prepares the SAP batch template for invoice entry.
    Clears prior data, sets required metadata fields, and invokes manual invoice selection.
    
    Workflow:
    - Opens batch template from configured path
    - Clears previous invoice lines starting from row 10
    - Fills in required metadata:
        - Document date
        - Client category (if provided)
        - Client code (placeholder; requires field verification)
    - Launches manual invoice selector from specified cell
    - Saves and closes the updated template
    
    Parameters:
    - doc_date (date): Date to assign to the SAP document fields
    - client_category (str, optional): SAP category indicator (e.g. 'D' for customer)
    - client_code (str, optional): SAP client code (currently inactive, pending field review)
    
    Returns:
    - str: Path to the modified batch template file
    """
    # Open Template file
    batch_template_path = Load_SAP_info.config["batch_template_path"]
    wb_template = xw.Book(batch_template_path)
    ws_template = wb_template.sheets[0]
   
    # Clear template before insert data
    temp_ini_row = 10
    
    temp_end_row = ws_template.range("D10").end("down").row
    if temp_ini_row < temp_end_row:
        ws_template.range(f'D{temp_ini_row}:D{temp_end_row}').clear_contents()
    
    # Complete all required fields before proceeding
    doc_date = datetime.strftime(doc_date,"%d.%m.%Y")
    ws_template.cells(2,5).value = doc_date
    ws_template.cells(2, 7).value = doc_date
    if client_category:
        ws_template.cells(6, 6).value = client_category
    
    if client_code:
        ws_template.cells().value = client_code # revisar
    
    # Callback the range selector Class to select the invoinces
    launch_range_selector(wb_template, 'D10')
    wb_template.save()
    wb_template.close()
    return batch_template_path
         
# ---------------------
# Main Programs
# ---------------------
def bank_file():
    """
    Prepares the daily bank movements file for SAP payment application.

    Cleans, formats, and validates today's bank data; reinserts not applied rows from the previous day;
    and prompts the user for row-level decisions to proceed with SAP posting.

    Workflow:
    - Open today's bank file and remove header rows
    - Replace dots from column C for cleaner matching
    - Open yesterday's payments file and identify "No Aplicado"(not apply) entries
    - Locate last matched payment from yesterday and truncate remaining rows from today's file
    - Remove unused columns to simplify layout
    - Set up dropdown validation for SAP action options
    - For each row:
        - Delete negative amounts and rows with unmatched descriptions
        - Assign concept, calculate text length, format assignment date
    - Create headers and apply conditional formatting
    - Reinsert "No Aplicado"(not apply) rows from previous day into today's sheet
    - Prompt user to review and begin row-by-row processing

    Returns:
    - None: updated Excel workbook is saved and ready for SAP interaction
    """
    # Open today's bank file
    bank_path = ask_open_file('Abre el fichero del banco de hoy')  
    wb = check_wb_open(bank_path)
    ws = wb.sheets[0]
    # Delete first 7 rows "bank headers"
    ws.range('1:7').delete()
    # Replace "." in column C
    col_c = ws.range('C:C')
    col_c.api.Replace(What=".", Replacement="", LookAt=2, MatchCase=False, SearchFormat=False, ReplaceFormat=False)
    # Open yesterday's payments file
    yest_file_path = ask_open_file('Abre los pagos del último día')
    wb_yest = check_wb_open(yest_file_path)
    ws_yest = wb_yest.sheets[0]
    # Get the Descriptio of the last payment
    last_pay = ws_yest.range('B2').value
    # Search for last payment in bank file and get its row
    find_result = col_c.api.Find(last_pay)
    if not find_result:
        print("Last payment not found.")
        return
    find_result_row = find_result.Row
    # Get No Aplicado rows
    dic_no_apply = {}
    end_row_yest = ws_yest.range("C1").end("down").row
    for i in range(2, end_row_yest + 1):
        if ws_yest.cells(i, 12).value == "No Aplicado":
            dic_no_apply[i] = i
    # Deletes all rows following the last known payment found in today's bank file
    last_row = ws.range("C1").end("down").row
    ws.range(f"{find_result_row}:{last_row}").delete()
    # Remove non-essential columns to simplify structure and setup headers
    setup_headers(ws, "bank_file")
    # Add validation list
    end_row = ws.range("A1").end("down").row
    val_list = ["SOLO", "TODO", "HASTA", "ENTRE", "RELACION", "REEMBOLSO", "A CUENTA", "FACTURA"]
    set_data_validation(ws,9,end_row,val_list,False)
    # Process rows: Delete negatives, format info
    for i in reversed(range(2, end_row + 1)):
        doc_date = datetime.strptime(ws.cells(i, 1).value, "%d/%m/%Y").date()
        assignment = doc_date.strftime("%Y%m%d")
        doc_date = doc_date.strftime("%d/%m/%Y")
        description = str(ws.cells(i, 2).value)
        amount = ws.cells(i, 3).value
        ws.cells(i, 7).formula = f"=LEN(F{i})"
        ws.cells(i, 8).value = assignment
        if amount < 0:
            ws.range(f"{i}:{i}").delete()
            continue
        concept =_new_concept(description, doc_date)
        if concept is None:
            ws.range(f"{i}:{i}").delete()
            continue
        ws.cells(i, 6).value = concept
    # Conditional formatting
    data_range = ws.range(f'G2:G{end_row}')
    data_range.api.FormatConditions.Add(1, 1, 1, "=50")  # xlCellValue, xlLessEqual
    data_range.api.FormatConditions(1).Font.Color = -11489280
    data_range.api.FormatConditions(1).Interior.Color = 13561798
    data_range.api.FormatConditions.Add(1, 1, 1, ">50")
    data_range.api.FormatConditions(2).Font.Color = -16776961
    data_range.api.FormatConditions(2).Interior.Color = 13551615
    # Reinsert No Aplicado rows
    insert_row = ws.range("A1").end("down").row
    if dic_no_apply:
        for row in dic_no_apply:
            ws_yest.range(f"{row}:{row}").copy()
            insert_row += 1
            ws.range(f"A{insert_row}").paste()
    wb_yest.close()
    wb.save()
    show_info("User Inputs", "Selecciona que hacer con cada pago")
    
    
def daily_payments():
    """    
    Automates SAP posting of daily bank payments listed in a treated Excel file.
    Interactively processes each payment row, determines SAP posting strategy,
    and executes entries using GUI scripting. ⚠️ Each accounting entry must be manually saved in SAP by the user.
    
    Workflow:
    - Load daily Excel file containing treated bank payments
    - Prompt for bank account configuration and client category
    - For each payment row:
        - Skip already marked rows
        - Determine posting logic based on the 'Acción' field (e.g., RELACION, FACTURA, TODO, etc.)
        - Load invoice detail or prepare manual input if needed
        - Apply debit and credit entries via SAP session
        - Handle custom scenarios like A CUENTA or REEMBOLSO
        - Simulate and confirm SAP transactions
        - Fill autogenerated accounting fields
        - Store SAP spool, mark row as 'Aplicado', and log entry number
    - Save the Excel workbook with status updates and notify user

    Returns:
    - None: entries are posted to SAP, results written to Excel, and confirmation shown at completion
    """
    # Flag
    Load_SAP_info.ContinueProgram = True
    # Default save path for SAP spool
    save_path = Load_SAP_info.config["spool_path"]
    # Prompt user to select treated daily bank file
    bank_path = ask_open_file("Abre el fichero del banco de hoy Tratado")
    if not bank_path:
        return
    # Load worksheet and determine number of data rows
    wb = check_wb_open(bank_path)
    ws = wb.sheets[0]  # Sheet index starts at 0
    end_row = ws.range("A1").end("down").row
    # SAP G/L Bank Account 
    bank_account = Load_SAP_info.config["bank_account"]
    # Prompt for vendor account context
    acc_confirmation = show_question("Confirmación","¿Hay alguna cuenta de Acreedor?")
    # Customer default category (Vendor category K, G/L category S)
    client_category = "D"
    # Iterate from bottom to top to preserve row integrity during actions
    for i in range(end_row, 1, -1):
        # Skip already processed rows
        if (ws.cells(i, 12).value or "").strip() == "Aplicado":
            show_info("Salto", f"Fila {i}: ya aplicado, se salta.")
            continue
        # Set payment metadata
        doc_date = datetime.strptime(ws.cells(i, 1).value,"%d/%m/%Y").date()
        due_date = datetime.strftime(doc_date,"%d.%m.%Y")
        assignment = ws.cells(i,8).value
        amount = float(ws.cells(i, 3).value)
        commentary = ws.cells(i, 6).value
        client_code = ws.cells(i, 4).value
        search_data1 = ws.cells(i, 10).value
        search_data2 = ws.cells(i, 11).value
        action = (ws.cells(i, 9).value or "").strip().upper()
        # Ask user for account category if needed
        if acc_confirmation == QMessageBox.Yes:
            cat, ok = QInputDialog.getText(None, "Categoría", f"Cliente Nº {client_code}\nD = Deudor / K = Acreedor:")
            if ok and cat.strip().upper() in ["D", "K"]:
                client_category = cat.strip().upper()
        # Handle missing action with status update
        if not action:
            title="Acción faltante"
            msg=f"Fila {i}: no se indicó acción."
            _pass_row(ws,i,title,msg)            
            continue
        # RELACION: load invoice details and handle manual entries if nedded
        elif action == "RELACION":
           # Confirm if invoinces must be selected
           no_pa = show_question("Confirmacion", "¿Es un pago sin PA?")
           # Confirm if the detail is in the Bank file (another worksheet)
           ask  = show_question("Confrimación", "¿Está la relación en el fichero del banco?")
           # Open file with the payment detail if ask match
           if ask == QMessageBox.No:
               payment_detail_path = ask_open_file(f"Abre el archivo con el detalle de Facturas {ws.range(f'f{i}').value} {ws.range(f'c{i}').value}")
               wb_payment_detail=check_wb_open(payment_detail_path)
           # Copy invoices into the template for SAP upload
           if no_pa == QMessageBox.No:
               batch_template_path = _load_template(doc_date,client_category)
               # Callback the SAP Transaction to load the Template
               batch_input(batch_template_path)
               if Load_SAP_info.ContinueProgram == False: pass
           # Process payments without associated invoices or item selection—typically used for manual input of multiple entries (Payment on account).
           else:
               call_transaction( "F-04")
           new_entry( "40", bank_account) # New Debit entry into G/L account
           if Load_SAP_info.ContinueProgram == False:
               _pass_row(ws,i)
               continue
           new_entry_add_data(amount, due_date,commentary,"-1",assignment)
           if Load_SAP_info.ContinueProgram == False:
               _pass_row(ws,i)
               continue
           # Loop for manual input Payment on account
           aux = True
           while aux:
               ask_SA = show_question("Confirmación", "¿Tiene algún apunte manual?")
               if ask_SA == QMessageBox.No:
                   aux = False
                   break
               ask_Rng =show_question("Confirmación", "¿Es un rango de apuntes?")
               # Apply manual entries
               if ask_Rng == QMessageBox.No:
                   VAC = ask_user_number("Introduce el importe del Apunte:")
                   if VAC is None:
                       _pass_row(ws, i)
                       break
                   if VAC > 0:
                       new_entry( "16", client_code) # New Credit entry into client account
                       if Load_SAP_info.ContinueProgram == False:
                           _pass_row(ws,i)
                           break
                       new_entry_add_data(VAC, due_date,commentary,"-1",assignment)
                       if Load_SAP_info.ContinueProgram == False:
                           _pass_row(ws,i)
                           break
                   elif VAC < 0:
                       VAC=VAC*-1
                       new_entry( "06", client_code) # New Debit entry into client account
                       if Load_SAP_info.ContinueProgram == False:
                           _pass_row(ws,i)
                           break
                       new_entry_add_data(VAC, due_date,commentary,"-1",assignment)
                       if Load_SAP_info.ContinueProgram == False:
                           _pass_row(ws,i)
                           break
                   ask_more=show_question("Confirmación", "¿Hay más apuntes manuales?")
                   if ask_more == QMessageBox.No:
                       aux =False
                       break
               # Request input range and apply selected amounts (Amouts are Selected, left cell commentary)
               else:
                   aux = False
                   if ask == QMessageBox.Yes:
                       range_selected=launch_range_selector(wb_payment_detail)
                   else:
                       range_selected=launch_range_selector(wb)    
                   # Iterating over individual cells
                   for row in range_selected.rows:
                        for cel in row:
                            ApVal = cel.value
                            try:
                                ComVal = cel.offset(0, -1).value
                            except Exception as e:
                                show_warning("Error", f"No se pudo obtener el comentario: {e}")
                                ComVal = ""
                            if ApVal > 0:
                                new_entry( "16", client_code)
                                if Load_SAP_info.ContinueProgram == False:
                                    _pass_row(ws,i)
                                    break
                                new_entry_add_data(ApVal, due_date,ComVal,"-1",assignment)
                                if Load_SAP_info.ContinueProgram == False:
                                    _pass_row(ws,i)
                                    break
                            elif ApVal < 0:
                                ApVal=ApVal*-1
                                new_entry( "06", client_code)
                                if Load_SAP_info.ContinueProgram == False:
                                    _pass_row(ws,i)
                                    break
                                new_entry_add_data(ApVal, due_date,ComVal,"-1",assignment)
                                if Load_SAP_info.ContinueProgram == False:
                                    _pass_row(ws,i)
                                    break
           if ask == QMessageBox.No:
               wb_payment_detail.close()
        # Other predefined actions
        # Handle specific 'Acción' scenarios like FACTURA, TODO, HASTA, SOLO, ENTRE, A CUENTA, REEMBOLSO
        else:
            # Call the add new entry SAP Transaction
            call_transaction("F-04")
            new_entry("40", bank_account, "", due_date) # New Debit entry into client account
            if Load_SAP_info.ContinueProgram == False:
                _pass_row(ws,i)
                continue
            new_entry_add_data(amount, due_date, commentary, "-1", assignment)
            if Load_SAP_info.ContinueProgram == False:
                _pass_row(ws,i)
                continue
            # Search for a specific Open Item (invoice) and select it
            if action == "FACTURA":
                search_items(client_category, 5, search_data1, "", client_code, search_data2) 
                if Load_SAP_info.ContinueProgram == False:
                    _pass_row(ws,i)
                    continue
            # Select all Open items in the account
            elif action == "TODO":
                search_items(client_category, 0, "", "", client_code) 
                if Load_SAP_info.ContinueProgram == False:
                    _pass_row(ws,i)
                    continue
            # Select all Open Items dated up to the specified due date
            elif action == "HASTA":
                date1 = search_data1.strftime("%d.%m.%Y")
                search_items(client_category, 16, "", "", client_code, date1)
                if Load_SAP_info.ContinueProgram == False:
                    _pass_row(ws,i)
                    continue
            # Select all Open Items with a specified due date
            elif action == "SOLO":
                date1 = search_data1.strftime("%d.%m.%Y")
                search_items(client_category, 16, date1, "", client_code)
                if Load_SAP_info.ContinueProgram == False:
                    _pass_row(ws,i)
                    continue
            # Select all Open Items between two specified dates
            elif action == "ENTRE":
                d1 = search_data1.strftime("%d.%m.%Y")
                d2 = search_data2.strftime("%d.%m.%Y")
                search_items(client_category, 16, d1, "", client_code, d2)
                if Load_SAP_info.ContinueProgram == False:
                    _pass_row(ws,i)
                    continue
            # A single On Account (Credit) entry 
            elif action == "A CUENTA":
                new_entry("16", client_code, "", due_date)
                if Load_SAP_info.ContinueProgram == False:
                    _pass_row(ws,i)
                    continue
                new_entry_add_data(amount, due_date, commentary, "-1", assignment)
                if Load_SAP_info.ContinueProgram == False:
                    _pass_row(ws,i)
                    continue
            # Add a Credit entry in favor of a given Vendor
            elif action == "REEMBOLSO":
                today = date.today()
                target_day = date(today.year + (1 if today.month == 12 else 0), (today.month % 12) + 1, 25) if today.day >= 8 else date(today.year, today.month, 25)
                fecha_reem = target_day.strftime("%d.%m.%Y")
                assignment_reem = target_day.strftime("%Y%m%d")
                commentary_reem = f"Tr. Reemb. Dronas OS {search_data1}"
                new_entry("36", client_code, "", due_date) # Credit Vendor account
                if Load_SAP_info.ContinueProgram == False:
                    _pass_row(ws,i)
                    continue
                new_entry_add_data(amount, fecha_reem, commentary_reem, "-1", assignment_reem)
                if Load_SAP_info.ContinueProgram == False:
                    _pass_row(ws,i)
                    continue
            # No additional actions are defined at this stage
            else:
                QMessageBox.warning(None, "Acción no válida", f"Fila {i}: Acción '{action}' no reconocida.")
                continue
        # Simulate accounting entry and generate final positions
        pos_ini, pos_fin = simulate(client_code,due_date)
        if Load_SAP_info.ContinueProgram == False:
            _pass_row(ws,i)
            continue       
        # Fill all autogenerated fields from simulated accounting data
        for j in range(pos_ini + 1, pos_fin+1):
            enter_position(j)
            new_entry_add_data(0,due_date,commentary,"-1",assignment)
            if Load_SAP_info.ContinueProgram == False:
                _pass_row(ws,i)
                continue
        # Manual intervention required: user must verify each entry and save in SAP manually (no automated save supported)
        save_confirmation()
        if Load_SAP_info.ContinueProgram == False:
            _pass_row(ws, i)
            continue
        # Retrieve the entry number generated during the accounting process
        entry_num = get_entry_number()
        if Load_SAP_info.ContinueProgram == False:
            _pass_row(ws,i)
            continue
        # Store the generated spool document on the server as a backup copy
        save_entry(save_path)
        if Load_SAP_info.ContinueProgram == False:
            _pass_row(ws,i)
            continue
        # Update row with new status
        ws.cells(i, 12).value = "Aplicado"
        ws.cells(i, 13).value = entry_num
        ws.range(f'{i}:{i}').api.Font.ColorIndex  = -4105
    # Save the workbook and prompt user to review everything
    ws.range("L:M").autofit()
    wb.save(bank_path)
    show_info("Finalizado", "Pagos diarios aplicados correctamente.\nRevisa los números de asiento.")

# ---------
# Debug
# ---------   
# Saveguard
if __name__ == "__main__":
    bank_file()
    print("Nice")