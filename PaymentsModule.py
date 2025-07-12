# -*- coding: utf-8 -*-
"""
@author: JesusMMA
"""

import xlwings as xw
import os
from PyQt5.QtWidgets import QMessageBox
from datetime import datetime
from SAPAux import (call_transaction, new_entry, new_entry_add_data,
                    search_items, simulate, enter_position, get_entry_number,
                    batch_input,back_to_main,enter_ajd,sap_data,items_found_sap,
                    handle_dif
)
from UserInputs import show_info,show_question,show_warning,save_confirmation
from Utilities import detail_handler
import Load_SAP_info



def payment_batch_template(clients_dic:dict, client_detail:dict, payment_dic:dict, wb, ws, payment_detail_path:str):
    """
    From validated detail file, populates and submits SAP batch entries for accounting,
    including debits, credits, and corrections. Finalizes results in a worksheet and performs cleanup.
    ⚠️ SAP entry confirmation and final save must be performed manually by the user.

    Workflow:
    - Extracts critical fields from preprocessed dictionaries
    - Connects to SAP and loads batch template
    - Inserts main payment entry with Special G/L indicator
    - Loops through clients to load debit and credit entries
    - Processes unmatched entries individually based on corporate name
    - Validates SAP data and resolves discrepancies if detected
    - Simulates final accounting positions and submits for confirmation
    - Retrieves SAP entry number, saves output file, and deletes original source

    Parameters:
    - clients_dic (dict): Dictionary mapping client names to their SAP codes
    - client_detail (dict): Dictionary with structural info from config about file layout and metadata
    - payment_dic (dict): Dictionary containing aggregated payment data
    - wb (Workbook): Workbook object of the source Excel file
    - ws (Worksheet): Worksheet object of the source Excel file
    - payment_detail_path (str): Path to the original file to be replaced

    Returns:
    - None: writes results to disk, logs SAP entry number, and removes source file
    """
    # Flag
    Load_SAP_info.ContinueProgram = True
    # Extract primary data and metadata needed for SAP entry
    client_name = payment_dic["client_name"] 
    due_date = payment_dic["due_date"] 
    due_date_assignment = payment_dic["due_date_assignment"] 
    doc_date = payment_dic["doc_date"] 
    doc_date_assignment = payment_dic["doc_date_assignment"] 
    payment_amount = payment_dic["payment_amount"] 
    invoices_amount = payment_dic["invoices_amount"] 
    debit_amounts = payment_dic["debit_amounts"] 
    credit_amounts = payment_dic["credit_amounts"] 
    entries_dic = payment_dic["entries_dic"] 
    invoices_dic = payment_dic["invoices_dic"]
    client_code = clients_dic[client_name.upper()]
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
    # Derive client-specific codes and values from config
    payment_method = client_detail["payment_method"]
    SGLIndicador = client_detail["SGLIndicator"]
    start_row = client_detail["start_row"]
    amount_col=client_detail["amount_col"]
    inv_ref_col = client_detail["inv_ref_col"]
    entry_comment = client_detail["entry_comment"]   
    # Manage all possibles corp_names data     
    corp_name_detail = client_detail["corp_name"]
    if isinstance(corp_name_detail, list):
        corp_name_row = corp_name_detail[0]
        corp_name_col = corp_name_detail[1]
        corp_name = (ws.cells(corp_name_row,corp_name_col).api.Text).strip().upper()
    elif isinstance(corp_name_detail, str):
        corp_name = corp_name_detail
    # Retrieve the payment number from the file
    payment_number = client_detail["payment_number"]
    if isinstance(payment_number, int):
        payment_number_col = payment_number
        payment_number_row = start_row
    elif isinstance(payment_number, list):
        payment_number_row = payment_number[0]
        payment_number_col = payment_number[1]
    payment_number = ws.cells(payment_number_row,payment_number_col).api.Text
    # Calculate commentary lines for clarity in SAP logs
    commentary = f"{payment_method}. {client_name} {payment_number} vto. {due_date}"
    debt_commentary = f"TOTAL CARGOS {client_name} {payment_number} vto. {due_date}"
    cred_commentary = f"TOTAL ABONOS {client_name} {payment_number} vto. {due_date}"
    ajd_comentary = f"GASTOS AJD {client_name} {payment_number} vto. {due_date}"
    # Callback the SAP Transaction to load the template
    batch_template_path = Load_SAP_info.config["batch_template_path"]
    batch_input(batch_template_path)
    if Load_SAP_info.ContinueProgram == False: return
    # Initiate promissory note debit with appropriate GL indicator
    if payment_method == "Pago Unif.":
        new_entry( "06", client_code)    
    else:
        new_entry( "09", client_code, SGLIndicador)
    if Load_SAP_info.ContinueProgram == False: return
    new_entry_add_data(payment_amount, due_date,commentary,"-1",doc_date_assignment)
    if Load_SAP_info.ContinueProgram == False: return
    # Iterate through each client and add appropiate entry matching debit/credit based on totals
    for key in clients_dic:
        debit = debit_amounts[key]
        credit = credit_amounts[key]
        client_code_loop = clients_dic[key]
        # Add debit entry to the appropriate client
        if debit != 0:
            debit = debit * -1
            new_entry( "06", client_code_loop)
            if Load_SAP_info.ContinueProgram == False: return
            new_entry_add_data(debit, due_date,debt_commentary,"-1",due_date_assignment)
            if Load_SAP_info.ContinueProgram == False: return
        # Add credit entry to the appropriate client
        if credit != 0:
            new_entry( "16", client_code_loop)
            if Load_SAP_info.ContinueProgram == False: return
            new_entry_add_data(credit, due_date,cred_commentary,"-1",due_date_assignment)
            if Load_SAP_info.ContinueProgram == False: return
    if isinstance(inv_ref_col, list):
        inv_ref_col = inv_ref_col[1]
    # Iterate through each direct entries and assign it to the corresponding client
    for entry_key in entries_dic:
        amount = float(ws.cells(entry_key,amount_col).value)
        inv_ref = ws.cells(entry_key,inv_ref_col).api.Text
        entry_commentary = f"CARGO {inv_ref} {entry_comment}"
        if isinstance(corp_name_detail, int):
            corp_name = (ws.cells(entry_key,corp_name_detail).api.Text).strip().upper()        
        for name_key in clients_dic:
            if name_key in corp_name:
                client_code_loop = clients_dic[name_key]
                # Add credit entry to the appropriate client
                if amount > 0:
                    new_entry( "16", client_code_loop)
                    if Load_SAP_info.ContinueProgram == False: return
                    new_entry_add_data(amount, due_date,entry_commentary,"-1",due_date_assignment)
                    if Load_SAP_info.ContinueProgram == False: return
                # Add debit entry to the appropriate client
                elif amount < 0:
                    amount = amount * -1
                    new_entry( "06", client_code_loop)
                    if Load_SAP_info.ContinueProgram == False: return
                    new_entry_add_data(amount, due_date,entry_commentary,"-1",due_date_assignment)
                    if Load_SAP_info.ContinueProgram == False: return
    # Apply ajd taxes if nedeed
    ajd_amount = payment_dic["ajd_amount"]
    ajd_assignment = client_detail["ajd_assignment"]
    if ajd_amount != 0:
        enter_ajd(ajd_amount,ajd_assignment,ajd_comentary,due_date)
        if Load_SAP_info.ContinueProgram == False: return
    # Retrieve critical data from the SAP entry
    result_data = sap_data()
    if Load_SAP_info.ContinueProgram == False: return
    dif_amount = result_data["dif_amount"]
    total_items_loaded = result_data["total_items_loaded"] 
    items_amount = result_data["items_amount"]
    total_amount = result_data["total_amount"]
    # Check for differences in the SAP entry and handle them accordingly
    if dif_amount != 0:
        total_invoices = len(invoices_dic)
        invoices_dif = invoices_amount - items_amount
        # Fix unloaded invoices by tracing missing references and add entry
        if  invoices_dif != 0 and total_invoices != total_items_loaded:
            ask = show_question("Confirmación",
                                f"Hay diferencia en las Facturas: {dif_amount}\n¿Quieres ajustar la diferencia?")
            if ask == QMessageBox.No:
                show_info("Cancelación", "No se ha ajustado la diferencia")
                return
            invoices_SAP_ref = items_found_sap()
            for inv_key in invoices_dic:
                if not inv_key in invoices_SAP_ref:
                    inv_row = invoices_dic[inv_key]
                    inv_amount = ws.cells(inv_row,amount_col).value
                    if isinstance(corp_name_detail, int):
                        inv_corp = (ws.cells(inv_row,corp_name_detail).api.Text).strip().upper()
                    else:
                        inv_corp = (ws.cells(inv_row,corp_name_col).api.Text).strip().upper()
                    for name_key in clients_dic:
                        if name_key in inv_corp:
                            client_code_loop = clients_dic[name_key]
                            if inv_amount > 0:
                                inv_commentary = f"PAGA FACTURA {inv_key}"
                                new_entry( "16", client_code_loop)
                                if Load_SAP_info.ContinueProgram == False: return
                                new_entry_add_data(inv_amount, due_date,inv_commentary,"-1",due_date_assignment)
                                if Load_SAP_info.ContinueProgram == False: return
                            elif inv_amount < 0:
                                inv_amount = inv_amount * -1
                                inv_commentary = f"SE DESCUENTA ABONO {inv_key}"
                                new_entry( "06", client_code_loop)
                                if Load_SAP_info.ContinueProgram == False: return
                                new_entry_add_data(inv_amount, due_date,inv_commentary,"-1",due_date_assignment)
                                if Load_SAP_info.ContinueProgram == False: return
        # If discrepancy is due to rounding, apply final adjustment to match totals
        elif invoices_dif != 0 and total_invoices == total_items_loaded:
            show_info("Diferencia", "La diferencia esta en centimos acumulados")
            handle_dif(dif_amount,client_code,due_date)
            if Load_SAP_info.ContinueProgram == False: return
    # Simulate accounting entry and generate final positions
    pos_ini, pos_fin = simulate(client_code,due_date)
    if Load_SAP_info.ContinueProgram == False: return
    # Fill all autogenerated fields from simulated accounting data
    for j in range(pos_ini + 1, pos_fin+1):
        enter_position(j)
        new_entry_add_data(0,due_date,commentary,"-1",due_date_assignment)
        if Load_SAP_info.ContinueProgram == False: return
    # Manual intervention required: user must verify each entry and save in SAP manually (no automated save supported)
    save_confirmation()
    if Load_SAP_info.ContinueProgram == False:
        show_info("Cancelación", "Se cancela el proceso")
        return
    # Retrieve the entry number generated during the accounting process
    entry_num = get_entry_number()
    if Load_SAP_info.ContinueProgram == False: return
    # Save updated Excel file using SAP entry number within the filename
    folder = os.path.dirname(payment_detail_path)
    new_path = os.path.join(folder, f"{entry_num} {client_name} {payment_amount}.xlsx")    
    wb.save(new_path)
    wb.close()
    show_info("Fin", f"Se ha aplicado el asiento {entry_num} y guardado el fichero.")
    # Delete original source file after successful completion
    if os.path.exists(payment_detail_path):
        os.remove(payment_detail_path)
        print("Old file deleted.")
    else:
        print("Old file not found.")     

def payment_search_amount(client_name,client_aux_name=""):
    """
    Automates the SAP entry process for promissory note payments.
    Performs SAP searches based on the lump-sum payment total previously grouped under 'PAGO UNIFICADO'.
    Includes AJD taxes added at the time each promissory note is created, which must be accounted for.

    Loads configuration parameters and prompts the user to prepare input data in a template file:
        A    ||       B         ||    C   ||        D      ||       E        ||     F     ||     G    ||          H         ||  I  ||       J       ||
    DOC DATE || ACCOUNTING DATE || AMOUNT || DOC ASSIGMENT || PAYMENT NUMBER || COMENTARY || DUE DATE || DUE DATE ASSIGMENT || AJD || SEARCH AMOUNT ||
    
    Then processes an Excel template,
    calculates relevant financial fields, executes SAP transaction F-04 with Special G/L indicators, and fills
    in generated accounting details. User is required to manually verify and confirm each transaction entry in SAP.
    Successful entries are saved to the worksheet including their corresponding SAP reference numbers.
   
    Workflow:
    - Loads client-specific configuration and input template
    - Clears previous entries and prompts user to prepare new input data
    - Iterates through each promissory note:
        - Parses key fields (dates, amounts, commentary)
        - Calculates search amount by adding AJD taxes
        - Fills tracking data into the template
        - Executes SAP entry flow:
            - Initiates F-04 transaction
            - Creates debit entry and adds data
            - Enters AJD tax values
            - Searches matching SAP entries
            - Simulates and confirms each accounting position
        - SAP reference number is saved to worksheet
 
    Parameters:
    - client_name (str): Client identifier used for template protection and naming
 
    Returns:
    - None: writes SAP entry results to Excel and displays warnings or messages as needed
    """
    try:
        # Flag
        Load_SAP_info.ContinueProgram = True
        # Load configuration and define SAP-relevant variables
        client_name_lower = client_name.lower()
        client_dic = Load_SAP_info.config[f"{client_name_lower}_dic"]
        client_code = client_dic[client_name.upper()]
        company_code = Load_SAP_info.config["company_code"]
        client_detail = Load_SAP_info.config[f"{client_name_lower}_pag_detail"]
        ajd_assignment = client_detail["ajd_assignment"]
        # Open and clear previous template data
        template_path = Load_SAP_info.config["unify_template_path"]
        wb = xw.Book(template_path)
        ws = wb.sheets[0]
        ws.api.Unprotect(Password=client_name)
        end_row = ws.range("A1").end("down").row
        ws.range(f"A2:K{end_row}").clear_content()
        ws.api.Protect(Password=client_name)
        # Prompt user to prepare the input worksheet before starting
        ask = show_question("Input de Usuario", "Prepare el fichero con los datos necesarios")
        # Exit early if user cancels or SAP session cannot be initiated
        if ask == QMessageBox.No:
            show_info("Cancelación", "Cancelado por el Usuario")
            return
        ws.api.Unprotect(Password=client_name)
        end_row = ws.range("A1").end("down").row
        # Iterate through rows and parse all relevant fields for each promissory note
        for i in range(2, end_row + 1):
            # SAP due_date
            due_date = datetime().strptime(ws.cells(i, 7).value,"%d/%m/%Y").date() # Due Date
            due_date_assignment = due_date.strftime("%Y%m%d")
            due_date = due_date.strftime("%d.%m.%Y") # Due Date
            # SAP entry date
            doc_date = datetime.strptime(ws.cells(i, 1).value,"%d/%m/%Y").date()
            doc_date_assignment = doc_date.strftime("%Y%m%d")
            doc_date = doc_date.strftime("%d.%m.%Y") # Due Date
            # AJD Taxes amount
            ajd = round(float(ws.cells(i, 9).value),2) # Promissory note Taxes
            # Promissory note Total amount
            amount = round(float(ws.cells(i, 3).value),2) # Promissory note Amount
            # Amount to search in open tems in SAP
            search_amount = round(amount + ajd,2) # Add AJD amount to payment value to calculate the searchable total
            payment_number = ws.cells(i, 5).value # Promissory note Number
            commentary = f"PAG. {client_name} {payment_number} VTO. {due_date}" # Commentary to add to the SAP entry
            ajd_comentary = f"GASTOS AJD {client_name} {payment_number} VTO. {due_date}"
            # Populate worksheet fields with derived values for traceability
            ws.cells(i, 2).value = doc_date
            ws.cells(i, 4).value = doc_date_assignment
            ws.cells(i, 6).value = commentary
            ws.cells(i, 8).value = due_date_assignment
            ws.cells(i, 10).value = search_amount
            # Begin SAP input flow
            # Call the new entry transaction
            call_transaction("F-04")
            if Load_SAP_info.ContinueProgram == False: continue
            # Debit into client Account with SGL Ind. (Special G/L indicator) Promissory note
            SGLIndicator = client_detail["SGLIndicator"]
            new_entry( "09", client_code,SGLIndicator,doc_date)
            if Load_SAP_info.ContinueProgram == False: continue
            # Add commentary and due dates for standard and tax entries
            new_entry_add_data(amount, due_date,commentary,"-1",doc_date_assignment)
            if Load_SAP_info.ContinueProgram == False: continue
            enter_ajd(ajd, ajd_assignment, ajd_comentary, due_date)
            if Load_SAP_info.ContinueProgram == False: continue
            # Search by amount
            client_category = client_detail["client_category"]
            search_items( client_category, 1,search_amount,company_code,client_code)
            if Load_SAP_info.ContinueProgram == False: continue
            # If user cancels during entry, back out and continue safely
            if Load_SAP_info.ContinueProgram == False:
                show_info("Cancelación","Se cancela el proceso")
                back_to_main()
                continue
            # Simulate accounting entry and generate final positions
            pos_ini, pos_fin = simulate(client_code,due_date)
            if Load_SAP_info.ContinueProgram == False: continue
            # Fill all autogenerated fields from simulated accounting data
            for j in range(pos_ini + 1, pos_fin+1):
                enter_position(j)
                new_entry_add_data(0,due_date,commentary,"-1",due_date_assignment)
                if Load_SAP_info.ContinueProgram == False: continue
            # Manual intervention required: user must verify each entry and save in SAP manually (no automated save supported)
            save_confirmation()
            if Load_SAP_info.ContinueProgram == False:
                show_info("Cancelación", "Se cancela el proceso")
                continue
            # Store SAP-generated reference number into the worksheet
            entry_num = get_entry_number()
            if Load_SAP_info.ContinueProgram == False: continue
            ws.cells(i, 11).value = entry_num
        # Finalize file and reapply protection
        ws.api.Protect(Password=client_name)
        wb.save()
    except Exception as e:
        if isinstance(e, OverflowError):
            pass
        else:
            show_warning("Error",f"{type(e).__name__} - {str(e)}")

# ------------------
# Specific Programs
# ------------------

def payment (client_name):
    """
    Entry point for processing a client's  payment workflow.
    Dynamically routes to the appropriate handler based on client identity and configuration.

    Workflow:
    - For 'Alcampo', initiates lump-sum payment automation and AJD handling via `_pag_search_amount`
    - For other clients, determines auxiliary configuration (e.g. 'Cecosa' → 'eroski')
    - Launches promissory note detail handler to categorize and prepare payments via `_pag_detail_handler`

    Parameters:
    - client_name (str): Name of the client to determine processing path

    Returns:
    - None: each subroutine performs its own SAP interaction and file handling
    """
    if client_name == "Alcampo":
        payment_search_amount(client_name)
    else:
        if client_name == "Cecosa":
            client_aux_name = "eroski"
        elif "El Corte Ingles" in client_name:
            client_name_aux = client_name.replace("El Corte Ingles", "ECI")
            client_name_aux = client_name_aux.replace(" ","_")
            client_name = "ECI"
        elif client_name == "Casa del Libro":
            client_name = "CDL"
        elif "Alcampo" in client_name:
            if "Pago Unif" in client_name:
                client_aux_name = "alcampo_pago_unif"
            else:
                client_aux_name ="alcampo_pag"
            client_name ="Alcampo"
        detail_handler(client_name,client_aux_name)
              
# ---------
# Debug
# ---------   
# Saveguard
if __name__ == "__main__":
    print("Nice")