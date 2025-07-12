# -*- coding: utf-8 -*-
"""
@author: JesusMMA
"""
import xlwings as xw
import Load_SAP_info
import os
import time
from SAPAux import call_transaction, call_variant, SAPSessionManager, back_to_main, run_background_job
from UserInputs import ask_user_date, ask_open_file, ask_user_string,show_info, ask_open_files,show_warning
from datetime import datetime, timedelta
from Utilities import (check_wb_open,split_by_filter,setup_headers,
                       merge_sheets,set_data_validation,
                       get_unique_column_values
)

def _export_sap_file(filename, folder_path):
    """
    Automates the export of an SAP financial report using transaction FBL5N and a predefined variant.

    Workflow:
    - Checks for an active SAP GUI session; initiates connection if missing
    - Launches transaction FBL5N in SAP
    - Applies the configured report variant for consistent filtering
    - Prompts user for the report’s end-of-month date
    - Populates the report field and triggers execution
    - Handles optional pop-up dialog if present
    - Navigates SAP menu to initiate export functionality
    - Sets export destination path and filename
    - Finalizes export process and returns to main SAP screen

    Parameters:
    - filename (str): Desired name of the exported report file
    - folder_path (str): Full path where the report should be saved

    Returns:
    - None: Export file to disk
    """
    # Ensure SAP GUI session is active
    session = SAPSessionManager.session
    if not session:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    # Open SAP transaction FBL5N
    call_transaction("FBL5N")
    if Load_SAP_info.ContinueProgram == False: return
    # Load variant for customized report filters/settings
    call_variant(r"\AM PA GESTOR")
    if Load_SAP_info.ContinueProgram == False: return
    # Ask user for report cutoff date
    work_date = ask_user_date("Introduce el último día del mes")
    # Fill date field and run report
    session.findById("wnd[0]/usr/ctxtPA_STIDA").text = work_date
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    # Handle optional pop-up window if present
    if session.Children.Count > 1:
        session.findById("wnd[1]").sendVKey(0)
    # Trigger export from SAP menu
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    # Provide export file details
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = folder_path
    # Confirm and finalize export
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    # Return to SAP main screen
    back_to_main()
    if Load_SAP_info.ContinueProgram == False: return



def _copy_previous_report(wb, report_path, sheets_to_copy):
    """
    Copies specific sheets from a previous workbook into the current one.
    
    Workflow:
    - Opens the previous workbook from the given path
    - Iterates through its sheets and copies only those that match `sheets_to_copy`
    - Inserts each copied sheet after the last sheet of the destination workbook
    - Closes the previous workbook to free resources
    
    Parameters:
    - wb: Excecl Workbook
    - report_path (str): path to the report to copy 
    - sheets_to_copy (list[str]): List of sheet names to copy
    
    Returns:
    - None: updates Excel file
    """
    # Copy sheets into the excel file
    wb_last = check_wb_open(report_path)
    for sheet in wb_last.sheets:
        if sheet.name in sheets_to_copy:
            sheet.copy(after=wb.sheets[-1])
    wb_last.close()



def _compare_and_copy(ws_ini, ws_base, wb):
    """
    Compares rows from an initial sheet with a base sheet using keys based on columns F, G, and J.
    Copies corresponding comments and account manager data into matching rows, and assigns account managers by client code.
    
    Workflow:
    - Calculates last filled rows in both sheets for accurate range limits
    - Builds dictionaries from `ws_base` keyed by F|G|J for lookups
    - Builds a secondary dictionary mapping client codes to row positions
    - Reads account manager assignments from the "CUENTAS CON GESTOR" sheet
    - For each row in `ws_ini`, checks for a matching composite key in `dict_base`
    - If found, transfers values from columns L, M, and O
    - Assigns account manager to column N based on client code via `gest_dict`
    
    Parameters:
    - ws_ini: Excel Worksheet with initial report data
    - ws_base: Excel Worksheet acting as the comparison reference
    - wb: Excel Workbook containing the account manager reference sheet
    
    Returns:
    - None: updates worksheet cells
    """
    last_ini_row = ws_ini.range("F" + str(ws_ini.cells.last_cell.row)).end("up").row
    last_base_row = ws_base.range("F" + str(ws_base.cells.last_cell.row)).end("up").row
    dict_base = {}
    dict_clients = {}
    for j in range(2, last_base_row + 1):
        cellF = ws_base.range(f"F{j}").value
        cellG = ws_base.range(f"G{j}").value
        cellJ = ws_base.range(f"J{j}").value
        key = f"{cellF}|{cellG}|{cellJ}"
        dict_base[key] = j
        dict_clients[cellG] = j
    gest_dict = {row[0]: row[2] for row in wb.sheets["CUENTAS CON GESTOR"].range("A1:C100").value if row and row[0]}
    for i in range(2, last_ini_row):
        cellF = ws_ini.range(f"F{i}").value
        cellG = ws_ini.range(f"G{i}").value
        cellJ = ws_ini.range(f"J{i}").value
        key = f"{cellF}|{cellG}|{cellJ}"
        if key in dict_base:
            row = dict_base[key]
            ws_ini.range(f"L{i}").value = ws_base.range(f"L{row}").value
            ws_ini.range(f"M{i}").value = ws_base.range(f"M{row}").value
            ws_ini.range(f"O{i}").value = ws_base.range(f"O{row}").value
        ws_ini.range(f"N{i}").value = gest_dict.get(ws_ini.range(f"G{i}").value, "")
    show_info("Fin", "Se han copiado los datos a la hoja {ws_ini.name}")



def large_format_retailers_file():
    """
     Orchestrates the entire report processing flow:
     loading templates, merging data, applying logic, splitting by account manager, and saving results.
    
     Workflow:
     - Constructs file path for today’s export
     - Optionally triggers SAP export (commented out)
     - Opens generated workbook and sets header formatting
     - Loads previous report and copies relevant sheets
     - Creates base sheet and merges account manager data
     - Compares and copies values between sheets
     - Deletes final empty row to clean up layout
     - Applies data validation to management columns
     - Splits data into individual sheets per account manager
     - Prompts user to rename the initial sheet and save the final output
     - Closes workbook and deletes temporary working file
    
     Parameters:
     - None (wrapped workflow with internal prompts and constants)
    
     Returns:
     - None: process concludes with saved Excel report
     """
    today_str = datetime.today().strftime("%d.%m.%Y")
    filename = f"fichero {today_str}.xlsx"
    folder_path = r"C:\\Users\\xexu_\\Desktop\\"
    full_path = os.path.join(folder_path, filename)
    _export_sap_file(filename, folder_path)
    time.sleep(5)
    wb = check_wb_open(full_path)
    ws_ini = wb.sheets[0]
    setup_headers(ws_ini,"large_retail_report")
    last_report_path = ask_open_file("Abre el Informe del Mes Anterior")
    _copy_previous_report(wb, last_report_path, Load_SAP_info.config["report_sheets_copy"])
    # Add sheet named 'Base datos'
    base_sheet = wb.sheets.add(after=wb.sheets[-1])
    base_sheet.name = "Base datos"
    ws_ini.range("1:1").copy(base_sheet.range("1:1"))
    merge_sheets(wb, base_sheet, Load_SAP_info.config["account_managers"])
    _compare_and_copy(ws_ini, base_sheet, wb)
    last_row = ws_ini.range("J" + str(ws_ini.cells.last_cell.row)).end("up").row
    ws_ini.range(f"{last_row}:{last_row}").delete()
    last_row = ws_ini.range("J" + str(ws_ini.cells.last_cell.row)).end("up").row
    # Get column indices and last row
    gest_col = ws_ini.range("1:1").value.index("GESTION") + 1
    cuadre_col = ws_ini.range("1:1").value.index("CUADRE") + 1
    # Reference ranges from another worksheet
    ws_acciones = wb.sheets["ACCIONES"]
    gest_range = ws_acciones.range("C2").expand("down")
    cuadre_range = ws_acciones.range("A2").expand("down")
    # Apply validations to each column
    set_data_validation(ws_ini, gest_col, last_row, gest_range, use_range=True)
    set_data_validation(ws_ini, cuadre_col, last_row, cuadre_range, use_range=True)
    split_by_filter(wb,ws_ini, 14)
    new_name = ask_user_string("Ingrese el nuevo nombre para la primera hoja: ")
    if new_name:
        ws_ini.name = new_name
        save_path = xw.apps.active.api.GetSaveAsFilename(FileFilter="Archivos de Excel (*.xlsx), *.xlsx")
        if save_path and isinstance(save_path, str):
            if not save_path.endswith(".xlsx"):
                save_path += ".xlsx"
            wb.save(save_path)
            show_info("Fin",f"Archivo guardado en '{save_path}'")
    wb.close()
    os.remove(full_path)

def generate_sap_files_balance_report():
    """
    SAP Balance Report Step 1:  
    Automates creation of monthly and yearly financial reports via SAP GUI for a selected year.
    
    Workflow:
    - Ensures SAP GUI session is active
    - Opens FBL5N transaction and loads predefined variant
    - Prompts user to input report year
    - Executes background job for full year range
    - Iteratively executes jobs for each calendar month
    - Confirms job generation and prompts user to check SM37
    
    Parameters:
    - None (user is prompted within the workflow)
    
    Returns:
    - None: jobs are triggered in SAP and status messages are shown
    """

    # Ensure SAP GUI session is active
    session = SAPSessionManager.session
    if not session:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    # Access transaction FBL5N and load preset variant
    call_transaction("FBL5N")
    if Load_SAP_info.ContinueProgram == False: return
    call_variant("am.fact.ctevta")
    if Load_SAP_info.ContinueProgram == False: return
    # Clear customer filter field and continue
    session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").Text = ""
    session.findById("wnd[0]").sendVKey(0)
    # Prompt user for year input and validate it    
    current_year = datetime.now().year
    while True:
        user_year = ask_user_string("Introduce el Año del informe (Formato AAAA): ")
        if user_year.isdigit() and len(user_year) == 4 and int(user_year) <= current_year:
            user_year = int(user_year)
            break
        show_warning("Año no válido, introduzca el año en Formato: AAAA")
    # Set date range for full year
    first_day = datetime(user_year, 1, 1).strftime("%d.%m.%Y")
    last_day = datetime(user_year, 12, 31).strftime("%d.%m.%Y")
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = first_day
    session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = last_day
    # Run job for full year period
    run_background_job()    
    time.sleep(5)
    # Run separate job for each month
    for i in range(1, 13):
        time.sleep(5)
        first_date = datetime(user_year, i, 1).strftime("%d.%m.%Y")
        if i < 12:
            last_date = datetime(user_year, i + 1, 1) - timedelta(days=1)
        else:
            last_date = datetime(user_year, 12, 31)
        last_date = last_date.strftime("%d.%m.%Y")
        
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = first_date
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = last_date
        run_background_job()
    # Notify completion and return to main menu
    show_info("Fin","✅ Jobs en fondo generados. Revisa SM37 y ejecuta el siguiente programa cuando estén listos.")
    back_to_main()
    if Load_SAP_info.ContinueProgram == False: return

def download_files_balance_report():
    """
    SAP Balance Report Step 2:  
    Automates download of completed spool files (TXT format) from SAP job monitor.
    
    Workflow:
    - Ensures SAP GUI session is active
    - Opens SM37 and filters by today's date
    - Filters jobs by 'Finished' status only
    - Iterates through job list to:
      - Select spool outputs
      - Export results as TXT files
      - Handle errors gracefully per row
    - Displays completion message once files are downloaded
    
    Parameters:
    - None (relies on today's date and internal logic)
    
    Returns:
    - None: spool files are saved locally by SAP GUI interaction
    """
    # Ensure SAP GUI session is active
    session = SAPSessionManager.session
    if not session:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    # Open SM37 job monitoring transaction
    call_transaction(session, "SM37")
    if Load_SAP_info.ContinueProgram == False: return
    # Filter jobs by today’s date
    today_str = datetime.now().strftime("%d.%m.%Y")
    session.findById("wnd[0]/usr/ctxtBTCH2170-FROM_DATE").Text = today_str
    session.findById("wnd[0]/usr/ctxtBTCH2170-TO_DATE").Text = today_str
    # Set status filters, only include finished jobs
    session.findById("wnd[0]/usr/chkBTCH2170-SCHEDUL").Selected = False
    session.findById("wnd[0]/usr/chkBTCH2170-READY").Selected = False
    session.findById("wnd[0]/usr/chkBTCH2170-RUNNING").Selected = False
    session.findById("wnd[0]/usr/chkBTCH2170-ABORTED").Selected = False
    session.findById("wnd[0]/usr/chkBTCH2170-FINISHED").Selected = True
    # Execute search
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    # Loop through found jobs and download their spool files
    for i in range(1, 14):  # 1 to 13
        time.sleep(3)
        job_row = 12 + i  # Adjusting for row index
        try:
            session.findById(f"wnd[0]/usr/chk[1,{job_row}]").Selected = True  # Select Job
            session.findById("wnd[0]/tbar[1]/btn[44]").press()  # Goto Spool
            session.findById("wnd[0]/usr/chk[1,3]").Selected = True  # Select Spool
            session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[3]").Select()  # Save as TXT
            session.findById("wnd[0]/tbar[0]/btn[12]").press()  # Confirm Save
            session.findById(f"wnd[0]/usr/chk[1,{job_row}]").Selected = False  # Deselect Job
        except Exception as e:
            show_warning("Error",f"[ERROR] in row {job_row}: {e}")
    # Final status update
    show_info("Fin","Ya se han descargado los ficheros.")


def create_balance_report():
    """
    SAP Balance Report Step 3:  
    Generates formatted Excel report by combining monthly TXT exports into structured sheets.
    
    Workflow:
    - Prompts user to select TXT files (one per month + annual)
    - Creates Excel workbook with named sheets per month
    - For each file:
      - Opens TXT as Excel
      - Filters column C to remove non-matching rows
      - Deletes visible filtered rows
      - Copies cleaned sheet to main workbook
    - Deletes initial blank sheet
    - Displays completion message and allows user to save
    
    Parameters:
    - None (file selection handled via GUI prompt)
    
    Returns:
    - None: process concludes with open Excel file ready for saving
    """
    # Sheet names corresponding to periods
    sheet_names = [
        "ANUAL", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO",
        "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
    ]
    # Prompt user to select downloaded TXT files
    file_paths = ask_open_files("Selecciona todos los TXT descargados.")
    if not file_paths:
        show_warning("Error","No se seleccionaron archivos o el proceso fue cancelado.")
        return
    # Create new Excel workbook
    app = xw.App(visible=True)
    wb = app.books.add()
    for i, path in enumerate(file_paths):
        try:
            temp_wb = app.books.open(path)
            temp_ws = temp_wb.sheets[0]
            # Identify range and apply filter
            last_row = temp_ws.range("O" + str(temp_ws.cells.last_cell.row)).end("up").row
            data_range = temp_ws.range(f"A9:P{last_row}")
            data_range.api.AutoFilter(Field=3, Criteria1="=")
            # Delete filtered visible rows
            try:
                visible_rows = data_range.offset(1, 0).special_cells(xw.constants.CellType.visible)
                visible_rows.api.EntireRow.Delete()
            except:
                pass  # No rows to delete
            # Disable filter and copy sheet to main workbook
            temp_ws.api.AutoFilterMode = False
            temp_ws.copy(after=wb.sheets[-1])
            wb.sheets[-1].name = sheet_names[i]
            temp_wb.close(False)
        except Exception as e:
            show_warning("Error",f"Error procesando {path}: {e}")
    # Delete default blank sheet
    if len(wb.sheets) > len(file_paths):
        wb.sheets[0].delete()
    # Final notification
    show_info("Fin","✅ Informe Generado. Guarde el fichero como quiera.")

def zaging_1():
    """
    Debt Aging step 1:    
    Automates cleaning and comparison of financial reports between Zaging and Standar Excel files.
    
    Workflow:
    - Prompts user to open Zaging and Standar Excel files
    - Cleans up Zaging sheet and sets up headers
    - Maps existing clients from Zaging
    - Iterates through Standar data:
      - Updates matching client info
      - Adds missing clients to Zaging
      - Computes key financial metrics (diff, vto-diff, vencido)
      - Highlights important differences with color
    - Formats final report:
      - Deletes unnecessary columns
      - Autofits and hides columns
      - Applies numeric formatting
    - Shows completion message and saves/cleans up workbook
    
    Parameters:
    - None (wrapped workflow with internal prompts and constants)
   
    Returns:
    - None: process concludes with saved Excel file
    """
    # Prompt for input files
    path_zaging = ask_open_file("Abre el fichero del Zaging")
    if not path_zaging:
        return
    path_standar = ask_open_file("Abre el fichero del Standar")
    if not path_standar:
        return
    # Load workbooks and first worksheets
    wb_zaging = check_wb_open(path_zaging)
    wb_standar = check_wb_open(path_standar)
    ws_zaging = wb_zaging.sheets[0]
    ws_standar = wb_standar.sheets[0]
    # Initial clean-up and headers
    ws_zaging.range("10:10").delete()
    ws_zaging.range("1:8").delete()
    setup_headers(ws_zaging, "zaging")
    # Create client dictionary to track existing rows
    dic_clients_zag = {}
    last_z_row = ws_zaging.range("A" + str(ws_zaging.cells.last_cell.row)).end('up').row
    last_st_row = ws_standar.range("A" + str(ws_standar.cells.last_cell.row)).end('up').row
    for i in range(2, last_z_row + 1):
        client = ws_zaging.range(f"A{i}").value
        dic_clients_zag[client] = i
    # Iterate Standar rows and update data into Zaging
    for i in range(2, last_st_row + 1):
        client = ws_standar.range(f"A{i}").value
        nombre = ws_standar.range(f"B{i}").value
        st_total = ws_standar.range(f"C{i}").value
        max360 = ws_standar.range(f"G{i}").value or 0
        if client in dic_clients_zag:
            row = dic_clients_zag[client]
        elif client not in dic_clients_zag:
            last_z_row += 1
            row = last_z_row
            ws_zaging.range(f"A{last_z_row}").value = client
            dic_clients_zag[client]=last_z_row
            ws_zaging.range(f"B{last_z_row}").value = nombre
            ws_zaging.range(f"A{last_z_row}").api.EntireRow.Interior.Color = 0xDCE4FA
        ws_zaging.range(f"T{row}").api.FormulaR1C1 = "=R[-0]C[-2]-R[-0]C[-1]"
        ws_zaging.range(f"U{row}").api.FormulaR1C1 = "=R[-0]C[-11]+R[-0]C[-1]"
        ws_zaging.range(f"O{row}").value = (ws_zaging.range(f"P{row}").value or 0) - max360
        col_sum = sum(cell or 0 for cell in ws_zaging.range((row, 3), (row, 9)).value)
        ws_zaging.range(f"J{row}").value = col_sum
        ws_zaging.range(f"R{row}").api.FormulaR1C1 = "=SUM(RC[-8]:RC[-1])"        
        ws_zaging.range(f"Q{row}").value = max360
        ws_zaging.range(f"S{row}").value = st_total
        dif = ws_zaging.range(f"T{row}").value
        if dif < -1 or dif > 1:
            color = ws_zaging.range(f"A{row}").api.EntireRow.Interior.Color
            if color != 0xDCE4FA:
                ws_zaging.range(f"A{row}").api.EntireRow.Interior.Color = 0xB4DFB4 
            maxstan = max360 - st_total
            if maxstan == 0 and ws_zaging.range(f"O{row}").value != 0:
                ws_zaging.range(f"O{row}").value = 0
                ws_zaging.range(f"O{row}").color = (255, 255, 0)
    # Final formatting            
    ws_zaging.range("P:P").delete()
    ws_zaging.range("A:T").autofit()
    ws_zaging.range("C:I").columns.hidden = True
    ws_zaging.range("S:S").number_format = "#0"
    # Shows completion message
    show_info("Fin","Fichero preparado para el Zaging")
    wb_standar.close()
    wb_zaging.save()
    wb_zaging.close()
    
def zaging_2():
    """
    Debt Aging step 2:    
    Automates updates Zaging file with SGL entries
    
    Workflow:
    - Prompts user to open Zaging and SGL Excel files
    - Iterates through SGL data:
      - Updates matching client info
      - Highlights important differences with color
    - Shows completion message and saves/cleans up workbook
    
    Parameters:
    - None (wrapped workflow with internal prompts and constants)
   
    Returns:
    - None: process concludes with saved Excel file
    """
    # Prompt for input files
    path_zaging = ask_open_file("Abre el fichero del Zaging")
    if not path_zaging:
        show_warning("Error","No se ha seleccionado el archivo. Se cancela el proceso")
        return
    path_sgl = ask_open_file("Abre el fichero de Partidas CME")
    if not path_sgl:
        show_warning("Error","No se ha seleccionado el archivo. Se cancela el proceso")
        return
    # Open workbooks
    wb_zaging = check_wb_open(path_zaging)
    wb_sgl = check_wb_open(path_sgl)
    ws_zaging = wb_zaging.sheets[0]
    ws_sgl = wb_sgl.sheets[0]
    # Read sgl data and build dictionary
    last_z_row=ws_zaging.range("J" + str(ws_zaging.cells.last_cell.row)).end('up').row
    last_sgl_row=ws_sgl.range("J" + str(ws_sgl.cells.last_cell.row)).end('up').row
    ws_sgl.range(f"{last_sgl_row}:{last_sgl_row}").delete()
    dic_clients_sgl = {}
    for i in range(2, last_sgl_row):
        client = ws_sgl.cells(i, 7).value
        amount = round(ws_sgl.cells(i, 10).value,2)
        if client in dic_clients_sgl:
            dic_clients_sgl[client] += amount
        else:
            dic_clients_sgl[client] = amount
    # Update Zaging workbook
    for i in range(2, last_z_row + 1):
        client = ws_zaging.cells(i, 1).api.Text
        if client in dic_clients_sgl:
            amount = dic_clients_sgl[client]
            current_val = ws_zaging.cells(i, 10).value or 0.0
            ws_zaging.cells(i, 10).value = current_val + amount
            ws_zaging.cells(i, 10).color = (255, 255, 0)  # Yellow
    # Close sgl and save Zaging
    show_info("Fin", "Confirmings incluidos")    
    wb_sgl.close()
    wb_zaging.save()
    wb_zaging.close()

def zaging_3():
    """
    Debt Aging step 3:
    Generates Zaging report by processing four Excel files:
    
    Workflow Parameters:
    - Prompts the user to select files:
        - Zaging report file
        - Open items file (Partidas Abiertas)
        - Cleared items file (Partidas Compensadas)
        - Modifications file (Modificaciones)
    - Updates Zaging report with recalculated financial values
    - Highlights discrepancies and modifications
    - Displays success or warning messages based on validation
    - Saves the updated report

    Returns:
        None; update zaging file and save it.
    """

    # Ask user to select Zaging file
    file_zaging = ask_open_file("Abre el fichero del Zaging")
    if not file_zaging:
        show_warning("Error","No se ha seleccionado el archivo. Se cancela el proceso")
        return
    
    # Ask user to select Open Items file
    file_pa = ask_open_file("Abre el fichero de Partidas Abiertas")
    if not file_pa:
        show_warning("Error","No se ha seleccionado el archivo. Se cancela el proceso")
        return
    
    # Ask user to select Cleared Items file
    file_pc = ask_open_file("Abre el fichero de Partidas Compensadas")
    if not file_pc:
        show_warning("Error","No se ha seleccionado el archivo. Se cancela el proceso")
        return

    # Ask user to select Modifications file
    file_modi = ask_open_file("Abre el fichero de Modificaciones")
    if not file_modi:
        show_warning("Error","No se ha seleccionado el archivo. Se cancela el proceso")
        return

    # Open the selected Excel files as workbooks
    wb_zaging = check_wb_open(file_zaging)
    wb_pa = check_wb_open(file_pa)
    wb_pc = check_wb_open(file_pc)
    wb_modi = check_wb_open(file_modi)

    # Select the first worksheet from each workbook
    ws_zaging = wb_zaging.sheets[0]
    ws_pa = wb_pa.sheets[0]
    ws_pc = wb_pc.sheets[0]
    ws_modi = wb_modi.sheets[0]

    # Initialize data structures for clients and documents
    dic_clients_pa = {}
    dic_clients_pc = {}
    dic_dev = {}
    dic_modi = {}
    clients = get_unique_column_values(ws_pa, 7)

    for client in clients:
        dic_clients_pa[client] = [0] * 7  # Aging buckets
        dic_clients_pc[client] = [0] * 7
        dic_dev[client] = []

    # Get last row of each worksheet
    zaging_rows = ws_zaging.range("A" + str(ws_zaging.cells.last_cell.row)).end('up').row
    pa_rows = ws_pa.range("A" + str(ws_pa.cells.last_cell.row)).end('up').row
    pc_rows = ws_pc.range("A" + str(ws_pc.cells.last_cell.row)).end('up').row
    modi_rows = ws_modi.range("A" + str(ws_modi.cells.last_cell.row)).end('up').row

    # Map document IDs to modification rows
    for row in range(2, modi_rows + 1):
        doc = ws_modi.range(f"K{row}").api.Text
        dic_modi[doc] = row

    print("OK")

    # Process Open Items (Partidas Abiertas)
    for i in range(2, pa_rows + 1):
        n_doc = ws_pa.cells(i, 6).api.Text
        client = ws_pa.cells(i, 7).api.Text
        doc_date = ws_pa.cells(i, 1).value.date()
        amount = ws_pa.cells(i, 10).value
        fy = str(doc_date.year + 1) if doc_date.month > 9 else str(doc_date.year)
        n_doc = f"{n_doc}{fy}"

        if n_doc in dic_modi:
            row = dic_modi[n_doc]
            ws_pa.cells(i, 3).value = ws_modi.cells(row, 9).value.date()
            ws_pa.cells(i, 3).color = (255, 0, 0)

        doc_class = ws_pa.cells(i, 9).value
        ref = ws_pa.cells(i, 8).value

        if doc_class in ["DA", "DB"] and ref:
            if ref not in dic_dev:
                dic_dev[client].append(ref)
        else:
            dif_day = ws_pa.cells(i, 15).value
            bucket_index = (
                0 if dif_day < 0 else
                1 if dif_day == 0 else
                min((dif_day // 90) + 2, 6)
            )
            dic_clients_pa[client][bucket_index] += amount

    print("OK")

    # Process Cleared Items (Partidas Compensadas)
    for i in range(pc_rows, 2, -1):
        client = ws_pc.cells(i, 7).api.Text
        n_doc_comp = ws_pc.cells(i, 12).api.Text

        if any(n_doc_comp in item for item in dic_dev[client]):
            n_doc = ws_pc.cells(i, 6).api.Text
            doc_date = ws_pc.cells(i, 1).value.date()
            amount = ws_pc.cells(i, 10).value
            fy = str(doc_date.year + 1) if doc_date.month > 9 else str(doc_date.year)
            n_doc = f"{n_doc}{fy}"

            if n_doc in dic_modi:
                row = dic_modi[n_doc]
                ws_pc.cells(i, 3).value = ws_modi.cells(row, 9).value.date()
                ws_pc.cells(i, 3).color = (255, 0, 0)

            doc_class = ws_pc.cells(i, 9).value
            ref = ws_pa.cells(i, 8).value

            if doc_class in ["DA", "DB"] and ref:
                if ref not in dic_dev:
                    dic_dev[client].append(ref)
            else:
                dif_day = ws_pc.cells(i, 15).value
                bucket_index = (
                    0 if dif_day < 0 else
                    1 if dif_day == 0 else
                    min((dif_day // 90) + 2, 6)
                )
                dic_clients_pc[client][bucket_index] += amount

    print("OK")

    # Reconcile Zaging Report
    for i in range(2, zaging_rows):
        client = ws_zaging.cells(i, 1).api.Text
        if client in clients:
            for j in range(7):
                pa_amount = dic_clients_pa[client][j]
                pc_amount = dic_clients_pc[client][j]
                temp_z_amount = round(pa_amount + pc_amount, 2)
                z_amount = round(ws_zaging.cells(i, j + 10).value or 0, 2)

                if temp_z_amount != z_amount:
                    ws_zaging.cells(i, j + 10).value = temp_z_amount
                    ws_zaging.cells(i, j + 10).color = (255, 255, 0)

            dif = abs(ws_zaging.cells(i, 19).value)
            if dif > 1:
                show_warning("Fallo", f"El zaging ha fallado en la fila {i}")
                ws_zaging.range(f"{i}:{i}").color = (255, 0, 0)

    # Final confirmation and save
    show_info("Fin", "Zaging generado")
    ws_zaging.save()

    file_zaging = ask_open_file("Abre el fichero del Zaging")
    if not file_zaging:
        show_warning("Error","No se ha seleccionado el archivo. Se cancela el proceso")
        return
    file_pa = ask_open_file("Abre el fichero de Partidas Abiertas")
    if not file_pa:
        show_warning("Error","No se ha seleccionado el archivo. Se cancela el proceso")
        return
    file_pc = ask_open_file("Abre el fichero de Partidas Compensadas")
    if not file_pc:
        show_warning("Error","No se ha seleccionado el archivo. Se cancela el proceso")
        return
    file_modi = ask_open_file("Abre el fichero de Modificaciones")
    if not file_modi:
        show_warning("Error","No se ha seleccionado el archivo. Se cancela el proceso")
        return
    # Open workbooks
    wb_zaging = check_wb_open(file_zaging)
    wb_pa = check_wb_open(file_pa)
    wb_pc = check_wb_open(file_pc)
    wb_modi = check_wb_open(file_modi)

    ws_zaging = wb_zaging.sheets[0]
    ws_pa = wb_pa.sheets[0]
    ws_pc = wb_pc.sheets[0]
    ws_modi = wb_modi.sheets[0]

    # Setup dictionaries
    dic_clients_pa = {}
    dic_clients_pc = {}
    dic_dev = {}
    dic_modi = {}
    clients = []
    clients = get_unique_column_values(ws_pa, 7)
    for client in clients:
        dic_clients_pa[client] = [0]*7
        dic_clients_pc[client] = [0]*7
        dic_dev[client] = []
        
    # Get row counts
    zaging_rows = ws_zaging.range("A" + str(ws_zaging.cells.last_cell.row)).end('up').row
    pa_rows = ws_pa.range("A" + str(ws_pa.cells.last_cell.row)).end('up').row
    pc_rows = ws_pc.range("A" + str(ws_pc.cells.last_cell.row)).end('up').row
    modi_rows = ws_modi.range("A" + str(ws_modi.cells.last_cell.row)).end('up').row
    for row in range(2,modi_rows + 1):
        doc = ws_modi.range(f"K{row}").api.Text
        dic_modi[doc] = row
    print("OK")
    for i in range(2, pa_rows + 1):
        n_doc = ws_pa.cells(i, 6).api.Text
        client = ws_pa.cells(i, 7).api.Text
        doc_date = ws_pa.cells(i, 1).value.date()
        amount = ws_pa.cells(i, 10).value
        if doc_date.month > 9:
            fy = str(doc_date.year + 1)
        else:
            fy = str(doc_date.year)
        n_doc = f"{n_doc}{fy}"
        if n_doc in dic_modi:
            row = dic_modi[n_doc]
            ws_pa.cells(i,3).value = ws_modi.cells(row,9).value.date() 
            ws_pa.cells(i,3).color =(255,0,0)
        doc_class = ws_pa.cells(i,9).value
        ref = ws_pa.cells(i,8).value
        if doc_class in ["DA","DB"] and ref:
            if not ref in dic_dev:
                dic_dev[client].append(ref)
        else:
            dif_day = ws_pa.cells(i,15).value
            if dif_day < 0:
                dic_clients_pa[client][0] += amount
            elif dif_day == 0:
                dic_clients_pa[client][1] += amount
            elif dif_day < 85:
                dic_clients_pa[client][2] += amount
            elif dif_day < 175:
                dic_clients_pa[client][3] += amount
            elif dif_day < 265:
                dic_clients_pa[client][4] += amount
            elif dif_day < 355:
                dic_clients_pa[client][5] += amount
            else:
                dic_clients_pa[client][6] += amount
    print("OK")
    for i in range(pc_rows,2,-1):
        client = ws_pc.cells(i, 7).api.Text
        n_doc_comp = ws_pc.cells(i, 12).api.Text
        if any(n_doc_comp in item for item in dic_dev[client]):
            n_doc = ws_pc.cells(i, 6).api.Text    
            doc_date = ws_pc.cells(i, 1).value.date()
            amount = ws_pc.cells(i, 10).value
            if doc_date.month > 9:
                fy = str(doc_date.year + 1)
            else:
                fy = str(doc_date.year)
            n_doc = f"{n_doc}{fy}"
            if n_doc in dic_modi:
                row = dic_modi[n_doc]
                ws_pc.cells(i,3).value = ws_modi.cells(row,9).value.date() 
                ws_pc.cells(i,3).color =(255,0,0)
            doc_class = ws_pc.cells(i,9).value
            ref = ws_pa.cells(i,8).value
            if doc_class in ["DA","DB"] and ref:
                if not ref in dic_dev:
                    dic_dev[client].append(ref)
            else:
                dif_day = ws_pc.cells(i,15).value
                if dif_day < 0:
                    dic_clients_pc[client][0] += amount
                elif dif_day == 0:
                    dic_clients_pc[client][1] += amount
                elif dif_day < 85:
                    dic_clients_pc[client][2] += amount
                elif dif_day < 175:
                    dic_clients_pc[client][3] += amount
                elif dif_day < 265:
                    dic_clients_pc[client][4] += amount
                elif dif_day < 355:
                    dic_clients_pc[client][5] += amount
                else:
                    dic_clients_pc[client][6] += amount
    print("OK")
    for i in range(2,zaging_rows):
        client = ws_zaging.cells(i,1).api.Text
        if client in clients:
            for j in range(0,7):
                pa_amount = dic_clients_pa[client][j]
                pc_amount = dic_clients_pc[client][j]
                temp_z_amount = round(pa_amount + pc_amount,2)
                z_amount = round(ws_zaging.cells(i, j+10).value or 0, 2)
                if temp_z_amount != z_amount:
                    ws_zaging.cells(i,j+10).value = temp_z_amount
                    ws_zaging.cells(i,j+10).color=(255,255,0)
            dif = abs(ws_zaging.cells(i,19).value)
            if dif > 1:
                show_warning("Fallo", f"El zaging a fallado en el la fila {i}")
                ws_zaging.range(f"{i}:{i}").color = (255,0,0)
    # Wrap-up
    show_info("Fin","Zaging generado")
    ws_zaging.save()   


# ---------
# Debug
# ---------   
# Saveguard
if __name__ == "__main__":
    zaging_3()