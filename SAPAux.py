# -*- coding: utf-8 -*-
"""
@author: JesusMMA
"""

import win32com.client
import gc
import pythoncom
from PyQt5.QtWidgets import QMessageBox
from UserInputs import (ask_user_date,ask_user_string,
                        show_question,show_info,show_warning,dif_popup
                        )                 
import Load_SAP_info
from datetime import datetime

# Global collections to store extracted SAP field data
gColl, nColl, tColl, typeColl = [], [], [], []

# -----------------------------------
# SAP connection Manager
# -----------------------------------
class SAPSessionManager:
    """
    Purpose:
    Centralized manager for handling SAP GUI scripting sessions using the win32com.client interface.
    Prevents duplicate connections, ensures clean disconnection, and provides a reusable interface
    for accessing the active SAP session across multiple modules.

    Scope:
    - Establish and reuse a single SAP GUI scripting session
    - Optionally close the SAP GUI window
    - Release all COM objects to avoid memory leaks or lingering sessions
    - Provide a clean interface for other modules to interact with SAP

    Workflow:
    - Call SAPSessionManager.connect() to establish or reuse an existing session
    - Use the returned session object to interact with SAP GUI scripting
    - Call SAPSessionManager.disconnect(close_window=True/False) to clean up
    
    Example Usage:
        from SAPAux import SAPSessionManager
    
        session = SAPSessionManager.connect()
        if session:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n000"
            session.findById("wnd[0]").sendVKey(0)
    
        SAPSessionManager.disconnect(close_window=True)
    """
    SapGuiAuto = None
    application = None
    connection = None
    session = None

    @classmethod
    def connect(cls):
        if cls.session:
            print("[INFO] Reusing existing SAP session.")
            return cls.session

        try:
            cls.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            cls.application = cls.SapGuiAuto.GetScriptingEngine
            cls.connection = cls.application.Children(0)
            cls.session = cls.connection.Children(0)
            print("[INFO] SAP session established.")
            return cls.session
        except Exception as e:
            print(f"[ERROR] Could not connect to SAP: {e}")
            Load_SAP_info.ContinueProgram = False
            return None

    @classmethod
    def disconnect(cls, close_window=False):
        try:
            if cls.session and close_window:
                try:
                    cls.session.findById("wnd[0]").Close()
                except Exception as e:
                    print(f"[WARNING] Could not close SAP window: {e}")

            # Release all COM objects
            for attr in ['session', 'connection', 'application', 'SapGuiAuto']:
                setattr(cls, attr, None)

            gc.collect()
            pythoncom.CoUninitialize()
            print("[INFO] SAP session disconnected.")
        except Exception as e:
            print(f"[ERROR] Error during SAP disconnect: {e}")

# -----------------------------------
# Inicializate SAP connection
# -----------------------------------

def _get_all(obj):
    """
    Recursively scans SAP GUI element tree to collect metadata for elements 
    matching the field name 'RFOPS_DK-XBLNR'.

    Workflow:
    - Traverses all children of the given SAP GUI object recursively
    - Checks for elements with name 'RFOPS_DK-XBLNR'
    - Appends matching elements' ID, name, text, and type to global collections

    Parameters:
    - obj: Root SAP GUI object to begin traversal

    Returns:
    - None: data is collected in global lists: gColl, nColl, tColl, typeColl
    """
    try:
        for i in range(obj.Children.Count):
            try:
                child = obj.Children.Item(i)
                _get_all(child)
                if child.Name == "RFOPS_DK-XBLNR":
                    gColl.append(child.Id)
                    nColl.append(child.Name)
                    tColl.append(child.Text)
                    typeColl.append(child.Type)
            except Exception as child_exception:
                print(f"[WARNING] Error processing child at index {i}: {child_exception}")
    except Exception as e:
        print(f"[ERROR] Failed to traverse SAP GUI tree: {e}")
        pass

def chk_window():
    """
    Returns the title of the current active SAP window.

    Parameters:
    - None (uses active session and SAP GUI commands internally)

    Returns:
    - str: Title text of the current SAP window
    - None: SAP element not found
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        return session.findById("wnd[0]").Text
    except Exception as e:
        print(f"[ERROR] Failed to retrieve SAP window title: {e}")
        Load_SAP_info.ContinueProgram = False
        return None

def chk_status_bar():
    """
    Reads the SAP status bar message and returns it.
    Displays warning if the message is classified as an error.

    Parameters:
    - None (uses active session and SAP GUI commands internally)

    Returns:
    - str: Text content of the status bar, or empty string if unavailable
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        msg = session.findById("wnd[0]/sbar").Text
        msg_type = session.findById("wnd[0]/sbar").MessageType
        if msg_type =="E":
            show_warning("Error", msg)
        return msg
    except Exception:
        return ""

def run_background_job():
    """
    SAP Background Job Launcher:  
    Automates the execution of a SAP report as a background job through GUI interaction.
    
    Workflow:
    - Ensures SAP GUI session is active
    - Navigates to report execution menu
    - Clears printer settings to avoid dialog conflicts
    - Triggers immediate execution ('SOFORT_PUSH')
    - Confirms and exits dialog windows
    
    Parameters:
    - None (uses active session and SAP GUI commands internally)
    
    Returns:
    - None: background job is launched and runs independently
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").SetFocus()
    session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").Key = ""
    session.findById("wnd[1]/tbar[0]/btn[13]").press()
    session.findById("wnd[1]/usr/btnSOFORT_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    
    
def call_transaction(txn_code:str):
    """
    Executes a given SAP transaction code from the Easy Access screen.
    Automatically resets to main menu if necessary.

    Parameters:
    - txn_code (str): SAP transaction code to execute (e.g. 'FB03', 'F-04')

    Returns:
    - None: launches specified transaction
    """
    session = SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
        
    try:
        current_title = chk_window()
        if current_title is None:
            raise RuntimeError("Unable to determine current SAP window.")
            if Load_SAP_info.ContinueProgram == False: return

        if "SAP Easy Access" not in current_title:
            back_to_main()
            if Load_SAP_info.ContinueProgram == False: return

        session.findById("wnd[0]/tbar[0]/okcd").Text = txn_code
        session.findById("wnd[0]").sendVKey(0)

    except Exception as e:
        print(f"[ERROR] Failed to execute transaction '{txn_code}': {e}")
        Load_SAP_info.ContinueProgram = False

def back_to_main():
    """
    Navigates back to SAP’s main menu screen from the current transaction.

    Parameters:
    - None (uses active session and SAP GUI commands internally)

    Returns:
    - None: clears current transaction with '/n00' and resets view
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n00"
        session.findById("wnd[0]").sendVKey(0)
    except Exception as e:
        print(f"[ERROR] Failed to return to SAP main menu: {e}")
        Load_SAP_info.ContinueProgram = False

def items_found_sap():
    """
    Extracts all open item identifiers from the 'Procesar partidas abiertas' SAP screen.
    Searches for entries labeled with 'RFOPS_DK-XBLNR' by scrolling through the item list 
    and compiles them into an ordered dictionary.

    Workflow:
    - Validates that the SAP session is in the correct view
    - Retrieves total number of expected open items
    - Recursively calls `_get_all()` to scrape current screen contents
    - Filters elements with name 'RFOPS_DK-XBLNR' and stores values
    - Scrolls through item pages using PAGE DOWN until all are collected
    - Navigates to summary screen and clears intermediate collections

    Parameters:
    - None (uses active session and SAP GUI commands internally)

    Returns:
    - Dict: Dictionary with open item identifiers as both keys and values
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        # Wait until correct window title appears
        attempts = 0
        while "Procesar partidas abiertas" not in chk_window():
            if Load_SAP_info.ContinueProgram == False: return
            if attempts > 5:
                raise TimeoutError("No se pudo acceder a la pantalla 'Procesar partidas abiertas'.")
                show_warning("Error", "No se pudo acceder a la pantalla 'Procesar partidas abiertas'.")
                Load_SAP_info.ContinueProgram = False
            session.findById("wnd[0]/tbar[1]/btn[16]").press()
            attempts += 1

        # Fetch total expected items
        try:
            total_items = int(session.findById(
                "wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-ANZPO"
            ).Text)
        except Exception as e:
            Load_SAP_info.ContinueProgram = False
            raise ValueError(f"No se pudo obtener el número total de partidas abiertas: {e}")
            show_warning("Error",f"No se pudo obtener el número total de partidas abiertas: {e}")

        collected = {}
        while len(collected) < total_items:
            _get_all()
            for name, value in zip(nColl, tColl):
                if name == "RFOPS_DK-XBLNR":
                    collected[value] = value
            session.findById("wnd[0]").sendVKey(82)  # PAGE DOWN

        session.findById("wnd[0]/tbar[1]/btn[14]").press()  # Navigate to summary screen
        clear_collections()
        return collected

    except Exception as e:
        show_warning("Error", f"[ERROR] Error during item extraction in 'items_found_sap': {e}")
        print(f"[ERROR] Error during item extraction in 'items_found_sap': {e}")
        Load_SAP_info.ContinueProgram = False
        return {}

def clear_collections():
    """
    Clears global scraping collections used by item retrieval functions.
    Ensures that previous session data does not pollute subsequent extractions.

    Globals:
    - gColl, nColl, tColl, typeColl: Lists containing GUI element metadata

    Returns:
    - None: performs in-place cleanup of global variables
    """
    try:
        global gColl, nColl, tColl, typeColl
        gColl.clear()
        nColl.clear()
        tColl.clear()
        typeColl.clear()
    except NameError as e:
        print(f"[WARNING] One or more global collections are not defined: {e}")
    except Exception as e:
        print(f"[ERROR] Failed to clear collections: {e}")
    
def batch_input(batch_template_path:str):
    """
    Executes a batch input transaction in SAP to upload the specified template file.
    Handles both PA-confirmed and PA-less workflows with user interaction.

    Workflow:
    - Validates that the SAP session is on the 'Easy Access' screen
    - Initiates custom transaction 'Z2S_K0021' for batch template loading
    - Sets input parameters and confirms template execution
    - Handles user decision if no PAs (payment agreements) were selected
    - If confirmed, bypasses PA selection and continues
    - Otherwise, selects all items and posts the batch entry

    Parameters:
    - batch_template_path (str): Full path to the template file to load

    Returns:
    - None: interacts directly with SAP and updates batch posting status
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        if "SAP Easy Access" not in chk_window():
            if Load_SAP_info.ContinueProgram == False: return
            back_to_main()

        call_transaction( "Z2S_K0021")  # Batch_input Transaction
        if Load_SAP_info.ContinueProgram == False: return

        session.findById("wnd[0]/usr/radP_CALLT").Select()
        session.findById("wnd[0]/usr/ctxtP_FILE").Text = batch_template_path
        session.findById("wnd[0]/tbar[1]/btn[8]").press()  # Run

        # Handle confirmation popup
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except Exception as popup_exception:
            print(f"[INFO] No confirmation popup appeared: {popup_exception}")

        sbar = chk_status_bar()
        if sbar == "Por favor, seleccione primero las partidas.":
            try:
                from PyQt5.QtWidgets import QMessageBox
                msg = QMessageBox()
                msg.setWindowTitle("Confirmación")
                msg.setText("No se han seleccionado PAs.\n¿Es un pago sin PA?")
                msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                result = msg.exec_()

                if result == QMessageBox.No:
                    QMessageBox.information(None, "Cancelación", "Se cancela el proceso.")
                    back_to_main()
                    if Load_SAP_info.ContinueProgram == False: return
                    Load_SAP_info.ContinueProgram = False
                    return
                session.findById("wnd[0]").sendVKey(12)  # Cancel PA selection

            except Exception as dialog_exception:
                print(f"[ERROR] Error en la interacción con el usuario: {dialog_exception}")
                Load_SAP_info.ContinueProgram = False
                return

        else:
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press()
            session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press()
            session.findById("wnd[0]/tbar[1]/btn[14]").press()

    except Exception as e:
        print(f"[ERROR] Error during batch input execution: {e}")
        Load_SAP_info.ContinueProgram = False

def sap_data():
    """
    Extracts key SAP fields from the 'Procesar partidas abiertas' view.
    Returns total item count, net item value, difference, and overall amount in a structured dictionary.

    Workflow:
    - Validates that SAP is in the correct open item view
    - Selects and expands all entries for visibility
    - Retrieves and converts numeric field data from GUI
    - Handles any parsing errors gracefully and defaults to 0.0
    - Navigates to summary view after extraction

    Parameters:
    - None (uses active session and SAP GUI commands internally)

    Returns:
    - Dict: Dictionary with:
        - 'total_items_loaded': Number of open items selected
        - 'items_amount': Sum of selected item values
        - 'dif_amount': Difference amount in the accounting entry
        - 'total_amount': Total accounting value
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        # Ensure we're on the correct screen
        attempts = 0
        while "Procesar partidas abiertas" not in chk_window():
            if Load_SAP_info.ContinueProgram == False: return
            if attempts > 5:
                raise TimeoutError("No se pudo acceder a la pantalla 'Procesar partidas abiertas'.")
            session.findById("wnd[0]/tbar[1]/btn[16]").press()
            attempts += 1

        # Select and expand all items
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press()
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press()

        # Define field mappings
        fields = {
            "total_items_loaded": "txtRF05A-ANZPO",
            "items_amount": "txtRF05A-NETTO",
            "dif_amount": "txtRF05A-DIFFB",
            "total_amount": "txtRF05A-BETRG",
        }

        data = {}
        for key, sap_id in fields.items():
            try:
                raw = session.findById(f"wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/{sap_id}").Text
                data[key] = float(raw.replace(".", "").replace(",", "."))
            except Exception as field_error:
                print(f"[WARNING] Could not parse field '{key}' ({sap_id}): {field_error}")
                data[key] = 0.0

        # Navigate to summary screen
        session.findById("wnd[0]/tbar[1]/btn[14]").press()
        return data

    except Exception as e:
        print(f"[ERROR] Failed to extract SAP data: {e}")
        Load_SAP_info.ContinueProgram = False
        return {
            "total_items_loaded": 0,
            "items_amount": 0.0,
            "dif_amount": 0.0,
            "total_amount": 0.0
        }

def new_entry(Posting_key:str, account:str, SGL_Ind:str="", doc_date:str=""):
    """
    Initiates a new SAP accounting entry using the specified posting key and account.
    Handles document header configuration and validates required fields before proceeding.

    Workflow:
    - Maps SAP GUI field IDs for header and entry-level inputs
    - If in compensation mode, prompts user for document date and fills header metadata
    - Validates posting key and checks if SGL indicator is required
    - Fills posting key, account, and SGL indicator into SAP entry fields
    - Sends confirmation to proceed and handles adjustment prompts from SAP

    Posting Key Logic:
    - '09': Customer Debit (requires SGL indicator)
    - '19': Customer Credit (requires SGL indicator)
    - '06': Customer Debit
    - '16': Customer Credit
    - '40': G/L Debit
    - '50': G/L Credit
    - '26': Vendor Debit
    - '36': Vendor Credit

    Parameters:
    - Posting_key (str): Code determining the transaction type (e.g. '40' for G/L Debit)
    - account (str): SAP account number for the posting
    - SGL_Ind (str, optional): Special G/L indicator, required for keys '09' and '19'
    - doc_date (str, optional): Accounting document date; prompted if missing in compensation mode

    Returns:
    - None: entry fields are set directly in SAP and validated via status bar response
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        Load_SAP_info.ContinueProgram = True

        fields = {
            "Posting_key": "wnd[0]/usr/ctxtRF05A-NEWBS",
            "account": "wnd[0]/usr/ctxtRF05A-NEWKO",
            "SGL_Ind": "wnd[0]/usr/ctxtRF05A-NEWUM",
            "doc_date": "wnd[0]/usr/ctxtBKPF-BLDAT",
            "acct_date": "wnd[0]/usr/ctxtBKPF-BUDAT",
            "soc_numb": "wnd[0]/usr/ctxtBKPF-BUKRS",
            "doc_type": "wnd[0]/usr/ctxtBKPF-BLART"
        }

        company_code = Load_SAP_info.config.get("company_code", "")
        if not company_code:
            raise ValueError("Company code not configured in Load_SAP_info.")

        # Header setup if in compensation mode
        if "Liquidar compensación: Datos cabecera" in chk_window():
            if Load_SAP_info.ContinueProgram == False: return
            session.findById("wnd[0]/usr/sub:SAPMF05A:0122/radRF05A-XPOS1[3,0]").Select()
            if not doc_date:
                doc_date = ask_user_date("Introduce la fecha contable")
            session.findById(fields["doc_date"]).Text = doc_date
            session.findById(fields["acct_date"]).Text = doc_date
            session.findById(fields["soc_numb"]).Text = company_code
            session.findById(fields["doc_type"]).Text = "SA"

        # Validate posting key
        valid_Posting_keys = {"09", "19", "06", "16", "40", "50", "26", "36"}
        """
        09 = Customer Debit, must add SGL indicator
        19 = Customer Credit, must add SGL indicator
        06 = Customer Debit
        16 = Customer Credit
        40 = G/L Debit
        50 = G/L Credit
        26 = Vendor Debit
        36 = Vendor Credit
        """
        if Posting_key not in valid_Posting_keys:
            show_warning("Error", "Esa clave no está incluida en el programa.")
            back_to_main()
            if Load_SAP_info.ContinueProgram == False: return
            Load_SAP_info.ContinueProgram = False
            return

        # Validate SGL indicator if required
        if Posting_key in {"09", "19"} and not SGL_Ind:
            QMessageBox.warning(None, "Advertencia", "Clave CME no introducida.")
            back_to_main()
            if Load_SAP_info.ContinueProgram == False: return
            Load_SAP_info.ContinueProgram = False
            return

        # Fill entry fields
        session.findById(fields["Posting_key"]).Text = Posting_key
        session.findById(fields["account"]).Text = account
        session.findById(fields["SGL_Ind"]).Text = SGL_Ind

        session.findById("wnd[0]").sendVKey(0)
        sbar = chk_status_bar()
        if "se adapta" in sbar:
            session.findById("wnd[0]").sendVKey(0)

    except Exception as e:
        print(f"[ERROR] Error during new SAP entry creation: {e}")
        Load_SAP_info.ContinueProgram = False

def new_entry_add_data(amount: float,due_date: str,commentary: str = "",
    payment_method: str = "",assignment: str = "",cost_center: str = ""
):
    """
    Completes an SAP accounting entry line with required and optional data.
    Automatically fills key fields such as amount, due date, commentary, business area,
    assignment, payment method, and cost center, guiding the user through input prompts if needed.
    
    Workflow:
    - Retrieves business area configuration from global settings
    - Formats default assignment from due date if missing
    - Identifies target SAP field IDs for all relevant input fields
    - Inputs amount and due date into appropriate SAP fields
    - Prompts user for commentary if missing, ensuring mandatory description
    - Validates and requests payment method if necessary, restricted to accepted values (2, 3, R, T)
    - Fills all applicable business area and cost center fields
    - Sends confirmation keystrokes and navigates to summary view
    - Checks status bar for validation or errors
    
    Parameters:
    - amount (float): Transaction amount for the accounting entry
    - due_date (str): SAP-formatted due date (e.g. 'DD.MM.YYYY')
    - commentary (str, optional): Description or label for the accounting line
    - payment_method (str, optional): Payment method code used in SAP (e.g. '2', '3', 'R', 'T')
    - assignment (str, optional): Assignment reference value; defaults to YYYYMMDD from due date if not provided
    - cost_center (str, optional): Optional cost center value used for expense categorization
    
    Returns:
    - None: entry is populated directly into SAP and navigated forward for confirmation
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        business_area = Load_SAP_info.config.get("business_area", "")
        if not business_area:
            raise ValueError("Business area not configured in Load_SAP_info.")

        # Generate default assignment from due date
        if not assignment:
            assignment = datetime.strptime(due_date, "%d.%m.%Y").strftime("%d/%m/%Y")
 
        # SAP Field IDs
        fields = {
            "amount": "wnd[0]/usr/txtBSEG-WRBTR",
            "due_date": "wnd[0]/usr/ctxtBSEG-ZFBDT",
            "payment_method": "wnd[0]/usr/ctxtBSEG-ZLSCH",
            "assignment": "wnd[0]/usr/txtBSEG-ZUONR",
            "commentary": "wnd[0]/usr/ctxtBSEG-SGTXT",
            "business_area": [
                "wnd[0]/usr/ctxtBSEG-GSBER",
                "wnd[0]/usr/subBLOCK:SAPLKACB:1007/ctxtCOBL-GSBER",
                "wnd[0]/usr/subBLOCK:SAPLKACB:1010/ctxtCOBL-GSBER",
                "wnd[1]/usr/ctxtCOBL-GSBER"
            ],
            "cost_center": [
                "wnd[0]/usr/subBLOCK:SAPLKACB:1010/ctxtCOBL-KOSTL",
                "wnd[0]/usr/subBLOCK:SAPLKACB:1007/ctxtCOBL-KOSTL",
                "wnd[1]/usr/ctxtCOBL-KOSTL"
            ]
        }
        # Fill core fields
        try:
            session.findById(fields["amount"]).Text = str(amount).replace(".", ",")
        except Exception as e:
            print(f"[WARNING] Could not set amount: {e}")

        try:
            session.findById(fields["due_date"]).Text = due_date
        except Exception as e:
            print(f"[WARNING] Could not set due date: {e}")

        session.findById(fields["assignment"]).Text = assignment

        # Ensure commentary is provided
        while not commentary:
            commentary = ask_user_string("Comentario para el apunte")
        session.findById(fields["commentary"]).Text = commentary

        # Handle payment method input and validation
        if not payment_method:
            answer = show_question("Confirmación de Vía de Pago", "¿El apunte tiene vía de pago?", QMessageBox.Yes | QMessageBox.No)
            if answer == QMessageBox.Yes:
                while not payment_method:
                    payment_method = ask_user_string("Introduce la vía de pago (2, 3, R, T)")
                    if payment_method not in {"2", "3", "R", "T"}:
                        show_warning("No válida", "Por favor introduce una vía de pago válida: 2, 3, R o T")
                        payment_method = ""
        try:
            session.findById(fields["payment_method"]).Text = payment_method
        except Exception as e:
            print(f"[WARNING] Could not set payment method: {e}")

        # Fill business area fields
        for div_field in fields["business_area"]:
            try:
                session.findById(div_field).Text = business_area
            except Exception:
                continue

        # Fill cost center if provided
        if cost_center:
            for cost_field in fields["cost_center"]:
                try:
                    session.findById(cost_field).Text = cost_center
                except Exception:
                    continue

        # Confirm and proceed
        session.findById("wnd[0]").sendVKey(0)
        session.ActiveWindow.sendVKey(0)

        # Check status bar
        chk_status_bar()

        # Navigate to summary
        session.findById("wnd[0]/tbar[1]/btn[14]").press()
        session.ActiveWindow.sendVKey(0)

    except Exception as e:
        print(f"[ERROR] Failed to complete SAP entry data: {e}")
        Load_SAP_info.ContinueProgram = False
       

def enter_position(position: str):
    """
    Navigates to and opens a specific accounting position within an active SAP document.
    Used during simulation to manually access and modify entry lines.

    Workflow:
    - Activates position field in SAP document view
    - Opens selection dialog for detailed entry by position number
    - Inputs the specified position and confirms selection

    Parameters:
    - position (str): Entry line position number to be accessed (e.g. '001', '002')

    Returns:
    - None: modifies SAP session state to focus on specified accounting position
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        session.findById("wnd[0]/usr/txtRF05A-ANZAZ").SetFocus()
        session.findById("wnd[0]").sendVKey(2)  # Open position selection dialog
        session.findById("wnd[1]/usr/txt*BSEG-BUZEI").Text = position
        session.findById("wnd[1]/tbar[0]/btn[13]").press()  # Confirm selection

    except Exception as e:
        print(f"[ERROR] Failed to enter SAP position '{position}': {e}")
        Load_SAP_info.ContinueProgram = False

def enter_ajd(amount: float, assignment: str, commentary: str, due_date: str):
    """
    Posts an SAP adjustment entry for AJD-related bank expenses and taxes.
    Selects G/L debit or credit depending on the amount and includes cost center allocation.

    Workflow:
    - Retrieves expense account and cost center from configuration
    - Determines if entry should be debit ('40') or credit ('50') based on amount sign
    - Converts negative amount to positive before posting (SAP expects absolute value)
    - Executes entry with associated due date, assignment, commentary, and cost center

    Parameters:
    - amount (float): Tax amount to be posted as an expense adjustment
    - assignment (str): SAP-formatted assignment reference (YYYYMMDD or similar)
    - commentary (str): Text label for the accounting line
    - due_date (str): Due date for the expense posting (DD.MM.YYYY)

    Returns:
    - None: updates SAP transaction directly with AJD entry details
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        config = Load_SAP_info.config
        expense_cost_center = config.get("expense_cost_center")
        expense_account = config.get("expense_account")

        if not expense_cost_center or not expense_account:
            raise ValueError("Missing 'expense_cost_center' or 'expense_account' in configuration.")

        # Determine posting key and normalize amount
        if amount > 0:
            posting_key = "40"  # G/L Debit
        else:
            posting_key = "50"  # G/L Credit
            amount = abs(amount)

        # Create entry and fill data
        new_entry(posting_key, expense_account)
        if Load_SAP_info.ContinueProgram == False: return
        new_entry_add_data(            
            amount=amount,
            due_date=due_date,
            commentary=commentary,
            payment_method="-1",
            assignment=assignment,
            cost_center=expense_cost_center
        )
        if Load_SAP_info.ContinueProgram == False: return

    except Exception as e:
        print(f"[ERROR] Failed to post AJD adjustment entry: {e}")
        Load_SAP_info.ContinueProgram = False
        

def search_items(category: str, position: int, search_data: str = "",
                 company_code: str = "", account: str = "", additional_data: str = ""):
    """
    Searches open SAP items in 'Visualizar Resumen' view using dynamic criteria.
    Handles selection modes for invoice references, amounts, due dates, and more. 
    Filters and applies results based on user input.
    
    Workflow:
    - Validates that SAP session is in 'Visualizar Resumen' view
    - Sets account category and optional company/account filters
    - Selects open items based on positional criteria:
        - Pos 0: Select all items
        - Pos 1: Filter by amount range
        - Pos 5: Filter by reference
        - Pos 16: Filter by net due date
    - Applies search data and additional filters where applicable
    - Handles user cancellation and missing views
    - Selects all matching open items for processing

    Supported Positions:
    -  0: Todas partidas abiertas 
    -  1: Importe 
    -  2: Nº documento
    -  3: Fecha contabilización
    -  4: Referencia a factura
    -  5: Referencia
    -  6: Clase de documento
    -  7: Indicador impuestos
    -  8: Solicitud acept. L/C
    -  9: Cta. subsidiaria
    - 10: Moneda
    - 11: Clave contabilización
    - 12: Fecha de documento
    - 13: Asignación
    - 14: Factura
    - 15: Posición
    - 16: Vencimiento neto

    Parameters:
    - category (str): SAP account category ('D' for customer, 'K' for vendor)
    - position (int): Selection mode for filtering open items
    - search_data (str, optional): Primary search value (e.g. amount, reference)
    - company_code (str, optional): SAP company code to restrict search
    - account (str, optional): SAP account number
    - additional_data (str, optional): Secondary search value (e.g. end date for range)

    Returns:
    - None: modifies SAP session view to select open items based on criteria
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        Load_SAP_info.ContinueProgram = True
    
        # Ensure we're in the correct view
        attempts = 0
        while "Visualizar Resumen" not in chk_window():
            if Load_SAP_info.ContinueProgram == False: return
            if attempts > 3:
                raise TimeoutError("No se pudo acceder a la vista 'Visualizar Resumen'.")
            response = show_question(
                "Confirmación",
                "No estás en la ventana apropiada.\nVe a Visualizar Resumen y presiona OK.",
                QMessageBox.Ok | QMessageBox.Cancel
            )
            if response == QMessageBox.Cancel:
                back_to_main()
                if Load_SAP_info.ContinueProgram == False: return
                Load_SAP_info.ContinueProgram = False
                return
            attempts += 1
    
        session.findById("wnd[0]/tbar[1]/btn[6]").press()
    
        # Set base filters
        session.findById("wnd[0]/usr/ctxtRF05A-AGKOA").Text = category
        if company_code:
            session.findById("wnd[0]/usr/ctxtRF05A-AGBUK").Text = company_code
        if account:
            session.findById("wnd[0]/usr/ctxtRF05A-AGKON").Text = account
    
        # Field mapping for supported positions
        field_map = {
            1: ("0730", "txtRF05A-VONWT[0,0]", "txtRF05A-BISWT[0,21]"),
            16: ("0732", "ctxtRF05A-VONDT[0,0]", "ctxtRF05A-BISDT[0,20]"),
            5: ("0731", "txtRF05A-SEL01[0,0]", "txtRF05A-SEL02[0,31]"),
        }
    
        if position == 0:
            session.findById("wnd[0]").sendVKey(0)  # Select all open items
        elif position in field_map:
            tab_code, field1, field2 = field_map[position]
            session.findById(f"wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[{position},0]").Select()
            session.findById("wnd[0]").sendVKey(0)
    
            if search_data:
                session.findById(f"wnd[0]/usr/sub:SAPMF05A:{tab_code}/{field1}").Text = search_data
                session.findById("wnd[0]").sendVKey(0)
            if additional_data:
                session.findById(f"wnd[0]/usr/sub:SAPMF05A:{tab_code}/{field2}").Text = additional_data
                session.findById("wnd[0]").sendVKey(0)
    
            session.findById("wnd[0]/tbar[1]/btn[16]").press()
        else:
            show_info("Cancelar", "No se ha contemplado esa selección.")
            back_to_main()
            if Load_SAP_info.ContinueProgram == False: return
            Load_SAP_info.ContinueProgram = False
            return
    
        # Check for no results
        msg = chk_status_bar()
        if "No se encontró" in msg:
            show_info("Aviso", msg)
    
        # Validate transition to open items screen
        if "Procesar partidas abiertas" not in chk_window():
            if Load_SAP_info.ContinueProgram == False: return
            show_warning("Error", "No se han encontrado Partidas Abiertas")
            back_to_main()
            if Load_SAP_info.ContinueProgram == False: return
            Load_SAP_info.ContinueProgram = False
            return
    
        # Select all matching items
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press()
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press()
    
    except Exception as e:
        print(f"[ERROR] Error during item search: {e}")
        Load_SAP_info.ContinueProgram = False
        
        
def get_entry_number() -> str:
    """
    Retrieves the document number of the currently active SAP accounting entry.
    Ensures the session is in the correct viewing mode before extracting the value.

    Workflow:
    - Checks if the SAP session is in 'Visualizar documento:Acceso' view
    - If not, resets the session to main screen and reopens document viewer (FB03)
    - Extracts document number from the input field

    Parameters:
    - None (uses active session and SAP GUI commands internally)


    Returns:
    - str: Document number of the active entry as displayed in SAP
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        if "Visualizar documento:Acceso" not in chk_window():
            if Load_SAP_info.ContinueProgram == False: return
            back_to_main()
            if Load_SAP_info.ContinueProgram == False: return
            call_transaction("FB03")
            if Load_SAP_info.ContinueProgram == False: return

        doc_number = session.findById("wnd[0]/usr/txtRF05L-BELNR").Text
        return doc_number

    except Exception as e:
        print(f"[ERROR] Failed to retrieve SAP document number: {e}")
        Load_SAP_info.ContinueProgram = False
        return ""

def save_entry(path: str):
    """
    Saves the PDF spool of the current SAP accounting entry to a specified file path.
    Navigates through SAP GUI to extract and store the spool document using entry metadata.

    Workflow:
    - Retrieves entry number from the active session
    - Clears status bar messages and resets SAP view
    - Navigates to spool print menu and disables print preview
    - Launches transaction SP01 to access spool
    - Filters spool by entry title to isolate desired output
    - Downloads PDF version of the spool to target directory with appropriate naming
    - Returns SAP to main screen

    Parameters:
    - path (str): Destination folder path where the PDF file will be saved

    Returns:
    - None: saves file directly to disk and resets SAP interface
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        entry_number = get_entry_number()
        if Load_SAP_info.ContinueProgram == False: return
        if not entry_number:
            raise ValueError("No se pudo obtener el número de documento.")

        # Clear status bar messages
        session.findById("wnd[0]").sendVKey(0)
        while chk_status_bar():
            session.findById("wnd[0]").sendVKey(0)

        # Open print menu and disable preview
        session.findById("wnd[0]/tbar[0]/btn[86]").press()
        session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").SetFocus()
        session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").Key = ""
        session.findById("wnd[1]/tbar[0]/btn[13]").press()
        session.findById("wnd[0]").sendVKey(0)

        # Navigate to SP01 and search for spool
        back_to_main()
        if Load_SAP_info.ContinueProgram == False: return
        call_transaction("SP01")
        if Load_SAP_info.ContinueProgram == False: return
        session.findById("wnd[0]").sendVKey(8)
        session.findById("wnd[0]/usr/lbl[3,3]").SetFocus()
        session.findById("wnd[0]").sendVKey(2)
        session.findById("wnd[0]/usr/txtTSP01_SP0R-RQTITLE").Text = entry_number
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Open spool and prepare for download
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
        session.findById("wnd[0]/usr/chk[1,3]").Selected = True
        session.findById("wnd[0]/usr/chk[1,3]").SetFocus()
        session.findById("wnd[0]/mbar/menu[0]/menu[2]/menu[2]").Select()

        # Set file path and name
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = f"{entry_number}.pdf"
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        back_to_main()
        if Load_SAP_info.ContinueProgram == False: return

    except Exception as e:
        print(f"[ERROR] Failed to save SAP entry PDF: {e}")
        Load_SAP_info.ContinueProgram = False

def call_variant(variant_name: str, variant_author: str = "", variant_modified_by: str = "",
                 variant_environment: str = "", variant_language: str = ""):
    """
    Loads and applies a SAP ALV (ABAP List Viewer) variant with optional filters.
    Enables customized layout or selection views during SAP reporting or data interaction.

    Workflow:
    - Opens the variant selection screen in SAP GUI
    - Fills in the desired variant name
    - Optionally filters by author, modifier, environment, or language
    - Executes selection and applies the variant

    Parameters:
    - variant_name (str): Name of the ALV variant to be applied
    - variant_author (str, optional): User ID of the variant creator
    - variant_modified_by (str, optional): User ID of the last person who modified it
    - variant_environment (str, optional): Environment associated with the variant (e.g. development, production)
    - variant_language (str, optional): Language code (e.g. 'EN' for English)

    Returns:
    - None: the variant is applied directly in the SAP session
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        # Open variant selection screen
        session.findById("wnd[0]/tbar[1]/btn[17]").press()
        # Fill in required and optional fields
        session.findById("wnd[1]/usr/txtV-LOW").Text = variant_name
        session.findById("wnd[1]/usr/ctxtENVIR-LOW").Text = variant_environment
        session.findById("wnd[1]/usr/txtENAME-LOW").Text = variant_author
        session.findById("wnd[1]/usr/txtAENAME-LOW").Text = variant_modified_by
        session.findById("wnd[1]/usr/txtMLANGU-LOW").Text = variant_language
        # Execute variant selection
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

    except Exception as e:
        print(f"[ERROR] Failed to apply variant '{variant_name}': {e}")
        Load_SAP_info.ContinueProgram = False

def to_account_dif(dif_sap:float, account:str, due_date:str, commentary:str="", due_date_assigment:str=""):
    """
    Handles SAP posting differences by allocating the amount directly to the client's account.
    Posts either a debit or credit entry depending on the sign of the difference.

    Workflow:
    - If difference is negative, posts debit to client account using transaction code '06'
    - If positive, posts credit to client account using code '16'
    - Applies commentary and due date assignment to the entry line
    - If no difference is present, alerts user that posting is not applicable

    Parameters:
    - dif_sap (float): Difference amount to be posted
    - account (str): SAP account number where the difference will be assigned
    - due_date (str): Due date for the posting entry
    - commentary (str, optional): Description for the accounting line
    - due_date_assigment (str, optional): SAP-formatted assignment date string

    Returns:
    - None: updates SAP entry based on correction strategy
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        if dif_sap > 0:
            posting_key = "16"  # Customer Credit
        elif dif_sap < 0:
            posting_key = "06"  # Customer Debit
            dif_sap = abs(dif_sap)
        else:
            show_info("Error", "Diferencia igual a cero: no se puede contabilizar.")
            return

        new_entry(posting_key, account)
        if Load_SAP_info.ContinueProgram == False: return
        new_entry_add_data(            
            amount=dif_sap,
            due_date=due_date,
            commentary=commentary,
            payment_method="",
            assignment=due_date_assigment
        )
        if Load_SAP_info.ContinueProgram == False: return

    except Exception as e:
        print(f"[ERROR] Failed to post account difference: {e}")
        Load_SAP_info.ContinueProgram = False

def round_dif(dif_sap: float, due_date:str, commentary:str="",due_date_assigment:str=""):
    """
    Rounds and compensates small SAP posting differences using a designated cost center.
    Adjusts either as debit or credit entry depending on the sign of the difference.

    Workflow:
    - Retrieves rounding account and cost center from configuration
    - Determines if difference is positive or negative
    - Posts the corresponding G/L entry using special transaction codes:
        - Code '50' for credit (when difference is negative)
        - Code '40' for debit (when difference is positive)
    - If no adjustment is necessary (i.e. difference is zero), informs the user and exits

    Parameters:
    - dif_sap (float): Difference amount to be rounded off
    - due_date (str): Due date for posting the entry
    - commentary (str, optional): Description to include with the adjustment entry
    - due_date_assigment (str, optional): SAP-formatted date for assignment field

    Returns:
    - None: entry is posted directly into SAP based on the specified difference
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        config = Load_SAP_info.config
        rounding_cost_center = config.get("rounding_cost_center")
        rounding_account = config.get("rounding_account")

        if not rounding_cost_center or not rounding_account:
            raise ValueError("Faltan 'rounding_cost_center' o 'rounding_account' en la configuración.")

        if dif_sap > 0:
            posting_key = "40"  # G/L Debit
        elif dif_sap < 0:
            posting_key = "50"  # G/L Credit
            dif_sap = abs(dif_sap)
        else:
            show_info("Error", "Diferencia igual a cero: no se puede redondear.")
            return

        new_entry(posting_key, rounding_account)
        if Load_SAP_info.ContinueProgram == False: return
        new_entry_add_data(
            amount=dif_sap,
            due_date=due_date,
            commentary=commentary,
            payment_method="-1",
            assignment=due_date_assigment,
            cost_center=rounding_cost_center
        )
        if Load_SAP_info.ContinueProgram == False: return

    except Exception as e:
        print(f"[ERROR] Error al contabilizar redondeo: {e}")
        Load_SAP_info.ContinueProgram = False

# Handler
def handle_dif(diff_sap:float, account:str, due_date:str, commentary:str="",due_date_assigment:str=""):
    """
    Handles SAP posting differences during simulation by prompting the user for correction strategy.
    Applies rounding adjustments or redirects difference to specified account, based on user selection.
    
    Workflow:
    - Displays popup with options for resolving the posting discrepancy
    - If user selects 'round_dif', applies correction via `round_dif()` handler
    - If user selects 'to_account', redirects difference to the designated account via `to_account_dif()`
    - Navigates to accounting summary view before each correction
    - If no valid selection is made, cancels the process and resets flow
    
    Parameters:
    - diff_sap (float): Difference amount to be resolved
    - account (str): Account number to redirect the difference to (if applicable)
    - due_date (str): Due date used in the corrective posting
    - commentary (str, optional): Commentary used in the posting line
    - due_date_assigment (str, optional): Assignment date format for SAP entry
    
    Returns:
    - None: modifies SAP transaction directly based on user choice
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        # Prompt user for correction strategy
        response = dif_popup(diff_sap)

        if response == "round_dif":
            session.findById("wnd[0]/tbar[1]/btn[14]").press()  # Proceed to accounting summary
            round_dif(diff_sap, due_date, commentary, due_date_assigment)
            if Load_SAP_info.ContinueProgram == False: return

        elif response == "to_account":
            session.findById("wnd[0]/tbar[1]/btn[14]").press()  # Proceed to accounting summary
            to_account_dif(diff_sap, account, due_date, commentary, due_date_assigment)
            if Load_SAP_info.ContinueProgram == False: return

        else:
            show_info("Error", "No se ha aplicado la diferencia.")
            back_to_main()
            if Load_SAP_info.ContinueProgram == False: return
            Load_SAP_info.ContinueProgram = False

    except Exception as e:
        print(f"[ERROR] Error al manejar la diferencia de contabilización: {e}")
        Load_SAP_info.ContinueProgram = False

def simulate(account:str,due_date:str) -> list[int] | None:
    """
    Executes SAP simulation via 'Visualizar Resumen' to validate accounting entries.
    Automatically detects and handles differences before re-simulation, then returns range of positions.
    
    ⚠️ If a large difference is detected, the function triggers correction via `handle_dif()` and reruns simulation.
    
    Workflow:
    - Ensures SAP window is in 'Visualizar Resumen' mode
    - Records initial accounting position from simulation input field
    - Initiates SAP simulation and checks status bar for discrepancies
    - If a difference is flagged:
        - Extracts and formats difference amount
        - Applies correction via `handle_dif()`
        - Re-runs simulation and captures updated position
    - Retrieves final accounting position after successful simulation
    
    Parameters:
    - account (str): SAP account number used for adjustment handling
    - due_date (str): Due date associated with the transaction entry
    
    Returns:
    - list[int] or None: A list with [initial_position, final_position] if simulation completes; None if canceled or error occurs
    """
    session=SAPSessionManager.session
    if session == None:
        SAPSessionManager.connect()
        session = SAPSessionManager.session
    try:
        Load_SAP_info.ContinueProgram = True

        # Ensure we're in 'Visualizar Resumen' view
        attempts = 0
        while "Visualizar Resumen" not in chk_window():
            if Load_SAP_info.ContinueProgram == False: return
            if attempts > 3:
                raise TimeoutError("No se pudo acceder a la vista 'Visualizar Resumen'.")
            session.findById("wnd[0]/tbar[1]/btn[14]").press()
            attempts += 1

        # Record initial position
        pos_ini = int(session.findById("wnd[0]/usr/txtRF05A-ANZAZ").Text)

        # Trigger simulation
        session.findById("wnd[0]/mbar/menu[0]/menu[3]").Select()
        status = chk_status_bar()

        # Handle large difference
        if status == "La diferencia es demasiado grande para una compensación":
            try:
                raw_diff = session.findById(
                    "wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-DIFFB"
                ).Text.replace(",", ".")
                if raw_diff.endswith("-"):
                    raw_diff = "-" + raw_diff[:-1]
                raw_diff = raw_diff.strip()
                raw_diff = raw_diff.replace(".", "")
                raw_diff = raw_diff.replace(",", ".")
                diff_sap = round(float(raw_diff), 2)
            except Exception as e:
                print(f"[ERROR] No se pudo leer el valor de diferencia: {e}")
                show_info("Error", "No se pudo leer el valor de diferencia.")
                Load_SAP_info.ContinueProgram = False
                back_to_main()
                return None

            # Apply correction and re-simulate
            handle_dif(diff_sap, account, due_date)
            if Load_SAP_info.ContinueProgram == False: return None

            pos_ini = int(session.findById("wnd[0]/usr/txtRF05A-ANZAZ").Text)
            session.findById("wnd[0]/mbar/menu[0]/menu[3]").Select()

        # Final position after simulation
        pos_end = int(session.findById("wnd[0]/usr/txtRF05A-ANZAZ").Text)
        return [pos_ini, pos_end]

    except Exception as e:
        print(f"[ERROR] Error durante la simulación: {e}")
        Load_SAP_info.ContinueProgram = False
        return None
# ---------
# Debug
# ---------   
# Saveguard
if __name__ == "__main__":
    print("Nice")
