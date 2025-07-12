# -*- coding: utf-8 -*-
"""
@author: JesusMMA
"""

from DiffUI import Ui_Form
import datetime
import sys
import Load_SAP_info
from PyQt5.QtWidgets import (
    QApplication, QMessageBox, QInputDialog, QFileDialog,
    QDialog
)
from PyQt5.QtCore import QLocale

# -----------------------------------------------
#  Centralized Dialog Helpers
# -----------------------------------------------
def show_info(title: str, message: str):
    """
    Displays an informational popup dialog using PyQt5.
    
    Parameters:
    - title (str): Title of the popup window
    - message (str): Message content to display
    
    Returns:
    - None: shows modal info dialog
    """
    QMessageBox.information(None, title, message)

def show_warning(title: str, message: str):
    """
    Displays a warning popup dialog using PyQt5.

    Parameters:
    - title (str): Title of the popup window
    - message (str): Warning message to display

    Returns:
    - None: shows modal warning dialog
    """
    QMessageBox.warning(None, title, message)

def show_question(title: str, message: str, buttons=QMessageBox.Yes | QMessageBox.No) -> int:
    """
    Displays a question dialog with customizable buttons.

    Parameters:
    - title (str): Window title
    - message (str): Question prompt
    - buttons (int, optional): Button set (e.g. Yes/No, Retry/Cancel)

    Returns:
    - int: User-selected button value
    """
    return QMessageBox.question(None, title, message, buttons)

# -----------------------------------------------
#  Retry Decorator
# -----------------------------------------------
def retry_input(func):
    """
    Decorator that re-prompts the user until valid input is provided or canceled.
    
    Parameters:
    - func: Function requiring validated input
    
    Returns:
    - Wrapper function that repeats until valid or canceled
    """
    def wrapper(*args, **kwargs):
        while True:
            result = func(*args, **kwargs)
            if result is not None:
                return result
            retry = show_question(
                "¿Reintentar?",
                "No se recibió una entrada válida.\n¿Desea intentarlo de nuevo?",
                QMessageBox.Retry | QMessageBox.Cancel
            )
            if retry == QMessageBox.Cancel:
                show_info("Cancelado", "Operación cancelada por el usuario.")
                Load_SAP_info.ContinueProgram = False
                return None
    return wrapper

# -----------------------------------------------
#  Utilities
# -----------------------------------------------
class DiffDialog(QDialog):
    """
    PyQt5 dialog for selecting SAP difference handling strategy (rounding or account assignment).
    
    Returns:
    - result: Set to 'round_dif' or 'to_account' based on user action
    """
    def __init__(self, diff_value):
        super().__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.ui.retranslateUi(self, diff_value)

        self.result:str |None = None
        self.ui.RoundBtn.clicked.connect(self.handle_round)
        self.ui.ToAccountBtn.clicked.connect(self.handle_to_account)

    def handle_round(self):
        self.result = "round_dif"
        self.accept()  # closes the dialog and sets result to Accepted

    def handle_to_account(self):
        self.result = "to_account"
        self.accept()


def dif_popup(dif):
    """
    Opens a modal dialog for SAP difference strategy selection.
    
    Parameters:
    - dif: Difference value displayed in dialog
    
    Returns:
    - str or None: Selected strategy ('round_dif', 'to_account') or None if canceled
    """
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)

    dialog = DiffDialog(dif)
    if dialog.exec_() == QDialog.Accepted:
        return dialog.result
    return None


def distinct_vals(ws, column_letter: str) -> list:
    """
    Returns a list of distinct non-empty values from a specified Excel column.

    Parameters:
    - ws: Excel worksheet object
    - column_letter (str): Column reference (e.g. 'D')

    Returns:
    - list: Unique, non-empty values from that column
    """
    try:
        # Expand the range starting from the header cell
        col_range = ws.range(f"{column_letter}1").end('down').value
        if not col_range or not isinstance(col_range, list):
            return []
        # Remove header and None values
        values = [v for v in col_range[1:] if v is not None]
        # Return unique values
        return list(set(values))
    except Exception as e:
        print(f"[ERROR] Failed to extract distinct values: {e}")
        Load_SAP_info.ContinueProgram = False
    return []

def ask_open_file(msg: str) -> str | None:
    """
    Prompts the user to select a file using QFileDialog.
    
    Parameters:
    - msg (str): Message shown in the dialog
    
    Returns:
    - str or None: Selected file path or None if canceled
    """
    try:
        while True:
            file_path, _ = QFileDialog.getOpenFileName(None, msg)
            if file_path:
                return file_path
            retry = show_question(
                "Confirmación",
                "No se ha seleccionado fichero.\n¿Desea continuar?",
                QMessageBox.Retry | QMessageBox.Cancel
            )
            if retry == QMessageBox.Cancel:
                show_info("Cancelado", "Proceso cancelado por el usuario.")
                Load_SAP_info.ContinueProgram = False
                return None
    except Exception as e:
        print(f"[ERROR] Error al seleccionar archivo: {e}")
        Load_SAP_info.ContinueProgram = False
        return None

def ask_open_files(msg: str) -> list[str] | None:
    """
    Prompts the user to select one or more files using QFileDialog.

    Parameters:
    - msg (str): Message shown in the dialog

    Returns:
    - list of str or None: Selected file paths or None if canceled
    """
    try:
        while True:
            file_paths, _ = QFileDialog.getOpenFileNames(None, msg)
            if file_paths:
                return file_paths
            retry = show_question(
                "Confirmación",
                "No se ha seleccionado ningún fichero.\n¿Desea continuar?",
                QMessageBox.Retry | QMessageBox.Cancel
            )
            if retry == QMessageBox.Cancel:
                show_info("Cancelado", "Proceso cancelado por el usuario.")
                Load_SAP_info.ContinueProgram = False
                return None
    except Exception as e:
        print(f"[ERROR] Error al seleccionar archivos: {e}")
        Load_SAP_info.ContinueProgram = False
        return None

@retry_input
def ask_user_date(prompt="Introduce la fecha (dd/mm/yyyy)") -> str | None:
    """
    Prompts user for a valid date string in dd/mm/yyyy format.

    Parameters:
    - prompt (str): Dialog message

    Returns:
    - str or None: Formatted date string (dd.mm.yyyy) or None
    """
    text, ok = QInputDialog.getText(None, "Introduce Fecha", prompt)
    if not ok:
        return None
    try:
        dt = datetime.datetime.strptime(text, "%d/%m/%Y")
        return dt.strftime("%d.%m.%Y")
    except ValueError:
        show_warning("Fecha inválida", "Introduce una fecha válida en formato dd/mm/yyyy.")
        return None

@retry_input
def ask_user_number(msg: str) -> float | None:
    """
    Prompts user to enter a numeric value with precision and locale control.

    Parameters:
    - msg (str): Context for input (e.g. purpose of the value)

    Returns:
    - float or None: Rounded numeric input or None
    """
    dialog = QInputDialog()
    dialog.setInputMode(QInputDialog.DoubleInput)
    dialog.setLabelText(f"Introduce el importe de {msg}:")
    dialog.setWindowTitle("Introduce importe")
    dialog.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))  # Forces dot as decimal
    dialog.setDoubleDecimals(2)
    dialog.setDoubleRange(0.0, 9999999999.0)
    dialog.setDoubleValue(0.0)

    if dialog.exec_() == QInputDialog.Accepted:
        return round(dialog.doubleValue(), 2)
    return None

@retry_input
def ask_user_string(msg: str) -> str | None:
    """
    Prompts user for a string input, ensuring non-empty response.

    Parameters:
    - msg (str): Descriptor for input prompt

    Returns:
    - str or None: Cleaned user string input or None
    """
    text, ok = QInputDialog.getText(None, "Introduce comentario", f"Introduce el {msg}:")
    return text.strip() if ok and text.strip() else None

def save_confirmation() -> bool:
    """
    Asks the user for final confirmation before saving data to SAP.

    Parameters:
    - None (uses active session and SAP GUI commands internally)

    Returns:
    - bool: True if user confirms, False if canceled
    """
    from SAPAux import chk_window
    win_text=chk_window()
    if Load_SAP_info.ContinueProgram == False: return
    while "Visualizar Resumen" in win_text:
        reply = show_question(
            "Confirmación",
            "¿Conforme con los apuntes?\n¿Desea continuar y guardar en SAP?"
        )
        if reply == QMessageBox.No:
            show_info("Cancelado", "Proceso cancelado por el usuario.")
            Load_SAP_info.ContinueProgram = False
            return False
        win_text=chk_window()
        if Load_SAP_info.ContinueProgram == False: return
    return True
# ---------
# Debug
# ---------   
# Saveguard
if __name__ == "__main__":
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)

    # Optional: test a dialog or widget
    show_info("Módulo cargado", "Este módulo se ejecuta directamente.")
