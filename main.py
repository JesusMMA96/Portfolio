# -*- coding: utf-8 -*-
"""
@author: JesusMMA
"""
import sys
import locale
# Set locale to Spanish (Spain)
locale.setlocale(locale.LC_ALL, 'es_ES.UTF-8')

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QMessageBox
)
from MainUI import Ui_MainUI
from AutoZagingUI import Ui_ZagingReportUI
from BalanceReportUI import Ui_BalanceReport
from DailyPaymentsModule import bank_file,daily_payments
from PaymentsModule import payment
from ReportsModule import (large_format_retailers_file,generate_sap_files_balance_report,
                           download_files_balance_report,create_balance_report,
                           zaging_1,zaging_2,zaging_3
)

from SAPAux import SAPSessionManager
class MainWindow(QMainWindow):
    """
    Main Application Window:  
    Acts as the primary UI interface for selecting reports and payment actions.
    
    Workflow:
    - Initializes tab navigation and connects buttons
    - Handles 'Aceptar' by checking selected items and routing to appropriate submodules:
      - Pagos Diarios → either Movimientos Bancarios or Pagos Diarios
      - Informes → Zaging, Informe de Saldos, Grandes Superficies
      - Other tabs → routed via payment function
    - Handles 'Cancelar' by prompting user confirmation and closing SAP session if active
    - Includes utility method for retrieving user selections based on current tab
    
    Parameters:
    - None (UI and behavior configured on initialization)
    
    Returns:
    - None: window drives further workflows and actions through submodules
    """

    def __init__(self):
        super().__init__()
        self.ui = Ui_MainUI()
        self.ui.setupUi(self)
        
        # Initialize sub-window references
        self.zag_wnd = None
        self.blnc_wnd = None

        # Connect buttons
        self.ui.OkBtn.clicked.connect(self.OkClick)
        self.ui.CancelBtn.clicked.connect(self.CancelClick)

    def OkClick(self):
        tab_index = self.ui.tabWidget.currentIndex()
        tab_name = self.ui.tabWidget.tabText(tab_index)
        selected_items = self.get_selected_items_for_tab(tab_name)
    
        if not selected_items:
            QMessageBox.warning(self, "Atención", "Debes seleccionar al menos un elemento.")
            return
        
        self.statusBar().showMessage(f"Procesando {tab_name}...", 3000)
        print("Pestaña:", tab_name)
        print("Elementos seleccionados:", selected_items)
        # Call the appropriate function
        if tab_name == "Pagos Diarios":
            print(tab_name)
            print(selected_items[0])
            if selected_items[0]  == "Movimientos Bancarios":
                bank_file()
            elif selected_items[0]  == "Pagos Diarios":
                daily_payments()
            else:
                print(selected_items[0])
        elif tab_name == "Informes":
            if selected_items[0] == "Zaging":
                self.zag_wnd = AutoZagingWindow()
                self.zag_wnd.show()
            elif selected_items[0] == "Informe de Saldos":
                self.blnc_wnd = BalanceReportWindow()
                self.blnc_wnd.show()
            elif selected_items[0] == "Fichero Grandes Superficies":
                large_format_retailers_file()
            else:
                print(selected_items[0])
        else:
            payment(selected_items[0])
        
    def CancelClick(self):
        reply = QMessageBox.question(
            self, "Salir", "¿Estás seguro que quieres salir?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            if SAPSessionManager.session:
                SAPSessionManager.disconnect(False)
            self.close()

    def get_selected_items_for_tab(self, tab_name):
        tab_mapping = {
            "Pagarés": self.ui.PagList,
            "Confirming": self.ui.ConfList,
            "Pagos Diarios": self.ui.DailyList,
            "Informes": self.ui.ReportsList,
        }
        list_widget = tab_mapping.get(tab_name)
        return [item.text() for item in list_widget.selectedItems()] if list_widget else []

class AutoZagingWindow(QMainWindow):
    """
    Debt Aging Report Submodule:  
    Launches individual steps of the automated Zaging process.
    
    Workflow:
    - Displays Zaging-specific interface
    - Connects buttons to sequential handlers:
      - Step 1 → Launches automated Zaging logic
      - Step 2 → Lauches add SGL entries to the report
      - Step 3 → Triggers final report creation
    
    Parameters:
    - None (activated from main window)
    
    Returns:
    - None: actions routed internally by button handlers
    """

    def __init__(self):
        super().__init__()
        self.ui = Ui_ZagingReportUI()
        self.ui.setupUi(self)

        # Connect buttons to their respective handlers
        self.ui.Zaging_1.clicked.connect(self.handle_zaging_1)
        self.ui.Zaging_2.clicked.connect(self.handle_zaging_2)
        self.ui.Zaging_3.clicked.connect(self.handle_zaging_3)

    def handle_zaging_1(self):
        zaging_1()
        print("Paso 1 ejecutado: Zaging automático iniciado.")

    def handle_zaging_2(self):
        zaging_2()
        print("Paso 2 ejecutado: Procesando etapa intermedia.")

    def handle_zaging_3(self):
        zaging_3()
        print("Paso 3 ejecutado: Generando informe.")

class BalanceReportWindow(QMainWindow):
    """
    Balance Report Submodule:  
    Handles generation and download of SAP balance reports via user-guided steps.
    
    Workflow:
    - Displays balance report interface
    - Connects buttons to corresponding handlers:
      - Step 1 → Triggers SAP report generation
      - Step 2 → Downloads completed spool files
      - Step 3 → Completes reporting logic
    
    Parameters:
    - None (activated from main window)
    
    Returns:
    - None: actions are SAP-triggered and handled internally
    """
    def __init__(self):
        super().__init__()
        self.ui = Ui_BalanceReport()
        self.ui.setupUi(self)

        # Connect buttons to their respective handlers
        self.ui.BalanceReport_1.clicked.connect(self.handle_BalanceReport_1)
        self.ui.BalanceReport_2.clicked.connect(self.handle_BalanceReport_2)
        self.ui.BalanceReport_3.clicked.connect(self.handle_BalanceReport_3)

    def handle_BalanceReport_1(self):
        generate_sap_files_balance_report()
        print("Paso 1 ejecutado: Generando ficheros.")

    def handle_BalanceReport_2(self):
        download_files_balance_report()
        print("Paso 2 ejecutado: Descargando ficheros.")

    def handle_BalanceReport_3(self):
        create_balance_report()
        print("Paso 3 ejecutado: Finalizando reporte de Saldos.")

def main():
    """
    Application Entry Point:  
    Initializes and launches the main application window.
    
    Workflow:
    - Ensures QApplication instance is active
    - Instantiates and displays MainWindow
    - Starts application event loop
    
    Parameters:
    - None (standard PyQt5 startup pattern)
    
    Returns:
    - None: enters blocking Qt exec loop
    """
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

