# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'MainUI.ui'
##
## Created by: Qt User Interface Compiler version 6.9.1
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PyQt5.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PyQt5.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PyQt5.QtWidgets import (QApplication, QFormLayout, QListWidget, QListWidgetItem,
    QMainWindow, QPushButton, QSizePolicy, QStatusBar,
    QTabWidget, QWidget)

class Ui_MainUI(object):
    def setupUi(self, MainUI):
        if not MainUI.objectName():
            MainUI.setObjectName(u"MainUI")
        MainUI.resize(500, 500)
        MainUI.setMinimumSize(QSize(500, 500))
        MainUI.setMaximumSize(QSize(500, 500))
        self.centralwidget = QWidget(MainUI)
        self.centralwidget.setObjectName(u"centralwidget")
        self.tabWidget = QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName(u"tabWidget")
        self.tabWidget.setGeometry(QRect(10, 10, 480, 400))
        font = QFont()
        font.setPointSize(13)
        font.setBold(True)
        self.tabWidget.setFont(font)
        self.PagTab = QWidget()
        self.PagTab.setObjectName(u"PagTab")
        self.formLayout = QFormLayout(self.PagTab)
        self.formLayout.setObjectName(u"formLayout")
        self.PagList = QListWidget(self.PagTab)
        QListWidgetItem(self.PagList)
        QListWidgetItem(self.PagList)
        QListWidgetItem(self.PagList)
        QListWidgetItem(self.PagList)
        QListWidgetItem(self.PagList)
        self.PagList.setObjectName(u"PagList")
        font1 = QFont()
        font1.setPointSize(20)
        font1.setBold(True)
        self.PagList.setFont(font1)

        self.formLayout.setWidget(0, QFormLayout.ItemRole.SpanningRole, self.PagList)

        self.tabWidget.addTab(self.PagTab, "")
        self.ConfTab = QWidget()
        self.ConfTab.setObjectName(u"ConfTab")
        self.ConfList = QListWidget(self.ConfTab)
        QListWidgetItem(self.ConfList)
        QListWidgetItem(self.ConfList)
        QListWidgetItem(self.ConfList)
        QListWidgetItem(self.ConfList)
        QListWidgetItem(self.ConfList)
        QListWidgetItem(self.ConfList)
        self.ConfList.setObjectName(u"ConfList")
        self.ConfList.setGeometry(QRect(9, 9, 456, 343))
        self.ConfList.setFont(font1)
        self.tabWidget.addTab(self.ConfTab, "")
        self.DailyTab = QWidget()
        self.DailyTab.setObjectName(u"DailyTab")
        self.DailyList = QListWidget(self.DailyTab)
        QListWidgetItem(self.DailyList)
        QListWidgetItem(self.DailyList)
        self.DailyList.setObjectName(u"DailyList")
        self.DailyList.setGeometry(QRect(9, 9, 456, 343))
        self.DailyList.setFont(font1)
        self.tabWidget.addTab(self.DailyTab, "")
        self.ReportsTab = QWidget()
        self.ReportsTab.setObjectName(u"ReportsTab")
        self.ReportsList = QListWidget(self.ReportsTab)
        QListWidgetItem(self.ReportsList)
        QListWidgetItem(self.ReportsList)
        QListWidgetItem(self.ReportsList)
        self.ReportsList.setObjectName(u"ReportsList")
        self.ReportsList.setGeometry(QRect(9, 9, 456, 343))
        self.ReportsList.setFont(font1)
        self.tabWidget.addTab(self.ReportsTab, "")
        self.OkBtn = QPushButton(self.centralwidget)
        self.OkBtn.setObjectName(u"OkBtn")
        self.OkBtn.setGeometry(QRect(10, 420, 230, 60))
        self.OkBtn.setFont(font1)
        self.CancelBtn = QPushButton(self.centralwidget)
        self.CancelBtn.setObjectName(u"CancelBtn")
        self.CancelBtn.setGeometry(QRect(260, 420, 230, 60))
        self.CancelBtn.setFont(font1)
        MainUI.setCentralWidget(self.centralwidget)
        self.statusbar = QStatusBar(MainUI)
        self.statusbar.setObjectName(u"statusbar")
        MainUI.setStatusBar(self.statusbar)

        self.retranslateUi(MainUI)

        self.tabWidget.setCurrentIndex(3)


        QMetaObject.connectSlotsByName(MainUI)
    # setupUi

    def retranslateUi(self, MainUI):
        MainUI.setWindowTitle(QCoreApplication.translate("MainUI", u"MainWindow", None))

        __sortingEnabled = self.PagList.isSortingEnabled()
        self.PagList.setSortingEnabled(False)
        ___qlistwidgetitem = self.PagList.item(0)
        ___qlistwidgetitem.setText(QCoreApplication.translate("MainUI", u"Alcampo", None));
        ___qlistwidgetitem1 = self.PagList.item(1)
        ___qlistwidgetitem1.setText(QCoreApplication.translate("MainUI", u"Alcampo Verdes", None));
        ___qlistwidgetitem2 = self.PagList.item(2)
        ___qlistwidgetitem2.setText(QCoreApplication.translate("MainUI", u"Carrefour", None));
        ___qlistwidgetitem3 = self.PagList.item(3)
        ___qlistwidgetitem3.setText(QCoreApplication.translate("MainUI", u"Cecosa", None));
        ___qlistwidgetitem4 = self.PagList.item(4)
        ___qlistwidgetitem4.setText(QCoreApplication.translate("MainUI", u"Eroski", None));
        
        self.PagList.setSortingEnabled(__sortingEnabled)

        self.tabWidget.setTabText(self.tabWidget.indexOf(self.PagTab), QCoreApplication.translate("MainUI", u"Pagar\u00e9s", None))

        __sortingEnabled1 = self.ConfList.isSortingEnabled()
        self.ConfList.setSortingEnabled(False)
        ___qlistwidgetitem5 = self.ConfList.item(0)
        ___qlistwidgetitem5.setText(QCoreApplication.translate("MainUI", u"Alcampo Pago Unif", None));
        ___qlistwidgetitem6 = self.ConfList.item(1)
        ___qlistwidgetitem6.setText(QCoreApplication.translate("MainUI", u"Casa del Libro", None));
        ___qlistwidgetitem7 = self.ConfList.item(2)
        ___qlistwidgetitem7.setText(QCoreApplication.translate("MainUI", u"Consum", None));
        ___qlistwidgetitem8 = self.ConfList.item(3)
        ___qlistwidgetitem8.setText(QCoreApplication.translate("MainUI", u"El Corte Ingles Codice", None));
        ___qlistwidgetitem9 = self.ConfList.item(4)
        ___qlistwidgetitem9.setText(QCoreApplication.translate("MainUI", u"El Corte Ingles Web", None));
        ___qlistwidgetitem10 = self.ConfList.item(5)
        ___qlistwidgetitem10.setText(QCoreApplication.translate("MainUI", u"FNAC", None));
        self.ConfList.setSortingEnabled(__sortingEnabled1)

        self.tabWidget.setTabText(self.tabWidget.indexOf(self.ConfTab), QCoreApplication.translate("MainUI", u"Confirming", None))

        __sortingEnabled2 = self.DailyList.isSortingEnabled()
        self.DailyList.setSortingEnabled(False)
        ___qlistwidgetitem11 = self.DailyList.item(0)
        ___qlistwidgetitem11.setText(QCoreApplication.translate("MainUI", u"Movimientos Bancarios", None));
        ___qlistwidgetitem12 = self.DailyList.item(1)
        ___qlistwidgetitem12.setText(QCoreApplication.translate("MainUI", u"Pagos Diarios", None));
        self.DailyList.setSortingEnabled(__sortingEnabled2)

        self.tabWidget.setTabText(self.tabWidget.indexOf(self.DailyTab), QCoreApplication.translate("MainUI", u"Pagos Diarios", None))

        __sortingEnabled3 = self.ReportsList.isSortingEnabled()
        self.ReportsList.setSortingEnabled(False)
        ___qlistwidgetitem13 = self.ReportsList.item(0)
        ___qlistwidgetitem13.setText(QCoreApplication.translate("MainUI", u"Fichero Grandes Superficies", None));
        ___qlistwidgetitem14 = self.ReportsList.item(1)
        ___qlistwidgetitem14.setText(QCoreApplication.translate("MainUI", u"Informe de Saldos", None));
        ___qlistwidgetitem15 = self.ReportsList.item(2)
        ___qlistwidgetitem15.setText(QCoreApplication.translate("MainUI", u"Zaging", None));
        self.ReportsList.setSortingEnabled(__sortingEnabled3)

        self.tabWidget.setTabText(self.tabWidget.indexOf(self.ReportsTab), QCoreApplication.translate("MainUI", u"Informes", None))
        self.OkBtn.setText(QCoreApplication.translate("MainUI", u"Continuar", None))
        self.CancelBtn.setText(QCoreApplication.translate("MainUI", u"Cancelar", None))
    # retranslateUi

