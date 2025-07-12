# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'BalanceReport.ui'
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
from PyQt5.QtWidgets import (QApplication, QPushButton, QSizePolicy, QWidget)

class Ui_BalanceReport(object):
    def setupUi(self, BalanceReport):
        if not BalanceReport.objectName():
            BalanceReport.setObjectName(u"BalanceReport")
        BalanceReport.resize(300, 300)
        BalanceReport.setMaximumSize(QSize(300, 300))
        BalanceReport.setBaseSize(QSize(300, 300))
        self.BalanceReport_1 = QPushButton(BalanceReport)
        self.BalanceReport_1.setObjectName(u"BalanceReport_1")
        self.BalanceReport_1.setGeometry(QRect(15, 5, 270, 95))
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.BalanceReport_1.sizePolicy().hasHeightForWidth())
        self.BalanceReport_1.setSizePolicy(sizePolicy)
        self.BalanceReport_1.setMinimumSize(QSize(270, 95))
        self.BalanceReport_1.setMaximumSize(QSize(270, 95))
        self.BalanceReport_1.setBaseSize(QSize(270, 95))
        font = QFont()
        font.setPointSize(14)
        font.setBold(True)
        self.BalanceReport_1.setFont(font)
        self.BalanceReport_2 = QPushButton(BalanceReport)
        self.BalanceReport_2.setObjectName(u"BalanceReport_2")
        self.BalanceReport_2.setGeometry(QRect(15, 100, 270, 95))
        sizePolicy.setHeightForWidth(self.BalanceReport_2.sizePolicy().hasHeightForWidth())
        self.BalanceReport_2.setSizePolicy(sizePolicy)
        self.BalanceReport_2.setMinimumSize(QSize(270, 95))
        self.BalanceReport_2.setMaximumSize(QSize(270, 95))
        self.BalanceReport_2.setBaseSize(QSize(370, 95))
        self.BalanceReport_2.setFont(font)
        self.BalanceReport_3 = QPushButton(BalanceReport)
        self.BalanceReport_3.setObjectName(u"BalanceReport_3")
        self.BalanceReport_3.setGeometry(QRect(15, 195, 270, 95))
        sizePolicy.setHeightForWidth(self.BalanceReport_3.sizePolicy().hasHeightForWidth())
        self.BalanceReport_3.setSizePolicy(sizePolicy)
        self.BalanceReport_3.setMinimumSize(QSize(270, 95))
        self.BalanceReport_3.setMaximumSize(QSize(270, 95))
        self.BalanceReport_3.setBaseSize(QSize(270, 95))
        self.BalanceReport_3.setFont(font)

        self.retranslateUi(BalanceReport)

        QMetaObject.connectSlotsByName(BalanceReport)
    # setupUi

    def retranslateUi(self, BalanceReport):
        BalanceReport.setWindowTitle(QCoreApplication.translate("BalanceReport", u"Balance Report GUI", None))
        self.BalanceReport_1.setText(QCoreApplication.translate("BalanceReport", u"Generar ficheros en SAP", None))
        self.BalanceReport_2.setText(QCoreApplication.translate("BalanceReport", u"Descargar ficheros de SAP", None))
        self.BalanceReport_3.setText(QCoreApplication.translate("BalanceReport", u"Genear Informe", None))
    # retranslateUi

