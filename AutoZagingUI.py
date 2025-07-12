# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'AutoZagingUI.ui'
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

class Ui_ZagingReportUI(object):
    def setupUi(self, ZagingReportUI):
        if not ZagingReportUI.objectName():
            ZagingReportUI.setObjectName(u"ZagingReportUI")
        ZagingReportUI.resize(300, 300)
        ZagingReportUI.setMaximumSize(QSize(300, 300))
        ZagingReportUI.setBaseSize(QSize(300, 300))
        self.Zaging_1 = QPushButton(ZagingReportUI)
        self.Zaging_1.setObjectName(u"Zaging_1")
        self.Zaging_1.setGeometry(QRect(15, 5, 270, 95))
        sizePolicy = QSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Maximum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.Zaging_1.sizePolicy().hasHeightForWidth())
        self.Zaging_1.setSizePolicy(sizePolicy)
        self.Zaging_1.setMinimumSize(QSize(270, 95))
        self.Zaging_1.setMaximumSize(QSize(270, 95))
        self.Zaging_1.setBaseSize(QSize(270, 95))
        font = QFont()
        font.setPointSize(14)
        font.setBold(True)
        self.Zaging_1.setFont(font)
        self.Zaging_2 = QPushButton(ZagingReportUI)
        self.Zaging_2.setObjectName(u"Zaging_2")
        self.Zaging_2.setGeometry(QRect(15, 100, 270, 95))
        sizePolicy.setHeightForWidth(self.Zaging_2.sizePolicy().hasHeightForWidth())
        self.Zaging_2.setSizePolicy(sizePolicy)
        self.Zaging_2.setMinimumSize(QSize(270, 95))
        self.Zaging_2.setMaximumSize(QSize(270, 95))
        self.Zaging_2.setBaseSize(QSize(370, 95))
        self.Zaging_2.setFont(font)
        self.Zaging_3 = QPushButton(ZagingReportUI)
        self.Zaging_3.setObjectName(u"Zaging_3")
        self.Zaging_3.setGeometry(QRect(15, 195, 270, 95))
        sizePolicy.setHeightForWidth(self.Zaging_3.sizePolicy().hasHeightForWidth())
        self.Zaging_3.setSizePolicy(sizePolicy)
        self.Zaging_3.setMinimumSize(QSize(270, 95))
        self.Zaging_3.setMaximumSize(QSize(270, 95))
        self.Zaging_3.setBaseSize(QSize(270, 95))
        self.Zaging_3.setFont(font)

        self.retranslateUi(ZagingReportUI)

        QMetaObject.connectSlotsByName(ZagingReportUI)
    # setupUi

    def retranslateUi(self, ZagingReportUI):
        ZagingReportUI.setWindowTitle(QCoreApplication.translate("ZagingReportUI", u"Auto Zaging Report GUI", None))
        self.Zaging_1.setText(QCoreApplication.translate("ZagingReportUI", u"Auto_Zaging: Paso 1", None))
        self.Zaging_2.setText(QCoreApplication.translate("ZagingReportUI", u"Auto_Zaging: Paso 2", None))
        self.Zaging_3.setText(QCoreApplication.translate("ZagingReportUI", u"Auto_Zaging: Paso 3", None))
    # retranslateUi

