# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'Diff.ui'
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
from PyQt5.QtWidgets import (QApplication, QLabel, QPushButton, QSizePolicy,
    QWidget)

class Ui_Form(object):
    def setupUi(self, Form):
        if not Form.objectName():
            Form.setObjectName(u"Form")
        Form.resize(300, 200)
        Form.setMinimumSize(QSize(300, 200))
        Form.setMaximumSize(QSize(300, 200))
        self.label = QLabel(Form)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(10, 12, 280, 100))
        font = QFont()
        font.setPointSize(13)
        font.setBold(True)
        self.label.setFont(font)
        self.label.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
        self.label.setAutoFillBackground(False)
        self.label.setScaledContents(False)
        self.label.setWordWrap(True)
        self.RoundBtn = QPushButton(Form)
        self.RoundBtn.setObjectName(u"RoundBtn")
        self.RoundBtn.setGeometry(QRect(10, 130, 135, 60))
        font1 = QFont()
        font1.setPointSize(14)
        font1.setBold(True)
        self.RoundBtn.setFont(font1)
        self.ToAccountBtn = QPushButton(Form)
        self.ToAccountBtn.setObjectName(u"ToAccountBtn")
        self.ToAccountBtn.setGeometry(QRect(155, 130, 135, 60))
        self.ToAccountBtn.setFont(font1)

        QMetaObject.connectSlotsByName(Form)
    # setupUi

    def retranslateUi(self, Form,Diff):
        Form.setWindowTitle(QCoreApplication.translate("Form", u"Diferencia", None))
        self.label.setText(QCoreApplication.translate("Form", f"La diferencia es {Diff} \u00bfquieres redondear o A cuenta?", None))
        self.RoundBtn.setText(QCoreApplication.translate("Form", u"Redondear", None))
        self.ToAccountBtn.setText(QCoreApplication.translate("Form", u"A cuenta", None))
    # retranslateUi

