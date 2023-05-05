# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ui_setting.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_SettingsWindow(object):
    def setupUi(self, SettingsWindow):
        SettingsWindow.setObjectName("SettingsWindow")
        SettingsWindow.resize(400, 600)
        self.centralwidget = QtWidgets.QWidget(SettingsWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.treeWidget = QtWidgets.QTreeWidget(self.centralwidget)
        self.treeWidget.setObjectName("treeWidget")
        self.verticalLayout.addWidget(self.treeWidget)
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setTitle("")
        self.groupBox.setObjectName("groupBox")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.groupBox)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.pushButton_Save_setting = QtWidgets.QPushButton(self.groupBox)
        self.pushButton_Save_setting.setObjectName("pushButton_Save_setting")
        self.horizontalLayout.addWidget(self.pushButton_Save_setting)
        self.pushButton_Reset_default = QtWidgets.QPushButton(self.groupBox)
        self.pushButton_Reset_default.setObjectName("pushButton_Reset_default")
        self.horizontalLayout.addWidget(self.pushButton_Reset_default)
        self.verticalLayout.addWidget(self.groupBox)
        SettingsWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(SettingsWindow)
        self.statusbar.setObjectName("statusbar")
        SettingsWindow.setStatusBar(self.statusbar)

        self.retranslateUi(SettingsWindow)
        QtCore.QMetaObject.connectSlotsByName(SettingsWindow)

    def retranslateUi(self, SettingsWindow):
        _translate = QtCore.QCoreApplication.translate
        SettingsWindow.setWindowTitle(_translate("SettingsWindow", "Settings"))
        self.label.setText(_translate("SettingsWindow", "Press \"Insert\" to insert item under node\n"
"         \"Enter\" to edit item\n"
"         \"Delete\" to delete item"))
        self.treeWidget.headerItem().setText(0, _translate("SettingsWindow", "Configuration"))
        self.pushButton_Save_setting.setText(_translate("SettingsWindow", "Save changes"))
        self.pushButton_Reset_default.setText(_translate("SettingsWindow", "Reset to default setting"))

