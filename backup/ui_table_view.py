# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ui_table_view.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_TableView(object):
    def setupUi(self, TableView):
        TableView.setObjectName("TableView")
        TableView.resize(876, 515)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(TableView.sizePolicy().hasHeightForWidth())
        TableView.setSizePolicy(sizePolicy)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(TableView)
        self.verticalLayout_2.setSizeConstraint(QtWidgets.QLayout.SetFixedSize)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.groupBox = QtWidgets.QGroupBox(TableView)
        self.groupBox.setTitle("")
        self.groupBox.setObjectName("groupBox")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.groupBox)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.groupBox_2 = QtWidgets.QGroupBox(self.groupBox)
        self.groupBox_2.setObjectName("groupBox_2")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.groupBox_2)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.CopyColumn = QtWidgets.QCheckBox(self.groupBox_2)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.CopyColumn.setFont(font)
        self.CopyColumn.setObjectName("CopyColumn")
        self.horizontalLayout_3.addWidget(self.CopyColumn)
        self.CopyIndex = QtWidgets.QCheckBox(self.groupBox_2)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.CopyIndex.setFont(font)
        self.CopyIndex.setObjectName("CopyIndex")
        self.horizontalLayout_3.addWidget(self.CopyIndex)
        self.horizontalLayout.addWidget(self.groupBox_2)
        self.groupBox_4 = QtWidgets.QGroupBox(self.groupBox)
        self.groupBox_4.setObjectName("groupBox_4")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.groupBox_4)
        self.verticalLayout.setObjectName("verticalLayout")
        self.pushButton_reset = QtWidgets.QPushButton(self.groupBox_4)
        self.pushButton_reset.setObjectName("pushButton_reset")
        self.verticalLayout.addWidget(self.pushButton_reset)
        self.horizontalLayout.addWidget(self.groupBox_4)
        self.groupBox_3 = QtWidgets.QGroupBox(self.groupBox)
        self.groupBox_3.setObjectName("groupBox_3")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.groupBox_3)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.search_lineEdit = QtWidgets.QLineEdit(self.groupBox_3)
        self.search_lineEdit.setObjectName("search_lineEdit")
        self.horizontalLayout_2.addWidget(self.search_lineEdit)
        self.radioButton_case = QtWidgets.QRadioButton(self.groupBox_3)
        self.radioButton_case.setObjectName("radioButton_case")
        self.horizontalLayout_2.addWidget(self.radioButton_case)
        self.pushButton_previous = QtWidgets.QPushButton(self.groupBox_3)
        self.pushButton_previous.setObjectName("pushButton_previous")
        self.horizontalLayout_2.addWidget(self.pushButton_previous)
        self.pushButton_next = QtWidgets.QPushButton(self.groupBox_3)
        self.pushButton_next.setObjectName("pushButton_next")
        self.horizontalLayout_2.addWidget(self.pushButton_next)
        self.label_search_result = QtWidgets.QLabel(self.groupBox_3)
        self.label_search_result.setObjectName("label_search_result")
        self.horizontalLayout_2.addWidget(self.label_search_result)
        self.horizontalLayout.addWidget(self.groupBox_3)
        self.verticalLayout_2.addWidget(self.groupBox)
        self.tableView = QtWidgets.QTableView(TableView)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.tableView.setFont(font)
        self.tableView.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContentsOnFirstShow)
        self.tableView.setAlternatingRowColors(False)
        self.tableView.setSortingEnabled(True)
        self.tableView.setObjectName("tableView")
        self.tableView.horizontalHeader().setStretchLastSection(True)
        self.verticalLayout_2.addWidget(self.tableView)

        self.retranslateUi(TableView)
        QtCore.QMetaObject.connectSlotsByName(TableView)

    def retranslateUi(self, TableView):
        _translate = QtCore.QCoreApplication.translate
        TableView.setWindowTitle(_translate("TableView", "Table_view"))
        self.groupBox_2.setTitle(_translate("TableView", "Copy"))
        self.CopyColumn.setText(_translate("TableView", "Column name"))
        self.CopyIndex.setText(_translate("TableView", "Index name"))
        self.groupBox_4.setTitle(_translate("TableView", "Reset to default table"))
        self.pushButton_reset.setText(_translate("TableView", "Reset"))
        self.groupBox_3.setTitle(_translate("TableView", "Search (support regular expressions)"))
        self.radioButton_case.setText(_translate("TableView", "case sensitive"))
        self.pushButton_previous.setText(_translate("TableView", "Previous"))
        self.pushButton_next.setText(_translate("TableView", "Next"))
        self.label_search_result.setText(_translate("TableView", "Search results cell: 0/0"))
