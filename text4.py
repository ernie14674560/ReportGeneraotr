import sys
from PyQt5.QtWidgets import (QWidget, QTableWidget, QHBoxLayout, QApplication, QTableWidgetItem )
from PyQt5.QtGui     import QBrush, QColor #,  QFont

class Table(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("QTableWidget - Example of a cell merge")
        self.resize(660,300 );
        conLayout = QHBoxLayout()

        tableWidget = QTableWidget()
        tableWidget.setRowCount(7)
        tableWidget.setColumnCount(6)
        conLayout.addWidget(tableWidget)

        # Hide headers
        tableWidget.horizontalHeader().setVisible(False)
        tableWidget.verticalHeader().setVisible(False)

        #tableWidget.setHorizontalHeaderLabels(['Column1','Column1','Column1'])

        # Sets the span of the table element at (row , column ) to the number of rows
        # and columns specified by (rowSpanCount , columnSpanCount ).
        tableWidget.setSpan(0, 0, 1, 6)
        newItem = QTableWidgetItem("tableWidget.setSpan(0, 0, 1, 6)")
        tableWidget.setItem(0, 0, newItem)

        tableWidget.setSpan(3, 0, 3, 1)
        newItem = QTableWidgetItem("tableWidget.setSpan(3, 0, 3, 1)")
        tableWidget.setItem(3, 0, newItem)

        newItem = QTableWidgetItem("Hello")
        newItem.setForeground(QBrush(QColor(0, 255, 0)))
        tableWidget.setItem(3, 1, newItem)

        newItem = QTableWidgetItem("pythoff")
        newItem.setForeground(QBrush(QColor(255, 0, 0)))
        tableWidget.setItem(3, 2, newItem)

        self.setLayout(conLayout)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    example = Table()
    example.show()
    sys.exit(app.exec_())