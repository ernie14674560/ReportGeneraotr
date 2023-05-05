#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import pandas as pd
from Database_query import recp_des_query
from app import TableView
from PyQt5 import QtCore, QtGui, QtWidgets
import random

try:
    from html import escape
except ImportError:
    from cgi import escape

words = [
    "Hello",
    "world",
    "Stack",
    "Overflow",
    "Hello world",
    """<font color="red">Hello world</font>""",
]


class HTMLDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, parent=None):
        super(HTMLDelegate, self).__init__(parent)
        self.doc = QtGui.QTextDocument(self)

    def paint(self, painter, option, index):
        substring = index.data(QtCore.Qt.UserRole)
        painter.save()
        options = QtWidgets.QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        res = ""
        color = QtGui.QColor("orange")
        if not pd.isnull(substring) and substring:
            substrings = options.text.split(substring)
            res = """<font color="{}">{}</font>""".format(
                color.name(QtGui.QColor.HexRgb), substring
            ).join(list(map(escape, substrings)))
        else:
            res = escape(options.text)
        self.doc.setHtml(res)

        options.text = ""
        style = (
            QtWidgets.QApplication.style()
            if options.widget is None
            else options.widget.style()
        )
        style.drawControl(QtWidgets.QStyle.CE_ItemViewItem, options, painter)

        ctx = QtGui.QAbstractTextDocumentLayout.PaintContext()
        if option.state & QtWidgets.QStyle.State_Selected:
            ctx.palette.setColor(
                QtGui.QPalette.Text,
                option.palette.color(
                    QtGui.QPalette.Active, QtGui.QPalette.HighlightedText
                ),
            )
        else:
            ctx.palette.setColor(
                QtGui.QPalette.Text,
                option.palette.color(QtGui.QPalette.Active, QtGui.QPalette.Text),
            )

        textRect = style.subElementRect(QtWidgets.QStyle.SE_ItemViewItemText, options)

        if index.column() != 0:
            textRect.adjust(5, 0, 0, 0)

        thefuckyourshitup_constant = 4
        margin = (option.rect.height() - options.fontMetrics.height()) // 2
        margin = margin - thefuckyourshitup_constant
        textRect.setTop(textRect.top() + margin)

        painter.translate(textRect.topLeft())
        painter.setClipRect(textRect.translated(-textRect.topLeft()))
        self.doc.documentLayout().draw(painter, ctx)

        painter.restore()

    def sizeHint(self, option, index):
        return QtCore.QSize(self.doc.idealWidth(), self.doc.size().height())


class Widget(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(Widget, self).__init__(parent)
        hlay = QtWidgets.QHBoxLayout()
        lay = QtWidgets.QVBoxLayout(self)

        self.le = QtWidgets.QLineEdit()
        self.button = QtWidgets.QPushButton("filter")
        self.table = QtWidgets.QTableWidget(5, 5)
        hlay.addWidget(self.le)
        hlay.addWidget(self.button)
        lay.addLayout(hlay)
        lay.addWidget(self.table)
        self.button.clicked.connect(self.find_items)
        self.table.setItemDelegate(HTMLDelegate(self.table))

        for i in range(self.table.rowCount()):
            for j in range(self.table.columnCount()):
                it = QtWidgets.QTableWidgetItem(random.choice(words))
                self.table.setItem(i, j, it)

    def find_items(self):
        text = self.le.text()
        # clear
        allitems = self.table.findItems("", QtCore.Qt.MatchContains)
        selected_items = self.table.findItems(self.le.text(), QtCore.Qt.MatchContains)
        for item in allitems:
            item.setData(QtCore.Qt.UserRole, text if item in selected_items else None)


# if __name__ == "__main__":
#     import sys
#
#     app = QtWidgets.QApplication(sys.argv)
#     w = Widget()
#     w.show()
#     sys.exit(app.exec_())
#
if __name__ == '__main__':
    import sys

    # Back up the reference to the exceptionhook
    sys._excepthook = sys.excepthook


    def my_exception_hook(exctype, value, traceback):
        # Print the error and traceback
        print(exctype, value, traceback)
        # Call the normal Exception hook after
        sys._excepthook(exctype, value, traceback)
        sys.exit(1)


    # Set the exception hook to our wrapping function
    sys.excepthook = my_exception_hook
    app = QtWidgets.QApplication(sys.argv)
    # dialog = TableFilteringDialog(recp_des_query(capability='1DANNL00'))
    dialog = TableView(df=recp_des_query(capability='1DANNL00'))
    dialog.show()
    try:
        sys.exit(app.exec_())
    except:
        print("Exiting")
