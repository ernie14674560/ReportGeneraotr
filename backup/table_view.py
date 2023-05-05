#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import re
import warnings

import numpy as np
import pandas as pd
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QDialogButtonBox, QAbstractItemView
from pandas.util._validators import validate_bool_kwarg

from ui_table_view import Ui_TableView

try:
    from html import escape
except ImportError:
    from cgi import escape


def selected_dropna(df, selected_idxs: list, axis=0, how="any", thresh=None, subset=None,
                    inplace=False):
    """

    :param df:
    :param selected_idxs: [(start, end), ....]
    :param axis:
    :param how:
    :param thresh:
    :param subset:
    :param inplace:
    :return:
    """
    inplace = validate_bool_kwarg(inplace, "inplace")

    if isinstance(axis, (tuple, list)):
        # GH20987
        msg = (
            "supplying multiple axes to axis is deprecated and "
            "will be removed in a future version."
        )
        warnings.warn(msg, FutureWarning, stacklevel=2)

        result = df
        for ax in axis:
            result = result.dropna(how=how, thresh=thresh, subset=subset, axis=ax)
    else:
        axis = df._get_axis_number(axis)
        agg_axis = 1 - axis

        agg_obj = df
        if subset is not None:
            ax = df._get_axis(agg_axis)
            indices = ax.get_indexer_for(subset)
            check = indices == -1
            if check.any():
                raise KeyError(list(np.compress(check, subset)))
            agg_obj = df.take(indices, axis=agg_axis)

        count = agg_obj.count(axis=agg_axis)

        if thresh is not None:
            mask = count >= thresh
        elif how == "any":
            mask = count == len(agg_obj._get_axis(agg_axis))
        elif how == "all":
            mask = count > 0
        else:
            if how is not None:
                raise ValueError("invalid how option: {h}".format(h=how))
            else:
                raise TypeError("must specify how or thresh")
        result_mask = pd.Series(data=True, index=mask.index)
        for start_idx, end_idx in selected_idxs:
            result_mask.loc[start_idx:end_idx] = mask.loc[start_idx:end_idx]
        result = df.loc(axis=axis)[result_mask]
    if inplace:
        df._update_inplace(result)
    else:
        return result


class HTMLHighlightDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, parent=None):
        super(HTMLHighlightDelegate, self).__init__(parent)
        self.doc = QtGui.QTextDocument(self)

    def paint(self, painter, option, index):
        substring = index.data(QtCore.Qt.UserRole)
        painter.save()
        options = QtWidgets.QStyleOptionViewItem(option)
        self.initStyleOption(options, index)
        # res = str()
        text_color = QtGui.QColor("red")
        bg_color = QtGui.QColor("yellow")
        if not pd.isnull(substring) and substring:
            substrings = options.text.split(substring)
            res = f"""<font style="color: {text_color.name(QtGui.QColor.HexRgb)}; background-color: {bg_color.name(
                QtGui.QColor.HexRgb)};">{substring}</font>""".join(list(map(escape, substrings)))
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

        some_constant = 4
        margin = (option.rect.height() - options.fontMetrics.height()) // 2
        margin = margin - some_constant
        textRect.setTop(textRect.top() + margin)

        painter.translate(textRect.topLeft())
        painter.setClipRect(textRect.translated(-textRect.topLeft()))
        self.doc.documentLayout().draw(painter, ctx)

        painter.restore()


class FilterHeader(QtWidgets.QHeaderView):
    filterActivated = QtCore.pyqtSignal()

    def __init__(self, parent):
        super().__init__(QtCore.Qt.Horizontal, parent)
        self._editors = []
        self._padding = 4
        self.parent_table = parent
        self.sectionResized.connect(self.adjustPositions)
        parent.horizontalScrollBar().valueChanged.connect(
            self.adjustPositions)
        self.setSectionsClickable(True)
        self.setHighlightSections(True)

    def setFilterBoxes(self, count):
        while self._editors:
            editor = self._editors.pop()
            editor.deleteLater()
        for index in range(count):
            editor = QtWidgets.QLineEdit(self.parent())
            editor.setPlaceholderText('Filter')
            editor.textEdited.connect(self.filterActivated.emit)  # determine how to connect to filter action
            self._editors.append(editor)
        self.adjustPositions()

    def sizeHint(self):
        size = super().sizeHint()
        if self._editors:
            height = self._editors[0].sizeHint().height()
            size.setHeight(size.height() + height + self._padding)
        return size

    def updateGeometries(self):
        if self._editors:
            height = self._editors[0].sizeHint().height()
            self.setViewportMargins(0, 0, 0, height + self._padding)
        else:
            self.setViewportMargins(0, 0, 0, 0)
        super().updateGeometries()
        self.adjustPositions()

    def adjustPositions(self):
        for index, editor in enumerate(self._editors):
            height = editor.sizeHint().height()
            vwidth = self.parent_table.verticalHeader().width()
            editor.move(
                self.sectionPosition(index) - self.offset() + vwidth,
                height + (self._padding // 2))
            editor.resize(self.sectionSize(index), height)

    def filterText(self, index):
        if 0 <= index < len(self._editors):
            return self._editors[index].text()
        return ''

    def setFilterText(self, index, text):
        if 0 <= index < len(self._editors):
            self._editors[index].setText(text)

    def clearFilters(self):
        for editor in self._editors:
            editor.clear()


class DataFrameModel(QtCore.QAbstractTableModel):
    """
    Dedicated model for inputted DataFrame, will transform input df dtype to string
    """
    DtypeRole = QtCore.Qt.UserRole + 1000
    ValueRole = QtCore.Qt.UserRole + 1001

    def __init__(self, df=pd.DataFrame(), parent=None):
        super(DataFrameModel, self).__init__(parent)
        # self._df_source = df.applymap(str)
        # self._df_display = self._df_source.copy()
        self._df_source = df
        self._df_display = df.applymap(lambda x: str(x) if pd.notnull(x) else '')
        self._df_delegate = pd.DataFrame().reindex_like(self._df_display)
        self._filters = {}
        self._sortBy = []
        self._sortDirection = []

    def getDataFreame(self):
        return self._df_source

    def setDataFrame(self, df):
        self.beginResetModel()
        # self._df_source = df.applymap(str)
        self._df_source = df
        self.endResetModel()

    def dataFrame(self):
        return self._df_source

    dataFrame = QtCore.pyqtProperty(pd.DataFrame, fget=dataFrame, fset=setDataFrame)

    @QtCore.pyqtSlot(int, QtCore.Qt.Orientation, result=str)
    def headerData(self, section: int, orientation: QtCore.Qt.Orientation, role: int = QtCore.Qt.DisplayRole):
        if role == QtCore.Qt.DisplayRole:
            if orientation == QtCore.Qt.Horizontal:
                try:
                    col_name = self._df_display.columns[section]
                except IndexError:
                    return QtCore.QVariant()
                return str(col_name)
            else:
                try:
                    idx_name = self._df_display.index[section]
                except IndexError:
                    return QtCore.QVariant()
                return str(idx_name)
        return QtCore.QVariant()

    def rowCount(self, parent=QtCore.QModelIndex()):
        if parent.isValid():
            return 0
        return len(self._df_source.index)

    def columnCount(self, parent=QtCore.QModelIndex()):
        if parent.isValid():
            return 0
        return self._df_source.columns.size

    def data(self, index, role=QtCore.Qt.DisplayRole):
        # alternate bgcolor by first column contents
        if role == QtCore.Qt.BackgroundRole:
            v = self.data(self.index(index.row(), 0), QtCore.Qt.DisplayRole)
            if not v:  # if empty
                return QtGui.QBrush(QtGui.QColor(53, 53, 53))  # alternateBase color

        if not index.isValid() or not (0 <= index.row() < self.rowCount()
                                       and 0 <= index.column() < self.columnCount()):
            return QtCore.QVariant()
        row = index.row()
        col = index.column()
        dt = self._df_display.iloc[:, col].dtype
        # val = self._df_display.values[row][col]
        try:
            val = self._df_display.values[row][col]
        except IndexError:
            return QtCore.QVariant()
        if role == QtCore.Qt.DisplayRole:
            return str(val)
        elif role == DataFrameModel.ValueRole:
            return val
        elif role == DataFrameModel.DtypeRole:
            return dt
        elif role == QtCore.Qt.UserRole:
            return self._df_delegate.values[row][col]
        return QtCore.QVariant()

    def setData(self, index, value, role=QtCore.Qt.EditRole):
        row = self._df_display.index[index.row()]
        col = self._df_display.columns[index.column()]
        if hasattr(value, "toPyObject"):
            # PyQt4 gets a QVariant
            value = value.toPyObject()
        else:
            # PySide gets an unicode
            dtype = self._df_display[col].dtype
            if dtype != object:
                value = None if value == "" else dtype.type(value)
        self._df_display.at[row, col] = value
        self.dataChanged.emit(index, index)
        return True

    def flags(self, index):
        flags = (
                QtCore.Qt.ItemIsSelectable
                | QtCore.Qt.ItemIsDragEnabled
                | QtCore.Qt.ItemIsEditable
                | QtCore.Qt.ItemIsEnabled
        )
        return flags

    def setFilters(self, filters):
        self.beginResetModel()
        self._filters = filters
        self.updateDisplay()
        self.endResetModel()

    def findItems(self, text, flags=re.IGNORECASE):
        """
            flags default: case in-sensitive
        """
        self.layoutAboutToBeChanged.emit()
        df = self._df_display
        search_text = r'({})'.format(text)
        mask = np.column_stack([df[col].str.extract(search_text, flags=flags, expand=False).fillna(0) for col in df])
        self._df_delegate = pd.DataFrame(mask, index=df.index, columns=df.columns)
        xy_coords = np.argwhere(mask)
        self.layoutChanged.emit()
        return xy_coords

    def sort(self, col=str(), order=QtCore.Qt.AscendingOrder, set_default=False):

        # Storing persistent indexes
        self.layoutAboutToBeChanged.emit()
        oldIndexList = self.persistentIndexList()
        oldIds = self._df_display.index.copy()

        # Sorting data
        if not set_default:
            column = self._df_display.columns[col]
            ascending = (order == QtCore.Qt.AscendingOrder)
            if column in self._sortBy:
                i = self._sortBy.index(column)
                self._sortBy.pop(i)
                self._sortDirection.pop(i)
            self._sortBy.insert(0, column)
            self._sortDirection.insert(0, ascending)

        self.updateDisplay(set_default=set_default)

        # Updating persistent indexes
        newIds = self._df_display.index
        newIndexList = []
        for index in oldIndexList:
            id = oldIds[index.row()]
            newRow = newIds.get_loc(id)
            newIndexList.append(self.index(newRow, index.column(), index.parent()))
        self.changePersistentIndexList(oldIndexList, newIndexList)
        self.layoutChanged.emit()
        self.dataChanged.emit(QtCore.QModelIndex(), QtCore.QModelIndex())

    def updateDisplay(self, set_default=False):

        dfDisplay1 = self._df_source.copy()

        # Filtering
        cond = pd.Series(True, index=dfDisplay1.index)

        for column, value in self._filters.items():
            cond = cond & (dfDisplay1[column].astype(str).str.lower().str.find(str(value).lower()) >= 0)
        dfDisplay1 = dfDisplay1[cond]
        if not set_default:
            # Sorting
            if len(self._sortBy) != 0:
                dfDisplay1.sort_values(by=self._sortBy,
                                       ascending=self._sortDirection,
                                       inplace=True)

        # Updating
        # self._df_display = dfDisplay1
        self._df_display = dfDisplay1.applymap(lambda x: str(x) if pd.notnull(x) else '')

    def roleNames(self):
        roles = {
            QtCore.Qt.DisplayRole: b'display',
            DataFrameModel.DtypeRole: b'dtype',
            DataFrameModel.ValueRole: b'value'
        }
        return roles


class TableView(QtWidgets.QWidget):
    def __init__(self, title='TableView', parent=None, df=pd.DataFrame(), model=None, filter_header=True,
                 resize_and_fixed=True, alternatingRowColors=True, init_show=True):
        super(TableView, self).__init__(parent=parent)
        # if ui is None:
        #     self.ui = Ui_TableView()
        # else:
        #     self.ui = ui()
        if parent is not None:
            parent.child_windows.append(self)
        self.ui = Ui_TableView()
        self.ui.setupUi(self)
        self.setWindowTitle(title)
        # self.view_height = 0
        if model is None:
            self.table_model = DataFrameModel(df)
        else:
            self.table_model = model
        self.clip = QtWidgets.QApplication.clipboard()

        self.search_matches = []
        self.search_order = 0
        self.search_total = 0
        # def input_table(self):
        if alternatingRowColors:
            self.ui.tableView.setAlternatingRowColors(True)
        if filter_header:
            self.header = FilterHeader(self.ui.tableView)
            self.ui.tableView.setHorizontalHeader(self.header)
            self.header.setFilterBoxes(self.table_model.columnCount())
            self.header.filterActivated.connect(self.handleFilterActivated)
        else:
            self.header = self.ui.tableView.horizontalHeader()
        self.ui.tableView.setModel(self.table_model)
        self.ui.tableView.setItemDelegate(HTMLHighlightDelegate(self.ui.tableView))
        self.ui.tableView.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        if resize_and_fixed:
            self.ui.tableView.resizeColumnsToContents()
            self.resizeViewToColumns()
            self.setFixedSize(self.size())
        self.set_search_button_enable(False)
        self.ui.pushButton_reset.clicked.connect(lambda: self.table_model.sort(set_default=True))
        self.ui.pushButton_next.clicked.connect(lambda: self.search_move_selection(forward=True))
        self.ui.pushButton_previous.clicked.connect(lambda: self.search_move_selection(forward=False))
        self.ui.search_lineEdit.setPlaceholderText('Search')
        self.ui.search_lineEdit.textEdited.connect(self.search_items)
        self.ui.search_lineEdit.returnPressed.connect(lambda: self.search_move_selection(forward=True))

        # show UI
        if init_show:
            self.show()

    def set_search_button_enable(self, t_or_f):
        for btn in [self.ui.pushButton_next, self.ui.pushButton_previous]:
            btn.setEnabled(t_or_f)

    def handleFilterActivated(self):
        filters = {}
        header = self.ui.tableView.horizontalHeader()
        for col_idx in range(header.count()):
            value = header.filterText(col_idx)
            if value:
                column = self.table_model.headerData(col_idx, QtCore.Qt.Horizontal, QtCore.Qt.DisplayRole)
                filters[column] = value
        self.table_model.setFilters(filters)

    def resizeViewToColumns(self):

        vwidth = self.ui.tableView.verticalHeader().width()
        hwidth = self.ui.tableView.horizontalHeader().length()
        swidth = self.ui.tableView.style().pixelMetric(QtWidgets.QStyle.PM_ScrollBarExtent)
        fwidth = self.ui.tableView.frameWidth() * 2
        desire_width = vwidth + hwidth + swidth + fwidth
        if desire_width > 1600:
            desire_width = 1600
        self.ui.tableView.setFixedWidth(desire_width)
        vlen = self.ui.tableView.verticalHeader().length()
        hlen = self.ui.tableView.horizontalHeader().height()
        desire_len = vlen + hlen + fwidth + swidth
        if desire_len > 800:
            desire_len = 800
        # if desire_len == self.view_height:
        #     desire_len += 1
        # self.view_height = desire_len
        self.ui.tableView.setFixedHeight(desire_len)

        margins = self.layout().contentsMargins()
        window_width = margins.left() + margins.right() + desire_width
        window_height = margins.top() + margins.bottom() + desire_len
        self.resize(window_width, window_height)

        # self.setFixedSize(QSize(window_width, window_height))  # setFixedSize(QSize)

    def keyPressEvent(self, e):
        if e.modifiers() and QtCore.Qt.ControlModifier:
            if e.key() == QtCore.Qt.Key_C:
                # copy
                self.copy_to_clipboard()

    def search_items(self):
        search_text = self.ui.search_lineEdit.text()
        matches = self.table_model.findItems(search_text,
                                             flags=0 if self.ui.radioButton_case.isChecked() else re.IGNORECASE)
        total = len(matches)
        self.search_matches = matches
        self.search_total = total
        if total and search_text:
            self.search_order = 0
            row, col = matches[0]
            self.ui.label_search_result.setText('Search result cell: 1/{}'.format(self.search_total))
            self.ui.label_search_result.repaint()
            self.select_item(row, col)
            self.set_search_button_enable(True)
        else:
            self.ui.label_search_result.setText('Search result cell: 0/0')
            self.ui.label_search_result.repaint()
            self.set_search_button_enable(False)
            self.ui.tableView.clearSelection()

    def search_move_selection(self, forward=True):
        if self.search_total:
            self.ui.tableView.clearSelection()
            if forward:
                self.search_order += 1
            else:
                self.search_order -= 1
            if abs(self.search_order) == self.search_total:
                self.search_order = 0
            display_order = self.search_order
            if display_order < 0:
                display_order += self.search_total
            display_order += 1
            self.ui.label_search_result.setText('Search result cell: {}/{}'.format(display_order, self.search_total))
            self.ui.label_search_result.repaint()
            row, col = self.search_matches[self.search_order]
            self.select_item(row, col)

    def select_item(self, row, col):
        index = self.table_model.index(row, col)
        self.ui.tableView.selectionModel().select(index, QtCore.QItemSelectionModel.Select)
        self.ui.tableView.scrollTo(index)

    def copy_to_clipboard(self):
        col_header = self.ui.CopyColumn.isChecked()
        row_idx = self.ui.CopyIndex.isChecked()
        selected = self.ui.tableView.selectionModel()
        selected_idx = selected.selectedIndexes()
        if len(selected_idx) > 0:
            # sort select indexes into rows and columns
            rows = sorted(index.row() for index in selected_idx)
            columns = sorted(index.column() for index in selected_idx)
            exist_rows = set(rows)
            exist_columns = set(columns)
            df_table = pd.DataFrame()
            for index in selected_idx:
                row = index.row()
                column = index.column()
                df_table.at[row, column] = index.data()
            if row_idx:
                idx_dict = {i: self.ui.tableView.verticalHeader().model().headerData(i, QtCore.Qt.Vertical) for i in
                            exist_rows}
                df_table.rename(index=idx_dict, inplace=True)
            if col_header:
                col_dict = {col: self.ui.tableView.horizontalHeader().model().headerData(col, QtCore.Qt.Horizontal) for
                            col in exist_columns}
                df_table.rename(columns=col_dict, inplace=True)
            table = df_table.to_csv(index=row_idx, header=col_header, sep='\t')
            self.clip.setText(table)


class BinSelectionView(TableView):
    def __init__(self, query_func, worker, df=pd.DataFrame(), title="Select your want", parent=None, init_show=False,
                 merge_cell=True):
        super(BinSelectionView, self).__init__(init_show=init_show, df=df, title=title, parent=parent)
        self.worker = worker
        self.search_hint = QtWidgets.QLabel()
        self.search_hint.setText('Select bin code for part to inspect, default is all')
        self.search_hint.setAlignment(QtCore.Qt.AlignBottom | QtCore.Qt.AlignLeft)
        self.ui.tableView.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.ui.verticalLayout_2.insertWidget(1, self.search_hint)
        self.buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.query_func = query_func
        self.buttonBox.accepted.connect(self.ok)
        self.buttonBox.rejected.connect(self.cancel)
        self.ui.verticalLayout_2.addWidget(self.buttonBox)
        if merge_cell:
            self.merge_cell()
        self.show()

    def merge_cell(self):
        """find duplicated value in col and merge them in table view"""
        df = self.table_model.getDataFreame()
        col_len = len(df.columns)
        df_shift = df.shift() != df
        idx_max = df_shift.last_valid_index()
        counters = {k: [] for k in range(col_len)}  # store "True" location in df for each col
        counters2 = {k: [] for k in range(col_len)}
        for y, row in enumerate(df_shift.itertuples(index=False)):
            for x, col in enumerate(row):
                if col:
                    counters[x].append(y)
        for col, values in counters.items():
            offsets = values, values[1:] + [idx_max + 1]
            for current, nxt in list(zip(*offsets)):
                # (current, next)
                if nxt - current > 1:
                    counters2[col].append((current, nxt))  # merge index range like (0, 3) for each col

        for col, lst in counters2.items():
            for current, nxt in lst:
                self.ui.tableView.setSpan(current, col, nxt - current, 1)

    def ok(self):
        df = self.table_model.getDataFreame()
        side_d = {'FS': 'fs_bins', 'BS': 'bs_bins', 'CP': 'cp_bins'}
        # use set to prevent some duplication index bug
        side_dict = {'fs_bins': set(), 'bs_bins': set(), 'cp_bins': set()}
        selected = self.ui.tableView.selectionModel()
        selected_idx = selected.selectedIndexes()
        if len(selected_idx) > 0:
            for index in selected_idx:
                idx = index.row()
                side = df.at[idx, 'Side']
                bin_code = df.at[idx, 'Bin']
                side_dict[side_d[side]].add(bin_code)
            selected_bin_codes = {k: tuple(v) for k, v in side_dict.items()}
            self.worker.fetch(self.query_func, kwargs=selected_bin_codes)

        else:
            self.worker.fetch(self.query_func)

    def cancel(self):
        self.close()
