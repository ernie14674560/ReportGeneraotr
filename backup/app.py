#!/usr/bin/env python
# _*_ coding:utf-8 _*_
import os
import pickle
import subprocess
import sys
import warnings

from wm_app import WaferMapApp

warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.filterwarnings("ignore", "(?s).*MATPLOTLIBDATA.*", category=UserWarning)

from PyQt5 import QtWidgets, QtGui, QtCore
from Configuration import cfg, dump_cfg, set_default_cfg, reset_cfg, display_cfg
from Custum_function import InputLotListToGet, recipe_query_by_title_des_cap, chk_jpg_and_convert_to_png
from Database_query import con_close
from DefCode_update import update_by_db, update_by_ini
from DefGroup_update import df_group_update
from Template_generator import NestedDict
from Weekly_update import update_all_report_until
from multi_thread import ThreadObject, Request
from myui import Ui_MainWindow
from table_view import TableView, BinSelectionView
from ui_setting import Ui_SettingsWindow
from wm_utils import export_to_excel

# import warnings
#
# warnings.filterwarnings("error")

# [(button name, display text, path to the file), ...]
open_menu_file = [('actionocap', 'open_ocap_template_xlsm', r'templates/ocap_template.xlsm'),
                  ('action_item_for_ppt', 'open_action_item_for_powerpoint', r'reference/action_item_for_ppt.xlsx')]

jpg_to_png_menu = [(f'action_{path}', path, path) for path in cfg['Convert jpg to png']]

# [(button name, display text, execute func), ...]
file_update_file_menu = [
    ('actionDefect_code_lists_by_Inspection_ini', 'Defect code lists by Inspection.ini', update_by_ini),
    ('actionDefect_code_lists_by_database', 'Defect code lists by database', update_by_db),
    ('actionDefectGroup_by_EDAS_ini', 'DefectGroup by EDAS.ini', df_group_update)]


class Error(Exception):
    def __init__(self, s: str):
        self.value = s

    def __str__(self):
        return self.value


def error_handler(value):
    try:
        return str(value)
    except Exception:
        pass
    try:
        value = value.decode()
        return value.encode("ascii", "backslashreplace")
    except Exception:
        pass
    return '<unprintable %s object>' % type(value).__name__


def open_file(filename, subprocesses):
    "explorer"
    # with subprocess.Popen([os.path.abspath(filename)], shell=True) as sub:
    #     subprocesses.append(sub)
    with subprocess.Popen(["explorer", os.path.abspath(filename)]) as sub:
        subprocesses.append(sub)


def open_wafer_map(result_dict):
    WaferMapApp(**result_dict)


def open_wafer_maps(result_dict, subprocesses):
    # def run(*popenargs, input=None, **kwargs):
    PIPE = -1
    in_put = pickle.dumps(result_dict)
    kwargs = {'stdout': subprocess.PIPE, 'stderr': subprocess.PIPE, 'shell': False}
    # popenargs = [sys.executable, 'open_wafer_map/open_wafer_map.exe']
    # stderr = ''
    if kwargs.get('stdin') is not None:
        raise ValueError('stdin and input arguments may not both be used.')
    kwargs['stdin'] = PIPE
    # with subprocess.Popen(popenargs, **kwargs) as process:
    # with subprocess.Popen(executable=r'open_wafer_map/open_wafer_map.exe', args="", **kwargs) as process:
    with subprocess.Popen(executable=r'open_wafer_map.exe', args="", **kwargs) as process:
        subprocesses.append(process)
        stdout, stderr = process.communicate(in_put)
    if stderr:
        raise Error(error_handler(stderr))


def fill_item(item, value):
    item.setExpanded(True)
    if type(value) is dict:
        for key, val in sorted(value.items()):
            child = QtWidgets.QTreeWidgetItem()
            child.setText(0, str(key))
            item.addChild(child)
            fill_item(child, val)
    elif type(value) is list:
        for val in value:
            child = QtWidgets.QTreeWidgetItem()
            item.addChild(child)
            if type(val) is dict:
                child.setText(0, '[dict]')
                fill_item(child, val)
            elif type(val) is list:
                child.setText(0, '[list]')
                fill_item(child, val)
            else:
                child.setText(0, str(val))
            child.setExpanded(True)
    else:
        child = QtWidgets.QTreeWidgetItem()
        child.setText(0, str(value))
        item.addChild(child)


def get_current_nodes(item):
    nodes = [item.text(0)]
    while True:
        item = item.parent()
        if item is not None:
            nodes.insert(0, item.text(0))
        else:
            break
    return nodes


class InputDialog(QtWidgets.QDialog):
    def __init__(self, title: str, string_items=None, list_items=None, parent=None):
        """
        :param title: title
        :param string_items: {display text: default str, ...}
        :param list_items: {display text: [default list], ...}
        :param parent:parent
        """
        super().__init__(parent)
        if string_items is None:
            string_items = {}
        if list_items is None:
            list_items = {}
        self.setWindowTitle(title)

        self.buttonBox = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel,
                                                    self)
        self.outputs1 = []
        self.outputs2 = []
        layout = QtWidgets.QFormLayout(self)
        # add combobox
        for item, l in list_items.items():
            output = QtWidgets.QComboBox(self)
            output.addItems(l)
            output.currentTextChanged.connect(self.disableButton)
            layout.addRow(item, output)
            self.outputs2.append(output)

        # add line edit
        for item, s in string_items.items():
            output = QtWidgets.QLineEdit(self)
            output.setText(s)
            output.textChanged.connect(self.disableButton)
            layout.addRow(item, output)
            self.outputs1.append(output)
        layout.addWidget(self.buttonBox)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)

        btn = self.buttonBox.button(QtWidgets.QDialogButtonBox.Ok)
        btn.setEnabled(False)

    def disableButton(self):
        # if len(self.leInput.text()) > 0:
        btn = self.buttonBox.button(QtWidgets.QDialogButtonBox.Ok)
        if all(v for v in self.getInputs()):

            btn.setEnabled(True)
        else:
            btn.setEnabled(False)

    def getInputs(self):
        result = []

        for i in self.outputs2:
            result.append(str(i.currentText()))
        for i in self.outputs1:
            result.append(i.text())
        return result


class SettingsViewTree(QtWidgets.QMainWindow):
    def __init__(self, parent=None, value=None):
        if value is None:
            value = {}
        super(SettingsViewTree, self).__init__(parent=parent)
        self.ui = Ui_SettingsWindow()
        self.ui.setupUi(self)
        self.tree = self.ui.treeWidget
        self.tree.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        insertKey = QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Insert), self.tree)
        insertKey.activated.connect(self.itemInsert)
        editKey = QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Return), self.tree)
        editKey.activated.connect(self.itemEdit)
        delkey = QtWidgets.QShortcut(QtGui.QKeySequence(QtCore.Qt.Key_Delete), self.tree)
        delkey.activated.connect(self.itemDel)
        self.nested_dict = NestedDict(value)
        fill_item(self.tree.invisibleRootItem(), value)

        self.ui.pushButton_Save_setting.clicked.connect(self.save_setting)
        self.ui.pushButton_Reset_default.clicked.connect(self.reset_to_default)

    def saved_info_MessageBox(self):
        QtWidgets.QMessageBox.information(self, "Save Complete!",
                                          "Changes have been saved to the configuration file!!"
                                          "\nThe application has to restart so the changes take effect.",
                                          QtWidgets.QMessageBox.Ok)

    def save_setting(self):
        ok = QtWidgets.QMessageBox.question(self, "Save Changes",
                                            "Will save the changes to the configuration file, Continue?",
                                            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)

        if ok == 16384:
            dump_cfg(self.nested_dict.dict)
            reset_cfg()
            self.saved_info_MessageBox()

    def reset_to_default(self):
        data = set_default_cfg()
        self.nested_dict = NestedDict(data)
        self.tree.clear()
        fill_item(self.tree.invisibleRootItem(), data)
        self.tree.update()
        self.saved_info_MessageBox()

    def itemInsert(self):
        """ key : Insert"""
        text, ok = QtWidgets.QInputDialog.getText(self, "Add Child", "Enter child name:")
        if ok and text != "":
            selected = self.tree.selectedItems()
            item = selected[0]

            if len(selected) > 0:
                nodes = get_current_nodes(item)
                leaf = self.nested_dict[tuple(nodes)]
                if isinstance(leaf, NestedDict):
                    text1, ok = QtWidgets.QInputDialog.getText(self, "Add Child Value", "Enter child value:")
                    if not ok:
                        return None
                    child = QtWidgets.QTreeWidgetItem()
                    child.setText(0, str(text))
                    item.addChild(child)
                    fill_item(child, text1)
                    nodes.append(text)
                    self.nested_dict[tuple(nodes)] = text1
                else:
                    if isinstance(leaf, list):
                        leaf.append(text)
                    elif isinstance(leaf, str):  # turn to list
                        self.nested_dict[tuple(nodes)] = [item.text(0), text]
                    QtWidgets.QTreeWidgetItem(item, [text])

            else:
                QtWidgets.QTreeWidgetItem(self.tree, [text])

    def itemEdit(self):
        """ key : Enter"""
        selected = self.tree.selectedItems()
        if selected:
            item = selected[0]
            text, ok = QtWidgets.QInputDialog.getText(self, "Edit Child", "Modify value:", QtWidgets.QLineEdit.Normal,
                                                      item.text(0))
            if ok and text != "":
                nodes = get_current_nodes(item)
                parents = nodes[:-1]
                leaf = self.nested_dict[tuple(parents)]
                if isinstance(leaf, list):
                    self.nested_dict[tuple(parents)] = [text if x == item.text(0) else x for x in leaf]
                elif isinstance(leaf, str):  # leaf is string
                    self.nested_dict[tuple(parents)] = text
                elif isinstance(leaf, NestedDict):
                    empty = not bool(parents)
                    parents.append(text)
                    if empty:  # empty
                        self.nested_dict[tuple(parents)] = self.nested_dict.pop(item.text(0))
                    else:
                        self.nested_dict[tuple(parents)] = self.nested_dict[tuple(parents[:-1])].pop(item.text(0))
                item.setText(0, text)

    def itemDel(self):
        """ key : Delete"""
        selected = self.tree.selectedItems()
        if selected:
            item = selected[0]
            current_item = self.tree.currentItem()
            ok = QtWidgets.QMessageBox.question(self, "Delete Child", "Will Delete The Selective Child, Continue?",
                                                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if ok == 16384:  # ok
                nodes = get_current_nodes(item)
                parents = nodes[:-1]
                leaf = self.nested_dict[tuple(parents)]
                if isinstance(leaf, list):
                    self.nested_dict[tuple(parents)] = [x for x in leaf if x != item.text(0)]

                else:  # str and nested dict
                    del_item = item.text(0)
                    if isinstance(leaf, str):
                        parents = nodes[:-2]
                        del_item = nodes[-2]
                    empty = not bool(parents)

                    if empty:  # empty
                        self.nested_dict.pop(del_item)
                    else:
                        self.nested_dict[tuple(parents)].pop(del_item, None)
                children = []
                for child in selected:
                    children.append(child)
                for child in children:
                    current_item.removeChild(child)
                if isinstance(leaf, str):
                    if current_item.parent().parent() is not None:
                        current_item.parent().parent().removeChild(current_item.parent())
                    else:
                        root = self.tree.invisibleRootItem()
                        root.removeChild(current_item.parent())


class EmittingStream(QtCore.QObject):
    textWritten = QtCore.pyqtSignal(str)

    def write(self, text):
        self.textWritten.emit(str(text))


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None, **kwargs):
        super(MainWindow, self).__init__(parent=parent)
        self.worker = ThreadObject()
        self.worker.data_finished.connect(self.result_classifier)
        self.worker.work_stopped_signal.connect(self.on_work_stopped)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        # self.setFixedSize(self.size())  # fix window size
        self.child_windows = []
        self.subprocesses = []
        self.settings_window = None
        sys.stdout = EmittingStream(textWritten=self.normalOutputWritten)
        sys.stderr = EmittingStream(textWritten=self.normalOutputWritten)

        self.setup_date()
        self.ui.target_date.dateChanged.connect(self.set_min_end_date)
        self.ui.end_date.dateChanged.connect(self.set_max_target_date)

        # default setting
        for b in [self.ui.button_First_5, self.ui.button_WaferID, self.ui.button_WaferID_2,
                  # self.ui.button_Final_yield,
                  self.ui.button_monthly_report_ref_last_year]:
            b.setChecked(True)
        self.lot_conditions_btn = [('wafer_id', self.ui.button_WaferID), ('first_5_char', self.ui.button_First_5),
                                   ('final_yield', self.ui.button_Final_yield)]
        self.monthly_lot_condition_btn = [('wafer_id', self.ui.button_WaferID_2)]

        for text_ui, btn in [(self.ui.columnstrTextEdit, self.ui.pushButton_query_lot),
                             # (self.ui.columnstrTextEdit_2, self.ui.pushButton_ignore_monthly_report),
                             ([self.ui.lineEdit_Description, self.ui.lineEdit_Description_2,
                               self.ui.lineEdit_2_Capability],
                              self.ui.pushButton_query_recipe_by_des_and_cap),
                             ]:
            self.set_ui_btn_connection(text_ui, btn)
        # button binding action
        self.ui.pushButton_weekly_update.clicked.connect(self.query_weekly_data)
        self.ui.pushButton_query_lot.clicked.connect(self.query_input_lot_list_to_get)
        self.ui.pushButton_ignore_monthly_report.clicked.connect(self.generate_modified_monthly_report)
        self.ui.pushButton_query_recipe_by_des_and_cap.clicked.connect(self.query_recipe_by_des_and_cap)
        # self.ui.pushButton_query_recipe_by_title_and_cap.clicked.connect(self.query_recipe_by_title_and_cap)
        self.ui.actionSettings.triggered.connect(self.settings_view)
        # update menu
        self.add_button_to_menu(self.ui.menuOpen, open_menu_file, func=self.open_file)
        self.add_button_to_menu(self.ui.menuUpdate_file, file_update_file_menu, func=self.initiate_funcion)
        self.add_button_to_menu(self.ui.menuConvert_jpg_to_png, jpg_to_png_menu, func=chk_jpg_and_convert_to_png)

    def __del__(self):
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__

    def normalOutputWritten(self, text):
        """Append text to the QTextEdit."""
        # Maybe QTextEdit.append() works as well, but this is how I do it:
        cursor = self.ui.console_output.textCursor()
        cursor.movePosition(QtGui.QTextCursor.End)
        cursor.insertText(text)
        self.ui.console_output.setTextCursor(cursor)
        self.ui.console_output.ensureCursorVisible()

    @staticmethod
    def add_button_to_menu(menu, add_btns, func):
        # for button_name, method_name, path_to_file in add_btns:
        for button_name, display_text, arg in add_btns:
            action = menu.addAction(button_name)
            action.setText(QtCore.QCoreApplication.translate("MainWindow", display_text))
            action.triggered.connect(lambda checked, f=func, a=arg: f(a))

    def _text_condition(self, text):
        content = str()
        if isinstance(text, QtWidgets.QPlainTextEdit):
            content = text.toPlainText()

        elif isinstance(text, QtWidgets.QLineEdit):
            content = text.text()
        elif isinstance(text, list):
            for t in text:
                content += self._text_condition(t)
        return content

    @staticmethod
    def query_input_lots(option, input_lot_condition, opt_args=None):
        query = InputLotListToGet(**input_lot_condition)
        opt_map = {'summary data': query.summary_data,
                   "wafers history search": query.open_wafer_history,
                   "lots history search": query.open_lots_history,
                   'cp reverse bias updated map': query.cp_reverse_bias_updated_map,
                   'max defect count': query.max_cat_count,
                   'inspection wafer map(s)': query.open_discrete_bin_map_selection,
                   'modify_monthly_data_and_report': query.modify_monthly_data_and_report,
                   'stacked wafer maps (multiple)': query.open_continuous_bin_map_selection,
                   "lot wafer part information": query.lot_wafer_part,
                   'cp value wafer map(s)': query.open_continuous_cp_value_map_selection}
        func = opt_map[option]
        return func() if opt_args is None else func(*opt_args)

    def set_query_button(self, text, button):
        """
        chk text/texts and disable or enable btn
        :param text: input or inputs(list of texts)
        :param button: disable or enable btn by text boolean value
        :return:
        """
        # content = self.ui.columnstrTextEdit.toPlainText()
        # button = self.ui.pushButton_query_lot
        # if isinstance(text, QtWidgets.QPlainTextEdit):
        #     content = text.toPlainText()
        #
        # elif isinstance(text, QtWidgets.QLineEdit):
        #     content = text.text()
        # elif isinstance(text, list):
        #     for t in text:
        content = self._text_condition(text)
        if content:
            button.setEnabled(True)
        else:
            button.setEnabled(False)

    def set_ui_btn_connection(self, text_ui, btn):
        if isinstance(text_ui, list):
            for t in text_ui:
                t.textChanged.connect(lambda: self.set_query_button(text_ui, btn))

        else:
            text_ui.textChanged.connect(lambda: self.set_query_button(text_ui, btn))
        btn.setEnabled(False)

    @staticmethod
    def on_work_stopped():
        # for btn1 in [self.ui.pushButton_weekly_update, self.ui.pushButton_query_lot,
        #              self.ui.pushButton_ignore_monthly_report]:
        #     btn1.setEnabled(True)
        # for text_edit in [self.ui.columnstrTextEdit, self.ui.columnstrTextEdit_2]:
        #     text_edit.setReadOnly(False)
        print('{0:=^70s}'.format('Finish!!!!!'))

    def setup_date(self):
        for d in {self.ui.target_date, self.ui.end_date, self.ui.ignore_target_month}:
            d.setDateTime(QtCore.QDateTime.currentDateTime())
            d.setMaximumDate(QtCore.QDate.currentDate())

    def set_max_target_date(self):
        self.ui.target_date.setMaximumDate(self.ui.end_date.date())

    def set_min_end_date(self):
        self.ui.end_date.setMinimumDate(self.ui.target_date.date())

    # button action ####################################################################################################
    def query_weekly_data(self):
        # self.worker.work_stopped_signal.connect(self.on_work_stopped)
        # self.ui.pushButton_weekly_update.setEnabled(False)
        self.worker.fetch(update_all_report_until,
                          [self.ui.target_date.date().toPyDate(), self.ui.end_date.date().toPyDate(),
                           self.ui.button_monthly_report_ref_last_year.isChecked()],
                          disable_ui={'buttons': [self.ui.pushButton_weekly_update]})

    def generate_modified_monthly_report(self):
        # self.worker.data_finished.connect(self.result_classifier)
        input_lot_condition = {key: btn.isChecked() for key, btn in self.monthly_lot_condition_btn}
        input_lot_condition['column_str'] = self.ui.columnstrTextEdit_2.toPlainText()
        # self.ui.pushButton_ignore_monthly_report.setEnabled(False)
        self.worker.fetch(self.query_input_lots, ['modify_monthly_data_and_report', input_lot_condition,
                                                  (self.ui.ignore_target_month.date().toPyDate(),
                                                   self.ui.button_monthly_report_ref_last_year.isChecked())],
                          disable_ui={'buttons': [self.ui.pushButton_ignore_monthly_report],
                                      'text_edits': [self.ui.columnstrTextEdit_2]})

    def query_input_lot_list_to_get(self):

        input_lot_condition = {key: btn.isChecked() for key, btn in self.lot_conditions_btn}
        input_lot_condition['column_str'] = self.ui.columnstrTextEdit.toPlainText()
        query_func = str(self.ui.query_func_comboBox.currentText())
        if 'wafer map' in query_func:
            input_lot_condition['wafer_map'] = True
        self.worker.fetch(self.query_input_lots, [query_func, input_lot_condition],
                          disable_ui={'buttons': [self.ui.pushButton_query_lot],
                                      'text_edits': [self.ui.columnstrTextEdit]})

    def query_recipe_by_des_and_cap(self):
        des = self.ui.lineEdit_Description.text()
        title = self.ui.lineEdit_Description_2.text()
        cap = self.ui.lineEdit_2_Capability.text()
        self.worker.fetch(recipe_query_by_title_des_cap, [title, des, cap],
                          disable_ui={'buttons': [self.ui.pushButton_query_recipe_by_des_and_cap],
                                      'text_edits': [self.ui.lineEdit_Description, self.ui.lineEdit_2_Capability
                                          , self.ui.lineEdit_Description_2]})

    # def query_recipe_by_title_and_cap(self):
    #     title = self.ui.lineEdit_Description_2.text()
    #     cap = self.ui.lineEdit_2_Capability_2.text()
    #     self.worker.fetch(recipe_query_by_title_and_cap, [title, cap],
    #                       {'buttons': [self.ui.pushButton_query_recipe_by_title_and_cap],
    #                        'text_edits': [self.ui.lineEdit_Description_2, self.ui.lineEdit_2_Capability_2]})

    # button action end ################################################################################################

    @QtCore.pyqtSlot(str)
    def open_file(self, path_to_file):
        self.worker.fetch(open_file, [path_to_file, self.subprocesses])

    @QtCore.pyqtSlot(dict)
    def single_open_wafer_map(self, result_dict):
        # self.worker.fetch(open_wafer_map, [result_dict, self.subprocesses])
        # self.worker.fetch(WaferMapApp, kwargs=result_dict)
        self.worker.fetch(open_wafer_map, [result_dict])

    @QtCore.pyqtSlot(dict)
    def multiple_open_wafer_maps(self, result_dict):

        self.worker.fetch(open_wafer_maps, [result_dict, self.subprocesses])

    def initiate_funcion(self, func):
        self.worker.fetch(func, [])

    def closeEvent(self, event):
        close = QtWidgets.QMessageBox.question(self,
                                               "QUIT",
                                               "Sure?",
                                               QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if close == QtWidgets.QMessageBox.Yes:
            con_close()
            QtWidgets.qApp.quit()
            event.accept()
            for sub in self.subprocesses:
                sub.terminate()
            for child in self.child_windows:
                child.close()

        else:
            event.ignore()

    def result_event_initiator(self, func, *func_args, **func_kwargs):
        func_event = func.__name__.replace('_', ' ')
        ok = QtWidgets.QMessageBox.question(self, func_event,
                                            f"Will {func_event}. Continue?",
                                            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if ok == 16384:
            self.worker.fetch(func, func_args, func_kwargs)
            # func(*func_args, **func_kwargs)

    def error_msg_window(self, error):
        window = QtWidgets.QMessageBox.warning(self, type(error).__name__,
                                               str(error),
                                               QtWidgets.QMessageBox.Ok)

    def finish_msg_window(self, msg):
        window = QtWidgets.QMessageBox.information(self, "Finish!!!!!",
                                                   msg,
                                                   QtWidgets.QMessageBox.Ok)

    def settings_view(self):
        self.settings_window = SettingsViewTree(value=display_cfg)
        self.settings_window.show()

    def result_classifier(self, result):
        val = result.val
        if val is not None and isinstance(val, tuple):
            purpose = val[0]
            result_data = val[1]
            if 'open table viewer' in purpose:
                df = result_data['df']
                title = result_data['title']
                child_window = TableView(title=title, df=df)
                self.child_windows.append(child_window)
            elif 'bin map selection' in purpose:
                # return by InputLotListToGet.open_discrete_bin_map_selection
                df = result_data['df']
                query_func = result_data['query_func']
                title = result_data['title']
                child_window = BinSelectionView(query_func, self.worker, df=df, title=title)
                self.child_windows.append(child_window)
            elif purpose == 'cp item map selection':
                query_func = result_data['query_func']

                dialog = InputDialog(result_data['title'],
                                     list_items={"CP item:": [str(n) for n in range(31)],
                                                 "CP stage:": [f'CP{n}' for n in range(1, 6)],
                                                 "Export to excel simultaneously:": ["No", "Yes"],
                                                 "Open wafer map window:": ["Yes", "No"],
                                                 "USL percentile": [str(n) for n in range(99, 50, -1)],
                                                 "LSL percentile": [str(n) for n in range(1, 50)],
                                                 "Use DPAT to find outlier": ["No", "Yes"]
                                                 },
                                     parent=self)
                if dialog.exec():
                    args = dialog.getInputs()
                    self.worker.fetch(query_func, args=tuple(args))
            elif 'open wafer map' in purpose:
                result_list = result_data
                window_num = len(result_list)
                if 'to excel' in purpose:

                    options = QtWidgets.QFileDialog.Options()
                    options |= QtWidgets.QFileDialog.DontUseNativeDialog
                    filename, _ = QtWidgets.QFileDialog.getSaveFileName(self, 'Save file', '',
                                                                        filter='*.xlsx',
                                                                        options=options)

                    if filename:
                        for i, result_dict in enumerate(result_list):
                            title = result_dict['title']
                            wafer = title.split(" ")[-1]
                            df_map = result_dict['wafer_detail']['df_map']
                            die_size = result_dict['die_size']
                            plot_range = result_dict['plot_range']
                            filename = export_to_excel(df_map, None, filename, die_size, return_writer=True,
                                                       wafer=wafer, plot_range=plot_range)
                        else:
                            filename.save()
                        self.result_event_initiator(self.open_file, filename.path)
                if 'no window' not in purpose:

                    for i, result_dict in enumerate(result_list):
                        if window_num > 1:
                            self.multiple_open_wafer_maps(result_dict)
                        else:
                            # self.multiple_open_wafer_maps(result_dict)
                            self.single_open_wafer_map(result_dict)

            elif purpose == 'open file':
                filename = result_data
                self.result_event_initiator(self.open_file, filename)
            elif purpose == 'error msg':
                error_msg = result_data
                self.error_msg_window(error_msg)
            elif purpose == 'finish':
                finish_msg = result_data
                self.finish_msg_window(finish_msg)


def set_dark_theme(app):
    app.setStyle('Fusion')
    palette = QtGui.QPalette()
    palette.setColor(QtGui.QPalette.Window, QtGui.QColor(53, 53, 53))
    palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.white)
    palette.setColor(QtGui.QPalette.Base, QtGui.QColor(15, 15, 15))
    palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(53, 53, 53))
    palette.setColor(QtGui.QPalette.ToolTipBase, QtCore.Qt.white)
    palette.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.white)
    palette.setColor(QtGui.QPalette.Text, QtCore.Qt.white)
    palette.setColor(QtGui.QPalette.Button, QtGui.QColor(53, 53, 53))
    palette.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.white)
    palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
    palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(142, 45, 197).lighter())
    palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.black)
    app.setPalette(palette)


def main():
    sys._excepthook = sys.excepthook

    def exception_hook(exctype, value, traceback):
        print(exctype, value, traceback)
        sys._excepthook(exctype, value, traceback)
        sys.exit(1)

    sys.excepthook = exception_hook
    app = QtWidgets.QApplication(sys.argv)
    app_icon = QtGui.QIcon('icon.png')
    app.setWindowIcon(app_icon)
    set_dark_theme(app)
    w = MainWindow()
    w.show()
    # sys.exit(app.exec_())
    try:
        sys.exit(app.exec_())
    except:
        print("Exiting")
    app.deleteLater()  # avoids some QThread messages in the shell on exit
    # cancel all running tasks avoid QThread/QTimer error messages
    # on exit
    Request.shutdown()


if __name__ == "__main__":
    main()
