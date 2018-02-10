# -*- coding: utf-8 -*-

import os
import sys
import time
import random
import chardet

from draw_ui import Ui_MainWindow
from PyQt5 import QtWidgets, QtCore, QtGui

TIME_INTERVAL = 0.01


def choice_one(list_):
    return random.choice(list_)


def get_text_list(file):
    with open(file, 'rb') as f:
        bdata = f.read()
        encode = chardet.detect(bdata)['encoding']
        return bdata.decode(encoding=encode).splitlines()


def get_execl_list(excel):
    import xlrd
    return xlrd.open_workbook(excel).sheet_by_index(0).col_values(0)


def get_file_list(file):
    return get_execl_list(file) if os.path.splitext(file)[1] in ['.xls', '.xlsx'] else get_text_list(file)


def save_file(list_, file):
    if os.path.exists(file):
        os.remove(file)
    with open(file, 'a') as f:
        for l in list_:
            f.write(l + '\n')


class MultiThread(QtCore.QThread):
    sig = QtCore.pyqtSignal()

    def __init__(self, list_, widget):
        super(MultiThread, self).__init__()
        self.list_ = list_
        self.widget = widget
        self.running_flag = True

    def run(self):
        self._roll_list(self.list_, self.widget)

    def stop(self):
        self.running_flag = False
        # self._stop_roll(self.list_, self.widget)

    def _stop_roll(self, list_, widget):
        t = TIME_INTERVAL
        while list_ and t < 0.5:
            t = t * 1.5
            time.sleep(t)
            widget.setText(random.choice(list_))
        self.sig.emit()

    def _roll_list(self, list_, widget):
        while list_ and self.running_flag:
            widget.setText(random.choice(list_))
            time.sleep(TIME_INTERVAL)
        self._stop_roll(self.list_, self.widget)


class Draw(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super(Draw, self).__init__()
        self.setupUi(self)
        self.setWindowTitle("抽抽乐")
        self.start_list = []
        self.status = 'Start'
        self._int_connect()
        self.thread = None
        self._reset_color(self.roll_label)

    def _int_connect(self):
        self.open_file_action.triggered.connect(self.choice_file_list)
        self.start_button.clicked.connect(self.start_clicked)
        self.next_round.clicked.connect(self.next_round_clicked)

    def choice_file_list(self):
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(self, u'选择员工列表')
        if filename:
            self.start_list = get_file_list(filename)
            self._set_remain_count()

    def start_clicked(self):
        self._start_clicked() if self.status == 'Start' else self._stop_clicked()
        self.open_file_action.setEnabled(False)
        self._save_unselected()
        return self._start_to_stop()

    def _start_clicked(self):
        self._reset_color(self.roll_label)
        if not self.start_list:
            self.choice_file_list()
        if self.start_list:
            self.thread = MultiThread(self.start_list, self.roll_label)
            self.thread.sig.connect(self._roll_stop)
            self.thread.start()

    def _roll_stop(self):
        self._handle_select(self.roll_label.text())
        self._set_color(self.roll_label)
        self.start_button.setEnabled(True)
        self.open_file_action.setEnabled(True)
        self.thread.sig.disconnect()

    def _stop_clicked(self):
        self.thread.stop()
        self.start_button.setEnabled(False if self.start_list else True)

    def next_round_clicked(self):
        self._clear_selected()

    # no use
    def _roll_list(self, list_, widget):
        while self.status == "FFF":
            widget.setText(random.choice(list_))
            time.sleep(1)
        self._handle_select(widget.text())

    @staticmethod
    def _set_color(widget):
        widget.setFont(QtGui.QFont("Microsoft YaHei UI", 30, QtGui.QFont.Bold))
        widget.setStyleSheet("QLabel {color : red}")

    @staticmethod
    def _reset_color(widget):
        widget.setFont(QtGui.QFont("Microsoft YaHei UI", 20))
        widget.setStyleSheet("QLabel {color : black}")

    def _handle_select(self, text):
        self.selected_list.addItem(text)
        self._remove_element(text, self.start_list)
        self._set_remain_count()

    @staticmethod
    def _remove_element(text, list_):
        list_.remove(text)

    def _clear_selected(self):
        self.selected_list.clear()
        self._reset_color(self.roll_label)
        self.roll_label.setText("???")
        self.start_button.setEnabled(True)
        self.open_file_action.setEnabled(True)

    def _set_remain_count(self):
        self.remain_label.setText("剩余人数：" + str(len(self.start_list)))

    def _start_to_stop(self):
        if self.status == 'Start':
            self._set_text(self.start_button, "停止")
            self.status = 'Stop'
        else:
            self._set_text(self.start_button, "开始")
            self.status = 'Start'

    @staticmethod
    def _set_text(widget, text):
        widget.setText(text)

    def _save_unselected(self):
        return save_file(self.start_list, 'unselect')


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = Draw()
    window.show()
    sys.exit(app.exec_())



