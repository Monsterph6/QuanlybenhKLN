# -*- coding: utf-8 -*-
"""
Ung dung quan ly & loc trung danh sach benh nhan Benh khong lay nhiem (KLN:
THA, dai thao duong, COPD/hen phe quan, ung thu...) - giao dien PyQt6.
Tang du lieu (SQLite, doc Excel, xuat CSV/Excel) nam trong core.py.
"""
import datetime
import os
import subprocess
import sys
import sqlite3

from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtCharts import (
    QChart, QChartView, QPieSeries, QHorizontalBarSeries, QBarSet,
    QBarCategoryAxis, QValueAxis,
)

try:
    import win32service
    import win32serviceutil
    HAS_WIN32_SERVICE = True
except ImportError:
    HAS_WIN32_SERVICE = False

import core
import netclient
import netserver
from core import (
    BASE_DIR, DB_PATH, COLUMNS, DEDUP_FIELDS, PAGE_SIZE,
    get_conn, init_db, record_count,
    MAPPING_FIELDS, list_sheet_names, detect_sheet_mapping, import_excel_mapped, at,
    write_export, export_query_to_file,
    count_swapped_gender_birthdate, fix_swapped_gender_birthdate,
    count_unparsed_kham_dates, fix_unparsed_kham_dates,
    count_unclassified_diseases, reclassify_diseases,
    DISEASE_CATEGORIES, DISEASE_OTHER_LABEL,
    build_dedup_key, scan_dedup_groups, group_detail_rows,
    delete_patients_by_ids,
    merge_specific_ids, merge_group, merge_all_groups,
    unique_rows_with_optional_history,
    add_dedup_exception, remove_dedup_exception, list_dedup_exceptions,
    backup_database, data_quality_summary, data_quality_rows_sql,
    has_password, set_password, verify_password, remove_password,
    get_local_version, check_latest_release,
    load_lan_config, save_lan_config,
    STATS_COLUMNS, stats_top_values, stats_birth_decade, stats_disease_counts,
)


# Ten Windows Service dung o vai tro May chu - phai khop dung
# service.QuanLyBenhNhanTHAService._svc_name_. Khai bao lai truc tiep (thay
# vi import service.py) de tranh keo theo pywin32.ServiceFramework ngay tu
# dau chuong trinh chinh - chi ServerTab (o duoi) moi thuc su can pywin32.
SERVER_SERVICE_NAME = "QuanLyBenhNhanTHA_Server"


def restart_app():
    """Khoi dong lai ung dung (dung sau khi doi cau hinh mang LAN de ap
    dung che do may chu / may tram tu dau, tranh phai xu ly hot-swap giua
    chung)."""
    os.execv(sys.executable, [sys.executable] + sys.argv[1:])

QT_EXPORT_FILTER = "Excel (*.xlsx);;CSV (*.csv)"


# ------------------------------------------------------------------
# Model bang du lieu dung chung
# ------------------------------------------------------------------

class TableModel(QtCore.QAbstractTableModel):
    def __init__(self, headers=None, rows=None):
        super().__init__()
        self._headers = headers or []
        self._rows = rows or []

    def rowCount(self, parent=QtCore.QModelIndex()):
        return len(self._rows)

    def columnCount(self, parent=QtCore.QModelIndex()):
        return len(self._headers)

    def data(self, index, role=QtCore.Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        if role == QtCore.Qt.ItemDataRole.DisplayRole:
            val = self._rows[index.row()][index.column()]
            return "" if val is None else str(val)
        return None

    def headerData(self, section, orientation, role=QtCore.Qt.ItemDataRole.DisplayRole):
        if role != QtCore.Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == QtCore.Qt.Orientation.Horizontal:
            return self._headers[section]
        return str(section + 1)

    def set_data(self, headers, rows):
        self.beginResetModel()
        self._headers = headers
        self._rows = rows
        self.endResetModel()

    def row_values(self, row_idx):
        return self._rows[row_idx]


class CheckableTableModel(QtCore.QAbstractTableModel):
    """Nhu TableModel nhung co them cot dau tien la checkbox de chon nhieu dong."""

    def __init__(self, headers=None, rows=None):
        super().__init__()
        self._headers = headers or []
        self._rows = rows or []
        self._checked = set()

    def rowCount(self, parent=QtCore.QModelIndex()):
        return len(self._rows)

    def columnCount(self, parent=QtCore.QModelIndex()):
        return len(self._headers) + 1

    def data(self, index, role=QtCore.Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        col = index.column()
        if col == 0:
            if role == QtCore.Qt.ItemDataRole.CheckStateRole:
                checked = index.row() in self._checked
                return QtCore.Qt.CheckState.Checked if checked else QtCore.Qt.CheckState.Unchecked
            return None
        if role == QtCore.Qt.ItemDataRole.DisplayRole:
            val = self._rows[index.row()][col - 1]
            return "" if val is None else str(val)
        return None

    def setData(self, index, value, role=QtCore.Qt.ItemDataRole.EditRole):
        if index.column() == 0 and role == QtCore.Qt.ItemDataRole.CheckStateRole:
            if value == QtCore.Qt.CheckState.Checked.value:
                self._checked.add(index.row())
            else:
                self._checked.discard(index.row())
            self.dataChanged.emit(index, index, [role])
            return True
        return False

    def flags(self, index):
        base = super().flags(index)
        if index.column() == 0:
            return base | QtCore.Qt.ItemFlag.ItemIsUserCheckable
        return base

    def headerData(self, section, orientation, role=QtCore.Qt.ItemDataRole.DisplayRole):
        if role != QtCore.Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == QtCore.Qt.Orientation.Horizontal:
            return "Chọn" if section == 0 else self._headers[section - 1]
        return str(section + 1)

    def set_data(self, headers, rows):
        self.beginResetModel()
        self._headers = headers
        self._rows = rows
        self._checked = set()
        self.endResetModel()

    def row_values(self, row_idx):
        return self._rows[row_idx]

    def checked_rows(self):
        return [self._rows[i] for i in sorted(self._checked)]

    def check_all(self, checked=True):
        self.beginResetModel()
        self._checked = set(range(len(self._rows))) if checked else set()
        self.endResetModel()


def make_table_view():
    view = QtWidgets.QTableView()
    view.setAlternatingRowColors(True)
    view.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
    view.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
    view.horizontalHeader().setStretchLastSection(True)
    view.horizontalHeader().setDefaultSectionSize(140)
    view.verticalHeader().setVisible(False)
    view.setSortingEnabled(False)
    return view


# ------------------------------------------------------------------
# Tab Nhap du lieu
# ------------------------------------------------------------------

def _col_letter(n):
    """Chi so cot 1-based -> ky hieu cot kieu Excel (1->A, 27->AA...)."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class ImportMappingDialog(QtWidgets.QDialog):
    """Hop thoai anh xa cot khi nhap Excel: ung dung tu doan truoc (theo ten
    cot, xem core.detect_sheet_mapping) - nguoi dung xem lai bang xem truoc,
    sua o nao bi sai/con trong ("(Không dùng)") roi moi bam nhap. File co
    nhieu sheet bo cuc khac nhau (thuong gap voi bao cao BKLN thuc te) thi
    doi sheet trong cung hop thoai nay va lap lai cho tung sheet can nhap -
    khong bat buoc phai nhap het moi sheet (vd bo qua sheet tong hop)."""

    def __init__(self, parent, path):
        super().__init__(parent)
        self.setWindowTitle("Ánh xạ cột khi nhập Excel")
        self.resize(900, 700)
        self.path = path
        self.sheet_names = list_sheet_names(path)
        self.current_sheet = None
        self.header_row = 1
        self.columns = {}
        self.field_combos = {}
        self.gender_male_combo = None
        self.gender_female_combo = None
        self.total_read = 0
        self.total_inserted = 0
        self.total_skipped = 0
        self._loading = False

        layout = QtWidgets.QVBoxLayout(self)

        top_row = QtWidgets.QHBoxLayout()
        top_row.addWidget(QtWidgets.QLabel("Sheet:"))
        self.sheet_combo = QtWidgets.QComboBox()
        self.sheet_combo.addItems(self.sheet_names)
        self.sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)
        top_row.addWidget(self.sheet_combo, stretch=1)
        top_row.addWidget(QtWidgets.QLabel("Dòng tiêu đề:"))
        self.header_row_spin = QtWidgets.QSpinBox()
        self.header_row_spin.setRange(1, 500)
        self.header_row_spin.valueChanged.connect(self.on_header_row_changed)
        top_row.addWidget(self.header_row_spin)
        layout.addLayout(top_row)

        note = QtWidgets.QLabel(
            "Ứng dụng đã tự đoán dòng tiêu đề và cột tương ứng cho từng trường bên dưới - "
            "kiểm tra bảng xem trước, sửa lại ô nào bị sai hoặc còn để \"(Không dùng)\" cho "
            "đúng, rồi bấm \"Nhập dữ liệu từ sheet này\". Nếu file có nhiều sheet bố cục khác "
            "nhau (ví dụ mỗi mặt bệnh 1 sheet riêng), đổi Sheet ở trên và lặp lại - sheet nào "
            "không phải danh sách bệnh nhân (ví dụ sheet tổng hợp số liệu) thì bỏ qua không cần nhập.")
        note.setWordWrap(True)
        layout.addWidget(note)

        layout.addWidget(QtWidgets.QLabel("Xem trước (5 dòng dữ liệu đầu theo dòng tiêu đề hiện tại):"))
        self.preview_table = QtWidgets.QTableWidget(0, 0)
        self.preview_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.preview_table.setMaximumHeight(140)
        layout.addWidget(self.preview_table)

        self.gender_split_check = QtWidgets.QCheckBox(
            "Giới tính & Năm sinh/Tuổi nằm trong 2 cột tách riêng (1 cột cho Nam, 1 cột cho Nữ)")
        self.gender_split_check.toggled.connect(self.on_gender_split_toggled)
        layout.addWidget(self.gender_split_check)

        self.mapping_form = QtWidgets.QFormLayout()
        mapping_widget = QtWidgets.QWidget()
        mapping_widget.setLayout(self.mapping_form)
        scroll = QtWidgets.QScrollArea()
        scroll.setWidget(mapping_widget)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll, stretch=1)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(90)
        layout.addWidget(self.log)

        btn_row = QtWidgets.QHBoxLayout()
        import_btn = QtWidgets.QPushButton("Nhập dữ liệu từ sheet này")
        import_btn.setObjectName("PrimaryButton")
        import_btn.clicked.connect(self.import_current_sheet)
        btn_row.addWidget(import_btn)
        btn_row.addStretch(1)
        close_btn = QtWidgets.QPushButton("Đóng")
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        self.load_sheet(self.sheet_names[0])

    def on_sheet_changed(self, idx):
        if idx < 0 or self._loading:
            return
        self.load_sheet(self.sheet_names[idx])

    def load_sheet(self, sheet_name, header_row=None):
        self._loading = True
        try:
            info = detect_sheet_mapping(self.path, sheet_name, header_row=header_row)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi", f"Không đọc được sheet \"{sheet_name}\":\n{e}")
            self._loading = False
            return
        self.current_sheet = sheet_name
        self.header_row = info["header_row"]
        self.columns = info["columns"]
        self.header_row_spin.blockSignals(True)
        self.header_row_spin.setValue(self.header_row)
        self.header_row_spin.blockSignals(False)
        self._rebuild_preview(info["preview_rows"])
        self._apply_mapping(info["mapping"], info["gender_split_cols"])
        self._loading = False

    def on_header_row_changed(self, value):
        if self._loading:
            return
        self.load_sheet(self.current_sheet, header_row=value)

    def on_gender_split_toggled(self, _checked):
        if self._loading:
            return
        self._rebuild_mapping_form()

    def _rebuild_preview(self, preview_rows):
        cols = sorted(self.columns.keys())
        self.preview_table.clear()
        self.preview_table.setColumnCount(len(cols))
        self.preview_table.setHorizontalHeaderLabels(
            [f"{_col_letter(c)}: {self.columns[c]}" for c in cols])
        self.preview_table.setRowCount(len(preview_rows))
        for r, row in enumerate(preview_rows):
            for ci, c in enumerate(cols):
                val = at(row, c)
                self.preview_table.setItem(r, ci, QtWidgets.QTableWidgetItem("" if val is None else str(val)))

    def _make_column_combo(self):
        combo = QtWidgets.QComboBox()
        combo.addItem("(Không dùng)", None)
        for c in sorted(self.columns.keys()):
            combo.addItem(f"Cột {_col_letter(c)}: {self.columns[c]}", c)
        return combo

    def _rebuild_mapping_form(self):
        while self.mapping_form.rowCount() > 0:
            self.mapping_form.removeRow(0)
        self.field_combos = {}
        self.gender_male_combo = None
        self.gender_female_combo = None

        split_on = self.gender_split_check.isChecked()
        for field, label, required in MAPPING_FIELDS:
            if split_on and field in ("gioi_tinh", "nam_sinh_raw"):
                continue
            combo = self._make_column_combo()
            self.field_combos[field] = combo
            self.mapping_form.addRow(label + (" *" if required else ""), combo)
        if split_on:
            self.gender_male_combo = self._make_column_combo()
            self.gender_female_combo = self._make_column_combo()
            self.mapping_form.addRow("Cột Nam (năm sinh/tuổi)", self.gender_male_combo)
            self.mapping_form.addRow("Cột Nữ (năm sinh/tuổi)", self.gender_female_combo)

    def _apply_mapping(self, mapping, gender_split_cols):
        self.gender_split_check.blockSignals(True)
        self.gender_split_check.setChecked(bool(gender_split_cols))
        self.gender_split_check.blockSignals(False)
        self._rebuild_mapping_form()
        for field, combo in self.field_combos.items():
            idx = combo.findData(mapping.get(field))
            combo.setCurrentIndex(idx if idx >= 0 else 0)
        if gender_split_cols:
            male_col, female_col = gender_split_cols
            self.gender_male_combo.setCurrentIndex(max(self.gender_male_combo.findData(male_col), 0))
            self.gender_female_combo.setCurrentIndex(max(self.gender_female_combo.findData(female_col), 0))

    def import_current_sheet(self):
        mapping = {field: combo.currentData() for field, combo in self.field_combos.items()}
        gender_split_cols = None
        if self.gender_split_check.isChecked():
            male_col = self.gender_male_combo.currentData()
            female_col = self.gender_female_combo.currentData()
            if not male_col or not female_col:
                QtWidgets.QMessageBox.warning(
                    self, "Thiếu cột",
                    "Đã bật tách Giới tính theo 2 cột Nam/Nữ - cần chọn cả 2 cột.")
                return
            gender_split_cols = (male_col, female_col)

        if not mapping.get("ho_ten"):
            QtWidgets.QMessageBox.warning(
                self, "Thiếu cột bắt buộc",
                "Cần chọn cột cho trường \"Họ và tên\" trước khi nhập.")
            return

        self.setCursor(QtCore.Qt.CursorShape.WaitCursor)
        try:
            total, inserted, skipped = import_excel_mapped(
                self.path, self.current_sheet, self.header_row, mapping, gender_split_cols)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi khi nhập dữ liệu", str(e))
            return
        finally:
            self.unsetCursor()

        self.total_read += total
        self.total_inserted += inserted
        self.total_skipped += skipped
        self.log.appendPlainText(
            f"Sheet \"{self.current_sheet}\": đọc {total:,} dòng, thêm mới {inserted:,}, "
            f"bỏ qua (trùng) {skipped:,}.")


class LoginDialog(QtWidgets.QDialog):
    """Man hinh dang nhap khi khoi dong app, chi hien neu da dat mat khau."""
    MAX_ATTEMPTS = 5

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Đăng nhập")
        self.setModal(True)
        self.attempts = 0

        layout = QtWidgets.QVBoxLayout(self)
        layout.addWidget(QtWidgets.QLabel("Nhập mật khẩu để mở ứng dụng:"))
        self.pw_edit = QtWidgets.QLineEdit()
        self.pw_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
        self.pw_edit.returnPressed.connect(self.try_login)
        layout.addWidget(self.pw_edit)

        self.error_label = QtWidgets.QLabel("")
        self.error_label.setStyleSheet("color: #c92a2a;")
        layout.addWidget(self.error_label)

        ok_btn = QtWidgets.QPushButton("Đăng nhập")
        ok_btn.setObjectName("PrimaryButton")
        ok_btn.clicked.connect(self.try_login)
        layout.addWidget(ok_btn)

    def try_login(self):
        if verify_password(self.pw_edit.text()):
            self.accept()
            return
        self.attempts += 1
        remaining = self.MAX_ATTEMPTS - self.attempts
        if remaining <= 0:
            QtWidgets.QMessageBox.critical(
                self, "Đăng nhập thất bại",
                "Sai mật khẩu quá số lần cho phép. Ứng dụng sẽ đóng.")
            self.reject()
            return
        self.error_label.setText(f"Sai mật khẩu. Còn {remaining} lần thử.")
        self.pw_edit.clear()
        self.pw_edit.setFocus()


class PasswordManageDialog(QtWidgets.QDialog):
    """Dat / doi / tat mat khau bao ve ung dung. Chi khoa giao dien - KHONG
    ma hoa file benh_nhan.db (ai co file van mo duoc bang cong cu SQLite khac)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Quản lý mật khẩu ứng dụng")
        self.resize(400, 240)
        layout = QtWidgets.QVBoxLayout(self)

        if has_password():
            layout.addWidget(QtWidgets.QLabel("Nhập mật khẩu hiện tại để tiếp tục:"))
            self.current_edit = QtWidgets.QLineEdit()
            self.current_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
            layout.addWidget(self.current_edit)
            unlock_btn = QtWidgets.QPushButton("Xác nhận")
            unlock_btn.clicked.connect(self.unlock)
            layout.addWidget(unlock_btn)

            self.action_area = QtWidgets.QWidget()
            self.action_area.setVisible(False)
            action_layout = QtWidgets.QVBoxLayout(self.action_area)
            action_layout.setContentsMargins(0, 8, 0, 0)
            action_layout.addWidget(QtWidgets.QLabel("Mật khẩu mới:"))
            self.new_edit = QtWidgets.QLineEdit()
            self.new_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
            action_layout.addWidget(self.new_edit)
            self.confirm_edit = QtWidgets.QLineEdit()
            self.confirm_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
            self.confirm_edit.setPlaceholderText("Nhập lại mật khẩu mới")
            action_layout.addWidget(self.confirm_edit)
            abtns = QtWidgets.QHBoxLayout()
            save_btn = QtWidgets.QPushButton("Lưu mật khẩu mới")
            save_btn.setObjectName("PrimaryButton")
            save_btn.clicked.connect(self.save_new_password)
            abtns.addWidget(save_btn)
            remove_btn = QtWidgets.QPushButton("Tắt mật khẩu")
            remove_btn.setObjectName("DangerButton")
            remove_btn.clicked.connect(self.disable_password)
            abtns.addWidget(remove_btn)
            action_layout.addLayout(abtns)
            layout.addWidget(self.action_area)
        else:
            note = QtWidgets.QLabel(
                "Đặt mật khẩu để yêu cầu nhập mật khẩu mỗi khi mở ứng dụng.\n"
                "Lưu ý: đây chỉ là khóa giao diện, KHÔNG mã hóa file benh_nhan.db.")
            note.setWordWrap(True)
            layout.addWidget(note)
            self.new_edit = QtWidgets.QLineEdit()
            self.new_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
            self.new_edit.setPlaceholderText("Mật khẩu mới")
            layout.addWidget(self.new_edit)
            self.confirm_edit = QtWidgets.QLineEdit()
            self.confirm_edit.setEchoMode(QtWidgets.QLineEdit.EchoMode.Password)
            self.confirm_edit.setPlaceholderText("Nhập lại mật khẩu mới")
            layout.addWidget(self.confirm_edit)
            save_btn = QtWidgets.QPushButton("Đặt mật khẩu")
            save_btn.setObjectName("PrimaryButton")
            save_btn.clicked.connect(self.save_new_password)
            layout.addWidget(save_btn)

        layout.addStretch(1)
        close_btn = QtWidgets.QPushButton("Đóng")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

    def unlock(self):
        if verify_password(self.current_edit.text()):
            self.action_area.setVisible(True)
            self.current_edit.setEnabled(False)
        else:
            QtWidgets.QMessageBox.critical(self, "Sai mật khẩu", "Mật khẩu hiện tại không đúng.")

    def save_new_password(self):
        pw = self.new_edit.text()
        confirm = self.confirm_edit.text()
        if not pw:
            QtWidgets.QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng nhập mật khẩu mới.")
            return
        if pw != confirm:
            QtWidgets.QMessageBox.warning(self, "Không khớp", "Mật khẩu nhập lại không khớp.")
            return
        set_password(pw)
        QtWidgets.QMessageBox.information(self, "Thành công", "Đã lưu mật khẩu mới.")
        self.accept()

    def disable_password(self):
        reply = QtWidgets.QMessageBox.question(
            self, "Xác nhận", "Tắt mật khẩu bảo vệ ứng dụng?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        remove_password()
        QtWidgets.QMessageBox.information(self, "Đã tắt", "Đã tắt mật khẩu bảo vệ ứng dụng.")
        self.accept()


class ImportTab(QtWidgets.QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(12)

        title = QtWidgets.QLabel("Bước 1: Chọn file Excel danh sách bệnh nhân")
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        row = QtWidgets.QHBoxLayout()
        self.path_edit = QtWidgets.QLineEdit()
        self.path_edit.setPlaceholderText("Chưa chọn file...")
        browse_btn = QtWidgets.QPushButton("Chọn file...")
        browse_btn.clicked.connect(self.browse)
        row.addWidget(self.path_edit)
        row.addWidget(browse_btn)
        layout.addLayout(row)

        self.import_btn = QtWidgets.QPushButton("Nhập vào cơ sở dữ liệu")
        self.import_btn.setObjectName("PrimaryButton")
        self.import_btn.clicked.connect(self.start_import)
        layout.addWidget(self.import_btn, alignment=QtCore.Qt.AlignmentFlag.AlignLeft)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log, stretch=1)

        quality_row = QtWidgets.QHBoxLayout()
        quality_btn = QtWidgets.QPushButton("Xuất báo cáo chất lượng dữ liệu (Excel/CSV)")
        quality_btn.clicked.connect(self.export_quality_report)
        quality_row.addWidget(quality_btn)
        quality_row.addStretch(1)
        layout.addLayout(quality_row)

        sep = QtWidgets.QFrame()
        sep.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        layout.addWidget(sep)

        row2 = QtWidgets.QHBoxLayout()
        reset_btn = QtWidgets.QPushButton("Xóa toàn bộ dữ liệu trong CSDL")
        reset_btn.setObjectName("DangerButton")
        reset_btn.clicked.connect(self.reset_db)
        row2.addWidget(reset_btn)

        fix_btn = QtWidgets.QPushButton("Sửa lỗi đảo cột Giới tính / Ngày sinh")
        fix_btn.clicked.connect(self.fix_swapped)
        row2.addWidget(fix_btn)

        fix_kham_btn = QtWidgets.QPushButton("Sửa lỗi định dạng Ngày khám bị bỏ sót")
        fix_kham_btn.clicked.connect(self.fix_kham_dates)
        row2.addWidget(fix_kham_btn)

        reclassify_btn = QtWidgets.QPushButton("Xác định lại Nhóm bệnh (KLN)")
        reclassify_btn.clicked.connect(self.reclassify_diseases_click)
        row2.addWidget(reclassify_btn)
        row2.addStretch(1)
        layout.addLayout(row2)

        row3 = QtWidgets.QHBoxLayout()
        backup_btn = QtWidgets.QPushButton("Sao lưu CSDL ngay")
        backup_btn.clicked.connect(self.backup_now)
        row3.addWidget(backup_btn)

        open_backup_btn = QtWidgets.QPushButton("Mở thư mục sao lưu")
        open_backup_btn.clicked.connect(self.open_backup_folder)
        row3.addWidget(open_backup_btn)
        row3.addStretch(1)
        layout.addLayout(row3)

        row4 = QtWidgets.QHBoxLayout()
        self.password_btn = QtWidgets.QPushButton()
        self.password_btn.clicked.connect(self.manage_password)
        row4.addWidget(self.password_btn)
        row4.addStretch(1)
        layout.addLayout(row4)
        self._refresh_password_button()

        default_xlsx = self._find_default_xlsx()
        if default_xlsx:
            self.path_edit.setText(default_xlsx)

    def _find_default_xlsx(self):
        for f in os.listdir(BASE_DIR):
            if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$"):
                return os.path.join(BASE_DIR, f)
        return None

    def browse(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Chọn file Excel", BASE_DIR, "Excel files (*.xlsx *.xls)")
        if path:
            self.path_edit.setText(path)

    def log_line(self, text):
        self.log.appendPlainText(text)

    def start_import(self):
        path = self.path_edit.text().strip()
        if not path or not os.path.isfile(path):
            QtWidgets.QMessageBox.critical(self, "Lỗi", "Vui lòng chọn một file Excel hợp lệ.")
            return
        try:
            dialog = ImportMappingDialog(self, path)
        except Exception as e:
            self.log_line(f"LỖI: {e}")
            QtWidgets.QMessageBox.critical(self, "Lỗi khi đọc file Excel", str(e))
            return
        dialog.exec()
        if dialog.total_read:
            self.import_done(dialog.total_read, dialog.total_inserted, dialog.total_skipped)

    def import_done(self, total, inserted, skipped):
        self.log_line(f"Đã đọc: {total:,} dòng")
        self.log_line(f"Thêm mới vào CSDL: {inserted:,} bản ghi")
        self.log_line(f"Bỏ qua (dòng dữ liệu giống hệt đã có, tránh nhập trùng khi nhập lại file): {skipped:,}")
        self.log_line("Lưu ý: đây chỉ là bỏ qua dòng nhập lại y hệt. Trùng bệnh nhân (nhiều lượt khám) "
                       "vẫn được giữ nguyên — dùng tab 'Lọc trùng' để xử lý.")

        summary = data_quality_summary()
        issues = [(label, n) for label, n in summary if n]
        if issues:
            self.log_line("Báo cáo chất lượng dữ liệu:")
            for label, n in issues:
                self.log_line(f"  - {label}: {n:,} dòng")
            self.log_line("Dùng nút 'Xuất báo cáo chất lượng dữ liệu' bên dưới để xem chi tiết từng dòng.")
        else:
            self.log_line("Không phát hiện vấn đề chất lượng dữ liệu nào.")
        self.log_line("-" * 60)
        self.main_window.on_data_changed()

    def reset_db(self):
        reply = QtWidgets.QMessageBox.question(
            self, "Xác nhận",
            "Xóa TOÀN BỘ dữ liệu hiện có trong cơ sở dữ liệu?\n"
            "Một bản sao lưu sẽ được tạo tự động trước khi xóa.",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        backup_path = backup_database(reason="truoc_khi_xoa_toan_bo")
        conn = get_conn()
        conn.execute("DELETE FROM patients")
        conn.commit()
        conn.close()
        self.log_line(f"Đã sao lưu vào: {backup_path}")
        self.log_line("Đã xóa toàn bộ dữ liệu trong CSDL.")
        self.main_window.on_data_changed()

    def fix_swapped(self):
        n = count_swapped_gender_birthdate()
        if n == 0:
            QtWidgets.QMessageBox.information(
                self, "Không có gì để sửa",
                "Không tìm thấy dòng nào bị đảo cột Giới tính / Ngày sinh.")
            return
        reply = QtWidgets.QMessageBox.question(
            self, "Xác nhận sửa dữ liệu",
            f"Phát hiện {n:,} dòng có cột Giới tính và Ngày sinh bị đảo chỗ cho nhau "
            "(lỗi từ file Excel nguồn, ví dụ Giới tính ghi '1958' còn Ngày sinh ghi 'Nam').\n\n"
            "Sửa lại các dòng này trong CSDL (hoán đổi 2 cột về đúng vị trí)? "
            "Một bản sao lưu sẽ được tạo tự động trước khi sửa.",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        backup_database(reason="truoc_khi_sua_gioitinh")
        fixed = fix_swapped_gender_birthdate()
        self.log_line(f"Đã sửa {fixed:,} dòng bị đảo cột Giới tính / Ngày sinh.")
        self.main_window.on_data_changed()

    def fix_kham_dates(self):
        n = count_unparsed_kham_dates()
        if n == 0:
            QtWidgets.QMessageBox.information(
                self, "Không có gì để sửa",
                "Không tìm thấy dòng nào bị lỗi định dạng Ngày khám.")
            return
        reply = QtWidgets.QMessageBox.question(
            self, "Xác nhận sửa dữ liệu",
            f"Phát hiện {n:,} dòng có Ngày khám không xác định được (ví dụ định dạng "
            "'HH:MM dd/mm/yyyy' thay vì 'dd/mm/yyyy HH:MM'). Tính lại cho các dòng này?\n"
            "Một bản sao lưu sẽ được tạo tự động trước khi sửa.",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        backup_database(reason="truoc_khi_sua_ngaykham")
        fixed = fix_unparsed_kham_dates()
        self.log_line(f"Đã tính lại Ngày khám cho {fixed:,} dòng.")
        self.main_window.on_data_changed()

    def reclassify_diseases_click(self):
        n = count_unclassified_diseases()
        if n == 0:
            QtWidgets.QMessageBox.information(
                self, "Không có gì để xác định",
                "Mọi dòng có Chẩn đoán đều đã được gán Nhóm bệnh (KLN).")
            return
        reply = QtWidgets.QMessageBox.question(
            self, "Xác nhận",
            f"Phát hiện {n:,} dòng có Chẩn đoán nhưng chưa được gán Nhóm bệnh "
            "(thường là dữ liệu nhập từ trước khi có tính năng này).\n\n"
            "Tự động dò từ khóa trong Chẩn đoán để gán Nhóm bệnh (Tăng huyết áp, "
            "Đái tháo đường, COPD/Hen phế quản, Ung thư, Tâm thần, hoặc 'Khác') cho các dòng đó? "
            "Các dòng đã có Nhóm bệnh từ trước (kể cả do tự sửa tay) sẽ được GIỮ NGUYÊN, "
            "không bị ghi đè. Một bản sao lưu sẽ được tạo tự động trước khi sửa.",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        backup_database(reason="truoc_khi_xac_dinh_nhom_benh")
        updated = reclassify_diseases(overwrite=False)
        self.log_line(f"Đã xác định Nhóm bệnh cho {updated:,} dòng.")
        self.log_line("Lưu ý: đây là dò theo từ khóa, không thay thế chẩn đoán chuyên môn — "
                       "xem lại và sửa tay ở tab 'Danh sách' nếu cần độ chính xác cao.")
        self.main_window.on_data_changed()

    def export_quality_report(self):
        summary = data_quality_summary()
        if not any(n for _, n in summary):
            QtWidgets.QMessageBox.information(
                self, "Không có vấn đề",
                "Không phát hiện vấn đề chất lượng dữ liệu nào trong CSDL hiện tại.")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Xuất báo cáo chất lượng dữ liệu",
            os.path.join(BASE_DIR, "bao_cao_chat_luong_du_lieu.xlsx"), QT_EXPORT_FILTER)
        if not path:
            return
        headers = [label for _, label in COLUMNS] + ["Loại lỗi"]
        n = export_query_to_file(data_quality_rows_sql(), [], path, headers=headers)
        QtWidgets.QMessageBox.information(self, "Xuất dữ liệu", f"Đã xuất {n:,} dòng có vấn đề ra:\n{path}")

    def backup_now(self):
        path = backup_database(reason="thu_cong")
        if not path:
            QtWidgets.QMessageBox.warning(self, "Chưa có dữ liệu", "Chưa có CSDL để sao lưu.")
            return
        where = "trên máy chủ" if core.is_remote() else ""
        self.log_line(f"Đã sao lưu {where} vào: {path}")
        QtWidgets.QMessageBox.information(self, "Sao lưu thành công", f"Đã sao lưu {where} vào:\n{path}")

    def open_backup_folder(self):
        if core.is_remote():
            QtWidgets.QMessageBox.information(
                self, "Máy trạm",
                "Đang ở chế độ máy trạm — thư mục sao lưu nằm trên máy chủ.\n"
                "Vui lòng mở thư mục backups/ tại máy chủ.")
            return
        os.makedirs(core.BACKUP_DIR, exist_ok=True)
        QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(core.BACKUP_DIR))

    def _refresh_password_button(self):
        if has_password():
            self.password_btn.setText("Đổi / Tắt mật khẩu ứng dụng")
        else:
            self.password_btn.setText("Đặt mật khẩu bảo vệ ứng dụng")

    def manage_password(self):
        dlg = PasswordManageDialog(self)
        dlg.exec()
        self._refresh_password_button()


# ------------------------------------------------------------------
# Tab Danh sach
# ------------------------------------------------------------------

class DataTab(QtWidgets.QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.page = 0
        self.total = 0

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(10)

        top = QtWidgets.QHBoxLayout()
        top.addWidget(QtWidgets.QLabel("Tìm (họ tên / CCCD / BHYT):"))
        self.search_edit = QtWidgets.QLineEdit()
        self.search_edit.setMaximumWidth(260)
        self.search_edit.returnPressed.connect(lambda: self.reload(reset_page=True))
        top.addWidget(self.search_edit)

        top.addWidget(QtWidgets.QLabel("Giới tính:"))
        self.gender_combo = QtWidgets.QComboBox()
        self.gender_combo.addItem("Tất cả")
        self.gender_combo.currentIndexChanged.connect(lambda _: self.reload(reset_page=True))
        top.addWidget(self.gender_combo)

        top.addWidget(QtWidgets.QLabel("Nhóm bệnh:"))
        self.disease_combo = QtWidgets.QComboBox()
        self.disease_combo.addItem("Tất cả")
        self.disease_combo.addItems([label for _, label, _ in DISEASE_CATEGORIES] + [DISEASE_OTHER_LABEL])
        self.disease_combo.currentIndexChanged.connect(lambda _: self.reload(reset_page=True))
        top.addWidget(self.disease_combo)

        search_btn = QtWidgets.QPushButton("Tìm kiếm")
        search_btn.clicked.connect(lambda: self.reload(reset_page=True))
        top.addWidget(search_btn)

        export_btn = QtWidgets.QPushButton("Xuất theo bộ lọc (Excel/CSV)")
        export_btn.clicked.connect(self.export_data)
        top.addWidget(export_btn)
        top.addStretch(1)
        layout.addLayout(top)

        self.model = TableModel([label for _, label in COLUMNS], [])
        self.table = make_table_view()
        self.table.setModel(self.model)
        layout.addWidget(self.table, stretch=1)

        nav = QtWidgets.QHBoxLayout()
        prev_btn = QtWidgets.QPushButton("< Trang trước")
        prev_btn.clicked.connect(self.prev_page)
        next_btn = QtWidgets.QPushButton("Trang sau >")
        next_btn.clicked.connect(self.next_page)
        self.page_label = QtWidgets.QLabel("")
        nav.addWidget(prev_btn)
        nav.addWidget(self.page_label)
        nav.addWidget(next_btn)
        nav.addStretch(1)
        layout.addLayout(nav)

        self.reload()

    def _where(self):
        clauses = []
        params = []
        s = self.search_edit.text().strip()
        if s:
            clauses.append("(ho_ten LIKE ? OR so_cccd LIKE ? OR ma_bhyt LIKE ?)")
            like = f"%{s}%"
            params += [like, like, like]
        g = self.gender_combo.currentText()
        if g and g != "Tất cả":
            clauses.append("gioi_tinh = ?")
            params.append(g)
        d = self.disease_combo.currentText()
        if d and d != "Tất cả":
            clauses.append("benh LIKE ?")
            params.append(f"%{d}%")
        where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
        return where, params

    def reload(self, reset_page=False):
        if reset_page:
            self.page = 0
        conn = get_conn()
        genders = [r[0] for r in conn.execute(
            "SELECT DISTINCT gioi_tinh FROM patients WHERE gioi_tinh <> '' ORDER BY gioi_tinh")]
        cur_gender = self.gender_combo.currentText()
        self.gender_combo.blockSignals(True)
        self.gender_combo.clear()
        self.gender_combo.addItems(["Tất cả"] + genders)
        idx = self.gender_combo.findText(cur_gender)
        self.gender_combo.setCurrentIndex(idx if idx >= 0 else 0)
        self.gender_combo.blockSignals(False)

        where, params = self._where()
        self.total = conn.execute(f"SELECT COUNT(*) FROM patients {where}", params).fetchone()[0]
        offset = self.page * PAGE_SIZE
        rows = conn.execute(
            f"SELECT {', '.join(c for c, _ in COLUMNS)} FROM patients {where} "
            f"ORDER BY id LIMIT ? OFFSET ?", params + [PAGE_SIZE, offset]
        ).fetchall()
        conn.close()

        self.model.set_data([label for _, label in COLUMNS], [tuple(r) for r in rows])

        pages = max(1, (self.total + PAGE_SIZE - 1) // PAGE_SIZE)
        self.page_label.setText(f"Trang {self.page + 1}/{pages}   (tổng {self.total:,} bản ghi khớp bộ lọc)")

    def prev_page(self):
        if self.page > 0:
            self.page -= 1
            self.reload()

    def next_page(self):
        pages = max(1, (self.total + PAGE_SIZE - 1) // PAGE_SIZE)
        if self.page + 1 < pages:
            self.page += 1
            self.reload()

    def export_data(self):
        where, params = self._where()
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Xuất dữ liệu", os.path.join(BASE_DIR, "danh_sach_loc.xlsx"), QT_EXPORT_FILTER)
        if not path:
            return
        sql = f"SELECT {', '.join(c for c, _ in COLUMNS)} FROM patients {where} ORDER BY id"
        headers = [label for _, label in COLUMNS]
        try:
            n = export_query_to_file(sql, params, path, headers=headers)
        except RuntimeError as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi", str(e))
            return
        QtWidgets.QMessageBox.information(self, "Xuất dữ liệu", f"Đã xuất {n:,} bản ghi ra:\n{path}")


# ------------------------------------------------------------------
# Tab Loc trung
# ------------------------------------------------------------------

class ExceptionsDialog(QtWidgets.QDialog):
    """Danh sach cac nhom da duoc nguoi dung xac nhan la KHONG trung, ung voi 1 to hop tieu chi."""

    def __init__(self, parent, key_type):
        super().__init__(parent)
        self.key_type = key_type
        self.setWindowTitle(f"Danh sách đã xác nhận KHÔNG trùng — {key_type}")
        self.resize(720, 420)

        layout = QtWidgets.QVBoxLayout(self)
        self.model = TableModel(["Giá trị khóa", "Họ tên (đại diện)", "Xác nhận lúc"], [])
        self.table = make_table_view()
        self.table.setModel(self.model)
        layout.addWidget(self.table, stretch=1)

        btns = QtWidgets.QHBoxLayout()
        remove_btn = QtWidgets.QPushButton("Bỏ xác nhận (đưa lại vào danh sách trùng)")
        remove_btn.clicked.connect(self.remove_selected)
        btns.addWidget(remove_btn)
        btns.addStretch(1)
        close_btn = QtWidgets.QPushButton("Đóng")
        close_btn.clicked.connect(self.accept)
        btns.addWidget(close_btn)
        layout.addLayout(btns)

        self.reload()

    def reload(self):
        rows = list_dedup_exceptions(self.key_type)
        self.model.set_data(
            ["Giá trị khóa", "Họ tên (đại diện)", "Xác nhận lúc"],
            [(r["key_value"], r["ten_dai_dien"], r["created_at"]) for r in rows])

    def remove_selected(self):
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            return
        key_val = self.model.row_values(sel[0].row())[0]
        remove_dedup_exception(self.key_type, key_val)
        self.reload()


class DedupTab(QtWidgets.QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.field_checks = {}
        self.current_key_expr = None
        self.current_key_where = None
        self.current_key_type = None

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(10)

        criteria_box = QtWidgets.QGroupBox(
            "Tiêu chí xác định trùng — chọn 1 hoặc nhiều trường (kết hợp AND, tất cả trường chọn phải khớp)")
        criteria_layout = QtWidgets.QHBoxLayout(criteria_box)
        for label, _, _ in DEDUP_FIELDS:
            cb = QtWidgets.QCheckBox(label)
            if label == "Số CCCD":
                cb.setChecked(True)
            self.field_checks[label] = cb
            criteria_layout.addWidget(cb)
        criteria_layout.addStretch(1)
        scan_btn = QtWidgets.QPushButton("Quét trùng")
        scan_btn.setObjectName("PrimaryButton")
        scan_btn.clicked.connect(self.scan)
        criteria_layout.addWidget(scan_btn)
        layout.addWidget(criteria_box)

        self.summary_label = QtWidgets.QLabel("Chưa quét.")
        layout.addWidget(self.summary_label)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Orientation.Vertical)

        group_box = QtWidgets.QWidget()
        group_layout = QtWidgets.QVBoxLayout(group_box)
        group_layout.setContentsMargins(0, 0, 0, 0)
        group_layout.addWidget(QtWidgets.QLabel("Danh sách nhóm trùng (bấm để xem chi tiết):"))
        self.group_model = TableModel(["Giá trị khóa", "Số bản ghi", "Họ tên (đại diện)"], [])
        self.group_table = make_table_view()
        self.group_table.setModel(self.group_model)
        self.group_table.selectionModel().selectionChanged.connect(self.show_detail)
        group_layout.addWidget(self.group_table)
        splitter.addWidget(group_box)

        detail_box = QtWidgets.QWidget()
        detail_layout = QtWidgets.QVBoxLayout(detail_box)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_top = QtWidgets.QHBoxLayout()
        detail_top.addWidget(QtWidgets.QLabel("Chi tiết các bản ghi trong nhóm được chọn (tích để chọn từng dòng):"))
        detail_top.addStretch(1)
        check_all_btn = QtWidgets.QPushButton("Tích tất cả")
        check_all_btn.clicked.connect(lambda: self.detail_model.check_all(True))
        uncheck_all_btn = QtWidgets.QPushButton("Bỏ tích tất cả")
        uncheck_all_btn.clicked.connect(lambda: self.detail_model.check_all(False))
        detail_top.addWidget(check_all_btn)
        detail_top.addWidget(uncheck_all_btn)
        detail_layout.addLayout(detail_top)

        self.detail_model = CheckableTableModel([label for _, label in COLUMNS], [])
        self.detail_table = make_table_view()
        self.detail_table.setModel(self.detail_model)
        self.detail_table.setColumnWidth(0, 46)
        detail_layout.addWidget(self.detail_table)
        splitter.addWidget(detail_box)

        layout.addWidget(splitter, stretch=1)

        keep_row = QtWidgets.QHBoxLayout()
        keep_row.addWidget(QtWidgets.QLabel("Bản ghi chính (giữ mọi thông tin + nhận lịch sử khám gộp):"))
        self.keep_combo = QtWidgets.QComboBox()
        self.keep_combo.addItems(["Ngày khám mới nhất", "Bản ghi đầu tiên"])
        keep_row.addWidget(self.keep_combo)
        keep_row.addStretch(1)
        layout.addLayout(keep_row)

        group_actions_box = QtWidgets.QGroupBox("Hành động với nhóm đang chọn ở trên")
        ga = QtWidgets.QHBoxLayout(group_actions_box)
        merge_checked_btn = QtWidgets.QPushButton("Gộp các bản ghi đã tích thành 1")
        merge_checked_btn.setObjectName("PrimaryButton")
        merge_checked_btn.clicked.connect(self.merge_checked)
        ga.addWidget(merge_checked_btn)

        merge_group_btn = QtWidgets.QPushButton("Gộp cả nhóm này thành 1 bản ghi")
        merge_group_btn.setObjectName("PrimaryButton")
        merge_group_btn.clicked.connect(self.merge_this_group)
        ga.addWidget(merge_group_btn)

        del_checked_btn = QtWidgets.QPushButton("Xóa hẳn các bản ghi đã tích (không gộp)")
        del_checked_btn.setObjectName("DangerButton")
        del_checked_btn.clicked.connect(self.delete_checked)
        ga.addWidget(del_checked_btn)

        confirm_btn = QtWidgets.QPushButton("Xác nhận: đây KHÔNG phải trùng")
        confirm_btn.clicked.connect(self.confirm_not_duplicate)
        ga.addWidget(confirm_btn)

        manage_btn = QtWidgets.QPushButton("Quản lý danh sách đã xác nhận...")
        manage_btn.clicked.connect(self.manage_exceptions)
        ga.addWidget(manage_btn)
        ga.addStretch(1)
        layout.addWidget(group_actions_box)

        global_box = QtWidgets.QGroupBox("Xuất / xử lý toàn bộ danh sách trùng")
        gl = QtWidgets.QHBoxLayout(global_box)

        self.history_check = QtWidgets.QCheckBox("Kèm cột lịch sử khám khi xuất")
        self.history_check.setChecked(True)
        gl.addWidget(self.history_check)

        exp_dup_btn = QtWidgets.QPushButton("Xuất danh sách trùng (Excel/CSV)")
        exp_dup_btn.clicked.connect(self.export_duplicates)
        gl.addWidget(exp_dup_btn)

        exp_uniq_btn = QtWidgets.QPushButton("Xuất danh sách đã lọc trùng - duy nhất (Excel/CSV)")
        exp_uniq_btn.clicked.connect(self.export_unique)
        gl.addWidget(exp_uniq_btn)

        merge_all_btn = QtWidgets.QPushButton("Gộp TẤT CẢ nhóm trùng (mỗi nhóm → 1 bản ghi + lịch sử khám)")
        merge_all_btn.setObjectName("PrimaryButton")
        merge_all_btn.clicked.connect(self.merge_all_groups_action)
        gl.addWidget(merge_all_btn)
        gl.addStretch(1)
        layout.addWidget(global_box)

    # ---------------- tieu chi & quet ----------------

    def _resolve_key(self):
        selected = [label for label, cb in self.field_checks.items() if cb.isChecked()]
        try:
            return build_dedup_key(selected)
        except ValueError as e:
            QtWidgets.QMessageBox.warning(self, "Chưa chọn tiêu chí", str(e))
            return None

    def scan(self):
        resolved = self._resolve_key()
        if not resolved:
            return
        key_expr, key_where, key_type = resolved
        self.current_key_expr = key_expr
        self.current_key_where = key_where
        self.current_key_type = key_type

        rows = scan_dedup_groups(key_expr, key_where, key_type)
        self.group_model.set_data(
            ["Giá trị khóa", "Số bản ghi", "Họ tên (đại diện)"],
            [(r["k"], r["n"], r["ten"]) for r in rows])
        self.detail_model.set_data([label for _, label in COLUMNS], [])

        total_extra = sum(r["n"] - 1 for r in rows)
        n_exceptions = len(list_dedup_exceptions(key_type))
        note = f"  (đã loại {n_exceptions:,} nhóm đã xác nhận không trùng)" if n_exceptions else ""
        self.summary_label.setText(
            f"Tiêu chí: {key_type}  —  Tìm thấy {len(rows):,} nhóm trùng, "
            f"dư thừa {total_extra:,} bản ghi.{note}")

    def show_detail(self, selected=None, deselected=None):
        sel = self.group_table.selectionModel().selectedRows()
        if not sel or not self.current_key_expr:
            self.detail_model.set_data([label for _, label in COLUMNS], [])
            return
        row_idx = sel[0].row()
        key_val = self.group_model.row_values(row_idx)[0]
        rows = group_detail_rows(self.current_key_expr, key_val)
        self.detail_model.set_data([label for _, label in COLUMNS], [tuple(r) for r in rows])

    def _require_scan(self):
        if self.group_model.rowCount() == 0:
            QtWidgets.QMessageBox.warning(self, "Chưa có dữ liệu", "Vui lòng bấm 'Quét trùng' trước.")
            return False
        return True

    def _selected_group(self):
        sel = self.group_table.selectionModel().selectedRows()
        if not sel:
            QtWidgets.QMessageBox.warning(self, "Chưa chọn nhóm", "Vui lòng chọn 1 nhóm trong danh sách ở trên trước.")
            return None
        return self.group_model.row_values(sel[0].row())

    def _order_for_keep(self):
        if self.keep_combo.currentText().startswith("Bản ghi đầu tiên"):
            return "ORDER BY id ASC"
        return "ORDER BY (ngay_kham_date IS NULL), ngay_kham_date DESC, id DESC"

    # ---------------- hanh dong theo nhom / theo dong tich ----------------

    def merge_checked(self):
        checked = self.detail_model.checked_rows()
        if len(checked) < 2:
            QtWidgets.QMessageBox.warning(
                self, "Cần chọn ít nhất 2 dòng",
                "Vui lòng tích chọn từ 2 bản ghi trở lên trong bảng chi tiết bên trên để gộp.")
            return
        id_col = [c for c, _ in COLUMNS].index("id")
        ids = [row[id_col] for row in checked]
        reply = QtWidgets.QMessageBox.question(
            self, "Xác nhận gộp",
            f"Gộp {len(ids)} bản ghi đã tích chọn thành 1 bản ghi (giữ '{self.keep_combo.currentText()}' "
            "làm bản ghi chính, các lượt khám còn lại được đưa vào cột 'Lịch sử khám')?\n"
            "Các bản ghi bị gộp sẽ không còn xuất hiện riêng lẻ nữa.",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        backup_database(reason="truoc_khi_gop")
        n = merge_specific_ids(ids, self._order_for_keep())
        QtWidgets.QMessageBox.information(self, "Hoàn tất", f"Đã gộp {n} bản ghi vào bản ghi chính.")
        self.main_window.on_data_changed()
        self.scan()

    def delete_checked(self):
        checked = self.detail_model.checked_rows()
        if not checked:
            QtWidgets.QMessageBox.warning(
                self, "Chưa chọn dòng nào",
                "Vui lòng tích chọn ít nhất 1 bản ghi trong bảng chi tiết bên trên.")
            return
        id_col = [c for c, _ in COLUMNS].index("id")
        ids = [row[id_col] for row in checked]
        reply = QtWidgets.QMessageBox.question(
            self, "Xác nhận xóa hẳn",
            f"Xóa HẲN {len(ids)} bản ghi đã tích chọn khỏi CSDL (không gộp, mất toàn bộ thông tin "
            "của các bản ghi này)?\nChỉ dùng khi đây thực sự là các dòng rác/lỗi, không phải lượt khám "
            "hợp lệ.\nHành động không thể hoàn tác.",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        backup_database(reason="truoc_khi_xoa_han")
        n = delete_patients_by_ids(ids)
        QtWidgets.QMessageBox.information(self, "Hoàn tất", f"Đã xóa {n} bản ghi.")
        self.main_window.on_data_changed()
        self.scan()

    def merge_this_group(self):
        if not self.current_key_expr:
            return
        g = self._selected_group()
        if not g:
            return
        key_val, count, ten = g
        if count < 2:
            QtWidgets.QMessageBox.information(self, "Không có gì để gộp", "Nhóm này chỉ có 1 bản ghi.")
            return
        reply = QtWidgets.QMessageBox.question(
            self, "Xác nhận gộp nhóm",
            f"Gộp {count} bản ghi của nhóm '{ten}' thành 1 bản ghi (giữ "
            f"'{self.keep_combo.currentText()}' làm bản ghi chính, các lượt khám còn lại được đưa "
            "vào cột 'Lịch sử khám')?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        backup_database(reason="truoc_khi_gop_nhom")
        n = merge_group(self.current_key_expr, key_val, self._order_for_keep())
        QtWidgets.QMessageBox.information(self, "Hoàn tất", f"Đã gộp {n} bản ghi vào bản ghi chính của nhóm.")
        self.main_window.on_data_changed()
        self.scan()

    def confirm_not_duplicate(self):
        if not self.current_key_type:
            return
        g = self._selected_group()
        if not g:
            return
        key_val, count, ten = g
        add_dedup_exception(self.current_key_type, key_val, ten or "")
        QtWidgets.QMessageBox.information(
            self, "Đã xác nhận",
            f"Đã đánh dấu nhóm '{ten}' là KHÔNG trùng.\nLần quét sau (với tiêu chí: "
            f"{self.current_key_type}) sẽ không hiển thị nhóm này nữa.")
        self.scan()

    def manage_exceptions(self):
        if not self.current_key_type:
            QtWidgets.QMessageBox.warning(self, "Chưa quét", "Vui lòng quét trùng trước để xác định tiêu chí.")
            return
        dlg = ExceptionsDialog(self, self.current_key_type)
        dlg.exec()
        self.scan()

    # ---------------- xuat / xoa toan bo danh sach trung ----------------

    def export_duplicates(self):
        if not self._require_scan():
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Xuất danh sách trùng", os.path.join(BASE_DIR, "danh_sach_trung.xlsx"), QT_EXPORT_FILTER)
        if not path:
            return
        key_expr, key_where = self.current_key_expr, self.current_key_where
        sql = f"""
            SELECT {', '.join(c for c, _ in COLUMNS)} FROM patients
            WHERE {key_expr} IN (
                SELECT {key_expr} FROM patients WHERE {key_where}
                GROUP BY {key_expr} HAVING COUNT(*) > 1
            )
            ORDER BY {key_expr}, id
        """
        headers = [label for _, label in COLUMNS]
        try:
            n = export_query_to_file(sql, [], path, headers=headers)
        except RuntimeError as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi", str(e))
            return
        QtWidgets.QMessageBox.information(self, "Xuất dữ liệu", f"Đã xuất {n:,} bản ghi trùng ra:\n{path}")

    def export_unique(self):
        resolved = self._resolve_key()
        if not resolved:
            return
        key_expr, key_where, key_type = resolved
        order = self._order_for_keep()
        include_history = self.history_check.isChecked()
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Xuất danh sách duy nhất", os.path.join(BASE_DIR, "danh_sach_duy_nhat.xlsx"), QT_EXPORT_FILTER)
        if not path:
            return
        headers, rows = unique_rows_with_optional_history(key_expr, key_where, order, include_history)
        try:
            n = write_export(path, headers, (tuple(r) for r in rows))
        except RuntimeError as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi", str(e))
            return
        QtWidgets.QMessageBox.information(
            self, "Xuất dữ liệu", f"Đã xuất {n:,} bản ghi (đã lọc trùng) ra:\n{path}")

    def merge_all_groups_action(self):
        if not self._require_scan():
            return
        order = self._order_for_keep()
        n_groups = self.group_model.rowCount()
        reply = QtWidgets.QMessageBox.question(
            self, "Xác nhận gộp TẤT CẢ",
            f"Gộp toàn bộ {n_groups:,} nhóm trùng trong CSDL (tiêu chí: {self.current_key_type}) — "
            f"mỗi nhóm còn lại đúng 1 bản ghi (giữ '{self.keep_combo.currentText()}' làm bản ghi chính), "
            "các lượt khám còn lại được đưa vào cột 'Lịch sử khám', không mất dữ liệu.\n\n"
            "Bạn có chắc chắn muốn tiếp tục?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        backup_database(reason="truoc_khi_gop_tat_ca")
        n_g, n_merged = merge_all_groups(
            self.current_key_expr, self.current_key_where, self.current_key_type, order)
        QtWidgets.QMessageBox.information(
            self, "Hoàn tất", f"Đã gộp {n_g:,} nhóm ({n_merged:,} bản ghi được đưa vào lịch sử khám).")
        self.main_window.on_data_changed()
        self.scan()


# ------------------------------------------------------------------
# Tab Truy van SQL
# ------------------------------------------------------------------

class FilterRowWidget(QtWidgets.QWidget):
    """1 dong dieu kien loc trong trinh tao cau lenh: Truong / Toan tu / Gia tri."""
    removed = QtCore.pyqtSignal(object)

    OPERATORS = [
        ("bằng", "eq"),
        ("khác", "ne"),
        ("chứa", "contains"),
        ("bắt đầu bằng", "startswith"),
        ("lớn hơn", "gt"),
        ("nhỏ hơn", "lt"),
        ("lớn hơn hoặc bằng", "gte"),
        ("nhỏ hơn hoặc bằng", "lte"),
        ("để trống", "empty"),
        ("không để trống", "not_empty"),
    ]

    def __init__(self):
        super().__init__()
        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.field_combo = QtWidgets.QComboBox()
        self.field_combo.addItems([label for _, label in COLUMNS])
        layout.addWidget(self.field_combo, stretch=2)

        self.op_combo = QtWidgets.QComboBox()
        self.op_combo.addItems([label for label, _ in self.OPERATORS])
        self.op_combo.currentIndexChanged.connect(self._update_value_visibility)
        layout.addWidget(self.op_combo, stretch=2)

        self.value_edit = QtWidgets.QLineEdit()
        self.value_edit.setPlaceholderText("Giá trị")
        layout.addWidget(self.value_edit, stretch=2)

        remove_btn = QtWidgets.QPushButton("Xóa")
        remove_btn.setFixedWidth(52)
        remove_btn.clicked.connect(lambda: self.removed.emit(self))
        layout.addWidget(remove_btn)

    def _update_value_visibility(self):
        op_key = self.OPERATORS[self.op_combo.currentIndex()][1]
        self.value_edit.setEnabled(op_key not in ("empty", "not_empty"))

    def build_condition(self):
        col = COLUMNS[self.field_combo.currentIndex()][0]
        op_key = self.OPERATORS[self.op_combo.currentIndex()][1]
        value = self.value_edit.text().strip()
        esc = value.replace("'", "''")
        if op_key == "empty":
            return f"({col} IS NULL OR TRIM({col}) = '')"
        if op_key == "not_empty":
            return f"({col} IS NOT NULL AND TRIM({col}) <> '')"
        if not value:
            return None
        if op_key == "eq":
            return f"{col} = '{esc}'"
        if op_key == "ne":
            return f"{col} <> '{esc}'"
        if op_key == "contains":
            return f"{col} LIKE '%{esc}%'"
        if op_key == "startswith":
            return f"{col} LIKE '{esc}%'"
        if op_key == "gt":
            return f"{col} > '{esc}'"
        if op_key == "lt":
            return f"{col} < '{esc}'"
        if op_key == "gte":
            return f"{col} >= '{esc}'"
        if op_key == "lte":
            return f"{col} <= '{esc}'"
        return None


class QueryBuilderBox(QtWidgets.QGroupBox):
    """Trinh tao cau lenh SQL bang giao dien (khong can go SQL): chon cot,
    dieu kien loc, nhom theo, sap xep, gioi han so dong - roi sinh ra cau
    SELECT tuong ung dua vao khung soan thao ben duoi de xem lai / chay."""
    build_requested = QtCore.pyqtSignal(str)

    def __init__(self):
        super().__init__("Trình tạo câu lệnh SQL (không cần biết cú pháp) — bấm để mở/thu gọn")
        self.setCheckable(True)
        self.setChecked(False)

        outer = QtWidgets.QVBoxLayout(self)
        self.content = QtWidgets.QWidget()
        self.content.setVisible(False)
        outer.addWidget(self.content)
        self.toggled.connect(self.content.setVisible)

        layout = QtWidgets.QVBoxLayout(self.content)

        layout.addWidget(QtWidgets.QLabel("Cột muốn hiển thị (bỏ trống = hiển thị tất cả):"))
        cols_widget = QtWidgets.QWidget()
        cols_grid = QtWidgets.QGridLayout(cols_widget)
        cols_grid.setContentsMargins(0, 0, 0, 0)
        self.column_checks = {}
        for i, (dbcol, label) in enumerate(COLUMNS):
            cb = QtWidgets.QCheckBox(label)
            self.column_checks[dbcol] = cb
            cols_grid.addWidget(cb, i // 5, i % 5)
        layout.addWidget(cols_widget)

        layout.addWidget(QtWidgets.QLabel("Điều kiện lọc (kết hợp AND):"))
        self.filters_container = QtWidgets.QVBoxLayout()
        layout.addLayout(self.filters_container)
        self.filter_rows = []

        add_filter_btn = QtWidgets.QPushButton("+ Thêm điều kiện lọc")
        add_filter_btn.clicked.connect(self.add_filter_row)
        layout.addWidget(add_filter_btn, alignment=QtCore.Qt.AlignmentFlag.AlignLeft)

        opts_row = QtWidgets.QHBoxLayout()
        opts_row.addWidget(QtWidgets.QLabel("Nhóm theo:"))
        self.group_combo = QtWidgets.QComboBox()
        self.group_combo.addItem("(không nhóm)")
        self.group_combo.addItems([label for _, label in COLUMNS])
        opts_row.addWidget(self.group_combo)

        opts_row.addWidget(QtWidgets.QLabel("Sắp xếp theo:"))
        self.order_combo = QtWidgets.QComboBox()
        self.order_combo.addItem("(không sắp xếp)")
        self.order_combo.addItems([label for _, label in COLUMNS])
        opts_row.addWidget(self.order_combo)
        self.order_dir_combo = QtWidgets.QComboBox()
        self.order_dir_combo.addItems(["Tăng dần", "Giảm dần"])
        opts_row.addWidget(self.order_dir_combo)

        opts_row.addWidget(QtWidgets.QLabel("Giới hạn số dòng:"))
        self.limit_spin = QtWidgets.QSpinBox()
        self.limit_spin.setRange(0, 1_000_000)
        self.limit_spin.setValue(0)
        self.limit_spin.setSpecialValueText("(không giới hạn)")
        opts_row.addWidget(self.limit_spin)
        opts_row.addStretch(1)
        layout.addLayout(opts_row)

        build_btn = QtWidgets.QPushButton("Tạo câu lệnh SQL")
        build_btn.setObjectName("PrimaryButton")
        build_btn.clicked.connect(self.build_sql)
        layout.addWidget(build_btn, alignment=QtCore.Qt.AlignmentFlag.AlignLeft)

        self.add_filter_row()

    def add_filter_row(self):
        row = FilterRowWidget()
        row.removed.connect(self.remove_filter_row)
        self.filter_rows.append(row)
        self.filters_container.addWidget(row)

    def remove_filter_row(self, row):
        if row in self.filter_rows:
            self.filter_rows.remove(row)
            row.setParent(None)
            row.deleteLater()

    def build_sql(self):
        selected_cols = [c for c, cb in self.column_checks.items() if cb.isChecked()]
        group_idx = self.group_combo.currentIndex()
        group_col = COLUMNS[group_idx - 1][0] if group_idx > 0 else None

        if group_col:
            select_cols = selected_cols or [group_col]
            if group_col not in select_cols:
                select_cols = [group_col] + select_cols
            cols_sql = ", ".join(select_cols) + ", COUNT(*) AS so_luong"
        else:
            cols_sql = ", ".join(selected_cols) if selected_cols else "*"

        conditions = []
        for row in self.filter_rows:
            cond = row.build_condition()
            if cond:
                conditions.append(cond)
        where_sql = f"WHERE {' AND '.join(conditions)}" if conditions else ""
        group_sql = f"GROUP BY {group_col}" if group_col else ""

        order_idx = self.order_combo.currentIndex()
        order_sql = ""
        if order_idx > 0:
            order_col = COLUMNS[order_idx - 1][0]
            direction = "ASC" if self.order_dir_combo.currentIndex() == 0 else "DESC"
            order_sql = f"ORDER BY {order_col} {direction}"
        elif group_col:
            order_sql = "ORDER BY so_luong DESC"

        limit_sql = f"LIMIT {self.limit_spin.value()}" if self.limit_spin.value() > 0 else ""

        parts = [f"SELECT {cols_sql}", "FROM patients"]
        if where_sql:
            parts.append(where_sql)
        if group_sql:
            parts.append(group_sql)
        if order_sql:
            parts.append(order_sql)
        if limit_sql:
            parts.append(limit_sql)
        self.build_requested.emit("\n".join(parts))


class SqlTab(QtWidgets.QWidget):
    QUICK_QUERIES = {
        "Tổng số bản ghi": "SELECT COUNT(*) AS tong_so_ban_ghi FROM patients",
        "Số bệnh nhân duy nhất (theo CCCD)":
            "SELECT COUNT(DISTINCT so_cccd) AS so_benh_nhan_duy_nhat FROM patients "
            "WHERE so_cccd IS NOT NULL AND TRIM(so_cccd) <> ''",
        "Thống kê theo giới tính":
            "SELECT gioi_tinh, COUNT(*) AS so_luong FROM patients "
            "GROUP BY gioi_tinh ORDER BY so_luong DESC",
        "Thống kê theo Tỉnh/TP":
            "SELECT tinh_tp, COUNT(*) AS so_luong FROM patients "
            "GROUP BY tinh_tp ORDER BY so_luong DESC",
        "Thống kê theo Phường/Xã (Top 30)":
            "SELECT phuong_xa, COUNT(*) AS so_luong FROM patients "
            "GROUP BY phuong_xa ORDER BY so_luong DESC LIMIT 30",
        "Top 20 chẩn đoán phổ biến":
            "SELECT chan_doan, COUNT(*) AS so_luong FROM patients "
            "GROUP BY chan_doan ORDER BY so_luong DESC LIMIT 20",
        "Tổ hợp Nhóm bệnh (KLN) - theo bản ghi":
            "SELECT benh, COUNT(*) AS so_luong FROM patients "
            "WHERE TRIM(IFNULL(benh, '')) <> '' "
            "GROUP BY benh ORDER BY so_luong DESC",
        "Thống kê theo năm sinh":
            "SELECT birth_year, COUNT(*) AS so_luong FROM patients "
            "WHERE birth_year IS NOT NULL GROUP BY birth_year ORDER BY birth_year",
    }

    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.last_sql = None

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(10)

        self.builder_box = QueryBuilderBox()
        self.builder_box.build_requested.connect(self.on_sql_generated)
        layout.addWidget(self.builder_box)

        top = QtWidgets.QHBoxLayout()
        top.addWidget(QtWidgets.QLabel("Câu lệnh nhanh:"))
        self.quick_combo = QtWidgets.QComboBox()
        self.quick_combo.addItem("")
        self.quick_combo.addItems(list(self.QUICK_QUERIES.keys()))
        self.quick_combo.currentTextChanged.connect(self.load_quick)
        top.addWidget(self.quick_combo)
        top.addStretch(1)
        layout.addLayout(top)

        layout.addWidget(QtWidgets.QLabel("Câu lệnh SQL (chỉ cho phép SELECT, bảng: patients):"))
        self.sql_edit = QtWidgets.QPlainTextEdit()
        self.sql_edit.setPlainText(self.QUICK_QUERIES["Tổng số bản ghi"])
        self.sql_edit.setMaximumHeight(110)
        layout.addWidget(self.sql_edit)

        btns = QtWidgets.QHBoxLayout()
        run_btn = QtWidgets.QPushButton("Chạy truy vấn")
        run_btn.setObjectName("PrimaryButton")
        run_btn.clicked.connect(self.run_query)
        btns.addWidget(run_btn)
        export_btn = QtWidgets.QPushButton("Xuất kết quả (Excel/CSV)")
        export_btn.clicked.connect(self.export_result)
        btns.addWidget(export_btn)
        btns.addStretch(1)
        layout.addLayout(btns)

        self.result_model = TableModel([], [])
        self.result_table = make_table_view()
        self.result_table.setModel(self.result_model)
        layout.addWidget(self.result_table, stretch=1)

        self.status_label = QtWidgets.QLabel("")
        layout.addWidget(self.status_label)

    def load_quick(self, text):
        q = self.QUICK_QUERIES.get(text)
        if q:
            self.sql_edit.setPlainText(q)
            self.run_query()

    def on_sql_generated(self, sql):
        self.sql_edit.setPlainText(sql)
        self.run_query()

    def _validate_select(self, sql):
        stripped = sql.strip().rstrip(";").strip()
        if not stripped.lower().startswith("select"):
            raise ValueError("Chỉ cho phép câu lệnh SELECT.")
        return stripped

    def run_query(self):
        raw = self.sql_edit.toPlainText()
        try:
            sql = self._validate_select(raw)
        except ValueError as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi", str(e))
            return
        try:
            conn = get_conn()
            cur = conn.execute(sql)
            rows = cur.fetchall()
            cols = [d[0] for d in cur.description]
            conn.close()
        except sqlite3.Error as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi SQL", str(e))
            return

        self.result_model.set_data(cols, [tuple(r) for r in rows])
        self.status_label.setText(f"{len(rows):,} dòng kết quả.")
        self.last_sql = sql

    def export_result(self):
        if not self.last_sql:
            QtWidgets.QMessageBox.warning(self, "Chưa có kết quả", "Vui lòng chạy truy vấn trước.")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Xuất kết quả truy vấn", os.path.join(BASE_DIR, "ket_qua_truy_van.xlsx"), QT_EXPORT_FILTER)
        if not path:
            return
        try:
            n = export_query_to_file(self.last_sql, [], path)
        except RuntimeError as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi", str(e))
            return
        QtWidgets.QMessageBox.information(self, "Xuất dữ liệu", f"Đã xuất {n:,} dòng ra:\n{path}")


# ------------------------------------------------------------------
# Tab Thong ke: bieu do truc quan (QtCharts) tren du lieu hien co trong
# CSDL - khong ve ban do dia ly (chua co du lieu ranh gioi hanh chinh
# chinh thuc dang tin cay de dua vao ung dung).
# ------------------------------------------------------------------

class StatsTab(QtWidgets.QWidget):
    # (nhan hien thi, cot CSDL hoac None cho "Nam sinh theo thap ky", kieu bieu do)
    # Cot "__disease__" la gia tri dac biet, khong phai ten cot CSDL that -
    # xem nhanh trong reload() ben duoi (dung stats_disease_counts() vi 1
    # benh nhan co the thuoc nhieu nhom benh cung luc, khac voi cac cot
    # thong ke thong thuong chi co 1 gia tri/dong).
    STAT_OPTIONS = [
        ("Giới tính", "gioi_tinh", "pie"),
        ("Nhóm bệnh (KLN)", "__disease__", "bar"),
        ("Tỉnh/Thành phố", "tinh_tp", "bar"),
        ("Phường/Xã (top 20)", "phuong_xa", "bar"),
        ("Chẩn đoán (top 15)", "chan_doan", "bar"),
        ("Năm sinh theo thập kỷ", None, "bar"),
    ]
    LIMITS = {"phuong_xa": 20, "chan_doan": 15}

    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(12)

        title = QtWidgets.QLabel("Thống kê trực quan")
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        row = QtWidgets.QHBoxLayout()
        row.addWidget(QtWidgets.QLabel("Loại thống kê:"))
        self.type_combo = QtWidgets.QComboBox()
        self.type_combo.addItems([label for label, _, _ in self.STAT_OPTIONS])
        self.type_combo.currentIndexChanged.connect(self.reload)
        row.addWidget(self.type_combo)

        refresh_btn = QtWidgets.QPushButton("Làm mới")
        refresh_btn.clicked.connect(self.reload)
        row.addWidget(refresh_btn)
        row.addStretch(1)
        layout.addLayout(row)

        self.chart_view = QChartView()
        self.chart_view.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        layout.addWidget(self.chart_view, stretch=1)

        self.summary_label = QtWidgets.QLabel()
        layout.addWidget(self.summary_label)

        self._loaded = False

    def showEvent(self, event):
        super().showEvent(event)
        if not self._loaded:
            self.reload()

    def mark_stale(self):
        """Goi khi du lieu CSDL thay doi o tab khac - hoan lai viec ve lai
        bieu do den lan sau tab nay duoc mo, thay vi ve lai ngay ca khi
        dang khong hien thi."""
        self._loaded = False

    def reload(self):
        self._loaded = True
        label, column, kind = self.STAT_OPTIONS[self.type_combo.currentIndex()]
        if column is None:
            data = stats_birth_decade()
        elif column == "__disease__":
            data = stats_disease_counts()
        else:
            data = stats_top_values(column, limit=self.LIMITS.get(column, 50))

        if not data:
            self.chart_view.setChart(QChart())
            self.summary_label.setText("Chưa có dữ liệu để thống kê.")
            return

        chart = self._make_pie_chart(label, data) if kind == "pie" else self._make_bar_chart(label, data)
        self.chart_view.setChart(chart)
        total = sum(n for _, n in data)
        self.summary_label.setText(f"Tổng {total:,} bản ghi trong {len(data)} nhóm đang hiển thị.")

    def _make_pie_chart(self, title, data):
        series = QPieSeries()
        for value, n in data:
            series.append(f"{value} ({n:,})", n)
        chart = QChart()
        chart.addSeries(series)
        chart.setTitle(title)
        chart.legend().setVisible(True)
        chart.legend().setAlignment(QtCore.Qt.AlignmentFlag.AlignRight)
        return chart

    def _make_bar_chart(self, title, data):
        # Bieu do ngang (thay vi doc) vi nhan (ten phuong/xa, chan doan...)
        # thuong la chuoi dai, de doc hon khi de o truc doc.
        ordered = list(reversed(data))  # gia tri lon nhat nam tren cung
        bar_set = QBarSet(title)
        categories = []
        for value, n in ordered:
            bar_set.append(n)
            categories.append(str(value) if value not in (None, "") else "(trống)")

        series = QHorizontalBarSeries()
        series.append(bar_set)

        chart = QChart()
        chart.addSeries(series)
        chart.setTitle(title)
        chart.legend().setVisible(False)

        axis_y = QBarCategoryAxis()
        axis_y.append(categories)
        chart.addAxis(axis_y, QtCore.Qt.AlignmentFlag.AlignLeft)
        series.attachAxis(axis_y)

        axis_x = QValueAxis()
        axis_x.setLabelFormat("%d")
        chart.addAxis(axis_x, QtCore.Qt.AlignmentFlag.AlignBottom)
        series.attachAxis(axis_x)

        return chart


# ------------------------------------------------------------------
# Tab Mang LAN: ket noi toi may chu chia se benh_nhan.db qua mang LAN
# noi bo (khong dung Internet/cloud). Ban than ung dung nay (app.py) chi
# dong vai tro "may tram" (client) - "may chu" la 1 goi rieng, nhe hon
# (khong dung PyQt6), chay nhu Windows Service - xem service.py,
# server_tray.py va muc "May chu chia se mang LAN" trong README.
# ------------------------------------------------------------------

class NetworkTab(QtWidgets.QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        cfg = load_lan_config()

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(12)

        title = QtWidgets.QLabel("Kết nối tới máy chủ chia sẻ dữ liệu qua mạng LAN")
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        note = QtWidgets.QLabel(
            "Dùng khi có 1 máy khác trong cùng mạng LAN nội bộ (không qua Internet/cloud) đã "
            "được cài làm \"máy chủ\" chia sẻ (xem gói cài đặt máy chủ riêng, không nằm trong "
            "ứng dụng này). Nhập địa chỉ máy chủ rồi lưu + khởi động lại để mọi thao tác (nhập "
            "Excel, lọc trùng, gộp, truy vấn SQL...) đọc/ghi trực tiếp qua mạng.\n\n"
            "CẢNH BÁO: chế độ này hiện KHÔNG yêu cầu mật khẩu — bất kỳ máy nào trong cùng mạng "
            "LAN cũng đọc/ghi được dữ liệu bệnh nhân qua cổng mạng của máy chủ. Chỉ dùng trong "
            "mạng đáng tin cậy (không có Wi-Fi khách).")
        note.setWordWrap(True)
        layout.addWidget(note)

        self.role_group = QtWidgets.QButtonGroup(self)
        self.rb_single = QtWidgets.QRadioButton("Một máy (mặc định, không chia sẻ)")
        self.rb_client = QtWidgets.QRadioButton("Máy trạm — kết nối tới máy chủ")
        for rb in (self.rb_single, self.rb_client):
            self.role_group.addButton(rb)
            layout.addWidget(rb)

        client_row = QtWidgets.QHBoxLayout()
        client_row.addWidget(QtWidgets.QLabel("Địa chỉ máy chủ:"))
        self.server_url_edit = QtWidgets.QLineEdit(cfg.get("server_url", ""))
        self.server_url_edit.setPlaceholderText("vd: http://192.168.1.10:8765")
        client_row.addWidget(self.server_url_edit, stretch=1)
        test_btn = QtWidgets.QPushButton("Kiểm tra kết nối")
        test_btn.clicked.connect(self.test_connection)
        client_row.addWidget(test_btn)
        layout.addLayout(client_row)

        save_btn = QtWidgets.QPushButton("Lưu cài đặt && khởi động lại ứng dụng")
        save_btn.setObjectName("PrimaryButton")
        save_btn.clicked.connect(self.save_and_restart)
        layout.addWidget(save_btn, alignment=QtCore.Qt.AlignmentFlag.AlignLeft)

        self.status_label = QtWidgets.QLabel()
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        layout.addStretch(1)

        role = cfg.get("role", "single")
        {"single": self.rb_single, "client": self.rb_client}.get(role, self.rb_single).setChecked(True)
        self._refresh_status(cfg)

    def _refresh_status(self, cfg):
        role = cfg.get("role", "single")
        if role == "client":
            self.status_label.setText(
                f"Đang ở chế độ MÁY TRẠM, kết nối tới: {cfg.get('server_url', '(chưa đặt)')}")
        else:
            self.status_label.setText("Đang ở chế độ MỘT MÁY (không chia sẻ qua mạng).")

    def test_connection(self):
        url = self.server_url_edit.text().strip()
        if not url:
            QtWidgets.QMessageBox.warning(self, "Thiếu địa chỉ", "Vui lòng nhập địa chỉ máy chủ.")
            return
        if not url.startswith("http"):
            url = "http://" + url
        ok, info = netclient.ping(url)
        if ok:
            QtWidgets.QMessageBox.information(
                self, "Kết nối thành công", f"Máy chủ phản hồi OK (phiên bản: {info or 'không rõ'}).")
        else:
            QtWidgets.QMessageBox.critical(self, "Không kết nối được", str(info))

    def save_and_restart(self):
        role = "client" if self.rb_client.isChecked() else "single"

        cfg = {"role": role}
        if role == "client":
            url = self.server_url_edit.text().strip()
            if not url:
                QtWidgets.QMessageBox.warning(self, "Thiếu địa chỉ", "Vui lòng nhập địa chỉ máy chủ.")
                return
            if not url.startswith("http"):
                url = "http://" + url
            cfg["server_url"] = url

        save_lan_config(cfg)
        reply = QtWidgets.QMessageBox.question(
            self, "Khởi động lại",
            "Đã lưu cài đặt. Cần khởi động lại ứng dụng để áp dụng đầy đủ. Khởi động lại ngay?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply == QtWidgets.QMessageBox.StandardButton.Yes:
            restart_app()


class ServerTab(QtWidgets.QWidget):
    """Quan ly may chu chia se mang LAN ngay trong ung dung chinh - chi hien
    khi may nay duoc cai voi vai tro "May chu" (lan_config.json co "port").
    Goi qua HTTP toi chinh Windows Service dang chay tren 127.0.0.1 (cac op
    "admin_*" cua netserver.py) - dung lai dung ham netclient.admin_* ma
    server_tray.py da dung, chi khac giao dien (PyQt6 thay vi tkinter) de
    khong bat buoc phai mo tray rieng moi xem/quan ly duoc."""

    REFRESH_MS = 10_000

    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        cfg = load_lan_config()
        self.port = cfg.get("port", 8765)
        self._session_by_row = {}

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(12)

        title = QtWidgets.QLabel("Quản lý máy chủ chia sẻ mạng LAN")
        title.setObjectName("SectionTitle")
        layout.addWidget(title)

        status_row = QtWidgets.QHBoxLayout()
        self.status_label = QtWidgets.QLabel()
        status_row.addWidget(self.status_label, stretch=1)
        self.start_btn = QtWidgets.QPushButton("Bật chia sẻ")
        self.start_btn.clicked.connect(self.start_service)
        status_row.addWidget(self.start_btn)
        self.stop_btn = QtWidgets.QPushButton("Dừng chia sẻ")
        self.stop_btn.clicked.connect(self.stop_service)
        status_row.addWidget(self.stop_btn)
        layout.addLayout(status_row)

        addr_row = QtWidgets.QHBoxLayout()
        addr_row.addWidget(QtWidgets.QLabel("Địa chỉ cung cấp cho máy trạm:"))
        self.addr_edit = QtWidgets.QLineEdit()
        self.addr_edit.setReadOnly(True)
        addr_row.addWidget(self.addr_edit, stretch=1)
        copy_btn = QtWidgets.QPushButton("Sao chép")
        copy_btn.clicked.connect(self.copy_address)
        addr_row.addWidget(copy_btn)
        open_btn = QtWidgets.QPushButton("Mở trong trình duyệt")
        open_btn.clicked.connect(self.open_address)
        addr_row.addWidget(open_btn)
        layout.addLayout(addr_row)

        if not HAS_WIN32_SERVICE:
            warn = QtWidgets.QLabel(
                "Không tìm thấy thư viện pywin32 trên máy này — chỉ xem được trạng thái/kết "
                "nối, không bật/dừng được dịch vụ từ tab này (dùng Windows Services "
                "(services.msc) hoặc biểu tượng khay hệ thống thay thế).")
            warn.setWordWrap(True)
            layout.addWidget(warn)
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(False)

        conn_title = QtWidgets.QLabel("Kết nối đang hoạt động")
        conn_title.setObjectName("SectionTitle")
        layout.addWidget(conn_title)

        self.conn_table = QtWidgets.QTableWidget(0, 3)
        self.conn_table.setHorizontalHeaderLabels(["Địa chỉ IP", "Kết nối lúc", "Rảnh (giây)"])
        self.conn_table.horizontalHeader().setStretchLastSection(True)
        self.conn_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.conn_table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        self.conn_table.setMaximumHeight(180)
        layout.addWidget(self.conn_table)

        conn_btn_row = QtWidgets.QHBoxLayout()
        refresh_btn = QtWidgets.QPushButton("Làm mới")
        refresh_btn.clicked.connect(self.refresh_connections)
        conn_btn_row.addWidget(refresh_btn)
        disconnect_btn = QtWidgets.QPushButton("Ngắt kết nối đã chọn")
        disconnect_btn.clicked.connect(self.disconnect_selected)
        conn_btn_row.addWidget(disconnect_btn)
        conn_btn_row.addStretch(1)
        layout.addLayout(conn_btn_row)

        acl_title = QtWidgets.QLabel("Giới hạn IP được phép kết nối (whitelist)")
        acl_title.setObjectName("SectionTitle")
        layout.addWidget(acl_title)

        self.acl_group = QtWidgets.QButtonGroup(self)
        self.rb_allow_all = QtWidgets.QRadioButton("Cho phép tất cả IP trong mạng LAN (mặc định)")
        self.rb_whitelist = QtWidgets.QRadioButton("Chỉ cho phép các IP trong danh sách dưới đây")
        for rb in (self.rb_allow_all, self.rb_whitelist):
            self.acl_group.addButton(rb)
            layout.addWidget(rb)

        acl_row = QtWidgets.QHBoxLayout()
        self.acl_list = QtWidgets.QListWidget()
        self.acl_list.setMaximumHeight(100)
        acl_row.addWidget(self.acl_list, stretch=1)
        acl_btns = QtWidgets.QVBoxLayout()
        self.acl_ip_edit = QtWidgets.QLineEdit()
        self.acl_ip_edit.setPlaceholderText("vd: 192.168.1.11")
        self.acl_ip_edit.returnPressed.connect(self.add_acl_ip)
        acl_btns.addWidget(self.acl_ip_edit)
        add_ip_btn = QtWidgets.QPushButton("Thêm IP")
        add_ip_btn.clicked.connect(self.add_acl_ip)
        acl_btns.addWidget(add_ip_btn)
        remove_ip_btn = QtWidgets.QPushButton("Xóa IP đã chọn")
        remove_ip_btn.clicked.connect(self.remove_acl_ip)
        acl_btns.addWidget(remove_ip_btn)
        acl_btns.addStretch(1)
        acl_row.addLayout(acl_btns)
        layout.addLayout(acl_row)

        save_acl_btn = QtWidgets.QPushButton("Lưu danh sách IP")
        save_acl_btn.setObjectName("PrimaryButton")
        save_acl_btn.clicked.connect(self.save_acl)
        layout.addWidget(save_acl_btn, alignment=QtCore.Qt.AlignmentFlag.AlignLeft)

        layout.addStretch(1)

        self.timer = QtCore.QTimer(self)
        self.timer.timeout.connect(self.refresh_status)
        self.timer.start(self.REFRESH_MS)

        self.refresh_status()
        self.refresh_connections()
        self.load_acl()

    def _base_url(self):
        return f"http://127.0.0.1:{self.port}"

    def service_status(self):
        if not HAS_WIN32_SERVICE:
            return "unknown"
        try:
            status = win32serviceutil.QueryServiceStatus(SERVER_SERVICE_NAME)[1]
        except Exception:
            return "unknown"
        return "running" if status == win32service.SERVICE_RUNNING else "stopped"

    def refresh_status(self):
        st = self.service_status()
        ip = netserver.get_local_ip()
        self.addr_edit.setText(f"http://{ip}:{self.port}")
        if st == "running":
            self.status_label.setText("● Dịch vụ đang CHIA SẺ dữ liệu.")
        elif st == "stopped":
            self.status_label.setText("● Dịch vụ đang DỪNG.")
        else:
            self.status_label.setText("● Không rõ trạng thái dịch vụ (chưa cài đặt, hoặc thiếu pywin32).")
        self.refresh_connections()

    def start_service(self):
        try:
            win32serviceutil.StartService(SERVER_SERVICE_NAME)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Không bật được dịch vụ",
                f"{e}\n\nCó thể cần chạy ứng dụng với quyền Administrator.")
        self.refresh_status()

    def stop_service(self):
        try:
            win32serviceutil.StopService(SERVER_SERVICE_NAME)
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Không dừng được dịch vụ",
                f"{e}\n\nCó thể cần chạy ứng dụng với quyền Administrator.")
        self.refresh_status()

    def copy_address(self):
        QtWidgets.QApplication.clipboard().setText(self.addr_edit.text())

    def open_address(self):
        QtGui.QDesktopServices.openUrl(QtCore.QUrl(self.addr_edit.text()))

    def refresh_connections(self):
        api_key = core.get_lan_api_key()
        try:
            conns = netclient.admin_list_connections(self._base_url(), api_key)
        except Exception:
            self.conn_table.setRowCount(0)
            self._session_by_row = {}
            return
        self.conn_table.setRowCount(len(conns))
        self._session_by_row = {}
        for row, c in enumerate(conns):
            since = datetime.datetime.fromtimestamp(c["connected_at"]).strftime("%H:%M:%S %d/%m/%Y")
            self.conn_table.setItem(row, 0, QtWidgets.QTableWidgetItem(c["ip"]))
            self.conn_table.setItem(row, 1, QtWidgets.QTableWidgetItem(since))
            self.conn_table.setItem(row, 2, QtWidgets.QTableWidgetItem(str(c["idle_seconds"])))
            self._session_by_row[row] = c["session_id"]

    def disconnect_selected(self):
        rows = sorted({idx.row() for idx in self.conn_table.selectedIndexes()})
        if not rows:
            return
        api_key = core.get_lan_api_key()
        for row in rows:
            sid = self._session_by_row.get(row)
            if not sid:
                continue
            try:
                netclient.admin_disconnect(self._base_url(), api_key, session_id=sid)
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Lỗi", f"Không ngắt được kết nối:\n{e}")
                return
        self.refresh_connections()

    def load_acl(self):
        api_key = core.get_lan_api_key()
        try:
            cfg = netclient.admin_get_acl(self._base_url(), api_key)
        except Exception:
            cfg = {"mode": "allow_all", "allowed_ips": []}
        (self.rb_whitelist if cfg.get("mode") == "whitelist" else self.rb_allow_all).setChecked(True)
        self.acl_list.clear()
        for ip in cfg.get("allowed_ips") or []:
            self.acl_list.addItem(ip)

    def add_acl_ip(self):
        ip = self.acl_ip_edit.text().strip()
        if ip:
            self.acl_list.addItem(ip)
            self.acl_ip_edit.clear()

    def remove_acl_ip(self):
        for item in self.acl_list.selectedItems():
            self.acl_list.takeItem(self.acl_list.row(item))

    def save_acl(self):
        mode = "whitelist" if self.rb_whitelist.isChecked() else "allow_all"
        allowed_ips = [self.acl_list.item(i).text() for i in range(self.acl_list.count())]
        if mode == "whitelist" and not allowed_ips:
            QtWidgets.QMessageBox.warning(
                self, "Thiếu IP",
                "Danh sách IP đang trống - nếu lưu ở chế độ này, KHÔNG máy trạm nào "
                "kết nối được (trừ chính máy chủ). Thêm ít nhất 1 IP hoặc chọn "
                "\"Cho phép tất cả IP\".")
            return
        api_key = core.get_lan_api_key()
        try:
            netclient.admin_set_acl(self._base_url(), api_key, mode=mode, allowed_ips=allowed_ips)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi", f"Không lưu được cấu hình:\n{e}")
            return
        QtWidgets.QMessageBox.information(self, "Đã lưu", "Đã cập nhật danh sách IP được phép.")


# ------------------------------------------------------------------
# Cua so chinh
# ------------------------------------------------------------------

class UpdateCheckWorker(QtCore.QThread):
    """Kiem tra ban cap nhat tren nen, khong chan giao dien. Chi phat tin hieu
    khi thuc su co ban moi hon; im lang neu khong co token / khong co mang."""
    result = QtCore.pyqtSignal(str)

    def run(self):
        local = get_local_version()
        remote, _url = check_latest_release()
        if not remote or not local:
            return
        remote_clean = remote.lstrip("vV")
        if remote_clean != local:
            self.result.emit(remote_clean)


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Quản lý & Lọc trùng danh sách bệnh nhân Bệnh không lây nhiễm (KLN)")
        self.resize(1280, 780)

        init_db()
        self._build_menu_bar()

        central = QtWidgets.QWidget()
        central_layout = QtWidgets.QVBoxLayout(central)
        central_layout.setContentsMargins(0, 0, 0, 0)
        central_layout.setSpacing(0)

        self.update_banner = QtWidgets.QWidget()
        self.update_banner.setObjectName("UpdateBanner")
        self.update_banner.setVisible(False)
        banner_layout = QtWidgets.QHBoxLayout(self.update_banner)
        banner_layout.setContentsMargins(14, 8, 14, 8)
        self.update_banner_label = QtWidgets.QLabel()
        banner_layout.addWidget(self.update_banner_label)
        banner_layout.addStretch(1)
        banner_close_btn = QtWidgets.QPushButton("Đóng")
        banner_close_btn.clicked.connect(lambda: self.update_banner.setVisible(False))
        banner_layout.addWidget(banner_close_btn)
        central_layout.addWidget(self.update_banner)

        tabs = QtWidgets.QTabWidget()
        self.import_tab = ImportTab(self)
        self.data_tab = DataTab(self)
        self.dedup_tab = DedupTab(self)
        self.sql_tab = SqlTab(self)
        self.stats_tab = StatsTab(self)

        tabs.addTab(self.import_tab, "Nhập dữ liệu")
        tabs.addTab(self.data_tab, "Danh sách")
        tabs.addTab(self.dedup_tab, "Lọc trùng")
        tabs.addTab(self.sql_tab, "Truy vấn SQL")
        tabs.addTab(self.stats_tab, "Thống kê")

        # May duoc cai vai tro May chu (lan_config.json co "port") luon doc/
        # ghi CSDL cuc bo cua chinh no ("role": "single") - khong can chon gi
        # o tab "Mạng LAN" nua (khong co gi de chon: khong the tu ket noi toi
        # chinh minh qua mang, va cung khong the doi sang vai tro khac tu
        # trong app). Thay tab do bang tab "Máy chủ" (quan ly viec chia se).
        self.network_tab = None
        self.server_tab = None
        if load_lan_config().get("port"):
            self.server_tab = ServerTab(self)
            tabs.addTab(self.server_tab, "Máy chủ")
        else:
            self.network_tab = NetworkTab(self)
            tabs.addTab(self.network_tab, "Mạng LAN")

        central_layout.addWidget(tabs)
        self.setCentralWidget(central)

        self.status_label = QtWidgets.QLabel()
        self.statusBar().addWidget(self.status_label)
        self.refresh_status()

        self.update_worker = UpdateCheckWorker()
        self.update_worker.result.connect(self.on_update_available)
        self.update_worker.start()

    def refresh_status(self):
        n = record_count()
        if core.is_remote():
            target = f"Máy trạm — máy chủ: {core.REMOTE_BASE_URL}"
        else:
            target = f"CSDL: {DB_PATH}"
        self.status_label.setText(f"  {target}   |   Tổng số bản ghi: {n:,}")

    def on_data_changed(self):
        self.refresh_status()
        self.data_tab.reload()
        self.stats_tab.mark_stale()

    def on_update_available(self, remote_version):
        self.update_banner_label.setText(
            f"Có bản cập nhật mới: v{remote_version}. Mở Start Menu → \"Kiểm tra cập nhật\" "
            "(hoặc chạy update.bat trong thư mục cài đặt) để cập nhật.")
        self.update_banner.setVisible(True)

    # ------------------------------------------------------------------
    # Thanh menu "Trợ giúp": kiểm tra/chạy cập nhật, hướng dẫn sử dụng.
    # ------------------------------------------------------------------

    def _build_menu_bar(self):
        help_menu = self.menuBar().addMenu("&Trợ giúp")

        check_update_action = QtGui.QAction("Kiểm tra cập nhật ngay", self)
        check_update_action.triggered.connect(self.check_update_now)
        help_menu.addAction(check_update_action)

        run_update_action = QtGui.QAction("Chạy cập nhật (update.bat)", self)
        run_update_action.triggered.connect(self.run_update_script)
        help_menu.addAction(run_update_action)

        help_menu.addSeparator()

        readme_action = QtGui.QAction("Hướng dẫn sử dụng (README)", self)
        readme_action.triggered.connect(self.open_readme)
        help_menu.addAction(readme_action)

        about_action = QtGui.QAction("Giới thiệu", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def check_update_now(self):
        self.setCursor(QtCore.Qt.CursorShape.WaitCursor)
        try:
            local = get_local_version()
            remote, _url = check_latest_release()
        finally:
            self.unsetCursor()

        if not remote:
            QtWidgets.QMessageBox.information(
                self, "Kiểm tra cập nhật",
                "Không kiểm tra được (không có mạng, hoặc chưa có bản phát hành nào).")
            return

        remote_clean = remote.lstrip("vV")
        if local and remote_clean == local:
            QtWidgets.QMessageBox.information(
                self, "Kiểm tra cập nhật", f"Bạn đang dùng bản mới nhất (v{local}).")
            return

        reply = QtWidgets.QMessageBox.question(
            self, "Có bản cập nhật mới",
            f"Phiên bản hiện tại: v{local or '?'}\nPhiên bản mới nhất: v{remote_clean}\n\n"
            "Chạy cập nhật ngay bây giờ?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply == QtWidgets.QMessageBox.StandardButton.Yes:
            self.run_update_script()

    def run_update_script(self):
        update_bat = os.path.join(BASE_DIR, "update.bat")
        if not os.path.exists(update_bat):
            QtWidgets.QMessageBox.warning(
                self, "Không tìm thấy update.bat",
                "Chỉ dùng được khi cài từ bộ cài đặt (Setup.exe) — không có khi chạy trực "
                "tiếp từ mã nguồn.")
            return

        is_server = bool(load_lan_config().get("port"))
        try:
            if sys.platform == "win32":
                # Vai tro May chu can quyen Administrator de dung/bat lai
                # Windows Service khi cap nhat - xem update.ps1.
                os.startfile(update_bat, "runas" if is_server else "open")
            else:
                subprocess.Popen(["bash", update_bat])
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Lỗi", f"Không chạy được update.bat:\n{e}")
            return

        QtWidgets.QMessageBox.information(
            self, "Đang cập nhật",
            "Đã mở cửa sổ cập nhật (update.bat). Làm theo hướng dẫn trong cửa sổ đó, "
            "sau đó mở lại ứng dụng.")

    def open_readme(self):
        readme = os.path.join(BASE_DIR, "README.md")
        if os.path.exists(readme):
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(readme))
        else:
            QtWidgets.QMessageBox.information(
                self, "Không tìm thấy README.md",
                "File hướng dẫn không có trong thư mục cài đặt.")

    def show_about(self):
        local = get_local_version() or "?"
        QtWidgets.QMessageBox.about(
            self, "Giới thiệu",
            "<b>Quản lý &amp; Lọc trùng danh sách bệnh nhân THA</b><br>"
            f"Phiên bản: v{local}<br>"
            "Nhà phát triển: Monsterph6<br>"
            "<a href='https://github.com/Monsterph6/QuanlybenhKLN'>"
            "github.com/Monsterph6/QuanlybenhKLN</a>")


STYLE_SHEET = """
QWidget {
    background: #f4f6fb;
    color: #1f2430;
    font-family: "Segoe UI", "Inter", sans-serif;
    font-size: 13px;
}
QMainWindow, QStatusBar {
    background: #f4f6fb;
}
QLabel#SectionTitle {
    font-size: 15px;
    font-weight: 600;
    color: #16233d;
}
QTabWidget::pane {
    border: 1px solid #dfe4ee;
    border-radius: 10px;
    background: #ffffff;
    top: -1px;
}
QTabBar::tab {
    background: transparent;
    color: #5b6472;
    padding: 9px 20px;
    margin-right: 4px;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    font-weight: 500;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: #1f5eff;
    border: 1px solid #dfe4ee;
    border-bottom: none;
}
QTabBar::tab:hover:!selected {
    color: #1f2430;
}
QPushButton {
    background: #ffffff;
    border: 1px solid #d3d9e6;
    border-radius: 8px;
    padding: 7px 16px;
    color: #1f2430;
}
QPushButton:hover {
    background: #eef2ff;
    border-color: #b8c4ef;
}
QPushButton:pressed {
    background: #e2e9ff;
}
QPushButton#PrimaryButton {
    background: #1f5eff;
    border: 1px solid #1f5eff;
    color: #ffffff;
    font-weight: 600;
}
QPushButton#PrimaryButton:hover {
    background: #1a4fe0;
}
QPushButton#DangerButton {
    background: #ffffff;
    border: 1px solid #f0b4b4;
    color: #c92a2a;
}
QPushButton#DangerButton:hover {
    background: #fff0f0;
}
QLineEdit, QPlainTextEdit, QComboBox {
    background: #ffffff;
    border: 1px solid #d3d9e6;
    border-radius: 7px;
    padding: 5px 8px;
    selection-background-color: #1f5eff;
}
QLineEdit:focus, QPlainTextEdit:focus, QComboBox:focus {
    border: 1px solid #1f5eff;
}
QComboBox::drop-down {
    border: none;
    width: 22px;
}
QTableView {
    background: #ffffff;
    alternate-background-color: #f7f9fd;
    gridline-color: #edf0f7;
    border: 1px solid #dfe4ee;
    border-radius: 8px;
    selection-background-color: #dbe6ff;
    selection-color: #16233d;
}
QHeaderView::section {
    background: #eef1f8;
    color: #46506a;
    padding: 6px 8px;
    border: none;
    border-bottom: 1px solid #dfe4ee;
    font-weight: 600;
}
QProgressBar {
    border: 1px solid #d3d9e6;
    border-radius: 6px;
    background: #eef1f8;
    height: 8px;
    text-align: center;
}
QProgressBar::chunk {
    background: #1f5eff;
    border-radius: 6px;
}
QSplitter::handle {
    background: #eef1f8;
}
QGroupBox {
    border: 1px solid #dfe4ee;
    border-radius: 8px;
    margin-top: 12px;
    padding: 10px 8px 8px 8px;
    font-weight: 600;
    color: #16233d;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 6px;
    color: #46506a;
}
QCheckBox {
    spacing: 6px;
}
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 1px solid #b8c0d4;
    border-radius: 4px;
    background: #ffffff;
}
QCheckBox::indicator:checked {
    background: #1f5eff;
    border-color: #1f5eff;
}
QDialog {
    background: #f4f6fb;
}
QStatusBar {
    border-top: 1px solid #dfe4ee;
    color: #5b6472;
}
QWidget#UpdateBanner {
    background: #fff4d6;
    border-bottom: 1px solid #f0d48a;
}
QWidget#UpdateBanner QLabel {
    color: #6b4f00;
    font-weight: 500;
}
QScrollBar:vertical {
    background: transparent;
    width: 10px;
}
QScrollBar::handle:vertical {
    background: #c7cfe0;
    border-radius: 5px;
    min-height: 24px;
}
QScrollBar::handle:vertical:hover {
    background: #a9b4cc;
}
"""


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLE_SHEET)

    lan_cfg = load_lan_config()
    if lan_cfg.get("role") == "client":
        core.configure_remote(lan_cfg.get("server_url", ""), lan_cfg.get("api_key", ""))

    if has_password():
        login = LoginDialog()
        if login.exec() != QtWidgets.QDialog.DialogCode.Accepted:
            sys.exit(0)

    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
