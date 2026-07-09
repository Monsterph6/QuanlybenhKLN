# -*- coding: utf-8 -*-
"""
Ung dung quan ly & loc trung danh sach benh nhan THA - giao dien PyQt6.
Tang du lieu (SQLite, doc Excel, xuat CSV/Excel) nam trong core.py.
"""
import os
import sys
import sqlite3

from PyQt6 import QtCore, QtGui, QtWidgets

import core
from core import (
    BASE_DIR, DB_PATH, COLUMNS, DEDUP_FIELDS, PAGE_SIZE,
    get_conn, init_db, record_count, import_excel,
    write_export, export_query_to_file,
    count_swapped_gender_birthdate, fix_swapped_gender_birthdate,
    build_dedup_key, scan_dedup_groups, group_detail_rows,
    delete_patients_by_ids,
    merge_specific_ids, merge_group, merge_all_groups,
    unique_rows_with_optional_history,
    add_dedup_exception, remove_dedup_exception, list_dedup_exceptions,
)

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

class ImportWorker(QtCore.QThread):
    finished_ok = QtCore.pyqtSignal(int, int, int)
    failed = QtCore.pyqtSignal(str)

    def __init__(self, path):
        super().__init__()
        self.path = path

    def run(self):
        try:
            total, inserted, skipped = import_excel(self.path)
            self.finished_ok.emit(total, inserted, skipped)
        except Exception as e:
            self.failed.emit(str(e))


class ImportTab(QtWidgets.QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.worker = None

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

        self.progress = QtWidgets.QProgressBar()
        self.progress.setVisible(False)
        self.progress.setRange(0, 0)
        layout.addWidget(self.progress)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log, stretch=1)

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
        row2.addStretch(1)
        layout.addLayout(row2)

        default_xlsx = self._find_default_xlsx()
        if default_xlsx:
            self.path_edit.setText(default_xlsx)

    def _find_default_xlsx(self):
        for f in os.listdir(BASE_DIR):
            if f.lower().endswith(".xlsx") and not f.startswith("~$"):
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
        self.import_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.log_line(f"Đang đọc file: {path}")

        self.worker = ImportWorker(path)
        self.worker.finished_ok.connect(self.import_done)
        self.worker.failed.connect(self.import_failed)
        self.worker.start()

    def import_done(self, total, inserted, skipped):
        self.progress.setVisible(False)
        self.import_btn.setEnabled(True)
        self.log_line(f"Đã đọc: {total:,} dòng")
        self.log_line(f"Thêm mới vào CSDL: {inserted:,} bản ghi")
        self.log_line(f"Bỏ qua (dòng dữ liệu giống hệt đã có, tránh nhập trùng khi nhập lại file): {skipped:,}")
        self.log_line("Lưu ý: đây chỉ là bỏ qua dòng nhập lại y hệt. Trùng bệnh nhân (nhiều lượt khám) "
                       "vẫn được giữ nguyên — dùng tab 'Lọc trùng' để xử lý.")
        self.log_line("-" * 60)
        self.main_window.on_data_changed()

    def import_failed(self, msg):
        self.progress.setVisible(False)
        self.import_btn.setEnabled(True)
        self.log_line(f"LỖI: {msg}")
        QtWidgets.QMessageBox.critical(self, "Lỗi khi nhập dữ liệu", msg)

    def reset_db(self):
        reply = QtWidgets.QMessageBox.question(
            self, "Xác nhận",
            "Xóa TOÀN BỘ dữ liệu hiện có trong cơ sở dữ liệu?\nHành động này không thể hoàn tác.",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        conn = get_conn()
        conn.execute("DELETE FROM patients")
        conn.commit()
        conn.close()
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
            "Sửa lại các dòng này trong CSDL (hoán đổi 2 cột về đúng vị trí)?",
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No)
        if reply != QtWidgets.QMessageBox.StandardButton.Yes:
            return
        fixed = fix_swapped_gender_birthdate()
        self.log_line(f"Đã sửa {fixed:,} dòng bị đảo cột Giới tính / Ngày sinh.")
        self.main_window.on_data_changed()


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
        n_g, n_merged = merge_all_groups(
            self.current_key_expr, self.current_key_where, self.current_key_type, order)
        QtWidgets.QMessageBox.information(
            self, "Hoàn tất", f"Đã gộp {n_g:,} nhóm ({n_merged:,} bản ghi được đưa vào lịch sử khám).")
        self.main_window.on_data_changed()
        self.scan()


# ------------------------------------------------------------------
# Tab Truy van SQL
# ------------------------------------------------------------------

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
# Cua so chinh
# ------------------------------------------------------------------

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Quản lý & Lọc trùng danh sách bệnh nhân THA")
        self.resize(1280, 780)

        init_db()

        tabs = QtWidgets.QTabWidget()
        self.import_tab = ImportTab(self)
        self.data_tab = DataTab(self)
        self.dedup_tab = DedupTab(self)
        self.sql_tab = SqlTab(self)

        tabs.addTab(self.import_tab, "Nhập dữ liệu")
        tabs.addTab(self.data_tab, "Danh sách")
        tabs.addTab(self.dedup_tab, "Lọc trùng")
        tabs.addTab(self.sql_tab, "Truy vấn SQL")
        self.setCentralWidget(tabs)

        self.status_label = QtWidgets.QLabel()
        self.statusBar().addWidget(self.status_label)
        self.refresh_status()

    def refresh_status(self):
        n = record_count()
        self.status_label.setText(f"  CSDL: {DB_PATH}   |   Tổng số bản ghi: {n:,}")

    def on_data_changed(self):
        self.refresh_status()
        self.data_tab.reload()


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
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
