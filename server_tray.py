# -*- coding: utf-8 -*-
"""
Tray helper cho may chu - CHUONG TRINH RIENG voi service.py. Chuong trinh
nay KHONG tu chay viec chia se (viec do la cua Windows Service that su
trong service.py, chay ngam ngay ca khi chua ai dang nhap) - day chi la
"bang dieu khien" nho chay trong phien dang nhap cua nguoi dung, hien
icon o khay he thong de:
  - xem dich vu dang chay hay dang dung, va dia chi IP:cong hien tai
  - bat / dung dich vu (can quyen Administrator - Windows se hoi neu
    tai khoan dang dung khong co san quyen do)
  - bat/tat tuy chon "khoi dong cung Windows" CHO CHINH TRAY NAY (dich
    vu that su tu khoi dong cung may qua Windows Service, khong phu
    thuoc vao tray - xem install_server.bat)
  - xem danh sach may tram dang ket noi toi, ngat 1 ket noi cu the, va
    quan ly danh sach IP duoc phep ket noi (whitelist) - cac thao tac
    nay goi qua HTTP toi chinh dich vu dang chay tren 127.0.0.1 (2 tien
    trinh khac nhau, khong dung chung bo nho) - xem netserver.py op
    "admin_*" va core.load_acl_config()/save_acl_config().

Dong cua so tray (nut "Thoat") KHONG lam dung viec chia se dang chay.

Can thu vien pystray + Pillow (xem requirements-server.txt). Cac cua so
quan ly ket noi/whitelist dung tkinter (co san trong thu vien chuan
Python, khong can cai them).
"""
import datetime
import sys
import threading
import tkinter as tk
import webbrowser
import winreg
from tkinter import messagebox, ttk

import pystray
import win32service
import win32serviceutil
from PIL import Image, ImageDraw

import core
import netclient
import netserver
from service import QuanLyBenhNhanTHAService

SERVICE_NAME = QuanLyBenhNhanTHAService._svc_name_
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_VALUE_NAME = "QuanLyBenhNhanTHA_ServerTray"
REFRESH_SECONDS = 10

icon = None
_stop_flag = threading.Event()


def _make_dot_icon(rgba):
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    ImageDraw.Draw(img).ellipse((8, 8, 56, 56), fill=rgba)
    return img


ICON_RUNNING = _make_dot_icon((34, 139, 34, 255))   # xanh la - dang chia se
ICON_STOPPED = _make_dot_icon((201, 42, 42, 255))   # do - dang dung
ICON_UNKNOWN = _make_dot_icon((150, 150, 150, 255))  # xam - chua ro / chua cai dat


def service_status():
    """Tra ve 'running' / 'stopped' / 'unknown' (vd: chua cai dat dich vu)."""
    try:
        status = win32serviceutil.QueryServiceStatus(SERVICE_NAME)[1]
    except Exception:
        return "unknown"
    return "running" if status == win32service.SERVICE_RUNNING else "stopped"


def is_autostart_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
            winreg.QueryValueEx(key, RUN_VALUE_NAME)
            return True
    except FileNotFoundError:
        return False


def enable_autostart():
    if getattr(sys, "frozen", False):
        cmd = f'"{sys.executable}"'
    else:
        cmd = f'"{sys.executable}" "{__file__}"'
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as key:
        winreg.SetValueEx(key, RUN_VALUE_NAME, 0, winreg.REG_SZ, cmd)


def disable_autostart():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, RUN_VALUE_NAME)
    except FileNotFoundError:
        pass


def _notify(message):
    if icon is not None:
        try:
            icon.notify(message, "Quản lý bệnh nhân THA - Máy chủ")
        except Exception:
            pass  # khong phai backend nao cung ho tro notify, bo qua neu loi


def start_service(_item=None):
    try:
        win32serviceutil.StartService(SERVICE_NAME)
    except Exception as e:
        _notify(f"Không bật được dịch vụ (cần quyền Administrator?): {e}")


def stop_service(_item=None):
    try:
        win32serviceutil.StopService(SERVICE_NAME)
    except Exception as e:
        _notify(f"Không dừng được dịch vụ (cần quyền Administrator?): {e}")


def open_address(_item=None):
    cfg = core.load_lan_config()
    port = cfg.get("port", 8765)
    ip = netserver.get_local_ip()
    webbrowser.open(f"http://{ip}:{port}")


def toggle_autostart(_item=None):
    if is_autostart_enabled():
        disable_autostart()
    else:
        enable_autostart()
    if icon is not None:
        icon.update_menu()


def quit_tray(_item=None):
    if icon is not None:
        icon.stop()


def _status_text(_item=None):
    st = service_status()
    if st == "running":
        cfg = core.load_lan_config()
        port = cfg.get("port", 8765)
        ip = netserver.get_local_ip()
        return f"Đang chia sẻ tại http://{ip}:{port}"
    if st == "stopped":
        return "Dịch vụ đang DỪNG"
    return "Không tìm thấy dịch vụ (chưa cài đặt?)"


def _icon_for_status():
    st = service_status()
    if st == "running":
        return ICON_RUNNING
    if st == "stopped":
        return ICON_STOPPED
    return ICON_UNKNOWN


def _refresh_loop():
    while not _stop_flag.is_set():
        if icon is not None:
            icon.icon = _icon_for_status()
            icon.title = _status_text()
        _stop_flag.wait(REFRESH_SECONDS)


# ------------------------------------------------------------------
# Cua so quan ly: ket noi dang hoat dong + whitelist IP. Goi qua HTTP toi
# chinh dich vu dang chay tren 127.0.0.1 (netserver.py o tien trinh khac).
# Cac ham nay chay trong luong (thread) callback cua pystray - tao rieng
# 1 cua so Tk() moi lan mo, dong xong huy luon, khong dung chung 1 root
# voi vong lap chinh cua tray.
# ------------------------------------------------------------------

def _admin_base_url():
    cfg = core.load_lan_config()
    port = cfg.get("port", 8765)
    return f"http://127.0.0.1:{port}"


def _run_dialog(build_fn):
    root = tk.Tk()
    root.title("Quản lý bệnh nhân THA - Máy chủ")
    try:
        build_fn(root)
        root.mainloop()
    finally:
        try:
            root.destroy()
        except Exception:
            pass


def show_connections(_item=None):
    if service_status() != "running":
        _notify("Dịch vụ chưa chạy, không có kết nối nào để xem.")
        return
    _run_dialog(_build_connections_window)


def _build_connections_window(root):
    root.title("Kết nối đang hoạt động")
    root.geometry("560x340")
    base_url = _admin_base_url()
    api_key = core.get_lan_api_key()
    session_by_iid = {}

    tk.Label(root, text="Các máy trạm đang kết nối tới máy chủ này:").pack(anchor="w", padx=8, pady=(8, 0))

    tree = ttk.Treeview(root, columns=("ip", "since", "idle"), show="headings")
    tree.heading("ip", text="Địa chỉ IP")
    tree.heading("since", text="Kết nối lúc")
    tree.heading("idle", text="Rảnh (giây)")
    tree.column("ip", width=180)
    tree.column("since", width=200)
    tree.column("idle", width=100)
    tree.pack(fill="both", expand=True, padx=8, pady=8)

    def refresh():
        for iid in tree.get_children():
            tree.delete(iid)
        session_by_iid.clear()
        try:
            conns = netclient.admin_list_connections(base_url, api_key)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không lấy được danh sách kết nối:\n{e}", parent=root)
            return
        for c in conns:
            since = datetime.datetime.fromtimestamp(c["connected_at"]).strftime("%H:%M:%S %d/%m/%Y")
            iid = tree.insert("", "end", values=(c["ip"], since, c["idle_seconds"]))
            session_by_iid[iid] = c["session_id"]

    def disconnect_selected():
        sel = tree.selection()
        if not sel:
            return
        sid = session_by_iid.get(sel[0])
        if not sid:
            return
        try:
            netclient.admin_disconnect(base_url, api_key, session_id=sid)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không ngắt được kết nối:\n{e}", parent=root)
            return
        refresh()

    btns = tk.Frame(root)
    btns.pack(fill="x", padx=8, pady=(0, 8))
    tk.Button(btns, text="Làm mới", command=refresh).pack(side="left")
    tk.Button(btns, text="Ngắt kết nối đã chọn", command=disconnect_selected).pack(side="left", padx=(8, 0))

    refresh()


def show_acl(_item=None):
    if service_status() != "running":
        _notify("Dịch vụ chưa chạy, không thể quản lý danh sách IP.")
        return
    _run_dialog(_build_acl_window)


def _build_acl_window(root):
    root.title("Quản lý IP được phép kết nối")
    root.geometry("440x400")
    base_url = _admin_base_url()
    api_key = core.get_lan_api_key()

    try:
        cfg = netclient.admin_get_acl(base_url, api_key)
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không lấy được cấu hình:\n{e}", parent=root)
        cfg = {"mode": "allow_all", "allowed_ips": []}

    mode_var = tk.StringVar(value=cfg.get("mode", "allow_all"))

    tk.Radiobutton(root, text="Cho phép tất cả IP trong mạng LAN (mặc định)",
                   variable=mode_var, value="allow_all").pack(anchor="w", padx=8, pady=(10, 0))
    tk.Radiobutton(root, text="Chỉ cho phép các IP trong danh sách dưới đây",
                   variable=mode_var, value="whitelist").pack(anchor="w", padx=8)

    tk.Label(root, text="Danh sách IP được phép:").pack(anchor="w", padx=8, pady=(8, 0))
    listbox = tk.Listbox(root)
    for ip in cfg.get("allowed_ips") or []:
        listbox.insert("end", ip)
    listbox.pack(fill="both", expand=True, padx=8, pady=4)

    entry_row = tk.Frame(root)
    entry_row.pack(fill="x", padx=8)
    entry = tk.Entry(entry_row)
    entry.pack(side="left", fill="x", expand=True)

    def add_ip():
        ip = entry.get().strip()
        if ip:
            listbox.insert("end", ip)
            entry.delete(0, "end")

    def remove_selected():
        for i in reversed(listbox.curselection()):
            listbox.delete(i)

    entry.bind("<Return>", lambda _e: add_ip())
    tk.Button(entry_row, text="Thêm", command=add_ip).pack(side="left", padx=(4, 0))

    btns = tk.Frame(root)
    btns.pack(fill="x", padx=8, pady=(4, 0))
    tk.Button(btns, text="Xóa IP đã chọn", command=remove_selected).pack(side="left")

    def save():
        allowed_ips = list(listbox.get(0, "end"))
        if mode_var.get() == "whitelist" and not allowed_ips:
            messagebox.showwarning(
                "Thiếu IP",
                "Danh sách IP đang trống - nếu lưu ở chế độ này, KHÔNG máy trạm nào "
                "kết nối được (trừ chính máy chủ). Thêm ít nhất 1 IP hoặc chọn "
                "\"Cho phép tất cả IP\".",
                parent=root)
            return
        try:
            netclient.admin_set_acl(base_url, api_key, mode=mode_var.get(), allowed_ips=allowed_ips)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không lưu được cấu hình:\n{e}", parent=root)
            return
        messagebox.showinfo("Đã lưu", "Đã cập nhật danh sách IP được phép.", parent=root)
        root.destroy()

    tk.Button(root, text="Lưu", command=save).pack(pady=10)


def main():
    global icon
    menu = pystray.Menu(
        pystray.MenuItem(_status_text, None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Mở địa chỉ máy chủ trong trình duyệt", open_address),
        pystray.MenuItem("Kết nối đang hoạt động...", show_connections),
        pystray.MenuItem("Quản lý IP được phép kết nối...", show_acl),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Bật chia sẻ", start_service),
        pystray.MenuItem("Dừng chia sẻ", stop_service),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Khởi động cùng Windows", toggle_autostart,
                          checked=lambda item: is_autostart_enabled()),
        pystray.MenuItem("Thoát (không dừng chia sẻ)", quit_tray),
    )
    icon = pystray.Icon("QuanLyBenhNhanTHA_Tray", ICON_UNKNOWN, "Quản lý bệnh nhân THA", menu)
    threading.Thread(target=_refresh_loop, daemon=True).start()
    icon.run()
    _stop_flag.set()


if __name__ == "__main__":
    main()
