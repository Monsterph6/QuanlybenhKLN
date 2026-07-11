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

Dong cua so tray (nut "Thoat") KHONG lam dung viec chia se dang chay.

Can thu vien pystray + Pillow (xem requirements-server.txt):
    pip install -r requirements-server.txt
"""
import sys
import threading
import webbrowser
import winreg

import pystray
import win32service
import win32serviceutil
from PIL import Image, ImageDraw

import core
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


def main():
    global icon
    menu = pystray.Menu(
        pystray.MenuItem(_status_text, None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Mở địa chỉ máy chủ trong trình duyệt", open_address),
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
