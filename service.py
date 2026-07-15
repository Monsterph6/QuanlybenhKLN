# -*- coding: utf-8 -*-
"""
Windows Service chia se benh_nhan.db qua mang LAN noi bo. Day la thanh
phan "may chu", TACH RIENG khoi app.py (ung dung chinh danh cho may
tram) - khong dung PyQt6 nen nhe hon nhieu, chay ngam hoan toan (khong
giao dien), tu khoi dong cung Windows va tu restart neu bi crash, chay
duoc ngay ca khi chua ai dang nhap vao may.

Can cai thu vien pywin32 (xem requirements-server.txt):
    pip install -r requirements-server.txt

Cai dat / go cai dat (PHAI mo Command Prompt hoac PowerShell voi quyen
Administrator - chuot phai bieu tuong, chon "Run as administrator"):

    python service.py --startup auto install
    python service.py start

    python service.py stop
    python service.py remove

Xem trang thai qua Windows "Services" (services.msc), ten dich vu:
"QuanLyBenhNhanTHA_Server". Muon xem trang thai / dia chi IP tien loi
hon (khong can quyen Administrator, khong can mo services.msc) thi chay
server_tray.py - xem file do.

Cong chia se doc tu file lan_config.json (cung thu muc voi file nay),
mac dinh 8765 neu chua co file - xem core.load_lan_config(). Sua file
nay bang Notepad roi chay lai "python service.py restart" de doi cong.
"""
import sys

import servicemanager
import win32event
import win32service
import win32serviceutil

import core
import netserver


class QuanLyBenhNhanTHAService(win32serviceutil.ServiceFramework):
    _svc_name_ = "QuanLyBenhNhanTHA_Server"
    _svc_display_name_ = "Quan ly benh nhan THA - May chu chia se LAN"
    _svc_description_ = (
        "Chia se file benh_nhan.db cho cac may tram trong cung mang LAN noi bo "
        "(khong dung Internet). Thanh phan may chu cua ung dung Quan ly & Loc "
        "trung danh sach benh nhan THA."
    )

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        netserver.stop_server()
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):
        cfg = core.load_lan_config()
        port = cfg.get("port", 8765)
        core.init_db()
        netserver.start_server(port)
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, f" - dang chia se tai cong {port}"),
        )
        win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)


if __name__ == "__main__":
    if len(sys.argv) == 1:
        # Duoc Windows Service Control Manager tu goi khi khoi dong dich vu.
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(QuanLyBenhNhanTHAService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        # Goi tu dong lenh: install / start / stop / remove / restart / debug ...
        win32serviceutil.HandleCommandLine(QuanLyBenhNhanTHAService)
