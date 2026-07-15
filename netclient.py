# -*- coding: utf-8 -*-
"""
Client ket noi toi netserver.py qua mang LAN. Bat chuoc lai giao dien
sqlite3.Connection/Cursor toi thieu ma core.py can (execute, executemany,
cursor, commit, close, fetchall/fetchone, description, lap qua tung dong),
de core.py va app.py khong phai sua logic doc/ghi khi chay o che do may
tram - xem core.get_conn().

Chi dung thu vien chuan (urllib) - khong them dependency moi.
"""
import json
import sqlite3
import urllib.error
import urllib.request

DEFAULT_TIMEOUT = 30


def _post(base_url, api_key, payload, timeout=DEFAULT_TIMEOUT):
    data = json.dumps(payload).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["X-Api-Key"] = api_key
    req = urllib.request.Request(f"{base_url}/rpc", data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            body = json.loads(e.read().decode("utf-8"))
            msg = body.get("error", str(e))
        except Exception:
            msg = str(e)
        raise sqlite3.Error(f"Máy chủ từ chối yêu cầu: {msg}")
    except urllib.error.URLError as e:
        raise sqlite3.Error(f"Không kết nối được máy chủ ({base_url}): {e.reason}")
    except (TimeoutError, OSError) as e:
        raise sqlite3.Error(f"Hết thời gian chờ máy chủ ({base_url}): {e}")


def ping(base_url, api_key="", timeout=5):
    """Kiem tra ket noi toi may chu. Tra ve (ok, thong_tin_hoac_thong_bao_loi)."""
    try:
        result = _post(base_url, api_key, {"op": "ping"}, timeout=timeout)
        return True, result.get("version")
    except sqlite3.Error as e:
        return False, str(e)


def request_backup(base_url, reason, keep, api_key=""):
    result = _post(base_url, api_key, {"op": "backup", "reason": reason, "keep": keep})
    return result.get("path")


class RemoteRow:
    """Tuong duong sqlite3.Row: truy cap duoc theo ten cot hoac chi so, va
    lap qua cac GIA TRI theo thu tu cot (khong phai lap qua ten cot)."""
    __slots__ = ("_keys", "_values")

    def __init__(self, keys, values):
        self._keys = keys
        self._values = values

    def keys(self):
        return list(self._keys)

    def __getitem__(self, key):
        if isinstance(key, str):
            return self._values[self._keys.index(key)]
        return self._values[key]

    def __iter__(self):
        return iter(self._values)

    def __len__(self):
        return len(self._values)

    def __repr__(self):
        return repr(dict(zip(self._keys, self._values)))


class RemoteCursor:
    def __init__(self, connection):
        self._connection = connection
        self._columns = []
        self._rows = []
        self._pos = 0
        self.rowcount = -1
        self.lastrowid = None

    @property
    def description(self):
        return [(c, None, None, None, None, None, None) for c in self._columns]

    def execute(self, sql, params=None):
        result = self._connection._call({
            "op": "execute",
            "session_id": self._connection.session_id,
            "sql": sql,
            "params": list(params) if params is not None else [],
        })
        self._apply(result)
        return self

    def executemany(self, sql, seq_params):
        result = self._connection._call({
            "op": "executemany",
            "session_id": self._connection.session_id,
            "sql": sql,
            "seq_params": [list(p) for p in seq_params],
        })
        self._apply(result)
        return self

    def _apply(self, result):
        self._columns = result.get("columns") or []
        raw_rows = result.get("rows") or []
        self._rows = [RemoteRow(self._columns, v) for v in raw_rows]
        self._pos = 0
        self.rowcount = result.get("rowcount", -1)
        self.lastrowid = result.get("lastrowid")

    def fetchall(self):
        rows = self._rows[self._pos:]
        self._pos = len(self._rows)
        return rows

    def fetchone(self):
        if self._pos >= len(self._rows):
            return None
        row = self._rows[self._pos]
        self._pos += 1
        return row

    def __iter__(self):
        return iter(self._rows)

    def __len__(self):
        return len(self._rows)


class RemoteConnection:
    """Thay the sqlite3.Connection khi ung dung o che do may tram, chuyen
    tiep moi lenh SQL qua mang toi netserver.py tren may chu. 1 phien (session)
    tren may chu tuong ung 1 doi tuong RemoteConnection - can thiet vi mot
    so ham trong core.py dung nhieu cau lenh (vd bang tam - TEMP TABLE) trong
    cung 1 "ket noi logic"."""

    def __init__(self, base_url, api_key="", timeout=DEFAULT_TIMEOUT):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key
        self.timeout = timeout
        self.row_factory = None  # giu de tuong thich API, khong dung toi
        result = self._call({"op": "connect"})
        self.session_id = result["session_id"]
        self._closed = False

    def _call(self, payload):
        return _post(self.base_url, self.api_key, payload, timeout=self.timeout)

    def cursor(self):
        return RemoteCursor(self)

    def execute(self, sql, params=None):
        return RemoteCursor(self).execute(sql, params)

    def executemany(self, sql, seq_params):
        return RemoteCursor(self).executemany(sql, seq_params)

    def commit(self):
        self._call({"op": "commit", "session_id": self.session_id})

    def close(self):
        if self._closed:
            return
        self._closed = True
        try:
            self._call({"op": "close", "session_id": self.session_id})
        except sqlite3.Error:
            pass  # dong ket noi la best-effort, khong lam vo luong xu ly dang chay
