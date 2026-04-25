"""
Global hotkey: Ctrl+Enter (Win32) with a dedicated message-pump thread.

RegisterHotKey + GetMessageW — no extra dependencies beyond the Windows C API.
"""

from __future__ import annotations

import ctypes
import sys
import threading
from collections.abc import Callable
from ctypes import wintypes

if sys.platform != "win32":

    def hotkey_ctrl_enter_available() -> bool:
        return False

    class CtrlEnterHotkey:
        def __init__(self, on_hotkey: Callable[[], None]) -> None:
            self._on_hotkey = on_hotkey

        @property
        def start_error(self) -> str:
            return ""

        def start(self) -> bool:
            return False

        def stop(self, timeout: float = 4.0) -> None:
            return

else:
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    WM_HOTKEY = 0x0312
    WM_QUIT = 0x0012
    MOD_CONTROL = 0x0002
    VK_RETURN = 0x0D
    PM_REMOVE = 0x0001

    class MSG(ctypes.Structure):
        _fields_ = [
            ("hwnd", wintypes.HWND),
            ("message", wintypes.UINT),
            ("wParam", wintypes.WPARAM),
            ("lParam", wintypes.LPARAM),
            ("time", wintypes.DWORD),
            ("pt", wintypes.POINT),
        ]

    def hotkey_ctrl_enter_available() -> bool:
        return True

    _LPMSG = ctypes.POINTER(MSG)
    _GetMessageW = user32.GetMessageW
    _GetMessageW.argtypes = [_LPMSG, wintypes.HWND, wintypes.UINT, wintypes.UINT]
    _GetMessageW.restype = wintypes.BOOL
    _TranslateMessage = user32.TranslateMessage
    _TranslateMessage.argtypes = [_LPMSG]
    _TranslateMessage.restype = wintypes.BOOL
    _DispatchMessageW = user32.DispatchMessageW
    _DispatchMessageW.argtypes = [_LPMSG]
    _DispatchMessageW.restype = ctypes.c_size_t
    _PostThreadMessageW = user32.PostThreadMessageW
    _PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    _PostThreadMessageW.restype = wintypes.BOOL

    class CtrlEnterHotkey:
        _HOTKEY_ID = 1

        def __init__(self, on_hotkey: Callable[[], None]) -> None:
            self._on_hotkey = on_hotkey
            self._thread: threading.Thread | None = None
            self._thread_id: int | None = None
            self._lock = threading.Lock()
            self._ready = threading.Event()
            self._register_ok: bool = False
            self._start_error: str = ""

        @property
        def start_error(self) -> str:
            return self._start_error

        def _loop(self) -> None:
            self._thread_id = int(kernel32.GetCurrentThreadId())
            if not user32.RegisterHotKey(
                None,
                self._HOTKEY_ID,
                MOD_CONTROL,
                VK_RETURN,
            ):
                self._start_error = f"RegisterHotKey failed (is Ctrl+Enter already in use?) winerr={kernel32.GetLastError()}"
                self._register_ok = False
                self._ready.set()
                return
            self._start_error = ""
            self._register_ok = True
            self._ready.set()
            msg = MSG()
            while True:
                r = int(_GetMessageW(ctypes.byref(msg), None, 0, 0))
                if r == 0:
                    break
                if r == -1:
                    break
                if msg.message == WM_HOTKEY:
                    try:
                        self._on_hotkey()
                    except Exception:
                        pass
                else:
                    _TranslateMessage(ctypes.byref(msg))
                    _DispatchMessageW(ctypes.byref(msg))
            user32.UnregisterHotKey(None, self._HOTKEY_ID)

        def start(self) -> bool:
            with self._lock:
                if self._thread is not None and self._thread.is_alive():
                    return self._register_ok
                self._ready.clear()
                self._register_ok = False
                self._start_error = ""
                self._thread = threading.Thread(target=self._loop, name="html-capture-hotkey", daemon=True)
                self._thread.start()
            if not self._ready.wait(5.0):
                with self._lock:
                    t = self._thread
                if t is not None and t.is_alive() and self._thread_id is not None:
                    _PostThreadMessageW(self._thread_id, WM_QUIT, 0, 0)
                    t.join(timeout=2.0)
                return False
            return self._register_ok

        def stop(self, timeout: float = 4.0) -> None:
            t: threading.Thread | None = None
            tid: int | None = None
            with self._lock:
                t = self._thread
                tid = self._thread_id
            if t is not None and tid and _PostThreadMessageW(tid, WM_QUIT, 0, 0):
                t.join(timeout=timeout)
            with self._lock:
                self._thread = None
                self._thread_id = None
