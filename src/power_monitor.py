"""
スリープ・ロック検知モジュール
Win32 APIを使用してPCのスリープとロック状態を監視
"""
import win32api
import win32con
import win32gui
import threading
import time


class PowerMonitor:
    """電源状態とロック状態を監視するクラス"""

    def __init__(self):
        self.is_sleeping = False
        self.is_locked = False
        self.callbacks = {
            'on_suspend': [],
            'on_resume': [],
            'on_lock': [],
            'on_unlock': []
        }

    def add_callback(self, event_type, callback):
        """
        イベントコールバックを追加

        Args:
            event_type (str): 'on_suspend', 'on_resume', 'on_lock', 'on_unlock'
            callback (callable): コールバック関数
        """
        if event_type in self.callbacks:
            self.callbacks[event_type].append(callback)

    def _trigger_callbacks(self, event_type):
        """コールバックを実行"""
        for callback in self.callbacks.get(event_type, []):
            try:
                callback()
            except Exception as e:
                print(f"Error in callback for {event_type}: {e}")

    def check_lock_status(self):
        """
        ロック状態をチェック

        Returns:
            bool: ロック中ならTrue
        """
        try:
            import ctypes
            from ctypes import wintypes

            user32 = ctypes.windll.user32

            hDesk = user32.OpenInputDesktop(0, False, win32con.MAXIMUM_ALLOWED)

            if hDesk == 0:
                self.is_locked = True
                return True

            user32.CloseDesktop(hDesk)
            self.is_locked = False
            return False

        except Exception as e:
            print(f"Error checking lock status: {e}")
            return False

    def start_monitoring(self):
        """監視を開始"""
        def monitor_thread():
            try:
                wc = win32gui.WNDCLASS()
                hinst = wc.hInstance = win32api.GetModuleHandle(None)
                wc.lpszClassName = "PowerMonitorWindowClass"
                wc.lpfnWndProc = self._wnd_proc

                classAtom = win32gui.RegisterClass(wc)

                hwnd = win32gui.CreateWindow(
                    classAtom,
                    "PowerMonitorWindow",
                    0,
                    0, 0,
                    win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT,
                    0, 0, hinst, None
                )

                win32gui.PumpMessages()

            except Exception as e:
                print(f"Error in monitor thread: {e}")

        thread = threading.Thread(target=monitor_thread, daemon=True)
        thread.start()

        lock_check_thread = threading.Thread(target=self._lock_check_loop, daemon=True)
        lock_check_thread.start()

    def _lock_check_loop(self):
        """ロック状態を定期的にチェック"""
        previous_lock_state = False

        while True:
            try:
                current_lock_state = self.check_lock_status()

                if current_lock_state and not previous_lock_state:
                    self._trigger_callbacks('on_lock')
                elif not current_lock_state and previous_lock_state:
                    self._trigger_callbacks('on_unlock')

                previous_lock_state = current_lock_state

            except Exception as e:
                print(f"Error in lock check loop: {e}")

            time.sleep(5)

    def _wnd_proc(self, hwnd, msg, wparam, lparam):
        """ウィンドウプロシージャ"""
        if msg == win32con.WM_POWERBROADCAST:
            if wparam == win32con.PBT_APMSUSPEND:
                self.is_sleeping = True
                self._trigger_callbacks('on_suspend')
                print("System is suspending")
            elif wparam == win32con.PBT_APMRESUMESUSPEND:
                self.is_sleeping = False
                self._trigger_callbacks('on_resume')
                print("System is resuming")

        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
