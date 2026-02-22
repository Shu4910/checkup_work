"""
アクティブウィンドウ情報取得モジュール
"""
import win32gui
import win32process
import psutil
import socket


def get_machine_name():
    """コンピュータ名を取得"""
    return socket.gethostname()


def get_user_name():
    """Windowsログインユーザー名を取得"""
    import os
    return os.getenv('USERNAME', 'Unknown')


def get_active_window_info():
    """
    アクティブウィンドウの情報を取得

    Returns:
        dict: {
            'process_name': str,
            'window_title': str,
            'pid': int
        }
    """
    try:
        hwnd = win32gui.GetForegroundWindow()
        if hwnd == 0:
            return None

        window_title = win32gui.GetWindowText(hwnd)

        _, pid = win32process.GetWindowThreadProcessId(hwnd)

        try:
            process = psutil.Process(pid)
            process_name = process.name()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            process_name = "Unknown"

        return {
            'process_name': process_name,
            'window_title': window_title,
            'pid': pid,
            'hwnd': hwnd
        }
    except Exception as e:
        print(f"Error getting active window info: {e}")
        return None


def is_system_locked():
    """
    システムがロック状態かどうかを確認

    Returns:
        bool: ロック状態ならTrue
    """
    try:
        import ctypes
        import win32ts

        session_id = win32ts.WTSGetActiveConsoleSessionId()

        if session_id == 0xFFFFFFFF:
            return True

        return False
    except Exception as e:
        print(f"Error checking lock status: {e}")
        return False
