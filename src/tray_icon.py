"""
システムトレイアイコンモジュール
"""
import os
import threading
import pystray
from PIL import Image, ImageDraw
import win32api
import win32con


class TrayIcon:
    """システムトレイアイコン管理クラス"""

    def __init__(self, app_name="Activity Monitor"):
        """
        Args:
            app_name (str): アプリケーション名
        """
        self.app_name = app_name
        self.icon = None
        self.status_callback = None
        self.quit_callback = None

    def create_image(self):
        """トレイアイコン用の画像を生成"""
        width = 64
        height = 64
        image = Image.new('RGB', (width, height), color='blue')
        dc = ImageDraw.Draw(image)

        dc.ellipse([10, 10, 54, 54], fill='white', outline='blue')

        dc.rectangle([28, 20, 36, 35], fill='blue')
        dc.ellipse([24, 37, 40, 50], fill='blue')

        return image

    def show_status(self, icon, item):
        """現在の状態を表示"""
        def _show():
            if self.status_callback:
                status_text = self.status_callback()
            else:
                status_text = "モニタリング実行中"
            win32api.MessageBox(
                0,
                status_text,
                self.app_name,
                win32con.MB_OK | win32con.MB_ICONINFORMATION | win32con.MB_TOPMOST
            )
        threading.Thread(target=_show, daemon=True).start()

    def quit_app(self, icon, item):
        """アプリケーションを終了"""
        def _quit():
            result = win32api.MessageBox(
                0,
                "アプリケーションを終了しますか？",
                "終了確認",
                win32con.MB_YESNO | win32con.MB_ICONQUESTION | win32con.MB_TOPMOST
            )
            if result == win32con.IDYES:
                if self.quit_callback:
                    self.quit_callback()
                icon.stop()
                os._exit(0)
        threading.Thread(target=_quit, daemon=True).start()

    def set_status_callback(self, callback):
        """
        ステータス表示用のコールバックを設定

        Args:
            callback (callable): 状態テキストを返す関数
        """
        self.status_callback = callback

    def set_quit_callback(self, callback):
        """
        終了時のコールバックを設定

        Args:
            callback (callable): 終了処理を行う関数
        """
        self.quit_callback = callback

    def run(self):
        """トレイアイコンを表示して実行"""
        menu = pystray.Menu(
            pystray.MenuItem("現在の状態を表示", self.show_status),
            pystray.MenuItem("終了", self.quit_app)
        )

        image = self.create_image()

        self.icon = pystray.Icon(
            self.app_name,
            image,
            self.app_name,
            menu
        )

        self.icon.run()
