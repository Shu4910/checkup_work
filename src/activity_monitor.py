"""
社員活動モニタリングアプリケーション メインモジュール
"""
import json
import threading
import time
from datetime import datetime
from pathlib import Path
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from window_tracker import get_active_window_info, get_machine_name, get_user_name
from browser_tracker import get_browser_info
from power_monitor import PowerMonitor
from csv_logger import CSVLogger
from tray_icon import TrayIcon


class ActivityMonitor:
    """アクティビティモニタリングメインクラス"""

    def __init__(self, config_path="config.json"):
        """
        Args:
            config_path (str): 設定ファイルのパス
        """
        self.config = self._load_config(config_path)
        self.polling_interval = self.config.get('polling_interval', 60)
        self.is_running = False
        self.is_paused = False

        self.logger = CSVLogger(
            self.config.get('log_directory', 'C:\\ProgramData\\ActivityMonitor\\logs'),
            self.config.get('log_retention_days', 90)
        )

        self.power_monitor = PowerMonitor()

        self.power_monitor.add_callback('on_suspend', self._on_suspend)
        self.power_monitor.add_callback('on_resume', self._on_resume)

        self.tray_icon = TrayIcon(
            app_name="Activity Monitor"
        )
        self.tray_icon.set_status_callback(self._get_status_text)
        self.tray_icon.set_quit_callback(self._quit)

        self.last_activity_time = None
        self.polling_thread = None

    def _load_config(self, config_path):
        """設定ファイルを読み込む"""
        try:
            config_file = Path(config_path)
            if not config_file.exists():
                config_file = Path(__file__).parent.parent / config_path

            with open(config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading config: {e}")
            return {}

    def _on_suspend(self):
        """スリープ時の処理"""
        print("System suspended - pausing monitoring")
        self.is_paused = True

    def _on_resume(self):
        """スリープ復帰時の処理"""
        print("System resumed - resuming monitoring")
        self.is_paused = False

    def _get_status_text(self):
        """ステータステキストを取得"""
        status_lines = [
            f"アプリケーション: Activity Monitor",
            f"状態: {'実行中' if self.is_running and not self.is_paused else '一時停止中' if self.is_paused else '停止'}",
            f"ポーリング間隔: {self.polling_interval}秒",
            f"ログ保存先: {self.config.get('log_directory')}",
        ]

        if self.last_activity_time:
            status_lines.append(f"最終記録時刻: {self.last_activity_time}")

        return "\n".join(status_lines)

    def _quit(self):
        """アプリケーションを終了"""
        print("Quitting application...")
        self.is_running = False

    def _collect_activity(self):
        """現在のアクティビティを収集"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        machine_name = get_machine_name()
        user_name = get_user_name()

        if self.power_monitor.is_locked or self.power_monitor.check_lock_status():
            activity_data = {
                'timestamp': timestamp,
                'machine_name': machine_name,
                'user_name': user_name,
                'process_name': '',
                'window_title': '',
                'url': '',
                'page_title': '',
                'status': 'locked'
            }
            return activity_data

        window_info = get_active_window_info()

        if window_info is None:
            return None

        process_name = window_info.get('process_name', '')
        window_title = window_info.get('window_title', '')
        hwnd = window_info.get('hwnd', 0)

        browser_info = get_browser_info(process_name, window_title, hwnd)

        activity_data = {
            'timestamp': timestamp,
            'machine_name': machine_name,
            'user_name': user_name,
            'process_name': process_name,
            'window_title': window_title,
            'url': browser_info.get('url', ''),
            'page_title': browser_info.get('page_title', ''),
            'status': 'active'
        }

        return activity_data

    def _polling_loop(self):
        """ポーリングループ"""
        while self.is_running:
            try:
                if not self.is_paused and not self.power_monitor.is_sleeping:
                    activity_data = self._collect_activity()

                    if activity_data:
                        self.logger.log_activity(activity_data)
                        self.last_activity_time = activity_data['timestamp']
                        print(f"Logged activity at {activity_data['timestamp']}")

                    if datetime.now().hour == 0 and datetime.now().minute == 0:
                        self.logger.clean_old_logs()

            except Exception as e:
                print(f"Error in polling loop: {e}")

            time.sleep(self.polling_interval)

    def start(self):
        """モニタリングを開始"""
        print("Starting Activity Monitor...")
        self.is_running = True

        self.power_monitor.start_monitoring()

        self.polling_thread = threading.Thread(target=self._polling_loop, daemon=True)
        self.polling_thread.start()

        self.tray_icon.run()


def main():
    """メイン関数"""
    config_path = Path(__file__).parent.parent / "config.json"

    monitor = ActivityMonitor(config_path)
    monitor.start()


if __name__ == "__main__":
    main()
