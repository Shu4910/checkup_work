"""
CSVログ保存モジュール
アクティビティログをCSV形式で保存
"""
import csv
import os
from datetime import datetime, timedelta
from pathlib import Path


class CSVLogger:
    """CSVログ管理クラス"""

    def __init__(self, log_directory, retention_days=90):
        """
        Args:
            log_directory (str): ログ保存先ディレクトリ
            retention_days (int): ログ保存期間（日数）
        """
        self.log_directory = Path(log_directory)
        self.retention_days = retention_days
        self._ensure_log_directory()

    def _ensure_log_directory(self):
        """ログディレクトリが存在することを確認"""
        try:
            self.log_directory.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"Error creating log directory: {e}")

    def _get_log_file_path(self, date=None):
        """
        ログファイルパスを取得

        Args:
            date (datetime): 日付（デフォルトは今日）

        Returns:
            Path: ログファイルパス
        """
        if date is None:
            date = datetime.now()

        filename = f"activity_log_{date.strftime('%Y-%m-%d')}.csv"
        return self.log_directory / filename

    def log_activity(self, activity_data):
        """
        アクティビティをログに記録

        Args:
            activity_data (dict): {
                'timestamp': str,
                'machine_name': str,
                'user_name': str,
                'process_name': str,
                'window_title': str,
                'url': str,
                'page_title': str,
                'status': str
            }
        """
        log_file = self._get_log_file_path()

        file_exists = log_file.exists()

        try:
            with open(log_file, 'a', newline='', encoding='utf-8-sig') as f:
                fieldnames = [
                    'timestamp', 'machine_name', 'user_name', 'process_name',
                    'window_title', 'url', 'page_title', 'status'
                ]
                writer = csv.DictWriter(f, fieldnames=fieldnames)

                if not file_exists:
                    writer.writeheader()

                writer.writerow(activity_data)

        except Exception as e:
            print(f"Error writing to log file: {e}")

    def clean_old_logs(self):
        """保存期間を過ぎた古いログファイルを削除"""
        try:
            cutoff_date = datetime.now() - timedelta(days=self.retention_days)

            for log_file in self.log_directory.glob("activity_log_*.csv"):
                try:
                    date_str = log_file.stem.replace("activity_log_", "")
                    file_date = datetime.strptime(date_str, "%Y-%m-%d")

                    if file_date < cutoff_date:
                        log_file.unlink()
                        print(f"Deleted old log file: {log_file}")

                except (ValueError, OSError) as e:
                    print(f"Error processing log file {log_file}: {e}")

        except Exception as e:
            print(f"Error cleaning old logs: {e}")

    def get_log_files(self):
        """
        すべてのログファイルを取得

        Returns:
            list: ログファイルパスのリスト
        """
        try:
            return sorted(self.log_directory.glob("activity_log_*.csv"))
        except Exception as e:
            print(f"Error getting log files: {e}")
            return []
