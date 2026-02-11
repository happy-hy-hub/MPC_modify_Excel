"""
設定ファイル管理モジュール

config.jsonの読み込みと設定値の提供を行います。
"""
import json
import logging
from pathlib import Path
from typing import Dict, Optional


class ConfigManager:
    """設定ファイルの管理クラス"""
    
    DEFAULT_CONFIG_PATH = "config.json"
    DEFAULT_EXCEL_PATH = "プロジェクト管理.xlsx"
    
    def __init__(self, config_path: str = DEFAULT_CONFIG_PATH):
        """
        ConfigManagerを初期化します。
        
        Args:
            config_path: 設定ファイルのパス（デフォルト: config.json）
        """
        self.config_path = Path(config_path)
        self.config = self.load_config()
    
    def load_config(self) -> Dict:
        """
        設定ファイルを読み込みます。
        
        Returns:
            設定情報の辞書
        
        Raises:
            FileNotFoundError: 設定ファイルが見つからない場合
            json.JSONDecodeError: JSONの形式が不正な場合
        """
        if not self.config_path.exists():
            logging.warning(f"設定ファイル {self.config_path} が見つかりません。デフォルト設定を使用します。")
            self.create_default_config()
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            logging.info(f"設定ファイルを読み込みました: {self.config_path}")
            return config
        except json.JSONDecodeError as e:
            logging.error(f"設定ファイルの形式が不正です: {e}")
            raise
    
    def get_excel_path(self) -> str:
        """
        Excelファイルのパスを取得します。
        
        Returns:
            Excelファイルの絶対パス
        """
        excel_path = self.config.get("excel_file_path", self.DEFAULT_EXCEL_PATH)
        
        # 相対パスの場合は絶対パスに変換
        path = Path(excel_path)
        if not path.is_absolute():
            path = Path.cwd() / path
        
        return str(path)
    
    def get_logging_config(self) -> Dict:
        """
        ログ設定を取得します。
        
        Returns:
            ログ設定の辞書
        """
        return self.config.get("logging", {
            "level": "INFO",
            "file": "logs/mpc_server.log",
            "max_bytes": 10485760,  # 10MB
            "backup_count": 5
        })
    
    def create_default_config(self):
        """
        デフォルトの設定ファイルを作成します。
        """
        default_config = {
            "excel_file_path": self.DEFAULT_EXCEL_PATH,
            "logging": {
                "level": "INFO",
                "file": "logs/mpc_server.log",
                "max_bytes": 10485760,
                "backup_count": 5
            }
        }
        
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=2)
            logging.info(f"デフォルト設定ファイルを作成しました: {self.config_path}")
            self.config = default_config
        except Exception as e:
            logging.error(f"設定ファイルの作成に失敗しました: {e}")
            self.config = default_config
