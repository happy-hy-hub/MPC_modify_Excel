"""
MPCサーバーメインモジュール

Claude Desktopとの通信を処理し、ツールを提供します。
"""
import asyncio
import json
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from .config import ConfigManager
from .excel_handler import (
    ExcelHandler,
    ExcelHandlerError,
    FileNotFoundError,
    FileAccessError,
    InvalidDataError,
    ProjectNotFoundError,
    DuplicateProjectError,
)


# ロギング設定
def setup_logging(config_manager: ConfigManager):
    """ロギングを設定します"""
    log_config = config_manager.get_logging_config()
    log_file = Path(log_config.get("file", "logs/mpc_server.log"))
    log_level = log_config.get("level", "INFO")
    max_bytes = log_config.get("max_bytes", 10485760)
    backup_count = log_config.get("backup_count", 5)
    
    # logsディレクトリを作成
    log_file.parent.mkdir(parents=True, exist_ok=True)
    
    # ロガー設定
    logger = logging.getLogger()
    logger.setLevel(getattr(logging, log_level))
    
    # ファイルハンドラー（ローテーション）
    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding='utf-8'
    )
    file_handler.setLevel(getattr(logging, log_level))
    
    # フォーマッター
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    file_handler.setFormatter(formatter)
    
    # ハンドラーを追加
    logger.addHandler(file_handler)
    
    logging.info("ロギングを設定しました")


# サーバーインスタンス
server = Server("project-mcp-server")

# グローバル変数（main()で初期化）
excel_handler: ExcelHandler = None


@server.list_tools()
async def handle_list_tools() -> list[Tool]:
    """
    利用可能なツール一覧を返します。
    
    Returns:
        ツール定義の配列
    """
    return [
        Tool(
            name="get_projects",
            description="全プロジェクトの一覧を取得します",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        Tool(
            name="get_project",
            description="特定のプロジェクトの詳細情報を取得します",
            inputSchema={
                "type": "object",
                "properties": {
                    "project_name": {
                        "type": "string",
                        "description": "取得するプロジェクト名"
                    }
                },
                "required": ["project_name"]
            }
        ),
        Tool(
            name="add_project",
            description="新しいプロジェクトを追加します",
            inputSchema={
                "type": "object",
                "properties": {
                    "project_name": {
                        "type": "string",
                        "description": "プロジェクト名"
                    },
                    "status": {
                        "type": "string",
                        "description": "進行状況（未着手、進行中、完了、中止）",
                        "enum": ["未着手", "進行中", "完了", "中止"],
                        "default": "未着手"
                    },
                    "deadline": {
                        "type": "string",
                        "description": "期限（任意）"
                    },
                    "assignee": {
                        "type": "string",
                        "description": "担当者（任意）"
                    },
                    "notes": {
                        "type": "string",
                        "description": "備考（任意）"
                    }
                },
                "required": ["project_name"]
            }
        ),
        Tool(
            name="update_project",
            description="既存のプロジェクト情報を更新します",
            inputSchema={
                "type": "object",
                "properties": {
                    "project_name": {
                        "type": "string",
                        "description": "更新するプロジェクト名"
                    },
                    "status": {
                        "type": "string",
                        "description": "進行状況（任意）",
                        "enum": ["未着手", "進行中", "完了", "中止"]
                    },
                    "deadline": {
                        "type": "string",
                        "description": "期限（任意）"
                    },
                    "assignee": {
                        "type": "string",
                        "description": "担当者（任意）"
                    },
                    "notes": {
                        "type": "string",
                        "description": "備考（任意）"
                    }
                },
                "required": ["project_name"]
            }
        ),
        Tool(
            name="delete_project",
            description="プロジェクトを削除します",
            inputSchema={
                "type": "object",
                "properties": {
                    "project_name": {
                        "type": "string",
                        "description": "削除するプロジェクト名"
                    }
                },
                "required": ["project_name"]
            }
        ),
        Tool(
            name="search_projects",
            description="条件に合致するプロジェクトを検索します",
            inputSchema={
                "type": "object",
                "properties": {
                    "status": {
                        "type": "string",
                        "description": "進行状況で絞り込み（任意）",
                        "enum": ["未着手", "進行中", "完了", "中止"]
                    },
                    "assignee": {
                        "type": "string",
                        "description": "担当者で絞り込み（任意）"
                    }
                },
                "required": []
            }
        ),
    ]


@server.call_tool()
async def handle_call_tool(name: str, arguments: dict) -> list[TextContent]:
    """
    ツール呼び出しを処理します。
    
    Args:
        name: ツール名
        arguments: ツールに渡す引数
    
    Returns:
        ツールの実行結果
    """
    logging.info(f"ツール呼び出し: {name}, パラメータ: {arguments}")
    
    try:
        if name == "get_projects":
            # 全プロジェクト取得
            projects = excel_handler.get_all_projects()
            return [TextContent(
                type="text",
                text=json.dumps(projects, ensure_ascii=False, indent=2)
            )]
        
        elif name == "get_project":
            # 特定プロジェクト取得
            project_name = arguments.get("project_name")
            project = excel_handler.get_project(project_name)
            
            if project:
                return [TextContent(
                    type="text",
                    text=json.dumps(project, ensure_ascii=False, indent=2)
                )]
            else:
                return [TextContent(
                    type="text",
                    text=f"プロジェクト '{project_name}' が見つかりません"
                )]
        
        elif name == "add_project":
            # プロジェクト追加
            project_data = {
                "プロジェクト": arguments.get("project_name"),
                "進行状況": arguments.get("status", "未着手"),
                "期限": arguments.get("deadline", ""),
                "担当者": arguments.get("assignee", ""),
                "備考": arguments.get("notes", "")
            }
            
            excel_handler.add_project(project_data)
            return [TextContent(
                type="text",
                text=f"プロジェクト '{project_data['プロジェクト']}' を追加しました"
            )]
        
        elif name == "update_project":
            # プロジェクト更新
            project_name = arguments.get("project_name")
            updates = {}
            
            if "status" in arguments:
                updates["進行状況"] = arguments["status"]
            if "deadline" in arguments:
                updates["期限"] = arguments["deadline"]
            if "assignee" in arguments:
                updates["担当者"] = arguments["assignee"]
            if "notes" in arguments:
                updates["備考"] = arguments["notes"]
            
            excel_handler.update_project(project_name, updates)
            return [TextContent(
                type="text",
                text=f"プロジェクト '{project_name}' を更新しました"
            )]
        
        elif name == "delete_project":
            # プロジェクト削除
            project_name = arguments.get("project_name")
            excel_handler.delete_project(project_name)
            return [TextContent(
                type="text",
                text=f"プロジェクト '{project_name}' を削除しました"
            )]
        
        elif name == "search_projects":
            # プロジェクト検索
            filters = {}
            
            if "status" in arguments:
                filters["進行状況"] = arguments["status"]
            if "assignee" in arguments:
                filters["担当者"] = arguments["assignee"]
            
            projects = excel_handler.search_projects(filters)
            return [TextContent(
                type="text",
                text=json.dumps(projects, ensure_ascii=False, indent=2)
            )]
        
        else:
            return [TextContent(
                type="text",
                text=f"不明なツールです: {name}"
            )]
    
    except FileNotFoundError as e:
        logging.error(f"ファイルエラー: {e}")
        return [TextContent(
            type="text",
            text=f"エラー: {str(e)}"
        )]
    
    except FileAccessError as e:
        logging.error(f"ファイルアクセスエラー: {e}")
        return [TextContent(
            type="text",
            text=f"エラー: {str(e)}"
        )]
    
    except InvalidDataError as e:
        logging.error(f"データエラー: {e}")
        return [TextContent(
            type="text",
            text=f"エラー: {str(e)}"
        )]
    
    except ProjectNotFoundError as e:
        logging.error(f"プロジェクト未検出: {e}")
        return [TextContent(
            type="text",
            text=f"エラー: {str(e)}"
        )]
    
    except DuplicateProjectError as e:
        logging.error(f"重複エラー: {e}")
        return [TextContent(
            type="text",
            text=f"エラー: {str(e)}"
        )]
    
    except Exception as e:
        logging.error(f"予期しないエラー: {e}", exc_info=True)
        return [TextContent(
            type="text",
            text=f"予期しないエラーが発生しました: {str(e)}"
        )]


async def main():
    """MPCサーバーのエントリーポイント"""
    global excel_handler
    
    # 設定読み込み
    config_manager = ConfigManager()
    
    # ロギング設定
    setup_logging(config_manager)
    
    logging.info("=" * 60)
    logging.info("プロジェクト管理MPCサーバーを起動しています...")
    logging.info("=" * 60)
    
    # Excelファイルパス取得
    excel_path = config_manager.get_excel_path()
    logging.info(f"Excelファイルパス: {excel_path}")
    
    # ExcelHandler初期化
    excel_handler = ExcelHandler(excel_path)
    
    # サーバー起動
    async with stdio_server() as (read_stream, write_stream):
        logging.info("MPCサーバーが起動しました")
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())
