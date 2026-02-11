# プロジェクト管理MPCサーバー

Claude DesktopからExcelファイルのプロジェクト管理を行うMPCサーバーです。

## 📋 機能

- ✅ プロジェクトの追加・更新・削除
- ✅ プロジェクト一覧の取得
- ✅ 条件検索（進行状況・担当者）
- ✅ 自然言語でのExcel操作

## 🚀 セットアップ

### 1. uvのインストール

```bash
# Windows (PowerShell)
irm https://astral.sh/uv/install.ps1 | iex

# macOS/Linux
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### 2. 依存関係のインストール

```bash
cd /c/MCP_Learning/Excel_admin
uv sync
```

### 3. config.jsonの設定

Excelファイルのパスを設定してください：

```json
{
  "excel_file_path": "C:/Users/YourName/Documents/プロジェクト管理.xlsx"
}
```

### 4. Claude Desktop Studioへの登録

Claude Desktopの設定ファイルに以下を追加：

```json
{
  "mcpServers": {
    "project-manager": {
      "command": "uv",
      "args": ["run", "project-mcp-server"],
      "cwd": "C:/MCP_Learning/Excel_admin"
    }
  }
}
```

### 5. Claude Desktopを再起動

サーバーが自動的に起動し、ツールが利用可能になります。

## 💬 使い方

Claude Desktopで以下のように話しかけてください：

- 「プロジェクト一覧を表示して」
- 「新しいプロジェクト『Webサイトリニューアル』を追加して、担当者は田中さん」
- 「『Webサイトリニューアル』の進行状況を『進行中』に更新して」
- 「完了したプロジェクトを教えて」

## 📂 プロジェクト構造

```
project-mcp-server/
├── pyproject.toml
├── config.json
├── README.md
├── .gitignore
├── src/
│   └── project_mcp_server/
│       ├── __init__.py
│       ├── __main__.py
│       ├── server.py
│       ├── excel_handler.py
│       └── config.py
├── logs/
└── docs/
```

## 🔧 開発

### ローカルでサーバーを起動

```bash
uv run project-mcp-server
```

### ログの確認

```bash
tail -f logs/mpc_server.log
```

## 📝 提供ツール

| ツール名 | 説明 |
|---------|------|
| `get_projects` | 全プロジェクト取得 |
| `get_project` | 特定プロジェクト取得 |
| `add_project` | プロジェクト追加 |
| `update_project` | プロジェクト更新 |
| `delete_project` | プロジェクト削除 |
| `search_projects` | 条件検索 |

## 📄 ライセンス

MIT License

## 👤 作成者

Your Name
