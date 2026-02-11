"""
MPCサーバーのエントリーポイント

`uv run project-mcp-server` で実行されます。
"""
import asyncio
import sys

# Python 3.10以降でのイベントループポリシー設定（Windows対応）
if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

from .server import main

if __name__ == "__main__":
    asyncio.run(main())
