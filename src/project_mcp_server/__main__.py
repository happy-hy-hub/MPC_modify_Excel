"""
MPCサーバーのエントリーポイント

`uv run project-mcp-server` で実行されます。
"""
import asyncio
from .server import main

if __name__ == "__main__":
    asyncio.run(main())
