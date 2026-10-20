"""CLI facade over the exact same handlers used by the MCP server."""
import argparse
import json
import sys
from pathlib import Path


def main():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    import project_mcp_server as server
    parser = argparse.ArgumentParser(description="Microsoft Project MCP 的命令行入口")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("tools", help="列出全部工具")
    call = sub.add_parser("call", help="以 UTF-8 JSON 参数文件调用 MCP 工具")
    call.add_argument("tool", choices=sorted(server.TOOL_MAP))
    call.add_argument("--args-file", required=True, type=Path)
    args = parser.parse_args()
    try:
        if args.command == "tools":
            result = [{"name": t["name"], "description": t["description"], "inputSchema": t["inputSchema"]}
                      for t in server.TOOLS]
        else:
            params = json.loads(args.args_file.read_text(encoding="utf-8-sig"))
            result = server.TOOL_MAP[args.tool]["handler"](params)
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
        return 0
    except Exception as exc:
        print(json.dumps({"error": str(exc), "type": type(exc).__name__}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
