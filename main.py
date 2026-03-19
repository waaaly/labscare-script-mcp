"""
LabsCare LIMS 报表脚本引擎 MCP Server  v2
知识库与框架分离：knowledge/*.json 存储所有规则，tools/ 模块处理逻辑。
新增 update_labscare_knowledge Tool，支持运行时热更新知识库。

传输层：同时支持三种模式，启动时用 --model 切换：
  - stdio           模式：python main.py
                    供 Claude Desktop / Cursor 等本地客户端
  - http            模式：python main.py --model http [--host 0.0.0.0] [--port 8000]
                    供 n8n MCP Client Tool 节点（Streamable HTTP，MCP 1.1+）
                    n8n 填写的 URL → http://<host>:<port>/mcp
  - sse             模式：python main.py --model sse [--host 0.0.0.0] [--port 8000]
                    供浏览器访问，通过 FastAPI 提供 SSE 传输层
                    浏览器访问 → http://<host>:<port>/sse
"""

import argparse
from starlette.routing import Mount
from typing import Dict, Any
from mcp.server import Server
from mcp.server.sse import SseServerTransport
from mcp import types
from mcp.types import Tool, TextContent
from fastapi import FastAPI, Request,Response 
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
from tools.knowledge import FIELD_DB, PATTERNS, DIAGNOSES, get_pattern_keys
from tools.handlers import (
    handle_generate_script,
    handle_explain_field,
    handle_debug_script,
    handle_get_pattern,
    handle_update_knowledge,
    handle_list_knowledge,
    handle_get_sampledata,
    handle_get_projectdata,
)
from tools.simulate import simulate_labscare_script
from tools.docx_parser import handle_parse_docx
from resources.script_spec import (
    script_spec
)

from prompts.labscare_component_spec import labscare_component_spec

mcp_server = Server("labscare-script")

@mcp_server.list_tools()
async def handle_list_tools():
    return [
        Tool(
            name="get_labscare_sampledata",
            description="获取 LabsCare LIMS 中某个项目的样本数据。",
            inputSchema={
                "type": "object",
                "properties": {
                    "lab": {"type": "string"},
                    "project_id": {"type": "string"}
                },
                "required": ["lab", "project_id"]
            }
        ),
        Tool(
            name="get_labscare_projectdata",
            description="获取 LabsCare LIMS 中某个项目的项目数据。",
            inputSchema={
                "type": "object",
                "properties": {
                    "lab": {"type": "string"},
                    "project_id": {"type": "string"}
                },
                "required": ["lab", "project_id"]
            }
        ),
        Tool(
            name="generate_labscare_script",
            description="根据用户描述的报表需求，生成完整的 LabsCare LIMS 报表 JS 脚本。",
            inputSchema={
                "type": "object",
                "properties": {
                    "report_type": {"type": "string"},
                    "data_sources": {"type": "array", "items": {"type": "string"}},
                    "placeholders": {"type": "array", "items": {"type": "string"}},
                    "special_needs": {"type": "object"},
                    "layer": {"type": "string"}
                },
                "required": ["report_type"]
            }
        ),
        Tool(
            name="explain_labscare_field",
            description="解释 LabsCare LIMS 中某个字段或占位符的含义、数据类型、取值方式。",
            inputSchema={
                "type": "object",
                "properties": {
                    "field_name": {"type": "string"},
                    "context": {"type": "string"}
                },
                "required": ["field_name"]
            }
        ),
        Tool(
            name="debug_labscare_script",
            description="诊断 LabsCare 报表脚本的常见问题，返回原因分析和修复方案。",
            inputSchema={
                "type": "object",
                "properties": {
                    "symptom": {"type": "string"},
                    "script_snippet": {"type": "string"},
                    "template_placeholders": {"type": "string"}
                },
                "required": ["symptom"]
            }
        ),
        Tool(
            name="get_labscare_pattern",
            description="获取 LabsCare 报表脚本的常用代码模式片段，可直接复制使用。",
            inputSchema={
                "type": "object",
                "properties": {
                    "pattern_name": {"type": "string"}
                },
                "required": ["pattern_name"]
            }
        ),
        Tool(
            name="update_labscare_knowledge",
            description="向 LabsCare 知识库添加、更新或删除条目，写入后立即热重载，无需重启服务器。",
            inputSchema={
                "type": "object",
                "properties": {
                    "operation": {"type": "string"},
                    "knowledge_type": {"type": "string"},
                    "key": {"type": "string"},
                    "content": {"type": "object"}
                },
                "required": ["operation", "knowledge_type", "key"]
            }
        ),
        Tool(
            name="list_labscare_knowledge",
            description="列出当前 LabsCare 知识库中的所有条目。",
            inputSchema={
                "type": "object",
                "properties": {
                    "knowledge_type": {"type": "string"}
                }
            }
        ),
        Tool(
            name="parse_labscare_docx",
            description="解析 LabsCare 报表脚本的注释，返回批注映射表。",
            inputSchema={
                "type": "object",
                "properties": {
                    "docx_path": {"type": "string"}
                },
                "required": ["docx_path"]
            }
        ),
        Tool(
            name="simulate_labscare_script",
            description="模拟 LabsCare 报表脚本的执行结果，返回模拟数据。",
            inputSchema={
                "type": "object",
                "properties": {
                    "js_code": {"type": "string"},
                    "lab": {"type": "string"},
                    "project_id": {"type": "string"}
                },
                "required": ["js_code", "lab", "project_id"]
            }
        )
    ]

@mcp_server.call_tool()
async def handle_call_tool(name: str, arguments: dict):
    if name == "get_labscare_sampledata":
        return await handle_get_sampledata(arguments)
    elif name == "get_labscare_projectdata":
        return await handle_get_projectdata(arguments)
    elif name == "generate_labscare_script":
        return handle_generate_script(arguments)
    elif name == "explain_labscare_field":
        return handle_explain_field(arguments)
    elif name == "debug_labscare_script":
        return handle_debug_script(arguments)
    elif name == "get_labscare_pattern":
        return handle_get_pattern(arguments)
    elif name == "update_labscare_knowledge":
        return handle_update_knowledge(arguments)
    elif name == "list_labscare_knowledge":
        return handle_list_knowledge(arguments)
    elif name == "parse_labscare_docx":
        return handle_parse_docx(arguments.get("docx_path", ""))
    elif name == "simulate_labscare_script":
        return await simulate_labscare_script(
            arguments.get("js_code", ""),
            arguments.get("lab", ""),
            arguments.get("project_id", "")
        )
    else:
        raise ValueError(f"Unknown tool: {name}")

@mcp_server.list_prompts()
async def handle_list_prompts():
    return [
        {
            "name": "labscare_component_spec",
            "description": "LabsCare 组件规格说明"
        }
    ]

@mcp_server.get_prompt()
async def handle_get_prompt(name: str, arguments: dict = None):
    if name == "labscare_component_spec":
        return {
            "messages": [
                {
                    "role": "user",
                    "content": {
                        "type": "text",
                        "text": labscare_component_spec()
                    }
                }
            ]
        }
    else:
        raise ValueError(f"Unknown prompt: {name}")

@mcp_server.list_resources()
async def handle_list_resources():
    return [
        {
            "uri": "labscare://script-spec",
            "name": "LabsCare Script Specification",
            "description": "LabsCare 报表脚本规范文档",
            "mimeType": "text/plain"
        }
    ]

@mcp_server.read_resource()
async def handle_read_resource(uri: str):
    if uri == "labscare://script-spec":
        return [TextContent(type="text", text=script_spec())]
    else:
        raise ValueError(f"Unknown resource: {uri}")

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

sse = SseServerTransport("/messages")

# 挂载 POST 消息接收端点
app.router.routes.append(
    Mount("/messages", app=sse.handle_post_message)
)

@app.get("/sse")
async def sse_endpoint(request: Request):
    async with sse.connect_sse(
        request.scope,
        request.receive,
        request._send
    ) as (read_stream, write_stream):
        await mcp_server.run(
            read_stream,
            write_stream,
            mcp_server.create_initialization_options()
        )

@app.get("/")
async def root():
    return {
        "message": "LabsCare MCP Server",
        "endpoints": {
            "sse": "/sse",
            "messages": "/messages",
            "health": "/health"
        }
    }

@app.get("/health")
async def health():
    from datetime import datetime
    return {"status": "healthy at :"+datetime.now().strftime("%Y-%m-%d %H:%M:%S")}



if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="LabsCare MCP Server")
    parser.add_argument("--model", choices=["stdio", "http", "sse"], default="stdio", help="传输模式：stdio（默认）、http（Streamable HTTP）或 sse（SSE 传输层）")
    parser.add_argument("--host", default="0.0.0.0", help="监听地址（默认 0.0.0.0）")
    parser.add_argument("--port", type=int, default=8000, help="监听端口（默认 8000）")
    args = parser.parse_args()
    
    if args.model == "sse":
        print(f"[LabsCare MCP] SSE 模式启动：http://{args.host}:{args.port}/sse")
        print(f"[LabsCare MCP] 浏览器访问：http://{args.host}:{args.port}/")
        uvicorn.run(app, host=args.host, port=args.port)
    elif args.model == "http":
        print(f"[LabsCare MCP] HTTP 模式启动：http://{args.host}:{args.port}/mcp")
        print(f"[LabsCare MCP] 供 n8n MCP Client Tool 节点使用")
        from mcp.server.fastmcp import FastMCP
        http_app = FastMCP("labscare-script")
        http_app.run(transport="streamable-http", host=args.host, port=args.port)
    else:
        print("[LabsCare MCP] stdio 模式启动")
        print("[LabsCare MCP] 供 Claude Desktop / Cursor 等本地客户端使用")
        from mcp.server.stdio import stdio_server
        async def main():
            async with stdio_server() as (read_stream, write_stream):
                try:
                    await asyncio.wait_for(
                        mcp_server.run(read_stream, write_stream, mcp_server.create_initialization_options()),
                        timeout=None  # 可选加超时
                    )
                except asyncio.CancelledError:
                    print("Client disconnected (cancelled)")
        import asyncio
        asyncio.run(main())
        