"""
LabsCare LIMS 报表脚本引擎 MCP Server  v2
知识库与框架分离：knowledge/*.json 存储所有规则，tools/ 模块处理逻辑。
新增 update_labscare_knowledge Tool，支持运行时热更新知识库。

传输层：同时支持两种模式，启动时用 --transport 切换：
  - stdio           模式：python main.py
                    供 Claude Desktop / Cursor 等本地客户端
  - http  模式：python main.py --transport http [--host 0.0.0.0] [--port 8000]
                    供 n8n MCP Client Tool 节点（Streamable HTTP，MCP 1.1+）
                    n8n 填写的 URL → http://<host>:<port>/mcp
"""

import argparse
from mcp.server.fastmcp import FastMCP
from mcp import types

from tools.knowledge import FIELD_DB, PATTERNS, DIAGNOSES, get_pattern_keys
from tools.handlers import (
    handle_generate_script,
    handle_explain_field,
    handle_debug_script,
    handle_get_pattern,
    handle_update_knowledge,
    handle_list_knowledge,
    handle_get_sampledata,
    handle_get_projectdata
)
from resources.script_spec import (
    script_spec
)

from prompts.labscare_component_spec import labscare_component_spec

app = FastMCP("labscare-script", stateless_http=True)

@app.tool(name="get_labscare_sampledata", description="获取 LabsCare LIMS 中某个项目的样本数据。")
async def tool_get_sampledata(lab: str, project_id: str) -> str:
    return await handle_get_sampledata({"lab": lab, "project_id": project_id})

@app.tool(name="get_labscare_projectdata", description="获取 LabsCare LIMS 中某个项目的项目数据。")
async def tool_get_projectdata(lab: str, project_id: str) -> str:
    return await handle_get_projectdata({"lab": lab, "project_id": project_id})

@app.tool(name="generate_labscare_script", description="根据用户描述的报表需求，生成完整的 LabsCare LIMS 报表 JS 脚本。")
def tool_generate(report_type: str, data_sources: list = None, placeholders: list = None, special_needs: dict = None, layer: str = "single") -> str:
    return handle_generate_script({"report_type": report_type, "data_sources": data_sources or ["getProjectData", "getProjectSamples"], "placeholders": placeholders or [], "special_needs": special_needs or {}, "layer": layer})

@app.tool(name="explain_labscare_field", description="解释 LabsCare LIMS 中某个字段或占位符的含义、数据类型、取值方式。")
def tool_explain(field_name: str, context: str = "") -> str:
    return handle_explain_field({"field_name": field_name, "context": context})

@app.tool(name="debug_labscare_script", description="诊断 LabsCare 报表脚本的常见问题，返回原因分析和修复方案。")
def tool_debug(symptom: str, script_snippet: str = "", template_placeholders: str = "") -> str:
    return handle_debug_script({"symptom": symptom, "script_snippet": script_snippet, "template_placeholders": template_placeholders})

@app.tool(name="get_labscare_pattern", description="获取 LabsCare 报表脚本的常用代码模式片段，可直接复制使用。")
def tool_pattern(pattern_name: str) -> str:
    return handle_get_pattern({"pattern_name": pattern_name})

@app.tool(name="update_labscare_knowledge", description="向 LabsCare 知识库添加、更新或删除条目，写入后立即热重载，无需重启服务器。")
def tool_update_knowledge(operation: str, knowledge_type: str, key: str, content: dict = None) -> str:
    return handle_update_knowledge({"operation": operation, "knowledge_type": knowledge_type, "key": key, "content": content or {}})

@app.tool(name="list_labscare_knowledge", description="列出当前 LabsCare 知识库中的所有条目。")
def tool_list_knowledge(knowledge_type: str = "all") -> str:
    return handle_list_knowledge({"knowledge_type": knowledge_type})

@app.resource("labscare://script-spec")
def resource_script_spec() -> str:
    return script_spec()

@app.prompt("labscare_component_spec")
def prompt_labscare_component_spec() -> str:
    return labscare_component_spec()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="LabsCare MCP Server")
    parser.add_argument("--transport", choices=["stdio", "http"], default="stdio", help="传输模式：stdio（默认）或 http（Streamable HTTP）")
    parser.add_argument("--host", default="0.0.0.0", help="监听地址（默认 0.0.0.0）")
    parser.add_argument("--port", type=int, default=8000, help="监听端口（默认 8000）")
    args = parser.parse_args()

    if args.transport == "http":
        print(f"HTTP[LabsCare MCP] Streamable HTTP 启动：http://{args.host}:{args.port}/mcp")
        app.run(transport="streamable-http")
    else:
        print("stdio[LabsCare MCP] stdio 模式启动", flush=True)
        app.run(transport="stdio")
