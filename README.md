

```markdown
# LabsCare Script MCP Server v2

一个专为 LabsCare/LIMS 报告引擎设计的 **MCP（Model Context Protocol）服务**，帮助 AI Agent（如 Grok、Claude、Cursor）自动/半自动生成 ES5 取数脚本。

核心特点：**知识库与框架完全分离**，所有字段解释、代码模式、问题诊断均存放在 `knowledge/*.json` 文件中，无需修改 Python 代码即可持续更新规则。

## 当前状态（2026年3月）
- 已实现：MCP 双传输（stdio + streamable-http）、6个核心工具 + 2个数据查询工具、知识库热更新、部分 Resources/Prompts 注册
- 正在完善：线下带批注 DOC 自动解析、JS 脚本沙箱模拟执行 & 覆盖率验证
- 目标：让 Agent 一键输入报告需求 + 带批注 DOC → 输出可直接预览通过的 JS 取数脚本

## 项目结构

```
labscare-script-mcp/
├── main.py                 # MCP 服务器主入口（FastMCP 实现）
├── tools/                  # 工具实现目录
│   └── handlers.py         # 各 tool 的具体处理函数（可继续拆分）
├── knowledge/              # 知识库（可直接编辑，热重载）
│   ├── fields.json         # 字段解释库
│   ├── patterns.json       # 代码模式片段库
│   └── diagnoses.json      # 常见问题诊断库
├── resources/              # 示例脚本、样本数据、带批注 DOC 等（部分已注册为 Resource）
├── prompts/                # MCP Prompt 模板（已注册 labscare_component_spec）
├── config.json             # 配置（可选）
├── get_docx_comment.py     # 提取 DOC 批注的辅助脚本（待包装成 Tool）
├── docx_to_xmreport.py     # DOC 转 XML 报告结构的辅助脚本
├── pyproject.toml          # uv 项目依赖管理
└── README.md
```

## 已实现的核心 Tools

| Tool 名                        | 功能描述                                      | 参数示例 |
|-------------------------------|-----------------------------------------------|----------|
| generate_labscare_script      | 根据需求生成完整 ES5 报表 JS 脚本             | report_type, data_sources, placeholders, special_needs |
| explain_labscare_field        | 解释某个字段的含义、取值方式                  | field_name |
| debug_labscare_script         | 静态诊断脚本常见问题 + 建议修复               | js_code |
| get_labscare_pattern          | 获取预定义的代码模式片段                      | pattern_key |
| update_labscare_knowledge     | 运行时添加/更新/删除知识库条目（热重载）      | operation, knowledge_type, key, content |
| list_labscare_knowledge       | 查看当前所有知识库条目                        | - |
| get_labscare_sampledata       | 通过接口获取样本数据（用于生成/验证参考）     | sample_id 等 |
| get_labscare_projectdata      | 通过接口获取项目/报告相关数据                 | project_id 等 |

## 安装与启动

1. 克隆仓库
   ```bash
   git clone https://github.com/waaaly/labscare-script-mcp.git
   cd labscare-script-mcp
   ```

2. 安装依赖（推荐使用 uv）
   ```bash
   uv sync          # 或 pip install -r requirements.txt（如果后续导出）
   ```

3. 启动（两种模式）

   **本地 stdio 模式**（推荐 Claude Desktop / Cursor 测试）
   ```bash
   python main.py --transport stdio
   ```

   **远程 HTTP 模式**（支持 Grok、n8n、自定义 Agent）
   ```bash
   python main.py --transport http --host 0.0.0.0 --port 8000
   ```
   - 然后用 ngrok 暴露：`ngrok http 8000`
   - server_url 示例：`https://xxxx.ngrok.io/mcp`

## 如何在 AI 工具中注册

### Claude Desktop / Cursor
在配置文件中添加：
```json
{
  "mcpServers": {
    "labscare-script": {
      "command": "python",
      "args": ["/path/to/labscare-script-mcp/main.py", "--transport", "stdio"]
    }
  }
}
```

### Grok / xAI 或其他支持 Remote MCP 的客户端
配置 server_url：
```json
{
  "tools": [
    {
      "type": "remote_mcp",
      "server_url": "https://your-ngrok-or-domain/mcp"
    }
  ]
}
```

## 使用示例（直接复制到 Grok/Claude/Cursor）

```
使用 labscare-script-mcp 帮我生成一个「血常规检测报告」的取数 JS 脚本。

报告类型：常规检验报告
需要的数据源：样本基本信息 + 检测结果 + 参考范围
特殊需求：结果异常时标红；按项目分组；输出格式符合报告引擎 {sections: [...], tables: [...]}

请先调用需要的工具收集情报（如 get_labscare_sampledata、explain_labscare_field），然后生成脚本，最后如果有 simulate 工具请验证一下。
```

## 如何添加/更新规则（知识库）

### 方式1：直接编辑 JSON 文件（最简单）
修改 `knowledge/fields.json` / `patterns.json` / `diagnoses.json`，保存后下次 Tool 调用自动加载（无需重启）。

### 方式2：让 AI 调用 update_labscare_knowledge
示例 Prompt：
```
t_patient_name 是患者姓名字段，在 formJs 中取值用 formJs['t_patient_name'].val
请调用 update_labscare_knowledge 添加这个字段解释。
```

Tool 会自动写入 JSON。

## 计划中的功能（欢迎 PR！）
- [Y] JS 脚本沙箱模拟执行 & 覆盖率验证 Tool（simulate_labscare_script）
- [Y] 完整测试用例 & mock 接口数据
- [ ] 带批注 DOC 自动解析 Tool（parse_annotated_doc）
- [ ] 更多 Resources 自动注册（示例脚本、样本 DOC）
- [ ] HTTP 模式下添加 Authorization 认证
- [ ] Dockerfile 一键部署

