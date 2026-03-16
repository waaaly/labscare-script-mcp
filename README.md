# LabsCare Script MCP Server v2

知识库与框架完全分离。规则存在 `knowledge/*.json`，无需改 Python 代码即可更新。

---

## 项目结构

```
labscare-mcp-v2/
├── server.py              ← 框架逻辑（只改这里来改行为）
├── requirements.txt
└── knowledge/
    ├── fields.json        ← 字段解释库（可直接编辑）
    ├── patterns.json      ← 代码模式库（可直接编辑）
    └── diagnoses.json     ← 问题诊断库（可直接编辑）
```

---

## 6 个 Tools

| Tool | 功能 |
|------|------|
| `generate_labscare_script` | 按需求生成完整报表 JS 脚本 |
| `explain_labscare_field` | 解释字段含义和取值方式 |
| `debug_labscare_script` | 诊断常见问题 + 静态代码检查 |
| `get_labscare_pattern` | 获取代码模式片段 |
| `update_labscare_knowledge` | **运行时更新知识库（热重载，无需重启）** |
| `list_labscare_knowledge` | 查看当前知识库所有条目 |

---

## 安装与启动

```bash
pip install mcp
python server.py
```

## 注册到 Claude Desktop

```json
{
  "mcpServers": {
    "labscare-script": {
      "command": "python",
      "args": ["/绝对路径/labscare-mcp-v2/server.py"]
    }
  }
}
```

---

## 如何添加新规则

### 方式 A：告诉 Claude，让它调用 Tool 写入

```
记住：t_newfield 是下拉字段，值格式 {val:'xxx'}，
在 form 表单中取值用 formJs['t_newfield'].val
```

Claude 会调用 `update_labscare_knowledge`：
```json
{
  "operation": "add",
  "knowledge_type": "field",
  "key": "t_newfield",
  "content": {
    "type": "用户配置标签 t_ | {val: string}",
    "desc": "xxx，下拉选择器，取值用 .val",
    "example": "formJs['t_newfield'].val"
  }
}
```

### 方式 B：直接编辑 JSON 文件

编辑 `knowledge/fields.json` / `patterns.json` / `diagnoses.json`，
下次调用 Tool 时自动加载最新内容（无需重启）。

### 各知识库的条目格式

**fields.json**（字段解释）：
```json
"字段名": {
  "type": "字段类型描述",
  "desc": "详细说明",
  "example": "JavaScript 示例代码"
}
```

**patterns.json**（代码模式）：
```json
"模式key": {
  "title": "模式标题",
  "code": "JavaScript 代码片段",
  "note": "注意事项"
}
```

**diagnoses.json**（问题诊断）：
```json
"症状关键词": {
  "causes": ["原因1", "原因2"],
  "fixes": ["修复方案1", "修复方案2"]
}
```

---

## 典型更新场景

| 场景 | 操作 |
|------|------|
| 新客户有特殊字段约定 | `update_labscare_knowledge` add field |
| 发现新的常见报错 | `update_labscare_knowledge` add diagnosis |
| 沉淀了新的代码写法 | `update_labscare_knowledge` add pattern |
| 已有约定有变化 | `update_labscare_knowledge` update |
| 某个约定已废弃 | `update_labscare_knowledge` delete |
