"""
LabsCare MCP Server - script-spec Resource 注册
将报告引擎 JS 规范文档以 Resource 形式暴露给 LLM
"""

SCRIPT_SPEC_CONTENT = """
# LabsCare 报告引擎 JS 脚本规范

## 1. 执行环境

- **运行时**：Nashorn / GraalJS（ES5 语法兼容）
- **禁止使用**：ES6+ 语法（`let`、`const`、箭头函数、模板字符串、`async/await` 等）
- **脚本出口**：脚本最后一行的表达式值即为导出给报告页面的数据对象，通常命名为 `tplData`

---

## 2. 固定头部写法

每份脚本必须以以下两行开头：

```javascript
load("/tools.js");
set("exceptionMsgLengthLimit", "10000");
```

- `load("/tools.js")`：导入工具函数库，使 getCheckBox()、headerUrl()、ossFileUrl() 等方法可用
- `set("exceptionMsgLengthLimit", "10000")`：设置异常信息最大长度

---

## 3. 运行时暴露的全局方法

| 方法 | 说明 |
|---|---|
| `get(key)` | 从运行时上下文中取值 |
| `set(key, value)` | 向运行时上下文写入配置项 |
| `load(path)` | 导入外部 JS 文件 |

获取核心对象：
```javascript
var helper = get("labscareHelper");
var projectId = get("projectId");
```

---

## 4. helper 对象方法

| 方法 | 返回 | 说明 |
|---|---|---|
| `helper.getProjectSamples(projectId)` | Java 对象 | 获取项目下所有样品数据 |
| `helper.getProjectData(projectId)` | Java Map | 获取项目流程数据 |

Java 对象必须序列化处理：
```javascript
var samples = helper.getProjectSamples(projectId);
samples = JSON.stringify(samples);
samples = JSON.parse(samples.replace(/null:/g, '"null":'));
// null: 是 Java Map null key 序列化产生的非法 JSON，需替换为 "null":
```

---

## 5. 流程数据遍历

用 indexOf 前缀匹配定位目标流程（ES5 无法用 startsWith）：
```javascript
var procedures = helper.getProjectData(projectId);
var processId = "";
var templateId = "";

for (var i in procedures) {
  if (i.indexOf("310") === 0) { processId = i; }
}

var procedure = procedures.get(processId);

if (procedure) {
  for (var i in procedure.processes) {
    if (i.indexOf("314") === 0) { templateId = i; }
  }
}
```

---

## 6. 表单数据获取

```javascript
var form = procedure.get("processes").get(templateId).get("form");
var formJs = JSON.parse(JSON.stringify(form));
```

---

## 7. tplData 组装与导出

```javascript
var samplesJs = JSON.parse(JSON.stringify(samples));
var reportNo = JSON.stringify(procedure.get('cateGroupName'));
var tplData = Object.assign({}, formJs, { 'samples': samplesJs });
tplData.reportNo = reportNo || formJs['<field_id>'] || '';

// 最后一行裸表达式 = 导出
tplData
```

---

## 8. tools.js 工具函数

| 函数 | 说明 |
|---|---|
| `getCheckBox(checked)` | 返回 ☑ 或 ☐ 方框字符 |
| `headerUrl(key)` | 返回签名图片 OSS 地址 |
| `ossFileUrl(key)` | 返回报告附件图片 OSS 地址 |
"""


def script_spec() -> str:
        """
        LabsCare 报告引擎 JS 脚本编写规范。
        包含执行环境说明、全局方法、helper 对象、
        流程数据遍历、tplData 组装等完整规范。
        """
        return SCRIPT_SPEC_CONTENT

    # 如果后续有更多 Resource 可在此继续注册
    # @mcp.resource("labscare://tools-api")
    # def tools_api() -> str:
    #     ...