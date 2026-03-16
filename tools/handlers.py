"""工具处理函数模块"""
import json
import httpx
from datetime import datetime
from tools.knowledge import (
    FIELD_DB, PATTERNS, DIAGNOSES,
    KB_FILE, _load, _save, reload_knowledge
)


def handle_generate_script(args: dict) -> str:
    report_type    = args.get("report_type", "")
    data_sources   = args.get("data_sources", ["getProjectData", "getProjectSamples"])
    placeholders   = args.get("placeholders", [])
    special        = args.get("special_needs", {})
    layer          = args.get("layer", "single")

    need_procedures = "getProjectData" in data_sources
    need_samples    = "getProjectSamples" in data_sources or "getSamples" in data_sources
    need_case       = "getCase" in data_sources

    checkboxes      = special.get("checkboxes", [])
    cb_strategy     = special.get("checkbox_strategy", "precompute")
    signatures      = special.get("signatures", [])
    sig_strategy    = special.get("signature_strategy", "precompute")
    subtables       = special.get("subtables", [])
    min_rows        = special.get("min_rows", 0)
    multipage       = special.get("multipage", False)
    retest          = special.get("retest", False)
    std_indicator   = special.get("standard_indicator", False)
    filter_term     = special.get("filter_terminated", True)
    dropdown_fields = special.get("dropdown_fields", [])
    extra           = special.get("extra_instructions", "")

    lines = []

    lines += [
        "//javascript",
        'load("/tools.js");',
        'set("exceptionMsgLengthLimit", "10000");',
    ]
    if checkboxes and cb_strategy == "inline_js":
        lines.append('set("getCheckBox", getCheckBox);')
    if signatures and sig_strategy == "inline_js":
        lines.append('set("signUrl", headerUrl);')
    lines += ["", 'var helper = get("labscareHelper");', ""]

    if need_samples:
        lines += [
            "var samples = helper.getProjectSamples(projectId);",
            'samples = JSON.parse(JSON.stringify(samples).replace(/null:/g, \'"null":\'));',
            "",
        ]

    if need_procedures:
        lines += [
            "var procedures = helper.getProjectData(projectId);",
            "",
            "var templateId = '';",
            "var processId = '';",
            "for (var i in procedures) {",
            "    if (i.indexOf('310') === 0) { processId = i; }",
            "}",
            "var procedure = procedures.get(processId);",
            "if (procedure) {",
            "    for (var i in procedure.processes) {",
            "        if (i.indexOf('314') === 0) { templateld = i; }",
            "    }",
            "}",
            "var form = procedure.get('processes').get(templateld).get('form');",
            'var formJs = JSON.parse(JSON.stringify(form).replace(/null:/g, \'"null":\'));',
            "",
        ]

    if need_case:
        lines += [
            "var cases;",
            "if (caseIdStr) { cases = helper.getCase(caseIdStr); }",
            'cases = JSON.parse(JSON.stringify(cases).replace(/null:/g, \'"null":\'));',
            "",
        ]

    need_lodash = (signatures and sig_strategy == "precompute") or subtables or dropdown_fields
    if need_lodash:
        lines += [
            "function lodashGet(object, path) {",
            "    if (object == null) return null;",
            "    var pathArray = Array.isArray(path) ? path : path",
            "        .replace(/\\[\"([^\"]+)\"\\]/g, '.$1')",
            "        .replace(/\\['([^']+)'\\]/g, '.$1')",
            "        .replace(/\\[([^\\]]+)\\]/g, '.$1')",
            "        .replace(/^\\./, '').split('.');",
            "    var result = object;",
            "    for (var j = 0; j < pathArray.length; j++) {",
            "        var key = pathArray[j];",
            "        if (result == null || result[key] === undefined) return null;",
            "        result = result[key];",
            "    }",
            "    return result !== undefined ? result : null;",
            "}",
            "",
        ]

    if checkboxes and cb_strategy == "precompute":
        lines += [
            "var SPACER = '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;';",
            "",
            "function buildMultiCheckbox(fieldArr, optionList) {",
            "    var strVal = [];",
            "    optionList.forEach(function(e) {",
            "        var index = -1;",
            "        if (fieldArr && fieldArr.length > 0) {",
            "            for (var i = 0; i < fieldArr.length; i++) {",
            "                if (fieldArr[i].val && e.indexOf(fieldArr[i].val) > -1) { index = i; break; }",
            "            }",
            "        }",
            "        strVal.push(index > -1 ? getCheckBox(index + 1) + e : getCheckBox() + e);",
            "    });",
            "    return strVal.join(SPACER);",
            "}",
            "",
            "function buildSingleCheckbox(fieldObj, optionList) {",
            "    var v = fieldObj && fieldObj.val ? fieldObj.val : '';",
            "    return optionList.map(function(e) {",
            "        return (e === v ? getCheckBox(fieldObj.val) : getCheckBox()) + e;",
            "    }).join(SPACER);",
            "}",
            "",
            "function getYesNoHTML(formData, keys, labels) {",
            "    var str = '';",
            "    keys.forEach(function(key, i) {",
            "        var val = formData[key] && formData[key].val ? formData[key].val : '';",
            "        str += getCheckBox(val === '是') + labels[i] + SPACER;",
            "    });",
            "    return str;",
            "}",
            "",
        ]

    if subtables:
        if min_rows > 0:
            lines += [
                f"function getTableData(formData) {{",
                f"    var origin = formData['{subtables[0]}'] || [];",
                f"    var count = origin.length < {min_rows} ? {min_rows} : origin.length;",
                "    var result = [];",
                "    for (var i = 0; i < count; i++) {",
                "        var row = origin[i];",
                "        if (row) { row['index'] = (i + 1) + ''; }",
                "        else { row = { index: '' }; }",
                "        result.push(row);",
                "    }",
                "    return result;",
                "}",
                "",
            ]
        else:
            lines += [
                "function getTableData(formData) {",
                f"    var origin = formData['{subtables[0]}'] || [];",
                "    return origin.map(function(row, i) {",
                "        return Object.assign({}, row, { index: (i + 1) + '' });",
                "    });",
                "}",
                "",
            ]

    if retest:
        lines += [
            "function isRetestItem(item) {",
            "    return item.t_sffc && item.t_sffc.val === '是'",
            "        && item.t_bgtqjg && item.t_bgtqjg.val === '复测结果';",
            "}",
            "",
        ]

    if std_indicator:
        lines += [
            "function getStandardIndicator(item) {",
            "    var lower = item.t_zxyxx, upper = item.t_zdyxx;",
            "    if (lower) {",
            "        var l = isNaN(parseFloat(lower)) ? lower : '≥' + lower;",
            "        return upper ? l + '-' + (isNaN(parseFloat(upper)) ? upper : '≤' + upper) : l;",
            "    }",
            "    return upper ? (isNaN(parseFloat(upper)) ? upper : '≤' + upper) : '/';",
            "}",
            "",
        ]

    if need_samples and filter_term and not multipage:
        lines += [
            "var outputData = [];",
            "",
            "for (var idx = 0; idx < samples.length; idx++) {",
            "    var sample = samples[idx];",
            "    if (sample['s_zhongzhi'] === '是') { continue; }",
            "",
            "    var rowData = {};",
            "",
        ]
        if any(p in ["factorName"] for p in placeholders):
            lines += [
                "    var factorNames = [];",
                "    var gtl = sample['gaugingTableList'] || [];",
                "    for (var j = 0; j < gtl.length; j++) {",
                "        if (gtl[j].factorName) factorNames.push(gtl[j].factorName.trim());",
                "    }",
                "    rowData.factorName = factorNames.join('、');",
                "",
            ]
        if signatures and sig_strategy == "precompute":
            for sig in signatures:
                lines.append(f"    rowData.{sig} = headerUrl + (lodashGet(sample, ['{sig}', 0, 'signUrl']) || '');")
            lines.append("")
        lines += [
            "    outputData.push(rowData);",
            "}",
            "",
            "outputData;",
        ]

    elif need_procedures and not multipage:
        overrides = []
        for f in dropdown_fields:
            overrides.append(f"    {f}: lodashGet(formJs, ['{f}', 'val']),")
        if signatures and sig_strategy == "precompute":
            for sig in signatures:
                overrides.append(f"    {sig}: headerUrl + (lodashGet(formJs, ['{sig}', 0, 'signUrl']) || ''),")
        if subtables:
            overrides.append("    tableData: getTableData(formJs),")

        if overrides:
            lines += ["var returnData = Object.assign({}, formJs, {"]
            lines += overrides
            lines += ["});", "", "returnData;"]
        else:
            lines += ["formJs;"]

    if multipage:
        lines = lines[:lines.index("")+2] + [
            "var outputData = {",
            "    'page1': { /* 第1页数据 */ },",
            "    'page2': { /* 第2页数据 */ },",
            "    'page':  { /* 所有页共享数据 */ },",
            "};",
            "outputData;",
        ]

    script = "\n".join(lines)

    notes = []
    if checkboxes and cb_strategy == "precompute":
        notes.append("复选框用脚本预计算，模板用 `${字段名_boxes}` 占位符。")
    if checkboxes and cb_strategy == "inline_js":
        notes.append("复选框逻辑写在模板内联 JS 中，已通过 `set('getCheckBox', getCheckBox)` 导出。")
    if signatures and sig_strategy == "precompute":
        notes.append("签名已预计算完整 URL，模板用 `=字段名` 嵌入图片。")
    if signatures and sig_strategy == "inline_js":
        notes.append("签名 URL 由模板内联 JS 拼接，已通过 `set('signUrl', headerUrl)` 导出。")
    if subtables:
        notes.append(f"子表格组件ID `{subtables[0]}`，请补充空行所需的列字段名。")
    if extra:
        notes.append(f"额外需求：{extra}")

    out = f"# LabsCare 报表脚本 — {report_type}\n\n"
    out += "```javascript\n" + script + "\n```\n"
    if notes:
        out += "\n**注意事项**\n" + "\n".join(f"- {n}" for n in notes) + "\n"
    out += (
        "\n**下一步**\n"
        "1. 补充 TODO 处的字段取值\n"
        "2. 检查占位符与脚本字段名是否一一对应\n"
        "3. 下拉选择器字段记得取 `.val`\n"
        "4. 使用 `debug_labscare_script` 工具排查问题\n"
    )
    return out


def handle_explain_field(args: dict) -> str:
    field   = args.get("field_name", "")
    context = args.get("context", "")

    info = FIELD_DB.get(field)
    if info:
        return (
            f"## 字段：`{field}`\n\n"
            f"**类型**：{info.get('type', '—')}\n\n"
            f"**说明**：{info.get('desc', '—')}\n\n"
            f"**示例**：\n```javascript\n{info.get('example', '—')}\n```\n"
        )

    prefix_map = {
        "s_": ("样品模板用户配置标签", "samplesJs 数组中的样品对象"),
        "g_": ("阈值模板用户配置标签", "sample.gaugingTableList 中的检测项目对象"),
        "c_": ("档案模板用户配置标签", "helper.getCase() 返回的档案对象"),
        "p_": ("步骤模板用户配置标签", "procedure 的 form 表单，通过 formJs 访问"),
        "t_": ("通用标签（跨模板）",   "formJs 或 sample 对象，取决于具体使用位置"),
    }
    for prefix, (tag_type, location) in prefix_map.items():
        if field.startswith(prefix):
            return (
                f"## 字段：`{field}`\n\n"
                f"**类型**：{tag_type}（用户自定义，因客户模板而异）\n\n"
                f"**数据位置**：{location}\n\n"
                f"**命名规则**：拼音首字母，重复时加 `_序号`\n\n"
                f"**注意**：字段值可能是纯文本、下拉对象 `{{val:'...'}}` 或签名数组，"
                f"需根据实际数据结构决定取值方式。\n"
                + (f"\n**上下文**：在 `{context}` 中访问此字段。" if context else "")
                + f"\n\n💡 使用 `update_labscare_knowledge` 可以为此字段添加详细说明到知识库。\n"
            )

    return (
        f"未找到字段 `{field}` 的说明。\n\n"
        "**可能的情况**：\n"
        "- 纯数字ID（如 `3151663582128416900`）：未配标签的子表格组件，用 `formJs['<数字ID>']` 访问\n"
        "- `s.` 前缀（如 `s.t_samplesDate`）：路径访问方式，表示从样品对象中读取通用标签\n"
        "- 系统固定字段（无前缀）：如 `tempId`、`factorName`、`caseName`\n\n"
        f"💡 使用 `update_labscare_knowledge` 添加此字段到知识库。\n"
    )


def handle_debug_script(args: dict) -> str:
    symptom      = args.get("symptom", "")
    snippet      = args.get("script_snippet", "")
    placeholders = args.get("template_placeholders", "")

    out = f"## 诊断：{symptom}\n\n"

    matched_key = next((k for k in DIAGNOSES if k in symptom), None)
    matched = DIAGNOSES.get(matched_key) if matched_key else None

    if matched:
        out += "### 可能原因\n"
        for i, c in enumerate(matched.get("causes", []), 1):
            out += f"{i}. {c}\n"
        out += "\n### 修复方案\n"
        for fix in matched.get("fixes", []):
            out += f"- {fix}\n"
    else:
        out += (
            "未匹配已知错误模式，通用排查清单：\n\n"
            "1. `load('/tools.js')` 和 `set('exceptionMsgLengthLimit', '10000')` 是否存在\n"
            "2. 所有 `JSON.parse` 前是否有 `.replace(/null:/g, '\"null\":')`\n"
            "3. `templateld`（末尾小写l）拼写是否正确\n"
            "4. Java Map 原始对象是否用 `.get('key')` 取值\n"
            "5. 下拉字段是否遗漏 `.val` 提取\n"
            "6. `<data>` 绑定的数组是否在返回对象顶层\n"
            "7. 模板内联 JS 用到的函数是否都通过 `set()` 导出\n"
        )

    if snippet:
        warnings = []
        if "templateId" in snippet and "templateld" not in snippet:
            warnings.append("⚠️ 发现 `templateId`（大写I），应为 `templateld`（小写l）")
        if "let " in snippet or "const " in snippet:
            warnings.append("⚠️ 发现 ES6 `let`/`const`，应改为 `var`")
        if "=>" in snippet:
            warnings.append("⚠️ 发现箭头函数 `=>`，建议改为 `function(){}` 保持 ES5 兼容")
        if "JSON.parse" in snippet and 'replace' not in snippet:
            warnings.append("⚠️ `JSON.parse` 前未做 null 键修复")
        if "getCheckBox" in snippet and "set(" not in snippet:
            warnings.append("⚠️ 使用了 `getCheckBox` 但未见 `set('getCheckBox', getCheckBox)`")
        if "signUrl" in snippet and "set(" not in snippet:
            warnings.append("⚠️ 模板使用了 `signUrl` 但未见 `set('signUrl', headerUrl)`")

        if warnings:
            out += "\n### 代码静态检查\n" + "\n".join(warnings) + "\n"
        else:
            out += "\n### 代码静态检查\n✅ 未发现常见语法问题\n"

    if placeholders:
        out += f"\n### 占位符分析：`{placeholders}`\n"
        if "${" in placeholders and ".val" not in placeholders and "lookup" not in placeholders:
            out += "- 若字段是下拉选择器，直接 `${字段名}` 会渲染为 `[object Object]`，建议改用 `${字段名.val ? 字段名.val : ''}` 或脚本预处理\n"
        if "=" in placeholders and "${" not in placeholders:
            out += "- `=字段名` 在普通区域嵌入图片，脚本需提供完整 URL 字符串\n"

    out += f"\n💡 若是新类型问题，可用 `update_labscare_knowledge` 将此诊断经验存入知识库。\n"
    return out


def handle_get_pattern(args: dict) -> str:
    pattern_name = args.get("pattern_name", "")
    p = PATTERNS.get(pattern_name)
    if not p:
        return (
            f"未找到模式：`{pattern_name}`\n\n"
            f"当前可用模式（{len(PATTERNS)} 个）：\n"
            + "\n".join(f"- `{k}`：{v.get('title','')}" for k, v in sorted(PATTERNS.items()))
        )
    out = f"## {p.get('title', pattern_name)}\n\n```javascript\n{p.get('code', '')}\n```\n"
    if p.get("note"):
        out += f"\n**注意**：{p['note']}\n"
    return out


def handle_update_knowledge(args: dict) -> str:
    operation      = args.get("operation", "add")
    knowledge_type = args.get("knowledge_type", "")
    key            = args.get("key", "").strip()
    content        = args.get("content", {})

    if not key:
        return "❌ key 不能为空"

    filename = KB_FILE.get(knowledge_type)
    if not filename:
        return f"❌ 未知 knowledge_type：{knowledge_type}"

    db = _load(filename)

    if operation == "delete":
        if key not in db:
            return f"⚠️ 条目 `{key}` 不存在于 {knowledge_type} 知识库中"
        del db[key]
        action_desc = f"已删除 `{key}`"

    elif operation in ("add", "update"):
        if not content:
            return "❌ add/update 操作需要提供 content"

        if operation == "add" and key in db:
            return (
                f"⚠️ 条目 `{key}` 已存在，如需覆盖请使用 operation='update'。\n"
                f"现有内容：\n```json\n{json.dumps(db[key], ensure_ascii=False, indent=2)}\n```"
            )

        if knowledge_type == "pattern" and "note" not in content:
            content["note"] = ""

        db[key] = content
        action_desc = f"{'新增' if operation == 'add' else '更新'}了 `{key}`"

    else:
        return f"❌ 未知操作：{operation}"

    _save(filename, db)
    reload_knowledge()

    summary = (
        f"✅ {action_desc}（{knowledge_type} 知识库）\n\n"
        f"**文件**：`knowledge/{filename}`\n"
        f"**时间**：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"**当前条目数**：{len(db)}\n\n"
        f"热重载完成，下次调用相关 Tool 时立即生效。\n"
    )
    if operation != "delete":
        summary += f"\n**写入内容**：\n```json\n{json.dumps(content, ensure_ascii=False, indent=2)}\n```\n"

    return summary


def handle_list_knowledge(args: dict) -> str:
    knowledge_type = args.get("knowledge_type", "all")

    sections = []

    def fmt_field(k, v):
        return f"- `{k}`：{v.get('type', '—')} — {v.get('desc', '')[:60]}{'…' if len(v.get('desc',''))>60 else ''}"

    def fmt_pattern(k, v):
        return f"- `{k}`：{v.get('title', '—')}"

    def fmt_diagnosis(k, v):
        causes_count = len(v.get("causes", []))
        fixes_count  = len(v.get("fixes", []))
        return f"- `{k}`：{causes_count} 个原因 / {fixes_count} 个修复方案"

    if knowledge_type in ("field", "all"):
        entries = "\n".join(fmt_field(k, v) for k, v in sorted(FIELD_DB.items()))
        sections.append(f"## 字段解释库（{len(FIELD_DB)} 条）\n\n{entries}")

    if knowledge_type in ("pattern", "all"):
        entries = "\n".join(fmt_pattern(k, v) for k, v in sorted(PATTERNS.items()))
        sections.append(f"## 代码模式库（{len(PATTERNS)} 条）\n\n{entries}")

    if knowledge_type in ("diagnosis", "all"):
        entries = "\n".join(fmt_diagnosis(k, v) for k, v in sorted(DIAGNOSES.items()))
        sections.append(f"## 问题诊断库（{len(DIAGNOSES)} 条）\n\n{entries}")

    if not sections:
        return f"❌ 未知 knowledge_type：{knowledge_type}"

    return "\n\n---\n\n".join(sections) + "\n\n💡 使用 `update_labscare_knowledge` 添加或修改条目。\n"

async def handle_get_sampledata(args: dict) -> str:
    lab = args.get("lab", "")
    project_id = args.get("project_id", "")
    if not project_id:
        return "❌ project_id 不能为空"
    # 读取项目根目录下的config.json文件
    import os
    config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config.json")
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except Exception as e:
        return f"❌ 读取config.json失败：{str(e)}"
    
    # 获取实验室配置
    lab_config = config.get("labs", {}).get(lab)
    if not lab_config:
        return f"❌ 未找到实验室配置：{lab}"
    
    # 使用配置中的token
    base_url = lab_config.get("base_sample_url", "")
    url = f"{base_url}/{project_id}/v1"
    token = lab_config.get("token", "")
    headers = {
        "token": f"{token}",
        "Content-Type": "application/json"
    }
     # 3. 发起请求
    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers=headers, timeout=10.0)
        
    if response.status_code == 200:
        return json.dumps(response.json())
    else:
        return f"LIMS 请求失败: {response.status_code}"

async def handle_get_projectdata(args: dict) -> str:
    lab = args.get("lab", "")
    project_id = args.get("project_id", "")
    if not project_id:
        return "❌ project_id 不能为空"
    # 读取项目根目录下的config.json文件
    import os
    config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config.json")
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except Exception as e:
        return f"❌ 读取config.json失败：{str(e)}"
    
    # 获取实验室配置
    lab_config = config.get("labs", {}).get(lab)
    if not lab_config:
        return f"❌ 未找到实验室配置：{lab}"
    
    # 使用配置中的token
    base_url = lab_config.get("base_project_url", "")
    url = f"{base_url}/{project_id}/v1"
    token = lab_config.get("token", "")
    headers = {
        "token": f"{token}",
        "Content-Type": "application/json"
    }
     # 3. 发起请求
    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers=headers, timeout=10.0)
        
    if response.status_code == 200:
        return json.dumps(response.json())
    else:
        return f"LIMS 请求失败: {response.status_code}"