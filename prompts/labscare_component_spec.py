"""
LabsCare MCP Server - 组件取值处理规范 Prompt 注册
"""


COMPONENT_SPEC_PROMPT = """
# LabsCare 报告组件取值处理规范

在 LabsCare 报告引擎脚本中，LIMS 表单组件的原始值需要按照以下规则转换后才能赋值给 tplData，供报告页面渲染。

---

## 1. 签名组件（Signature）

### 接口返回：对象形式
```json
{ "url": "xxxx" }
```
**处理结果**（字符串）：
```javascript
// 字段名示例：signField
var result = headerUrl(formJs['signField'].url);
```

### 接口返回：数组形式
```json
[ { "url": "xxxx" }, { "url": "yyyy" } ]
```
**处理结果**（对象数组）：
```javascript
var rawArr = formJs['signField'];  // 原始数组
var result = [];
for (var i = 0; i < rawArr.length; i++) {
  result.push({ url: headerUrl(rawArr[i].url) });
}
// result 形如：[{ url: "https://..." }, { url: "https://..." }]
```

> **判断逻辑**：用 `Array.isArray()` 或 `instanceof Array` 区分对象与数组，ES5 环境推荐用 `Object.prototype.toString.call(val) === '[object Array]'`。

---

## 2. 图片组件（Image）

### 接口返回：逗号分隔的字符串（单张或多张）
```
"fileKey1,fileKey2,fileKey3"
```
**处理结果**（对象数组）：
```javascript
var rawStr = formJs['imgField'] || '';
var keys = rawStr.split(',');
var result = [];
for (var i = 0; i < keys.length; i++) {
  if (keys[i]) {
    result.push({ url: ossFileUrl(keys[i]) });
  }
}
// result 形如：[{ url: "https://..." }, { url: "https://..." }]
```

> **注意**：需过滤空字符串，防止末尾逗号产生空项。

---

## 3. 单选组件（Radio）

### 接口返回：对象形式，包含选中项文本
```json
{ "val": "是" }
```
**处理结果**（HTML 字符串）：

将所有可选项逐一渲染，选中项调用 `getCheckBox(true)`，其余调用 `getCheckBox(false)`，拼接为 HTML 字符串。

```javascript
// 可选项需从 tplData 上下文或已知配置中获取
// 示例：可选项为 ['是', '否']，选中值为 formJs['radioField'].val

var selectedVal = formJs['radioField'].val;
var options = ['是', '否'];  // 所有可选项，需根据实际字段配置填写
var result = '';
for (var i = 0; i < options.length; i++) {
  var isChecked = (options[i] === selectedVal);
  result += getCheckBox(isChecked) + options[i];
}
// 选中"是"时，result 形如：☑是☐否
```

> **注意**：`getCheckBox(true)` 返回 `☑`，`getCheckBox(false)` 返回 `☐`，最终拼接为 HTML 字符串直接嵌入报告模板。

---

## 4. 各组件处理方式汇总

| 组件类型 | 接口返回类型 | 处理方法 | 最终结果类型 |
|---|---|---|---|
| 签名（对象） | `{ url: string }` | `headerUrl(val.url)` | 字符串 |
| 签名（数组） | `Array<{ url: string }>` | 遍历调用 `headerUrl` | `Array<{ url: string }>` |
| 图片 | 逗号分隔字符串 | `split(',')` + `ossFileUrl` | `Array<{ url: string }>` |
| 单选 | `{ val: string }` | 遍历可选项 + `getCheckBox` | HTML 字符串 |

---

## 5. 组件处理代码的位置

所有组件处理逻辑应在 `tplData` 组装之前完成，处理后的结果直接挂载到 `tplData`：

```javascript
// 签名处理
var signRaw = formJs['signField'];
if (Object.prototype.toString.call(signRaw) === '[object Array]') {
  var signResult = [];
  for (var i = 0; i < signRaw.length; i++) {
    signResult.push({ url: headerUrl(signRaw[i].url) });
  }
  tplData['signField'] = signResult;
} else {
  tplData['signField'] = headerUrl(signRaw.url);
}

// 图片处理
var imgRaw = formJs['imgField'] || '';
var imgResult = [];
var imgKeys = imgRaw.split(',');
for (var i = 0; i < imgKeys.length; i++) {
  if (imgKeys[i]) { imgResult.push({ url: ossFileUrl(imgKeys[i]) }); }
}
tplData['imgField'] = imgResult;

// 单选处理
var selectedVal = formJs['radioField'].val;
var options = ['是', '否'];
var radioResult = '';
for (var i = 0; i < options.length; i++) {
  radioResult += getCheckBox(options[i] === selectedVal) + options[i];
}
tplData['radioField'] = radioResult;
```
"""


def labscare_component_spec() -> str:
    """
    LabsCare 报告组件取值处理规范。
    描述签名组件、图片组件、单选组件等在报告引擎脚本中的
    原始值格式及对应的 JS 转换处理方式。
    生成报告脚本前应先读取此规范。
    """
    return COMPONENT_SPEC_PROMPT