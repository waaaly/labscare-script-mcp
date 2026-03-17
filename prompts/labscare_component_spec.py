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

## 6. 下拉单选组件（Dropdown Radio）

### 接口返回：对象形式，包含选中项文本（注意字段名可能有空格：`"val "`）
```json
{ "val ": "选项1", "id": "xxx", "oldId": "xxx" }
```
**处理结果**（HTML 字符串）：

处理方式与单选组件相同，将所有可选项逐一渲染，选中项调用 `getCheckBox(true)`，其余调用 `getCheckBox(false)`，拼接为 HTML 字符串。

```javascript
// 可选项需从 tplData 上下文或已知配置中获取
// 示例：可选项为 ['选项1', '选项2', '选项3']，选中值为 formJs['下拉单选组件'].val

var selectedVal = formJs['下拉单选组件'].val;  // 注意字段名可能有空格
var options = ['选项1', '选项2', '选项3'];  // 所有可选项，需根据实际字段配置填写
var result = '';
for (var i = 0; i < options.length; i++) {
  var isChecked = (options[i] === selectedVal);
  result += getCheckBox(isChecked) + options[i];
}
// 选中"选项1"时，result 形如：☑选项1☐选项2☐选项3
```

---

## 7. 多选组件（Checkbox）

### 接口返回：数组形式，每个元素包含 `val` 字段
```json
[
  { "val": "选项1", "checked": "0", "id": "xxx", "oldId": "xxx" },
  { "val": "选项2", "checked": "0", "id": "xxx", "oldId": "xxx" }
]
```
**处理结果**（HTML 字符串）：

多选组件返回的是当前已勾选的选项数组。渲染时需要将所有可选项逐一展示，对于已勾选的选项调用 `getCheckBox(true)`，未勾选的调用 `getCheckBox(false)`。

```javascript
// 可选项需从 tplData 上下文或已知配置中获取
// 示例：可选项为 ['选项1', '选项2', '选项3']，当前已勾选的值为 formJs['多选组件'][i].val

var selectedItems = formJs['多选组件'];  // 当前已勾选的选项数组
var options = ['选项1', '选项2', '选项3'];  // 所有可选项，需根据实际字段配置填写
var result = '';
for (var i = 0; i < options.length; i++) {
  var isChecked = selectedItems.some(function(item) { return item.val === options[i]; });
  result += getCheckBox(isChecked) + options[i];
}
// 若选中"选项1"和"选项3"，result 形如：☑选项1☐选项2☑选项3
```

---

## 8. 下拉多选组件（Dropdown Checkbox）

### 接口返回：数组形式，每个元素包含 `val` 字段
```json
[
  { "val": "选项2", "isTitle": "0", "checked": "0", "id": "xxx", "oldId": "xxx" },
  { "val": "选项3", "isTitle": "0", "checked": "0", "id": "xxx", "oldId": "xxx" }
]
```
**处理结果**（HTML 字符串）：

处理方式与多选组件相同，将所有可选项逐一渲染，已勾选的选项调用 `getCheckBox(true)`，其余调用 `getCheckBox(false)`，拼接为 HTML 字符串。

```javascript
// 可选项需从 tplData 上下文或已知配置中获取
// 示例：可选项为 ['选项1', '选项2', '选项3']，当前已勾选的值为 formJs['下拉多选组件'][i].val

var selectedItems = formJs['下拉多选组件'];  // 当前已勾选的选项数组
var options = ['选项1', '选项2', '选项3'];  // 所有可选项，需根据实际字段配置填写
var result = '';
for (var i = 0; i < options.length; i++) {
  var isChecked = selectedItems.some(function(item) { return item.val === options[i]; });
  result += getCheckBox(isChecked) + options[i];
}
// 若选中"选项2"和"选项3"，result 形如：☑选项2☑选项3☐选项1
```

---

## 9. 附件组件（Attachment）

### 接口返回：逗号分隔的字符串（单个或多个文件路径）
```
"/lab/2108593480794423566/all/2015319562860418758/20260317092500/test.png,/lab/2108593480794423566/all/2015319562860418758/20260317095847/systeminfo.png"
```
**处理结果**（对象数组）：

```javascript
var rawStr = formJs['附件组件'] || '';
var paths = rawStr.split(',');
var result = [];
for (var i = 0; i < paths.length; i++) {
  if (paths[i]) {
    result.push({ url: ossFileUrl(paths[i]) });
  }
}
// result 形如：[{ url: "https://oss/.../test.png" }, { url: "https://oss/.../systeminfo.png" }]
```

> **注意**：需过滤空字符串，防止末尾逗号产生空项。

---

## 10. 文本类组件（单行文本、多行文本、时间组件、日期组件）

### 接口返回：字符串形式
- 单行文本：`"文本内容"`
- 多行文本：`"多行\n内容"`
- 时间组件：`"00:00:03"`
- 日期组件：`"2026-03-17"`

**处理结果**（字符串）：

除非用户提示词有特殊处理要求，否则直接返回原值：

```javascript
tplData['单行文本'] = formJs['单行文本'] || '';
tplData['多行文本'] = formJs['多行文本'] || '';
tplData['时间组件'] = formJs['时间组件'] || '';
tplData['日期组件'] = formJs['日期组件'] || '';
```

---

## 11. 表格组件（Table）

### 接口返回：数组形式，每行是一个对象
```json
[
  {
    "下拉单选组件": "{}",
    "下拉多选组件": "[]"
  },
  {
    "下拉单选组件": "{}",
    "下拉多选组件": "[]"
  }
]
```
**处理结果**（对象数组）：

表格组件的每一行包含多个列组件，每列的处理方式按照其对应的组件类型进行处理：

```javascript
var tableRaw = formJs['表格组件'] || [];
var tableResult = [];
for (var i = 0; i < tableRaw.length; i++) {
  var row = {};
  
  // 下拉单选组件处理
  var dropdownRadioRaw = tableRaw[i]['下拉单选组件'];
  if (typeof dropdownRadioRaw === 'string' && dropdownRadioRaw !== '{}') {
    var selectedVal = JSON.parse(dropdownRadioRaw).val;
    var options = ['选项1', '选项2', '选项3'];
    var radioResult = '';
    for (var j = 0; j < options.length; j++) {
      radioResult += getCheckBox(options[j] === selectedVal) + options[j];
    }
    row['下拉单选组件'] = radioResult;
  }
  
  // 下拉多选组件处理
  var dropdownCheckboxRaw = tableRaw[i]['下拉多选组件'];
  if (typeof dropdownCheckboxRaw === 'string' && dropdownCheckboxRaw !== '[]') {
    var selectedItems = JSON.parse(dropdownCheckboxRaw);
    var options = ['选项1', '选项2', '选项3'];
    var checkboxResult = '';
    for (var j = 0; j < options.length; j++) {
      var isChecked = selectedItems.some(function(item) { return item.val === options[j]; });
      checkboxResult += getCheckBox(isChecked) + options[j];
    }
    row['下拉多选组件'] = checkboxResult;
  }
  
  tableResult.push(row);
}
tplData['表格组件'] = tableResult;
```

> **注意**：表格组件中的字段值可能是 JSON 字符串（如 `"{}"` 或 `"[]"`），需要先用 `JSON.parse()` 解析后再按对应组件类型处理。

---

## 4. 各组件处理方式汇总

| 组件类型 | 接口返回类型 | 处理方法 | 最终结果类型 |
|---|---|---|---|
| 签名（对象） | `{ url: string }` | `headerUrl(val.url)` | 字符串 |
| 签名（数组） | `Array<{ url: string }>` | 遍历调用 `headerUrl` | `Array<{ url: string }>` |
| 图片 | 逗号分隔字符串 | `split(',')` + `ossFileUrl` | `Array<{ url: string }>` |
| 单选 | `{ val: string }` | 遍历可选项 + `getCheckBox` | HTML 字符串 |
| 下拉单选 | `{ "val ": string }` | 遍历可选项 + `getCheckBox` | HTML 字符串 |
| 多选组件 | `Array<{ val: string, ... }>` | 遍历可选项 + `getCheckBox` | HTML 字符串 |
| 下拉多选 | `Array<{ val: string, ... }>` | 遍历可选项 + `getCheckBox` | HTML 字符串 |
| 附件组件 | 逗号分隔字符串 | `split(',')` + `ossFileUrl` | `Array<{ url: string }>` |
| 单行文本 | 字符串 | 原值返回 | 字符串 |
| 多行文本 | 字符串 | 原值返回 | 字符串 |
| 时间组件 | 字符串 | 原值返回 | 字符串 |
| 日期组件 | 字符串 | 原值返回 | 字符串 |
| 表格组件 | `Array<Object>` | 逐行处理，按列组件类型处理 | `Array<Object>` |

---

## 12. 组件处理代码的位置

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

// 下拉单选处理
var dropdownRadioRaw = formJs['下拉单选组件'];
var dropdownRadioVal = dropdownRadioRaw.val;
var dropdownOptions = ['选项1', '选项2', '选项3'];
var dropdownRadioResult = '';
for (var i = 0; i < dropdownOptions.length; i++) {
  dropdownRadioResult += getCheckBox(dropdownOptions[i] === dropdownRadioVal) + dropdownOptions[i];
}
tplData['下拉单选组件'] = dropdownRadioResult;

// 多选组件处理
var checkboxRaw = formJs['多选组件'];
var checkboxOptions = ['选项1', '选项2', '选项3'];
var checkboxResult = '';
for (var i = 0; i < checkboxOptions.length; i++) {
  var isChecked = checkboxRaw.some(function(item) { return item.val === checkboxOptions[i]; });
  checkboxResult += getCheckBox(isChecked) + checkboxOptions[i];
}
tplData['多选组件'] = checkboxResult;

// 下拉多选组件处理
var dropdownCheckboxRaw = formJs['下拉多选组件'];
var dropdownCheckboxOptions = ['选项1', '选项2', '选项3'];
var dropdownCheckboxResult = '';
for (var i = 0; i < dropdownCheckboxOptions.length; i++) {
  var isChecked = dropdownCheckboxRaw.some(function(item) { return item.val === dropdownCheckboxOptions[i]; });
  dropdownCheckboxResult += getCheckBox(isChecked) + dropdownCheckboxOptions[i];
}
tplData['下拉多选组件'] = dropdownCheckboxResult;

// 附件组件处理
var attachmentRaw = formJs['附件组件'] || '';
var attachmentResult = [];
var attachmentPaths = attachmentRaw.split(',');
for (var i = 0; i < attachmentPaths.length; i++) {
  if (attachmentPaths[i]) {
    attachmentResult.push({ url: ossFileUrl(attachmentPaths[i]) });
  }
}
tplData['附件组件'] = attachmentResult;

// 文本类组件处理（原值返回）
tplData['单行文本'] = formJs['单行文本'] || '';
tplData['多行文本'] = formJs['多行文本'] || '';
tplData['时间组件'] = formJs['时间组件'] || '';
tplData['日期组件'] = formJs['日期组件'] || '';

// 表格组件处理
var tableRaw = formJs['表格组件'] || [];
var tableResult = [];
for (var i = 0; i < tableRaw.length; i++) {
  var row = {};
  
  var dropdownRadioRaw = tableRaw[i]['下拉单选组件'];
  if (typeof dropdownRadioRaw === 'string' && dropdownRadioRaw !== '{}') {
    var selectedVal = JSON.parse(dropdownRadioRaw).val;
    var options = ['选项1', '选项2', '选项3'];
    var radioResult = '';
    for (var j = 0; j < options.length; j++) {
      radioResult += getCheckBox(options[j] === selectedVal) + options[j];
    }
    row['下拉单选组件'] = radioResult;
  }
  
  var dropdownCheckboxRaw = tableRaw[i]['下拉多选组件'];
  if (typeof dropdownCheckboxRaw === 'string' && dropdownCheckboxRaw !== '[]') {
    var selectedItems = JSON.parse(dropdownCheckboxRaw);
    var options = ['选项1', '选项2', '选项3'];
    var checkboxResult = '';
    for (var j = 0; j < options.length; j++) {
      var isChecked = selectedItems.some(function(item) { return item.val === options[j]; });
      checkboxResult += getCheckBox(isChecked) + options[j];
    }
    row['下拉多选组件'] = checkboxResult;
  }
  
  tableResult.push(row);
}
tplData['表格组件'] = tableResult;
```
"""


def labscare_component_spec() -> str:
    """
    LabsCare 报告组件取值处理规范。
    描述签名组件、图片组件、单选组件、下拉单选、多选组件、下拉多选、
    附件组件、文本类组件（单行文本、多行文本、时间组件、日期组件）、
    表格组件等在报告引擎脚本中的原始值格式及对应的 JS 转换处理方式。
    生成报告脚本前应先读取此规范。
    """
    return COMPONENT_SPEC_PROMPT