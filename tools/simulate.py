# tools/simulate.py
from py_mini_racer import MiniRacer
from py_mini_racer.py_mini_racer import JSParseException
from typing import Dict, Any
import json
from pathlib import Path

def load_report_structure() -> Dict:
    """加载预期 tplData 结构（后续可按项目类型扩展）"""
    path = Path("knowledge/report_structure.json")
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except:
            pass
    # 兜底结构（你后面根据真实报告慢慢补充）
    return {
        "required_keys": [
            "patient_name", "sample_id", "test_item", "result", "unit",
            "reference_range", "abnormal_flag", "samples", "reportNo", "processes"
        ]
    }

def to_python(obj):
    # 获取类名字符串，避免直接引用不存在的 JSArray/JSObject 类
    type_name = type(obj).__name__
    
    if type_name == 'JSArray':
        return [to_python(item) for item in obj]
    elif type_name == 'JSObject':
        # 兼容处理：将 JSObject 转为字典
        return {k: to_python(obj[k]) for k in obj.keys()}
    else:
        # 基础类型直接返回
        return obj
async def simulate_labscare_script(
    js_code: str,
    lab: str,
    project_id: str
) -> Dict[str, Any]:
    """
    LabsCare ES5 脚本沙箱模拟执行（专为真实报告引擎设计）
    
    参数：
        js_code:    生成的完整 JS 脚本（必须包含最后 tplData）
        project_id: 项目ID（会同时调用两个真实接口取数）
    
    返回：
        {
            "success": bool,
            "output": dict,           # tplData 实际内容
            "errors": list[str],
            "coverage": float,        # 业务字段覆盖率
            "coverage_details": dict,
            "message": str
        }
    """
    ctx = MiniRacer()
    errors: list[str] = []
    output = None
    coverage = 0.0
    required = set()
    actual = set()

    try:
        # ==================== 1. 调用真实接口取数（同时调用两个） ====================
        from tools.handlers import handle_get_sampledata, handle_get_projectdata
        
        # 调用样本接口
        samples_raw = handle_get_sampledata({"lab": lab, "project_id": project_id})  # ← 按你实际 handler 签名调整
        if hasattr(samples_raw, "__await__"):
            samples_raw = await samples_raw
        if isinstance(samples_raw, str):
            samples_raw = json.loads(samples_raw)
        
        # 调用项目/流程接口
        procedures_raw = handle_get_projectdata({"lab": lab, "project_id": project_id})
        if hasattr(procedures_raw, "__await__"):
            procedures_raw = await procedures_raw
        if isinstance(procedures_raw, str):
            procedures_raw = json.loads(procedures_raw)
        
        # 确保 procedures 是一个对象
        if not isinstance(procedures_raw, dict):
            procedures_raw = procedures_raw.get("procedures", {}) if hasattr(procedures_raw, "get") else {}
        
        # 构造注入结构（完全匹配你脚本里的写法）
        data_to_inject = {
            "samples": samples_raw if isinstance(samples_raw, list) else samples_raw.get("samples", []),
            "procedures": procedures_raw
        }

        # ==================== 2. 注入引擎全局环境 ====================
        ctx.eval(f"""
            var projectId = "{project_id}";
            var mockData = {json.dumps(data_to_inject, ensure_ascii=False)};
            
            // 确保 procedures 是一个真正的对象（而不是字符串）
            if (typeof mockData.procedures === 'string') {{
                try {{
                    mockData.procedures = JSON.parse(mockData.procedures);
                }} catch(e) {{
                    mockData.procedures = {{}};
                }}
            }}
            
            // 递归函数：为对象及其所有子对象添加 .get() 方法
            function addGetMethod(obj) {{
                if (!obj || typeof obj !== 'object') {{
                    return;
                }}
                if (!obj.get) {{
                    obj.get = function(key) {{
                        return Object.prototype.hasOwnProperty.call(this, key) ? this[key] : null;
                    }};
                }}
                for (var key in obj) {{
                    if (Object.prototype.hasOwnProperty.call(obj, key)) {{
                        var value = obj[key];
                        if (value && typeof value === 'object' && !Array.isArray(value)) {{
                            addGetMethod(value);
                        }}
                    }}
                }}
            }}
            
            // 模拟报告引擎运行时（完全兼容你脚本里的 helper）
            // 让 procedures 对象及其所有子对象支持 .get() 方法
            addGetMethod(mockData.procedures);
            function load(path) {{ return true; }}
            function set(key, value) {{}}
            function get(name) {{
                if (name === "labscareHelper") {{
                    return {{
                        getProjectSamples: function(pid) {{ return mockData.samples; }},
                        getProjectData: function(pid) {{ return mockData.procedures; }}
                    }};
                }}
                return null;
            }}
          
        """)

        # ==================== 3. 包裹脚本 + 捕获最后 tplData ====================
        # 检测最后一行变量名（不一定是 tplData）
        lines = js_code.strip().split('\n')
        last_line = lines[-1].strip()
        # 去掉可能的注释和分号
        last_line = last_line.split('//')[0].strip().rstrip(';')
        last_var = last_line

        # ==================== 4. 执行沙箱 ====================
        # 在最后添加 JSON.stringify 来确保返回字符串
        full_script = f"""
                    {js_code}
                    JSON.stringify({last_var});
                    """
        js_result = ctx.eval(full_script)
        python_result = json.loads(js_result) # 转为 Python 字典
        output = python_result if isinstance(python_result, dict) else {"raw": python_result}

        # ==================== 5. 业务覆盖率计算 ====================
        structure = load_report_structure()
        required = set(structure.get("required_keys", []))
        actual = set(output.keys()) if isinstance(output, dict) else set()
        
        hit = len(required & actual)
        coverage = round((hit / len(required) * 100), 2) if required else 100.0

        success = (coverage >= 95.0) and ("error" not in output)

    except JSParseException as e:
        errors.append(f"JS 执行错误: {str(e)}")
        success = False
    except Exception as e:
        errors.append(f"模拟器异常: {str(e)}")
        success = False

    return {
        "success": success,
        "output": output,
        "errors": errors,
        "coverage": coverage,
        "coverage_details": {
            "missing_keys": list(required - actual),
            "hit_rate": f"{coverage}%"
        },
        "message": "✅ 模拟通过！tplData 结构正确，可直接预览" if success else "⚠️ 需要优化"
    }