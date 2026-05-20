# 木牛流马 — 工程行业 AI 技能工具箱

工程行业通用智能体技能集，通过 MCP 协议对接任意智能体框架。
**37 个 MCP 工具，5 个独立模块。**

## 架构原则

```
纯文本任务                           多模态任务
（模板分析/数据匹配/报告生成）        （照片分析/影像分类）
        ↓                                  ↓
  智能体自己思考决策                  MCP 内嵌多模态 LLM
  MCP 只做数据读写                   智能体只发起调用
```

## 快速接入（给其他智能体使用）

```bash
git clone https://github.com/lvtaiyang/muniuliuma.git && cd muniuliuma
python check_env.py                                    # 检测环境
pip install mcp openai pyyaml httpx                    # 核心依赖
pip install python-docx openpyxl                       # 文档处理（可选）
# 编辑 config.yaml，填入 DeepSeek 和 Qwen 的 api_key
```

MCP 配置：
```json
{"mcpServers": {"muniuliuma": {"command": "python", "args": ["/path/to/muniuliuma/mcp_server.py"]}}}
```

---

# ⚠️ 智能体行为规范（必读）

## 规则 1：永远不要假设 API Key 已配置

**每次对话开始**，先用 `get_config` 检查 API Key 是否已填。如果没有：
- 纯文本任务 → 走纯 I/O 路径（智能体自己分析）
- 多模态任务 → 告诉用户需要配置 vision API Key

## 规则 2：模板中数据只是示例

xlsx 模板中的一切数值（0.98、98.1%、K9+370.58）都是示例/占位数据。
**不要提取模板中的数值当作真实数据**。真实数据必须从台账获取。

## 规则 3：默认走纯 I/O 路径

以下工具不调 LLM，是默认选择：

| 工具 | 用途 |
|------|------|
| `read_template_cells` | 读 xlsx 全部单元格位置+内容 |
| `read_ledger` (通过 ledger_reader) | 读台账数据 |
| `fill_template_cells` | 写单元格值并另存 |

以下工具内嵌 LLM，**仅在以下情况使用**：
- `analyze_experiment_template` — 模板>3个sheet 或 >50处合并区域，智能体难以直接理解时
- `generate_experiment_report` — 快速原型/草稿
- `organize_images` / `run_monitor` — 需要多模态分析照片

## 规则 4：生成报告的标准流程

```
步骤1: read_template_cells(模板路径)
       → 返回所有工作表的所有单元格(位置+内容)
       → 你自己看哪些是固定文本、哪些要填充、数据表在哪

步骤2: 读取台账（用户提供或已有）
       → 分析台账的列: 序号/桩号/部位/设计指标/厚度/施工日期/检测日期

步骤3: 你自己做数据匹配和计算:
       - 固定文本 → 直接用模板中的标签/标题值
       - 台账有的列 → 精确提取
       - 需要计算的 → 你自己算(平均值=sum/count, 合格率=合格数/总数)
       - 检测结论 → 你自己写(参考JTG F80/1-2017格式)
       - 台账确实没有的 → "见原始记录"

步骤4: fill_template_cells(模板路径, 输出路径, {"Sheet1": {"A3": "值", "B5": "值", ...}})
       → COM写入并另存为新文件
```

**不要**调用 `generate_experiment_report`，除非用户明确要求"快速生成个草稿看看"。

## 规则 5：项目初始化必须先做

首次对话先 `get_activity_summary`，没有项目的话按 `docs/modules/project_initializer.md` 流程初始化。

---

详细模块说明见 `docs/modules/` 目录。
