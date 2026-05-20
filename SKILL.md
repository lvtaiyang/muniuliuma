# 木牛流马 — 工程行业 AI 技能工具箱

工程行业通用智能体技能集，通过 MCP 协议对接任意智能体框架（Claude Code / Cursor / Dify 等）。
**30 个 MCP 工具，5 个独立模块。**

## 前置条件

- Python >= 3.10
- 纯文本 LLM API Key（OpenAI 兼容协议，推荐 DeepSeek v4 Flash）— 模板分析/文档生成/数据提取
- 多模态 LLM API Key（OpenAI 兼容协议，推荐通义千问 qwen3.6-plus）— 照片分析/影像分类
- Node.js >= 14（仅模块 2 微信监控需要，自动检测安装）
- Windows + Office/WPS（仅 Word/Excel 模板填充需要，非 Windows 环境走 openpyxl 回退）

## 快速接入（给其他智能体使用）

### 1. 下载安装

```bash
git clone https://github.com/lvtaiyang/muniuliuma.git
cd muniuliuma
# 跨平台全量（Linux/Mac/Windows 通用）
pip install -e ".[full]"
# Windows 用户推荐加装 COM 自动化（完美保留 Word/Excel 格式）
pip install -e ".[full,win32]"
```

按需安装：
```bash
pip install -e .                  # 核心
pip install -e ".[docx]"          # Word 模板填充（回退模式）
pip install -e ".[xlsx]"          # Excel 回退模式
pip install -e ".[win32]"         # Windows COM 自动化（推荐）
pip install -e ".[pdf]"           # PDF 文档分析
```

### 2. 配置 API Key

编辑 `config.yaml`，填入自己的模型 key：

```yaml
llm:
  text:       # 纯文本
    api_key: sk-your-deepseek-key
  vision:     # 多模态
    api_key: sk-your-qwen-key
```

MCP Server 内部自动按任务类型选择模型，一次配置所有调用方共用。主智能体不需要关心模型的事。

### 3. 在智能体框架中注册

**Claude Code** — 命令行添加：
```bash
claude mcp add muniuliuma -- python /path/to/muniuliuma/mcp_server.py
```
或手动编辑 `~/.claude/mcp.json`：
```json
{
  "mcpServers": {
    "muniuliuma": {
      "command": "python",
      "args": ["/path/to/muniuliuma/mcp_server.py"]
    }
  }
}
```

**Cursor** — 编辑 `.cursor/mcp.json`：
```json
{
  "mcpServers": {
    "muniuliuma": {
      "command": "python",
      "args": ["/path/to/muniuliuma/mcp_server.py"]
    }
  }
}
```

**Dify / 其他支持 MCP 的框架** — 同样的 stdio MCP 协议，配置格式一致。

注册后重启智能体，即可自动发现全部 30 个工具。

---

## 模块 1: 建设期项目初始化 (project_initializer)
**9 tools** — 对话式创建项目，全流程 7 阶段，模板上传/分析/复用

| 工具 | 用途 |
|------|------|
| `project_init` | 创建项目 + 7 阶段目录骨架 |
| `project_list/load/update` | 浏览/加载/更新项目 |
| `project_discover` | 从目录反向查找关联项目 |
| `template_analyze` | LLM 分析上传的 Word/Excel/PDF 模板 |
| `template_save/list/apply` | 保存/列表/应用模板生成文档 |

## 模块 2: 微信隐蔽验收影像归档 (wechat_monitor)
**8 tools** — 监控微信群图片，多模态分析，分类归档

| 工具 | 用途 |
|------|------|
| `check_wechat_setup` | 检测 wechat-cli 环境 |
| `install_wechat_cli` | 自动安装 wechat-cli |
| `list_wechat_groups` | 列出微信群聊 |
| `setup_monitoring` | 配置群聊/项目/模型/归档路径 |
| `run_monitor` | 执行监控管线：拉图→分析→分类→归档 |
| `get_summary/get_config/update_config` | 汇总/配置管理 |

## 模块 3: 现有影像资料整理 (image_organizer)
**2 tools** — 扫描已有图片目录，分析/分类/归档

| 工具 | 用途 |
|------|------|
| `scan_directory` | 预览目录图片列表 |
| `organize_images` | 扫描→多模态分析→分类→归档 |

## 模块 4: 施工日志生成 (construction_log)
**4 tools** — 从影像或台账生成施工日志，填入 Word 模板

| 工具 | 用途 |
|------|------|
| `generate_daily_log` | 从归档照片生成施工日志（多模态） |
| `generate_log_from_ledgers` | 从台账文件生成施工日志（纯文本） |
| `list_logs/read_log` | 浏览/读取日志 |

## 模块 5: 实验检测报告自动生成 (experiment_report)
**7 tools** — 从实验台账 + xlsx 模板生成标准检测报告

### 核心流程

```
上传 xlsx 模板
    ↓
analyze_experiment_template   ← LLM 逐单元格分析：含义/数据来源/逻辑关系
    ↓                              不确定处生成具体问题
返回分析结果 + 待确认问题
    ↓
用户回答问题
    ↓
confirm_template_analysis     ← 合并答案，确认后保存模板定义
    ↓
generate_experiment_report    ← 台账→LLM提取→填充xlsx→另存
    ↓
输出 .xlsx 实验报告 + Markdown 预览
```

### 工具

| 工具 | 用途 |
|------|------|
| `analyze_experiment_template` | 逐单元格深度分析 xlsx 模板，识别数据来源和逻辑关系 |
| `confirm_template_analysis` | 回答不确定问题，保存确认后的模板定义 |
| `list_experiment_templates` | 列出项目可用模板定义 |
| `generate_experiment_report` | 用模板+台账批量生成实验报告（xlsx） |
| `generate_single_experiment_report` | 为特定检测项目生成单份报告 |
| `list_experiment_reports` | 列出已生成的报告 |
| `read_experiment_report` | 读取报告内容（xlsx 以 Markdown 预览） |

### 关键设计

- **逐单元格理解**：LLM 分析每个有值单元格的精确行列号、合并区域、数据来源
- **逻辑关系识别**：自动发现单元格间的计算/判定逻辑（如 "实测值 vs 标准要求 → 合格/不合格"）
- **不确定必问**：任何拿不准的单元格含义或逻辑关系，生成具体问题让用户确认
- **win32com 填充**：Windows 上驱动本地 Excel/WPS 原生操作，完美保留合并单元格、边框、字体等格式；非 Windows 走 openpyxl 回退

---

## 首次使用流程

```
智能体收到用户请求后：

1. check_wechat_setup → 检测环境
   ├─ 就绪 → 下一步
   └─ 缺失 → install_wechat_cli 自动安装

2. project_init → 创建项目工作区

3. 用户上传模板 → template_analyze → 确认 → template_save
   或实验报告模板 → analyze_experiment_template → confirm_template_analysis

4. setup_monitoring → 绑定项目+群聊+模型

5. run_monitor / organize_images / generate_daily_log / generate_experiment_report → 日常使用
```

## 项目结构

```
muniuliuma/
├── SKILL.md
├── config.yaml
├── mcp_server.py              # MCP Server 入口（30 tools）
├── pyproject.toml
├── test_experiment_report.py  # 实验报告功能测试脚本
├── vendor/wechat-cli/         # wechat-cli 内置 npm 包
└── src/
    ├── config.py              # 配置管理（text/vision 双模型）
    ├── win32_helper.py        # COM 自动化 + openpyxl 回退（Word/Excel 操作）
    ├── wechat_monitor/        # 模块1: 微信影像归档
    ├── project_initializer/   # 模块2: 项目初始化与模板管理
    ├── image_organizer/       # 模块3: 现有资料整理
    ├── construction_log/      # 模块4: 施工日志生成
    └── experiment_report/     # 模块5: 实验报告生成
        ├── template_analyzer.py   # xlsx 模板逐单元格分析
        └── report_generator.py    # 台账数据提取 + xlsx 填充
```
