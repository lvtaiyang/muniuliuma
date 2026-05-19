# 木牛流马 — 工程行业 AI 技能工具箱

工程行业通用智能体技能集，通过 MCP 协议对接任意智能体框架（Claude Code / OpenClaw / Hermes / Dify 等）。
23 个 MCP 工具，4 个独立模块。

## 前置条件

- Python >= 3.10
- 多模态 LLM API Key（OpenAI 兼容协议，推荐通义千问 VL / 混元 Vision）
- Node.js >= 14（仅模块 2 微信监控需要，自动检测安装）

## 安装

```bash
git clone <repo-url> muniuliuma && cd muniuliuma

# 核心安装
pip install -e .

# 全部可选依赖（Word/Excel/PDF 处理）
pip install -e ".[full]"

# 或按需安装
pip install -e ".[docx]"    # Word 模板填充
pip install -e ".[xlsx]"    # Excel 台账读取
pip install -e ".[pdf]"     # PDF 文档分析
```

### MCP 配置

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

---

## 模块 1: 建设期项目初始化 (project_initializer)
**9 tools** — 对话式创建工程项目，全流程 7 阶段工作区，模板上传/分析/复用

| 工具 | 用途 |
|------|------|
| `project_init` | 创建项目 + 7 阶段目录 |
| `project_list/load/update` | 浏览/加载/更新项目 |
| `project_discover` | 从目录反向查找项目 |
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
| `run_monitor` | 执行监控管线 |
| `get_summary/get_config/update_config` | 汇总/配置管理 |

## 模块 3: 现有影像资料整理 (image_organizer)
**2 tools** — 扫描已有图片目录，分析/分类/归档

| 工具 | 用途 |
|------|------|
| `scan_directory` | 预览目录图片 |
| `organize_images` | 扫描→分析→分类→归档 |

## 模块 4: 施工日志生成 (construction_log)
**4 tools** — 从影像或台账生成施工日志，填入 Word 模板

| 工具 | 用途 |
|------|------|
| `generate_daily_log` | 从影像生成日志 |
| `generate_log_from_ledgers` | 从台账生成日志 |
| `list_logs/read_log` | 浏览/读取日志 |

---

## 首次使用流程

```
智能体收到用户请求后：

1. 调用 check_wechat_setup → 检测环境
   ├─ 就绪 → 下一步
   └─ 缺失 → install_wechat_cli 自动安装

2. project_init → 创建项目工作区

3. 用户上传模板 → template_analyze → 确认 → template_save

4. setup_monitoring → 绑定项目+群聊+模型

5. run_monitor / organize_images / generate_daily_log → 日常使用
```

## 项目结构

```
muniuliuma/
├── SKILL.md
├── config.yaml
├── mcp_server.py              # MCP Server（23 tools）
├── pyproject.toml
├── vendor/wechat-cli/         # wechat-cli 内置 npm 包
└── src/
    ├── config.py
    ├── wechat_monitor/        # 模块1: 微信影像归档
    ├── project_initializer/   # 模块2: 项目初始化
    ├── image_organizer/       # 模块3: 现有资料整理
    └── construction_log/      # 模块4: 施工日志
```
