# 模块: 施工日志 (construction_log)

## 用途

从影像资料（多模态）或台账（纯文本）自动生成施工日志。

## 工具

| 工具 | 说明 | LLM |
|------|------|-----|
| `generate_daily_log` | 从归档照片生成日志 | 内嵌多模态 |
| `generate_log_from_ledgers` | 从台账生成日志 | 内嵌纯文本 |
| `list_logs` | 列出已有日志 | 无 |
| `read_log` | 读日志内容 | 无 |

## 标准流程

### 从照片生成（需要多模态 Key）

```
1. 确认 vision API key 已配置(get_config)
2. 确认 project_init 已完成
3. generate_daily_log(project_name="项目名", log_date="2024-07-27")
   → MCP 内嵌多模态 LLM 分析当天归档照片
   → 生成包含: 施工进度/材料进场/质量检验/安全检查的结构化日志
   → 保存到 04_施工实施/施工日志/
```

### 从台账生成（需要 text Key）

```
1. generate_log_from_ledgers(
     project_name="项目名",
     ledger_paths_json='["/path/材料台账.xlsx", "/path/检测台账.csv"]',
     log_date="2024-07-27"
   )
   → LLM 汇总台账数据
   → 生成日志并保存
```
