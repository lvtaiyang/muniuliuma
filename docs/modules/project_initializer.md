# 模块: 项目初始化 (project_initializer)

## 用途

创建项目工作区、初始化分部分项目录结构。

## 工具

| 工具 | 说明 |
|------|------|
| `project_init` | 创建项目 + 阶段目录 |
| `project_list` | 列出所有项目 |
| `project_load` | 加载项目详情 |
| `project_update` | 更新项目字段 |
| `project_discover` | 从目录反向查找项目 |
| `parse_project_document` | LLM 解析合同/分部分项文档 |
| `parse_project_text` | 文本描述解析项目结构 |
| `get_project_wbs` | 获取项目分部分项树 |

## 标准流程

### 首次使用

```
1. get_activity_summary → 看有没有已创建项目
   没有 → 继续

2. 询问用户:
   - 项目名称
   - 项目类型（住宅/工业/市政/公路/水利/其他）
   - 当前阶段:
       前期 → 01_前期决策 / 02_设计准备 / 03_招标合同
       施工 → 04_施工实施
       竣工 → 05_竣工验收 / 06_结算审计 / 07_后评估
   - 工作区目录路径

3. 如果用户有分部分项清单/合同文档:
   parse_project_document(文档路径) → 解析出 WBS 树
   project_init(wbs_json=..., contract_json=..., current_stage="04_施工实施")
   → 自动按分部分项创建资料子目录

4. 如果没有分部分项文档:
   project_init(name="项目名", workspace="/path/", current_stage="04_施工实施")
   → 施工阶段默认创建: 隐蔽验收影像资料/施工日志/实验检测报告/变更签证/材料台账

5. 告诉用户:
   - 目录结构已创建
   - 可以上传实验报告xlsx模板做分析
   - 可以上传台账数据
```

### 已创建项目

```
1. project_list → 确认项目名
2. project_load(name) → 获取工作区/阶段/WBS
3. 根据用户需求找到对应子目录操作
```
