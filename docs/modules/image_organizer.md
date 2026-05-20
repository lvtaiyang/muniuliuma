# 模块: 影像资料整理 (image_organizer)

## 用途

扫描已有图片目录，多模态分析、分类、归档。**需要 vision API Key。**

## 工具

| 工具 | 说明 | LLM |
|------|------|-----|
| `scan_directory` | 扫描目录图片列表 | 无 |
| `organize_images` | 扫描→多模态分析→分类→归档 | 内嵌多模态 |

## 标准流程

```
1. scan_directory(source_dir="/path/to/images")
   → 预览图片列表，确认哪些要处理

2. organize_images(source_dir="/path/to/images", limit="50")
   → 逐张多模态分析(工程部位/验收类别/施工内容)
   → 自动分类归档到项目工作区
   → 更新汇总
```
