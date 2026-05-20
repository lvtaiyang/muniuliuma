# 模块: 微信影像归档 (wechat_monitor)

## 用途

监控微信群聊中的施工照片，自动多模态分析、分类、归档。**需要 vision API Key + Node.js。**

## 工具

| 工具 | 说明 |
|------|------|
| `check_wechat_setup` | 检测 wechat-cli 环境 |
| `install_wechat_cli` | 自动安装 wechat-cli |
| `list_wechat_groups` | 列出微信群聊 |
| `setup_monitoring` | 配置群聊/项目/模型 |
| `run_monitor` | 执行一次监控管线 |
| `get_summary` | 获取归档汇总 |

## 标准流程

```
1. check_wechat_setup → 检测 Node.js 和 wechat-cli
   缺失 → install_wechat_cli

2. list_wechat_groups → 列出群聊，让用户选监控哪些

3. setup_monitoring(
     groups='["群名1", "群名2"]',
     project_name="项目名",
     api_key="sk-xxx",
     model="qwen3.6-plus"
   )

4. run_monitor → 拉取群聊图片→多模态分析→分类→归档→更新汇总
```
