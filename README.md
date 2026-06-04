# meeting-minutes



> Hermes Agent skill · v

## 功能

1|---
2|name: meeting-minutes
3|description: 公文排版格式规范 + 行文风格指南。基于国企/党政机关公文标准，涵盖字体选型、字号层级、行间距、段落间距等排版规范，以及正式语体、句式结构、逻辑组织等风格要求。适用于会议记录、营销报告、数据分析报告等各类中文公文文档排版
4|version: "2.5"
5|changelog: |
6|  v2.5 — 三步分流工作流重写：从双叉（归档+wiki）升级为四叉戟模式（归档/wiki/运行/删除）
7|  v2.5 — 归档目录映射重写：匹配输出/ 新结构（01-脚本/ 02-舆情/ 03-数据/）
8|  v2.5 — 新增目录编号规则：连续不跳号，按用途命名，无内容即删不保留空壳
9|  v2.4 — 新增归档后清理规则（8.6）：输出稳定归档到源文件后删除输出目录对应文件
10|  v2.3 — 新增追加模式（8.8）、敏感性分析表格式（3.3）、报告追加流程文档化
11|  v2.2 — 新增分析逻辑参考小节，关联maoxuan-skill；快速检查清单拆分为排版格式/内容规范两组，新增连字符/em dash/AI-isms逐项检查
12|  v2.1 — 新增落款右对齐后处理；新增连字符禁用规则；脚本自动识别公司名+日期行设右对齐；新增中文公文AI-isms参考（references/chinese-gongwen-ai-isms.md）
13|  v2.0 — 新增文件命名规范（八、文件命名规范），标准化所有输出文件的日期/项目/版本命名体系
14|  v1.0 — 初始版本
15|trigger: user asks to format a .docx document, standardize document formatting, crea

## 安装

此技能通过 Hermes Agent 管理。在 Hermes 配置中启用即可：

```bash
hermes skill enable meeting-minutes
```

## 依赖

- Python 3.10+
- Hermes Agent

## 许可证

MIT
