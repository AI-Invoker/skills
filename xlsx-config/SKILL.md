---
name: "xlsx-config"
description: "聚焦 xlsx 配表结构与 win32com 表格自动化。Invoke when user asks about工作表初始化、字段设计、sheet 组织或用 Python 自动生成/格式化 Excel；默认唯一使用 win32com。"
version: "1.0.0"
---

# xlsx 表格配置

> **一句话**: 先明确工作簿结构和任务入口，再用 `win32com` 以打开态安全方式处理 Excel
> **版本**: v1.0.0
> **用途**: 指导工作簿结构设计、工作表初始化、Excel 自动化与 `excel-scripts` 技术路线统一
> **适用范围**: 工作簿 / sheet 结构、字段设计、样式配置、`win32com` 自动化、已打开工作簿修改

## 触发示例

- “帮我设计一个新的 xlsx 工作簿和字段结构”
- “初始化一张模板表，并按样式配置写进去”
- “修改我当前打开的工作簿里的指定单元格”
- “用 Python 自动生成并格式化这个 Excel 文件”

## @工作流: xlsx 配表与自动化路由

<!-- @类型: 主工作流 -->
<!-- @目的: 将 xlsx 相关需求路由到初始化、结构设计、打开态修改或自动化执行路径 -->
<!-- @场景: 用户要求设计工作簿、初始化表格、修改当前工作簿或通过脚本处理 Excel -->
<!-- @前置条件: 已明确任务与 Excel / xlsx 相关 -->
<!-- @后置验证: 已命中正确路径，并明确后续要读取的参考文档 -->
<!-- @触发条件: 当用户提出 xlsx 配表结构或 Excel 自动化需求时 -->
<!-- @ID: wf-xlsx-config-routing -->

## 快速路由

### @步骤40: 用决策矩阵命中当前任务

<!-- @类型: 决策步骤 -->
<!-- @优先级: 必须 -->
<!-- @验证点: 已判断当前是初始化、结构设计、打开态修改还是脚本入口类任务 -->
<!-- @验证方式: 能将用户问题稳定归入矩阵中的一类 -->
<!-- @ID: step-route-xlsx-task -->

| 场景 | 命中信号 | 跳转到 |
|---|---|---|
| 初始化工作区 Excel 能力 | `初始化 excel-scripts`、`第一次做 Excel 脚本`、`建主控入口` | 初始化路径 + [initialization-and-entrypoint.md](references/initialization-and-entrypoint.md) |
| 设计工作簿 / sheet / 字段结构 | `怎么拆 sheet`、`字段怎么设计`、`工作簿怎么组织` | 结构设计路径 + [workbook-design-patterns.md](references/workbook-design-patterns.md) |
| 修改已打开工作簿 | `改当前打开文件`、`写指定单元格`、`修用户正在看的工作簿` | 打开态路径 + [win32com-automation-rules.md](references/win32com-automation-rules.md) |
| 按配置自动化生成或格式化 | `用 Python 生成 xlsx`、`格式化工作表`、`按规范执行技能` | 自动化路径 + [style-config-rules.md](references/style-config-rules.md) |

### @步骤41: 强规则摘要

<!-- @类型: 规则步骤 -->
<!-- @优先级: 必须 -->
<!-- @验证点: 已应用 xlsx-config 的固定技术与执行约束 -->
<!-- @验证方式: 输出和执行中明确体现 `win32com`、打开态检查、样式配置驱动和主控入口约束 -->
<!-- @ID: step-apply-xlsx-hard-rules -->

- @动作: 当工作区采用本 skill 处理 Excel 时，默认唯一技术路线是 `win32com`
- @动作: 若目标工作簿可能已经打开，优先连接已打开工作簿，不直接走文件级覆盖
- @动作: `excel-scripts/main.py` 是默认主控入口，具体业务逻辑收敛在 `excel-scripts/scripts/`
- @动作: 样式配置必须来自 `excel_format_config.json` 的真实输入，而不是占位或写死样式
- @动作: 当任务同时涉及创造性内容与表格写入时，脚本只负责结构与写入工具，不默认承担创意内容数据源职责
- @动作: 若任务本质是职业、卡池、系统内容设计而不是纯表格结构，先交由对应领域 skill 定义内容，再由本 skill 处理工作台和自动化
- @警告: 不要在未明确说明的情况下静默切换到 `openpyxl` 或其他并行实现

## @工作流: xlsx 配表与自动化主流程

<!-- @类型: 主工作流 -->
<!-- @目的: 用统一顺序处理初始化、结构设计、打开态修改和自动化执行 -->
<!-- @场景: 需要新建、读取、写入、格式化或调用 Excel 自动化逻辑 -->
<!-- @前置条件: 已确认当前任务属于 xlsx 配表或 Excel 自动化范畴 -->
<!-- @后置验证: 输出结构设计或自动化方案可执行，且符合 `win32com`、主控入口和打开态规则 -->
<!-- @触发条件: 命中任一路由路径后 -->
<!-- @ID: wf-xlsx-config-main -->

### @步骤0: 先判断是否需要初始化

<!-- @类型: 操作步骤 -->
<!-- @优先级: 必须 -->
<!-- @验证点: 已确认当前工作区是否需要先补 `excel-scripts` 基础结构 -->
<!-- @验证方式: 输出中明确说明工作区已有结构还是需要初始化 -->
<!-- @ID: step-check-initialization -->

- @动作: 先检查当前主目录是否已有 `excel-scripts/`
- @动作: 若没有，再补 `main.py`、`excel_format_config.json` 和 `scripts/` 目录
- @动作: 若已存在，则继续沿用现有结构，不重复新建平行目录
- @提示: 初始化细节与默认目录结构见 [initialization-and-entrypoint.md](references/initialization-and-entrypoint.md)

### @步骤1: 再判断当前是读、写还是调用入口

<!-- @类型: 操作步骤 -->
<!-- @优先级: 必须 -->
<!-- @验证点: 已确定当前任务属于读取、写入、结构设计还是调用脚本入口 -->
<!-- @验证方式: 输出中明确当前要走的能力类别与入口 -->
<!-- @ID: step-classify-read-write-call -->

- @动作: 读取类任务优先说明工作簿、sheet、范围、表头和结构检查
- @动作: 写入类任务优先说明目标工作簿、目标 sheet、目标区域和样式应用
- @动作: 调用入口类任务优先收敛到 `main.py + 参数`
- @动作: 结构设计类任务优先回答工作簿拆分、字段设计、主键和引用关系

### @步骤2: 若涉及现有工作簿，优先做打开态检查

<!-- @类型: 操作步骤 -->
<!-- @优先级: 必须 -->
<!-- @验证点: 已在修改现有 `.xlsx` 前判断目标文件是否已被打开 -->
<!-- @验证方式: 输出或执行中明确体现“先枚举已打开 Workbooks，再决定是否 Open 文件” -->
<!-- @ID: step-check-open-workbook -->

- @动作: 优先尝试连接当前活动的 Excel 实例，而不是直接新开 Excel 进程
- @动作: 先枚举当前已打开的 `Workbooks`，按 `FullName` 精确匹配目标路径
- @动作: 如果目标工作簿已打开，后续读写直接作用于这本已打开的工作簿
- @动作: 完成写入后，再次校验目标单元格的 `Formula` 或 `Value`
- @提示: 打开态、安全读写和精确匹配规则见 [win32com-automation-rules.md](references/win32com-automation-rules.md)

### @步骤3: 若任务是结构设计，先回答工作簿为什么这样拆

<!-- @类型: 操作步骤 -->
<!-- @优先级: 必须 -->
<!-- @验证点: 已先解释工作簿、sheet、字段和表间关系为什么这样设计 -->
<!-- @验证方式: 结构设计输出不只给字段表，还说明拆表原则和维护边界 -->
<!-- @ID: step-design-workbook-structure -->

- @动作: 先说明工作簿和 sheet 的职责拆分，再列字段结构
- @动作: 先说明主键、枚举、索引、外键关系和命名规范，再进入具体列定义
- @动作: 若任务只停在字段表，但没有解释拆表原则和维护边界，视为不完整结构设计
- @提示: 结构设计模式见 [workbook-design-patterns.md](references/workbook-design-patterns.md)

### @步骤4: 若任务是自动化生成或格式化，按配置驱动执行

<!-- @类型: 操作步骤 -->
<!-- @优先级: 必须 -->
<!-- @验证点: 样式和工作表结构由配置驱动，且真正写入目标工作簿 -->
<!-- @验证方式: 输出或执行中体现先读取配置，再映射样式值，再应用到工作表和数据区 -->
<!-- @ID: step-apply-style-config -->

- @动作: 必须先读取 `excel_format_config.json`
- @动作: 必须把配置映射成 `win32com` 可执行的样式值，而不是只创建配置文件
- @动作: 样式处理顺序固定为：先给 `Cells` 应用 `sheet_background_color`，再覆盖数据区与表头
- @动作: 生成完成后，校验目标工作簿包含预期 sheet，且样式已写入目标单元格或 sheet 属性
- @提示: 样式配置、校验口径和写入顺序见 [style-config-rules.md](references/style-config-rules.md)

## 附录

### 快速参考

- 初始化目录结构与 `main.py` 主控入口：见 [initialization-and-entrypoint.md](references/initialization-and-entrypoint.md)
- 打开态检查、连接已打开工作簿和安全写入：见 [win32com-automation-rules.md](references/win32com-automation-rules.md)
- 样式配置驱动与校验口径：见 [style-config-rules.md](references/style-config-rules.md)
- 工作簿 / sheet / 字段 / 引用关系设计：见 [workbook-design-patterns.md](references/workbook-design-patterns.md)

## 版本历史

- **v1.0.0** (2026-05-27) - 按 `kz-skill-creator` 结构重整：补齐 version 与版本历史，将主文档收口为路由层，并把初始化、自动化、样式和结构设计说明下沉到 `references/`
