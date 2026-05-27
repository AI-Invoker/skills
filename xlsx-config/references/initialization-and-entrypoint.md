---
name: "xlsx-config-initialization-and-entrypoint"
description: "xlsx-config 初始化与主控入口参考。用于建立 excel-scripts 目录结构并统一 main.py 入口职责。"
version: "1.0.0"
---

# 初始化与主控入口参考

<!-- @类型: 参考指南 -->
<!-- @目的: 为 xlsx-config 的初始化流程和主控入口职责提供展开版规则 -->
<!-- @场景: 第一次在工作区使用 Excel 自动化、初始化 excel-scripts、整理 main.py 职责 -->
<!-- @触发条件: 当主 SKILL.md 命中“初始化工作区 Excel 能力”路径时 -->

> **一句话**: 先补齐 `excel-scripts` 基础结构，再把所有 Excel 任务统一收敛到 `main.py`
> **版本**: v1.0.0
> **用途**: 规范初始化目录、入口职责和后续脚本放置位置
> **适用范围**: Excel 脚本初次落地、主控入口整理、目录职责统一

## @工作流: 初始化 excel-scripts 结构

<!-- @类型: 主工作流 -->
<!-- @目的: 在工作区中建立稳定的 Excel 自动化目录和主控入口 -->
<!-- @场景: 当前主目录下不存在 `excel-scripts`，或需要统一入口职责 -->
<!-- @前置条件: 已确认当前任务确实需要 Excel 自动化能力 -->
<!-- @后置验证: 已形成可扩展的 `excel-scripts` 目录结构和统一入口 -->
<!-- @触发条件: 当工作区第一次使用 xlsx-config，或现有入口结构混乱时 -->
<!-- @ID: wf-init-excel-scripts -->

### @步骤1: 先建立最小目录结构

<!-- @类型: 操作步骤 -->
<!-- @优先级: 必须 -->
<!-- @验证点: `excel-scripts` 最小目录结构完整存在 -->
<!-- @验证方式: 目录中存在 `main.py`、`excel_format_config.json` 和 `scripts/` -->
<!-- @ID: step-create-minimal-structure -->

- @动作: 在主目录下建立 `excel-scripts/`
- @动作: 在 `excel-scripts/` 下建立 `scripts/`
- @动作: 在 `excel-scripts/` 下建立 `main.py`
- @动作: 在 `excel-scripts/` 下建立 `excel_format_config.json`

### @步骤2: 不重复创建平行目录

<!-- @类型: 操作步骤 -->
<!-- @优先级: 必须 -->
<!-- @验证点: 若已有 `excel-scripts`，后续能力在现有结构上扩展，而不是再造平行目录 -->
<!-- @验证方式: 输出或执行中明确说明沿用现有目录结构 -->
<!-- @ID: step-reuse-existing-entrypoint -->

- @动作: 如果 `excel-scripts` 已存在，继续沿用，不新建并行入口目录
- @动作: 后续所有 Excel 处理脚本都放在 `excel-scripts/scripts/`

### @步骤3: 把 `main.py` 维持为主控入口

<!-- @类型: 操作步骤 -->
<!-- @优先级: 必须 -->
<!-- @验证点: `main.py` 负责命令解析和分发，而不是承担单一业务逻辑 -->
<!-- @验证方式: 新能力优先以子命令或参数形式接入 `main.py` -->
<!-- @ID: step-keep-main-py-thin -->

- @动作: `main.py` 只负责命令解析、参数接收和任务分发
- @动作: 具体业务逻辑收敛在 `scripts/` 下的模块函数中
- @动作: 不把 `main.py` 写成一次性固定执行某一个脚本的入口
- @动作: 新增能力时，优先为 `main.py` 增加子命令或参数

### @步骤4: 默认样式配置应真实可用

<!-- @类型: 操作步骤 -->
<!-- @优先级: 可选 -->
<!-- @验证点: `excel_format_config.json` 是后续样式处理的真实输入，而不是占位文件 -->
<!-- @验证方式: 配置文件结构完整，且后续样式处理会实际读取它 -->
<!-- @ID: step-prepare-style-config -->

- @动作: `excel_format_config.json` 至少包含 `sheet_background_color`、`tab_background_color`、`table_background_color` 和 `table_header_background_color`
- @动作: 不把样式文件写成只存在但不参与执行的占位物

## 版本历史

- **v1.0.0** (2026-05-27) - 初始版本，承接 `xlsx-config` 的初始化流程和 `main.py` 主控入口规则
