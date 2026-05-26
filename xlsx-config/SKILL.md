---
name: "xlsx-config"
description: "聚焦 xlsx 配表结构与 win32com 表格自动化。Invoke when user asks about工作表初始化、字段设计、sheet 组织或用 Python 自动生成/格式化 Excel；默认唯一使用 win32com。"
---

# xlsx 表格配置

## 作用定位

初始化工作表单元格字体字号、单元格背景

这个 skill 用于 xlsx 表格配置与表驱动工作流设计，聚焦：

- 工作簿 / sheet 结构
- 配置表拆分原则
- 字段、主键、索引、枚举、外键关系
- 初始化工作表模板
- xlsx 导入、读取、校验、回写
- `win32com` 与 Excel 自动化
- Python 驱动 Excel 的建表、格式化与导出流程
- 工作区内 Excel 技术路线统一

它适合把“策划要怎么建表、怎么分 sheet、怎么定义列、怎么让脚本处理这些表”整理成一套统一规范。



## 什么时候调用

当用户有以下需求时，调用这个 skill：

- 设计新的 xlsx 工作簿、sheet 和字段结构
- 初始化一张配置表、模板表、映射表或中间表
- 设计表间引用关系、主键、枚举、索引和命名规范
- 把策划规则整理为可落地的 xlsx 配表结构
- 说明如何用 `win32com` 调用本机 Excel 处理当前工作簿
- 说明如何用 Python 自动生成、读取、格式化并导出 `.xlsx`
- 检查配表是否可维护、可扩展、可校验

## 技能初始化

当用户还没有使用过这套 Excel 配表能力，或当前主目录下还不存在 `excel-scripts` 时，先执行初始化流程。

初始化目标目录结构如下：

```text
excel-scripts/
├─ main.py
├─ excel_format_config.json
└─ scripts/
```

其中：

- `main.py` 是主控脚本，负责解析参数并分发具体命令
- `scripts/` 下放具体业务脚本，不应把单一业务逻辑硬编码在 `main.py`

初始化要求：

- 在主目录新建 `excel-scripts`
- 在 `excel-scripts` 下新建脚本目录 `scripts`
- 在 `excel-scripts` 下创建主文件 `main.py`
- 在 `excel-scripts` 下创建表格样式文件 `excel_format_config.json`

默认样式文件内容如下：

```json
{
  "excel_format_style": {
    "sheet_background_color": "绿色(淡色，80%)",
    "tab_background_color": "绿色(淡色，60%)",
    "tab_text_color": "绿色(深色，25%)",
    "table_background_color": "金色(淡色，80%)",
    "table_header_background_color": "橙色(淡色，60%)",
    "default_horizontal_alignment": "left",
    "default_vertical_alignment": "top"
  }
}
```

初始化判断原则：

- 如果用户是第一次使用此 skill，优先先检查并补齐上述目录和文件
- 如果 `excel-scripts` 已存在，则在现有结构上继续扩展，不重复新建平行目录
- 后续所有表格处理相关的 Python 脚本都放在 `excel-scripts/scripts/`
- 后续整理出的业务表格结构都直接写入本 skill 文档

## 技能流程

使用这个 skill 时，默认按以下顺序思考：

1. 先判断是否需要执行“技能初始化”
2. 再判断用户当前要做的是“读取”“写入”还是“调用函数”
3. 如果任务目标是现有 Excel 文件，先检查用户是否已经在本机 Excel 中打开该工作簿
4. 最后再补充具体的工作簿结构、sheet 划分、字段定义与样式约束

## 主控入口约束

当工作区存在 `excel-scripts/main.py` 时，默认应把它作为主控入口，通过命令参数分发具体 Excel 任务。

- `main.py` 应负责命令解析、参数接收和任务分发
- 具体业务逻辑应收敛在 `scripts/` 下的模块函数中
- 不要把 `main.py` 写成“固定只执行某一个脚本”的一次性入口
- 新增能力时，优先为 `main.py` 增加子命令或参数，而不是再创建平行入口文件
- 当用户明确要求“通过 `/excel-scripts` 处理”时，应优先走 `main.py + 参数` 这条链路

## 技术路线总原则

当工作区已经采用本 skill 处理 Excel 时，默认唯一技术路线是 `win32com`。

- 新建 Excel 生成脚本时，默认只能选择 `win32com`
- 修改已有 Excel 自动化脚本时，默认应继续沿用并收敛到 `win32com`
- 不允许在同一工作区内为同类 Excel 任务继续新增 `openpyxl` 版本、`pandas + openpyxl` 版本或其他并行实现
- 如果工作区中已经存在旧的 `openpyxl` Excel 脚本，默认动作应是迁移到 `win32com`，而不是继续扩写旧链路
- 只有在用户明确指定“不依赖本机 Excel”“不能使用 COM”“只接受纯文件读写”时，才允许偏离 `win32com`
- 一旦发生偏离，必须先向用户明确说明原因、限制和影响，不能静默切换到其他技术路线

## 打开态检查

当任务涉及“修改已有 `.xlsx`”“给指定单元格写公式”“修正用户当前正在查看的工作簿”时，必须先判断目标文件是否已经被用户打开。

- 优先尝试连接当前活动的 Excel 实例，而不是直接新开一个独立 Excel 进程
- 必须先枚举当前已打开的 `Workbooks`，按 `FullName` 精确匹配目标文件路径
- 如果目标工作簿已经打开，后续读写必须直接作用于这本已打开的工作簿，再执行保存
- 如果目标工作簿未打开，才允许通过 `Workbooks.Open()` 打开文件进行处理
- 如果文件正被用户打开但脚本又试图用文件级方式直接覆盖保存，必须优先切换为“连接已打开工作簿”的处理方式
- 完成写入后，必须再次校验目标单元格的 `Formula` 或 `Value`，确保修改已经落到用户当前看到的那一份工作簿
- 如果任务目标是“当前打开文件中的指定单元格”，默认应优先使用已打开工作簿路径做精确匹配，避免误改同名副本
- 如果未显式指定工作表，默认先尝试当前活动工作表；若当前活动工作表没有源数据，再自动查找首个存在源数据的工作表

## 样式配置强约束

当任务涉及“新建 `.xlsx`”“回写 `.xlsx`”“初始化工作表模板”时，必须把 `excel_format_config.json` 作为真实输入而不是占位文件。

- 默认唯一实现方式是 `win32com`；除非用户明确要求，否则不得改用 `openpyxl`
- 必须先读取 `excel_format_config.json`
- 必须把配置项映射为脚本可执行的样式值，例如映射为 `win32com` 可用的 Excel 颜色值
- 必须把配置实际应用到工作表，而不是只创建配置文件
- 至少要覆盖 `sheet_background_color`、`table_background_color`、`table_header_background_color`、`tab_background_color`
- 样式处理顺序必须固定为：先对工作表 `Cells` 统一应用 `sheet_background_color`，再单独覆盖数据区与表头格式
- 当采用 `win32com` 驱动本机 Excel 时，允许直接对工作表 `Cells` 做统一格式处理，效果等价于手动全选后设置底色
- 如果 `tab_text_color` 受 Excel 对象模型限制无法直接生效，必须明确说明限制，不能忽略不说明
- 如果生成脚本里存在写死样式，必须优先替换为配置驱动；只有缺省值才允许代码兜底
- 生成完成后，必须校验输出工作簿是否已包含预期 sheet，且样式已写入目标单元格 / sheet 属性

如果本次任务要求“按规范执行技能”，则默认表示需要完整执行上述配置驱动流程，而不是仅完成目录初始化。



## 功能

这个 skill 的默认功能入口收敛为三类：

### 1. 读取

用于读取 Excel 工作簿、sheet 或指定范围数据。

- 读取当前工作簿
- 读取指定 `.xlsx` 文件
- 读取指定 sheet
- 将 sheet 读取为 DataFrame
- 检查当前表头、主键、字段和基础结构

### 2. 写入

用于把处理后的结果、模板结构或导出表写回 Excel。

- 新建或覆盖指定 sheet
- 把 DataFrame 写回工作表
- 初始化表头、数据区、校验区
- 先对工作表 `Cells` 全量写入淡绿色背景，再按配置覆盖数据区底色、表头底色与标签页颜色
- 如果目标文件已被用户打开，先连接已打开工作簿再写入，避免修改落不到用户当前看到的文件
- 如果发现已有同类 `openpyxl` 写入脚本，默认应迁移为 `win32com`，而不是继续复用旧实现
- 写入导出结果或中间结果

### 3. 调用函数

用于通过 `win32com` 或 Python 脚本入口触发表格相关逻辑。

- 从 `main.py` 暴露宏入口
- 调用 `scripts/` 下的业务脚本
- 调用通用表处理函数
- 触发读取、写入、测试或按钮逻辑
- 当函数作用于现有工作簿时，先检查该工作簿是否已在 Excel 中打开

## 禁止事项

- 不要在用户未明确要求的情况下新写 `openpyxl` 版本 Excel 生成脚本
- 不要为同一个 `.xlsx` 目标同时维护 `win32com` 和 `openpyxl` 两套生产实现
- 不要因为 `openpyxl` 写起来更快就绕过当前 skill 已确定的 `win32com` 路线
- 不要在未告知用户的情况下把已采用 `win32com` 的流程切回纯文件写入方案
