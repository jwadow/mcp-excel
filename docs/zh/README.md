<div align="center">

# 📊 Excel MCP Server

**专为 AI 智能体设计的快速高效电子表格分析工具，基于原子操作构建**

[🇬🇧 English](../../README.md) • [🇷🇺 Русский](../ru/README.md) • 🇨🇳 中文 • [🇪🇸 Español](../es/README.md) • [🇯🇵 日本語](../ja/README.md) • [🇧🇷 Português](../pt/README.md)

由 [@Jwadow](https://github.com/jwadow) 用 ❤️ 制作

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io)
[![Sponsor](https://img.shields.io/badge/💖_赞助-支持开发-ff69b4)](#-支持项目)

**通过原子操作让 AI 智能体分析 Excel 表格 — 无需将数据加载到上下文中**

*适用于 OpenCode、Claude Code、Codex app、Cursor、Cline、Roo Code、Kilo Code 及其他兼容 MCP 的 AI 智能体*

[为什么需要这个](#-为什么需要这个) • [你的智能体能做什么](#-你的智能体能做什么) • [安装与配置](#%EF%B8%8F-安装与配置) • [可用工具](#%EF%B8%8F-可用工具) • [💖 赞助](#-支持项目)

</div>

---

## 🤨 为什么需要这个

**问题：** 大多数 AI 的 Excel 工具会将原始表格数据直接倾倒到智能体的上下文中。这会塞满上下文窗口，拖慢一切，而且 AI 在大型数据集中仍然可能计算错误或混淆。

**本项目：** 把它想象成 Excel 的 SQL。你的 AI 智能体组合原子操作（`filter_and_count`、`aggregate`、`group_by`），获得精确结果 — 而不是成千上万行数据。

智能体在**不看数据**的情况下分析数据。结果以数字、公式和洞察的形式返回。

> *"这就像通过 SQL 操作数据库，而不是把所有东西拖进内存。"*
> — 某 AI 智能体在分析生产环境表格后的评价

### 🔌 什么是 MCP？

[Model Context Protocol](https://modelcontextprotocol.io) 是一个开放标准，允许 AI 智能体使用外部工具。

本项目就是这样一个工具。当你将此服务器连接到 AI 智能体（OpenCode、Claude Code、Codex app、Cursor、Cline、Roo Code、Kilo Code 等）时，你的智能体会获得大量用于处理 Excel 文件的新命令 — 过滤、计数、聚合、分析。

**关键优势：** 你的 AI 不会将数千行表格数据加载到内存中。相反，它提出具体问题并获得精确答案。更快、更准确、不会溢出上下文。

---

## 💬 AI 智能体怎么说

来自在生产环境中使用此 MCP 服务器的 AI 智能体的真实反馈：

> *"分析了 34,211 行数据而没有将数据加载到上下文中。每个操作只返回结果 — 计数、总和、平均值。上下文保持干净。无论文件大小如何，操作都在 25-45 毫秒内执行。"*

> *"这是 Excel 的 SQL。查询、过滤、聚合 — 无需将数据倾倒到上下文中。分析任务的可靠工具。"*

> *"过滤系统能很好地处理复杂逻辑。嵌套的 AND/OR 组、12 个运算符、无限条件。在不编写代码的情况下构建了多类别分类。"*

> *"批量操作很高效。一次 `filter_and_count_batch` 调用代替多个单独请求。文件加载一次，应用所有过滤器，结果一起返回。"*

*是的，智能体现在会写评论了。这些是 AI 智能体分析真实世界表格数据的实际反思。欢迎来到 2026 年。*

---

## 🚀 你的智能体能做什么

连接后，你的 AI 智能体将获得大量用于分析表格数据的专业工具。智能体只接收精确的查询和可靠的结果。

### 📊 数据探索
- **检查文件** - 结构、工作表、列、数据类型（自动检测混乱的标题）
- **列分析** - 统计信息、空值计数、热门值、一次调用的数据质量
- **查找数据** - 跨多个工作表搜索，在任何地方定位列

### 🔍 过滤与查询
- **12 个过滤运算符** - `==`、`!=`、`>`、`<`、`>=`、`<=`、`in`、`not_in`、`contains`、`startswith`、`endswith`、`regex`
- **复杂逻辑** - 嵌套的 AND/OR 组、NOT 运算符、无限条件
- **批量操作** - 一次请求将数据分类到多个类别（快 6 倍）
- **重叠分析** - 维恩图、交集计数、集合运算

### 📈 聚合与分析
- **8 个聚合函数** - sum、mean、median、min、max、std、var、count
- **分组** - 具有多个分组列的数据透视表
- **统计分析** - 相关性（Pearson/Spearman/Kendall）、异常值检测（IQR/Z-score）
- **时间序列** - 环比增长、移动平均、累计总和

### 🏆 高级操作
- **排名** - 前 N 名、后 N 名、百分位排名（支持分组）
- **计算列** - 列之间的算术表达式
- **数据验证** - 查找重复项、空值、数据质量检查
- **工作表比较** - 版本之间的差异、查找更改

### ⚡ 性能特性
- **原子操作** - 20-50 毫秒内得到结果，无论文件大小
- **智能缓存** - 文件加载一次，所有操作重用
- **示例行** - 预览过滤数据而无需完全检索
- **上下文保护** - 智能限制防止 AI 上下文溢出

### 📋 Excel 集成
- **公式生成** - 每个结果都包含用于动态更新的 Excel 公式
- **TSV 输出** - 将结果直接复制粘贴到 Excel
- **旧版支持** - 适用于旧的 .xls 文件（Excel 97-2003）
- **多工作表** - 分析一个文件中的多个工作表

**你的智能体现在可以处理的示例查询：**
- *"显示收入前 10 名的客户"*
- *"查找第四季度金额 > $1000 的所有订单"*
- *"计算每个产品类别的月环比增长"*
- *"哪些客户既是 VIP 又活跃？（重叠分析）"*
- *"在电子邮件列中查找重复项"*

## ⚙️ 安装与配置

### 前置要求

**Python 3.10 或更高版本** — [在此下载](https://www.python.org/downloads/)

### 步骤 1：克隆仓库

```bash
git clone https://github.com/jwadow/mcp-excel.git
cd mcp-excel
```

*没有 Git？点击此仓库页面顶部的 "Code" → "Download ZIP"，解压并在该文件夹中打开终端。*

### 步骤 2：选择安装方法

<details>
<summary><b>🎯 选项 A：Poetry（推荐）</b></summary>

Poetry 是现代 Python 依赖管理器（替代 pip+venv+requirements.txt）。
[安装它](https://python-poetry.org/docs/#installation)：`pip install poetry` 或 `pipx install poetry`

**安装依赖：**
```bash
poetry install
```

**配置你的 AI 智能体：**

将此添加到你的 MCP 设置（JSON 配置）：
```json
{
  "mcpServers": {
    "excel": {
      "command": "poetry",
      "args": ["run", "python", "-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

**重要：** 将 `C:/path/to/mcp-excel` 替换为克隆仓库的实际路径。

</details>

<details>
<summary><b>📦 选项 B：使用虚拟环境的 pip</b></summary>

**安装依赖：**
```bash
# Windows
python -m venv venv
venv\Scripts\activate
pip install -e .

# Linux/Mac
python -m venv venv
source venv/bin/activate
pip install -e .
```

**查找 venv 中的 Python 路径：**
```bash
# Windows
where python

# Linux/Mac
which python
```

**配置你的 AI 智能体：**

将此添加到你的 MCP 设置（JSON 配置）：
```json
{
  "mcpServers": {
    "excel": {
      "command": "C:/path/to/mcp-excel/venv/Scripts/python.exe",
      "args": ["-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

**重要：**
- 将 `C:/path/to/mcp-excel/venv/Scripts/python.exe` 替换为 `where python` 命令的实际路径
- 在 Linux/Mac 上使用 `which python` 的路径（例如 `/path/to/mcp-excel/venv/bin/python`）

</details>

<details>
<summary><b>🐍 选项 C：系统 Python（不推荐）</b></summary>

**全局安装依赖：**
```bash
pip install "mcp>=1.1.0" "pandas>=2.2.0" "pydantic>=2.10.0" "xlrd>=2.0.1" "openpyxl>=3.1.0" "psutil>=6.1.0" "python-dateutil>=2.9.0"
```

**配置你的 AI 智能体：**
```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

⚠️ **警告：** 这会污染你的全局 Python 环境。请改用 Poetry 或 venv。

</details>

### 步骤 3：验证安装

重启你的 AI 智能体并测试：
```
"分析位于 C:/Users/YourName/Documents/test.xlsx 的 Excel 文件"
```

如果有效 - 完成！如果没有，请检查：
- `cwd` 中的仓库路径是否正确
- `command` 中的 Python 路径是否正确（对于 pip 方法）
- 所有依赖项是否已安装

### 支持的 AI 智能体

适用于任何兼容 MCP 的 AI 智能体。

⚠️ **重要：** 这是一个 MCP 服务器。它在你的 AI 智能体需要时自动运行。不要在终端中手动运行它。

## 💡 使用方法

配置后，重启你的 AI 智能体并要求它分析 Excel 文件：

```
"分析位于 C:/Users/YourName/Documents/sales.xls 的 Excel 文件"
"显示 sales.xlsx 中收入前 10 名的客户"
"在 contacts.xlsx 的'电子邮件'列中查找重复项"
"从 revenue.xls 计算月环比增长"
```

## 🛠️ 可用工具

<details>
<summary><b>📋 完整工具参考（25 个工具）- 点击展开</b></summary>

### 📊 文件检查（5 个工具）

#### `inspect_file`
获取文件结构概览 - 工作表、尺寸、格式。
**用于：** 初始文件探索、工作表发现、格式验证
**返回：** 工作表列表、行/列计数、文件元数据

#### `get_sheet_info`
详细的工作表分析，带自动标题检测。
**用于：** 理解数据结构、列类型、示例预览
**返回：** 列名称/类型、行数、示例数据（3 行）、标题检测信息

#### `get_column_names`
快速列枚举，无需加载完整数据。
**用于：** 架构验证、过滤器构建、列可用性检查
**返回：** 列名称列表、列计数

#### `get_data_profile`
全面的列分析 - 类型、统计、空值、热门值。
**用于：** 初始数据探索、质量评估、分布分析
**返回：** 每列：类型、空值 %、唯一计数、统计（数值）、前 N 个值
**效率：** 替代 10+ 个单独调用（get_column_stats + get_value_counts + find_nulls）

#### `find_column`
跨多个工作表定位列。
**用于：** 多工作表导航、数据发现、跨工作表分析
**返回：** 带列位置、索引、行数的工作表列表（不区分大小写）

---

### 📥 数据检索（3 个工具）

#### `get_unique_values`
从列中提取唯一值。
**用于：** 数据探索、过滤器构建、唯一值发现、数据质量检查
**返回：** 唯一值列表、计数、截断标志（如果超过限制）
**默认限制：** 100 个值

#### `get_value_counts`
频率分析 - 前 N 个最常见的值。
**用于：** 分布分析、识别主导类别、数据不平衡检测
**返回：** 值 → 计数字典、总计数、TSV 输出
**默认：** 前 10 个值

#### `filter_and_get_rows`
检索带分页的过滤行。
**用于：** 数据提取、示例检查、详细分析、导出
**返回：** 过滤行（字典列表）、总计数、TSV 输出
**分页：** 支持 limit/offset

---

### 🔍 过滤与计数（3 个工具）

#### `filter_and_count`
使用 14 个运算符计数匹配条件的行。
**运算符：** `==`、`!=`、`>`、`<`、`>=`、`<=`、`in`、`not_in`、`contains`、`startswith`、`endswith`、`regex`、`is_null`、`is_not_null`
**逻辑：** 嵌套的 AND/OR 组、NOT 运算符、无限条件
**用于：** 分类、分段、数据验证、类别计数
**返回：** 计数 + Excel 公式（COUNTIFS）、可选示例行

#### `filter_and_count_batch`
一次调用将数据分类到多个类别（快 6 倍）。
**用于：** 多类别分类、市场细分、质量控制
**返回：** 每个类别的计数 + 公式、Excel 的 TSV 表
**效率：** 加载文件一次，应用所有过滤器，返回所有结果

#### `analyze_overlap`
维恩图分析 - 交集、并集、独占区域。
**用于：** 重叠分析、交叉销售机会、数据一致性检查
**返回：** 集合计数、成对交集（A ∩ B）、并集、维恩数据（2-3 个集合）
**示例：** VIP 且活跃的客户、产品类别重叠、已完成订单但无完成日期

---

### 📈 聚合与分析（2 个工具）

#### `aggregate`
使用可选过滤器执行聚合（8 个操作）。
**操作：** `sum`、`mean`、`median`、`min`、`max`、`std`、`var`、`count`
**用于：** 总计、平均值、最小/最大值、统计摘要、条件聚合、KPI 计算
**返回：** 聚合值 + Excel 公式（SUMIF、AVERAGEIF 等）
**特殊：** 自动将文本存储的数字转换为数值

#### `group_by`
具有多列分组的数据透视表。
**用于：** 类别分析、分层分组、按地区/产品的销售
**返回：** 带聚合值的分组数据、TSV 输出
**支持：** 多个分组列、所有 8 个聚合操作

---

### 📊 统计（3 个工具）

#### `get_column_stats`
统计摘要 - 计数、平均值、中位数、标准差、四分位数。
**用于：** 分布分析、数据分析、异常值检测准备
**返回：** 完整统计（min、max、mean、median、std、Q1、Q3）、空值计数、TSV 输出

#### `correlate`
2+ 列之间的相关矩阵。
**方法：** Pearson（线性）、Spearman（基于秩）、Kendall（基于秩）
**用于：** 关系分析、变量依赖性、特征选择
**返回：** 相关矩阵（-1 到 1）、TSV 输出

#### `detect_outliers`
使用 IQR 或 Z-score 方法进行异常检测。
**方法：** IQR（稳健）、Z-score（假设正态分布）
**用于：** 欺诈检测、传感器错误、数据质量、异常值识别
**返回：** 带索引的异常值行、计数、使用的方法/阈值

---

### ✅ 数据验证（2 个工具）

#### `find_duplicates`
根据指定列查找重复行。
**用于：** 数据质量、去重规划、完整性检查
**返回：** 所有重复行（包括第一次出现）、计数、索引
**注意：** 使用 `duplicated(keep=False)` 标记所有重复项

#### `find_nulls`
查找带详细统计的空值/空值。
**用于：** 完整性检查、缺失值分析、数据清理
**返回：** 每列：空值计数、百分比、索引（前 100 个）
**注意：** 占位符（"."、"-"）不是空值 - 使用 `==` 或 `in` 运算符

---

### 🔄 多工作表操作（2 个工具）

#### `search_across_sheets`
跨所有工作表搜索值。
**用于：** 跨工作表搜索、值跟踪、数据定位
**返回：** 带匹配计数的工作表列表、总匹配数
**支持：** 数值和字符串值

#### `compare_sheets`
使用键列比较两个工作表之间的差异。
**用于：** 版本比较、更改检测、对账、审计跟踪
**返回：** 带差异的行、状态（only_in_sheet1/sheet2/different_values）、并排比较

---

### 📅 时间序列（3 个工具）

#### `calculate_period_change`
环比增长分析。
**周期：** month、quarter、year
**用于：** 趋势分析、增长跟踪、季节性比较、同比分析
**返回：** 带值的周期、绝对/百分比变化、Excel 公式

#### `calculate_running_total`
带可选分组的累计总和。
**用于：** 累积分析、进度跟踪、余额计算、现金流
**返回：** 带累计总和的行、Excel 公式（SUM($B$2:B2)）
**支持：** 分组（每组重置累计总和）

#### `calculate_moving_average`
使用指定窗口大小进行平滑。
**用于：** 趋势检测、降噪、模式识别
**返回：** 带移动平均值的行、Excel 公式（AVERAGE(B1:B7)）
**示例：** 7 天移动平均、30 天股价平滑

---

### 🏆 高级操作（2 个工具）

#### `rank_rows`
按列值排名，带前 N 名过滤。
**方向：** desc（最高优先）、asc（最低优先）
**用于：** 排行榜、前/后分析、百分位排名
**返回：** 带排名编号的排名行、Excel 公式（RANK）
**支持：** 前 N 名过滤、组内排名

#### `calculate_expression`
列之间的算术表达式。
**操作：** `+`、`-`、`*`、`/`、括号
**用于：** 派生指标、财务计算、比率分析、KPI 计算
**返回：** 计算值、Excel 公式（例如 =A2*B2）
**示例：** 收入 = 价格 * 数量、利润率 = (收入 - 成本) / 收入

</details>

## 🗺️ 路线图

### 📁 文件格式支持

**当前支持：**
- ✅ **XLS** - Excel 97-2003（只读）
- ✅ **XLSX** - Excel 2007+（只读）

**计划中：**
- 🔜 **XLSM** - 支持宏的 Excel
- 🔜 **CSV** - 逗号分隔值
- 🔜 **TSV** - 制表符分隔值
- 🔜 **ODS** - OpenDocument 电子表格
- 🔜 **Parquet** - 列式存储格式

### 🚀 功能

- **写入操作** - 修改电子表格文件（创建计算列、更新值）
- **SSE 传输模式** - 用于远程访问的服务器发送事件
- **高级公式生成** - 具有嵌套函数的更复杂 Excel 公式
- **数据导出** - 将过滤/聚合结果导出到新文件

---

## 📜 许可证

本项目根据 **GNU Affero General Public License v3.0 (AGPL-3.0)** 授权。

这意味着：
- ✅ 你可以使用、修改和分发此软件
- ✅ 你可以将其用于商业目的
- ⚠️ **分发软件时必须公开源代码**
- ⚠️ **网络使用即分发** — 如果你在服务器上运行修改版本并让其他人与之交互，你必须提供源代码
- ⚠️ 修改必须在相同许可证下发布

完整许可证文本请参见 [LICENSE](../../LICENSE) 文件。

### 为什么选择 AGPL-3.0？

AGPL-3.0 确保对此软件的改进使整个社区受益。如果你修改此服务器并将其部署为服务，你必须与用户分享你的改进。

---

## 💖 支持项目

<div align="center">

<img src="https://raw.githubusercontent.com/Tarikul-Islam-Anik/Animated-Fluent-Emojis/master/Emojis/Smilies/Smiling%20Face%20with%20Hearts.png" alt="Love" width="80" />

**如果这个项目为你节省了时间或金钱，请考虑支持它！**

每一份贡献都有助于保持项目的活力和发展

<br>

### 🤑 捐赠

[**☕ 一次性捐赠**](https://app.lava.top/jwadow?tabId=donate) • [**💎 每月支持**](https://app.lava.top/jwadow?tabId=subscriptions)

<br>

### 🪙 或发送加密货币

| 货币 | 网络 | 地址 |
|:--------:|:-------:|:--------|
| **USDT** | TRC20 | `TSVtgRc9pkC1UgcbVeijBHjFmpkYHDRu26` |
| **BTC** | Bitcoin | `12GZqxqpcBsqJ4Vf1YreLqwoMGvzBPgJq6` |
| **ETH** | Ethereum | `0xc86eab3bba3bbaf4eb5b5fff8586f1460f1fd395` |
| **SOL** | Solana | `9amykF7KibZmdaw66a1oqYJyi75fRqgdsqnG66AK3jvh` |
| **TON** | TON | `UQBVh8T1H3GI7gd7b-_PPNnxHYYxptrcCVf3qQk5v41h3QTM` |

</div>

---

## 🤝 贡献

欢迎贡献！请确保：

1. 所有依赖项与 AGPL 兼容
2. 代码遵循现有风格
3. 新功能包含测试
4. 文档已更新

有关问题、错误或贡献，请在 GitHub 上打开 issue。

---

## 💬 需要帮助？

有问题？发现错误？有功能想法？我们在这里帮助！

**👉 [在 GitHub 上打开 Issue](https://github.com/jwadow/mcp-excel/issues/new)**

无论你是在安装时遇到困难、发现了问题，还是只是想提出改进建议 — GitHub Issues 就是你要去的地方。如果你是 GitHub 新手也不用担心，只需点击上面的链接并描述你的情况。我们会一起解决。

---

<div align="center">

**[⬆ 返回顶部](#-excel-mcp-server)**

</div>
