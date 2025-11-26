# Excel File Batch Processor

一个命令行批处理工具，用于自动化处理Excel文件（.xls和.xlsx格式），将数据加载到SQLite数据库。

## 功能特性

- 📊 **批量处理**: 自动扫描并处理`new_payroll`和`old_payroll`文件夹中的所有Excel文件
- 🔄 **智能分割**: 自动识别工作表内的多个数据区域并分割为独立表格
- 📋 **数据验证**: 自动验证列结构并应用特殊逻辑处理
- ⚡ **高效处理**: 支持.xls和.xlsx两种格式的Excel文件
- 💾 **数据库集成**: 自动将处理后的数据加载到SQLite数据库

## 安装要求

### 系统要求
- Python 3.7+
- Windows/Linux/macOS

### Python依赖包

```bash
pip install openpyxl xlrd pandas numpy
```

## 项目结构

```
payroll_communication/
├── batch_process.py          # 主批处理程序
├── check_payroll_database.py # 数据库检查工具
├── check_load_log_table.py   # 加载日志检查工具
├── README.md                 # 说明文档
├── DAILY_LOG.md              # 开发日志
├── .gitignore                # Git忽略规则文件
├── new_payroll/              # 新工资表文件夹（被.gitignore排除）
│   ├── 202501.xls
│   ├── 202502.xls
│   └── ...
├── old_payroll/              # 旧工资表文件夹（被.gitignore排除）
│   ├── 202001.xls
│   ├── 202002.xls
│   └── ...
├── original_files/           # 原始文件文件夹（被.gitignore排除）
├── ../payroll_database.db    # SQLite数据库文件（位于父目录，被.gitignore排除）
├── one_time_pgms/            # 一次性处理程序
│   ├── excel_viewer.py       # Web-based Excel文件查看器
│   ├── templates/            # Web界面模板
│   │   └── index.html        # 主界面模板
│   ├── run_excel_viewer.bat  # Windows启动脚本
│   └── debug_file_matching.py # 文件匹配调试工具
└── excel_processor/          # 核心处理模块
    ├── config.py             # 全局配置和日志设置
    ├── sheet_gen.py          # Excel文件生成器
    ├── sheet_processor.py    # 工作表处理器
    ├── df_gen.py             # 数据框生成器
    └── special_logic.py      # 特殊逻辑处理
```

## 使用方法

### 批处理模式（推荐）

#### 1. 处理所有Excel文件

```bash
python batch_process.py
```

此命令将自动处理 `new_payroll` 和 `old_payroll` 文件夹中的所有Excel文件，并将数据加载到SQLite数据库。

#### 2. 处理单个Excel文件

```bash
python batch_process.py 201406.xls
```

此命令仅处理指定的Excel文件，适用于调试或特定文件处理。

#### 3. 数据库检查工具

```bash
# 检查工资数据表
python check_payroll_database.py

# 检查加载日志表
python check_load_log_table.py
```

### 处理流程

批处理程序自动执行以下步骤：
1. **扫描文件**: 自动扫描 `new_payroll` 和 `old_payroll` 文件夹中的所有Excel文件
2. **处理工作表**: 使用智能算法分割工作表为多个数据表格
3. **数据验证**: 自动验证列结构并应用特殊逻辑处理
4. **数据库加载**: 将处理后的数据加载到SQLite数据库
5. **日志记录**: 所有处理过程记录到 `log_batch.txt` 文件

### 输出文件

- **payroll_database.db**: SQLite数据库文件，包含处理后的工资数据
- **log_batch.txt**: 处理日志文件，记录详细的处理过程
- **special_logic_applied.log**: 特殊逻辑应用日志文件（汇总格式）

## 文件处理逻辑

### 文件扫描
- 自动扫描`new_payroll`和`old_payroll`文件夹
- 支持`.xls`和`.xlsx`格式文件
- 按文件名降序排序

### 文件处理
- 使用`openpyxl`处理`.xlsx`文件
- 使用`xlrd`处理`.xls`文件
- 自动识别工作表名称
- 显示所有数据

### 状态管理
- 使用内部状态跟踪处理进度
- 实时更新文件处理状态
- 防止重复处理

## 故障排除

### 常见问题

1. **依赖包安装失败**
   ```bash
   # 如果遇到numpy版本冲突
   pip uninstall numpy -y
   pip install numpy==1.24.3
   ```

2. **文件无法读取**
   - 确保Excel文件没有损坏
   - 检查文件权限
   - 确认文件格式支持（.xls或.xlsx）

3. **批处理程序启动失败**
   - 检查Python版本（需要3.7+）
   - 确认所有依赖包已正确安装
   - 检查文件路径和权限设置

### 错误处理
- 文件读取错误时会显示错误信息
- 处理失败的文件状态会重置为"待处理"
- 应用程序会继续处理下一个文件

## 开发说明

### 主要函数

#### `process_one_file(file_path, file_name)`
处理单个Excel文件的核心函数：
- 根据文件扩展名选择适当的库
- 返回工作簿对象和工作表名称列表
- 包含完整的错误处理

### 会话状态管理
- `processed_files`: 已处理的文件列表
- `current_file_index`: 当前处理的文件索引
- `file_status`: 文件状态字典
- `current_workbook`: 当前工作簿对象
- `selected_sheet`: 当前选中的工作表

## SQLite数据库说明

### 数据库结构

应用程序使用SQLite数据库进行数据持久化存储，包含以下两个主要表：

#### 1. `payroll_details` 表 - 工资数据存储
存储从Excel文件加载的所有工资记录，包含完整的源数据跟踪信息。

**表结构：**
```sql
CREATE TABLE payroll_details (
    文件名 CHAR(100),        -- 源Excel文件名
    sheet名 CHAR(100),       -- 源工作表名
    职员全名 CHAR(20),       -- 员工姓名
    日期 CHAR(20),          -- 工作日期
    客户名称 CHAR(60),       -- 客户名称
    型号 CHAR(100),         -- 产品型号
    工序全名 CHAR(100),     -- 完整工序名称
    工序 CHAR(100),         -- 工序简称
    计件数量 NUMERIC(10,2), -- 计件数量（浮点数，2位小数）
    系数 NUMERIC(10,2),     -- 系数（浮点数，2位小数）
    定额 NUMERIC(10,2),     -- 定额（浮点数，2位小数）
    金额 NUMERIC(10,2),     -- 金额（浮点数，2位小数）
    备注 CHAR(100),         -- 备注信息
    代码 CHAR(12),          -- 代码（外键，引用quota表）
    FOREIGN KEY (代码) REFERENCES quota(代码)
);
```

**数据跟踪功能：**
- 每条记录都包含源文件和工作表信息
- 支持完整的数据溯源和审计
- 便于识别数据来源和处理历史

#### 2. `load_log` 表 - 数据处理日志
记录数据处理过程中被丢弃的列信息，用于数据质量监控和问题排查。

**表结构：**
```sql
CREATE TABLE load_log (
    file_name CHAR(50),         -- 文件名
    sheet_name CHAR(50),        -- 工作表名
    table_index INT,            -- 表格索引（1=表一，2=表二，...）
    discarded_columns CHAR(200), -- 被丢弃的列名（逗号分隔）
    discarded_cols_num INT      -- 被丢弃列的数量
);
```

**日志功能：**
- 记录不符合预期列结构的列名
- 帮助识别数据格式问题
- 支持数据质量分析和改进

### 数据库操作

#### 数据加载
- 使用 `load_df_to_db()` 函数将DataFrame数据加载到数据库
- 自动过滤不符合预期列结构的列
- 自动添加源文件和工作表信息
- 支持批量数据插入

#### 数据查询
可以通过SQLite命令行工具查询数据库内容：

```bash
# 查看表结构
sqlite3 payroll_database.db ".schema"

# 查询工资记录
sqlite3 payroll_database.db "SELECT * FROM payroll_details LIMIT 10;"

# 查询加载日志
sqlite3 payroll_database.db "SELECT * FROM load_log;"

# 统计记录数量
sqlite3 payroll_database.db "SELECT COUNT(*) FROM payroll_details;"
```

### 数据库文件管理

- **文件位置**: `../payroll_database.db`（父目录）
- **自动创建**: 应用程序首次运行时自动创建数据库和表
- **数据持久化**: 所有处理的数据都会持久化保存
- **备份建议**: 定期备份数据库文件

## Batch Processing Module

### Overview

`batch_process.py` 是一个命令行批处理程序，用于自动化处理 `new_payroll` 和 `old_payroll` 文件夹中的所有Excel文件。该程序使用生成器函数实现高效的数据流处理，支持完整的端到端数据处理管道。

### 核心功能

#### 1. 模块化架构
- **sheet_gen**: Excel文件生成器，从Excel文件生成工作表内容
- **df_gen**: 数据框生成器，从工作表内容生成分割的数据框
- **数据库集成**: 自动将处理的数据加载到SQLite数据库

#### 2. 灵活的运行模式
- **默认模式**: 处理所有Excel文件
- **单文件模式**: 处理指定的单个Excel文件
- **测试模式**: 快速验证生成器函数逻辑

#### 3. 智能数据处理
- 自动跳过包含"汇总"的工作表
- 智能分割工作表为多个数据表格
- 完整的错误处理和日志记录

### 使用方法

#### 模式1: 默认批处理（处理所有文件）
```bash
python batch_process.py
```
- 处理 `new_payroll` 和 `old_payroll` 文件夹中的所有Excel文件
- 显示详细的处理进度和统计信息
- 包含数据库加载结果统计

#### 模式2: 单文件处理
```bash
python batch_process.py 201406.xls
```
- 仅处理指定的Excel文件
- 验证文件存在于 `new_payroll` 或 `old_payroll` 文件夹
- 提供详细的处理日志和数据库加载结果

#### 模式3: 错误处理（文件不存在）
```bash
python batch_process.py nonexistent_file.xlsx
```
- 显示清晰的错误信息
- 列出所有可用的Excel文件供参考
- 优雅的错误处理

### 生成器函数测试

#### 测试 sheet_gen 函数
```bash
cd excel_processor
python sheet_gen.py
```
- 测试Excel文件扫描和读取功能
- 验证工作表内容生成逻辑
- 显示文件统计和处理结果

#### 测试 df_gen 函数
```bash
cd excel_processor
python df_gen.py
```
- 测试工作表分割和数据框生成功能
- 支持实际数据和模拟数据测试
- 显示表格分割结果和数据结构

#### 集成测试
```bash
cd excel_processor
python test_integration.py
```
- 测试完整的处理管道：sheet_gen → df_gen
- 验证模块间集成和数据处理流程
- 显示处理统计和性能指标

### 技术架构

#### 数据流处理
```
Excel文件 → sheet_gen → 工作表内容 → df_gen → 分割数据框 → 数据库
```

#### 核心组件
1. **sheet_gen.py** - Excel文件处理生成器
   - `get_excel_files()`: 获取所有Excel文件
   - `sheet_gen()`: 生成工作表内容
   - `test_sheet_gen()`: 测试函数

2. **df_gen.py** - 数据框生成器
   - `df_gen()`: 从工作表生成分割数据框
   - `test_df_gen()`: 实际数据测试
   - `test_df_gen_with_mock_data()`: 模拟数据测试

3. **batch_process.py** - 主批处理程序
   - `batch_process_main()`: 完整批处理逻辑
   - `batch_process_test_limited()`: 有限测试模式
   - `process_single_file()`: 单文件处理模式

### 输出示例

#### 正常处理输出
```
Found 152 Excel files to process
Processing file: 201406.xls
Processing sheet: 201406.xls - 14年6月精加工
  Generated dataframe: Table 1, Shape=(85, 10)
  ✓ Successfully loaded to database: Successfully loaded 85 rows to database
Batch process completed. Processed 3 sheets and 46 dataframes.
Database loading results: 16 successful, 30 failed
```

#### 错误处理输出
```
Error: File 'nonexistent_file.xlsx' not found in new_payroll or old_payroll folders
Available files in new_payroll and old_payroll folders:
  - 201406.xls
  - 201407_1.xls
  - 201408.xls
  - ...
```

### 优势特点

1. **高效处理**: 使用生成器实现内存高效的数据流处理
2. **模块化设计**: 功能分离，便于维护和扩展
3. **灵活配置**: 支持多种运行模式满足不同需求
4. **完整测试**: 提供全面的测试覆盖和验证工具
5. **错误恢复**: 优雅的错误处理和用户友好的错误信息

## 程序调用关系和架构

### 模块调用关系图

```
batch_process.py (主批处理程序)
    ↓
excel_processor/ (核心处理模块)
    ├── config.py (全局配置和日志设置)
    ├── sheet_gen.py (Excel文件生成器)
    ├── sheet_processor.py (工作表处理器)
    ├── df_gen.py (数据框生成器)
    └── special_logic.py (特殊逻辑处理)
```

### 详细调用栈

#### 1. batch_process.py 调用流程
```python
batch_process_main()
    ↓
get_excel_files()  # 从 sheet_gen.py 获取文件列表
    ↓
sheet_gen()  # 从 sheet_gen.py 生成工作表内容
    ↓
df_gen()  # 从 df_gen.py 生成分割数据框
    ↓
load_df_to_db()  # 从 sheet_processor.py 加载数据到数据库
```

#### 2. 模块间数据流
```
Excel文件 → sheet_gen() → SheetContents → df_gen() → SplitDataFrame → load_df_to_db() → SQLite数据库
```

#### 3. 全局日志配置
- **位置**: `excel_processor/config.py` 中的 `setup_global_logging()` 函数
- **功能**: 统一配置所有模块的日志输出
- **输出**: 所有日志写入 `log_batch.txt` 文件
- **格式**: `%(asctime)s - %(name)s - %(levelname)s - %(message)s`
- **级别**: INFO 级别，包含控制台和文件输出

### 模块职责说明

#### batch_process.py
- 主批处理程序入口
- 支持命令行参数处理
- 协调各模块间的调用流程
- 提供单文件和批量处理模式

#### excel_processor/config.py
- 全局配置管理
- 共享常量定义 (`expected_columns`, `COMMON_COL_COUNT`)
- 全局日志配置 (`setup_global_logging()`)

#### excel_processor/sheet_gen.py
- Excel文件扫描和读取
- 工作表内容生成器
- 文件路径管理
- 工作表过滤逻辑

#### excel_processor/sheet_processor.py
- 工作表数据处理
- 数据框分割和验证
- 数据库加载功能
- 列验证和智能表头检测

#### excel_processor/df_gen.py
- 数据框生成器
- 表格分割逻辑
- 数据框验证和过滤

#### excel_processor/special_logic.py
- 特殊数据处理规则
- 列名标准化
- 数据格式转换

## Excel文件查看器

### 功能特性

- 🌐 **Web界面**: 基于Flask的Web应用程序，提供直观的文件比较界面
- 📊 **并排比较**: 左侧显示old_folder文件，右侧显示new_folder文件，便于对比分析
- 🔄 **文件导航**: 提供Previous/Next按钮，支持浏览所有文件（202001-202012）
- 📋 **工作表选择**: 三个工作表按钮：绕嵌排、精加工、喷漆装配
- 🔍 **智能匹配**: 自动处理工作表名称变体（'金加工'→'精加工'，'装配喷漆'→'喷漆装配'）
- 📱 **响应式设计**: 使用Bootstrap 5实现现代化自适应界面
- ⚡ **动态加载**: AJAX技术实现无刷新页面切换

### 使用方法

#### 1. 启动Web服务器
```bash
# 方法1: 直接运行Python脚本
python one_time_pgms/excel_viewer.py

# 方法2: 使用Windows批处理脚本
one_time_pgms/run_excel_viewer.bat
```

#### 2. 访问Web界面
- 打开浏览器访问：`http://localhost:5000`
- 系统自动显示第一个文件（202001）的"绕嵌排"工作表

#### 3. 功能操作
- **工作表选择**: 点击顶部的三个按钮切换不同工作表
- **文件导航**: 使用Previous/Next按钮浏览所有文件
- **数据查看**: 左右两侧分别显示old_folder和new_folder的相同文件
- **滚动浏览**: 使用滚动条查看完整的表格数据

### 技术架构

#### 后端技术
- **Flask**: Python Web框架，提供RESTful API
- **pandas**: Excel文件读取和数据处理
- **openpyxl/xlrd**: Excel文件格式支持

#### 前端技术
- **Bootstrap 5**: 现代化响应式UI框架
- **JavaScript**: 动态数据加载和交互功能
- **AJAX**: 异步数据请求，无需页面刷新

#### 文件处理
- **路径管理**: 自动定位new_payroll和old_payroll文件夹
- **文件匹配**: 智能识别相同文件名的文件
- **工作表读取**: 支持.xls和.xlsx格式文件
- **错误处理**: 优雅的错误处理和用户提示

### 安全特性

- **文件隔离**: 系统只处理new_payroll和old_payroll根目录的文件
- **placeholder保护**: old_payroll/placeholder/文件夹中的文件不会被处理
- **路径验证**: 严格的路径验证，防止目录遍历攻击

### 重要发现

**文件处理决策**: 通过使用excel_viewer.py进行文件比较，我们发现new_folder中的所有2020xx.xls文件的金额都比old_folder中的对应文件大。因此，决定将old_folder中的这些文件移动到old_folder\placeholder文件夹中，以确保这些文件不会被批处理系统处理。

**决策依据**:
- 通过excel_viewer.py的并排比较功能，确认new_folder中的文件金额更大
- 为避免数据重复和确保数据准确性，将old_folder中的对应文件移至placeholder
- placeholder文件夹中的文件被系统自动排除，不会被批处理程序处理

## 扩展功能

程序可以轻松扩展以下功能：
- 数据导出功能
- 自定义数据处理逻辑
- 批量数据转换
- 数据验证和清洗
- 报表生成

## TODO - 未来改进

### ✅ COMPLETED - Excel Sheet Processing Enhancement
- **功能**: 改进 `process_excel_sheet` 函数，使其能够从一个工作表生成多个DataFrame
- **实现状态**: 已完成
- **实现内容**:
  - ✅ 检测工作表内的空行作为数据区域分隔符
  - ✅ 自动识别工作表内的多个数据区域
  - ✅ 返回两个结果：`df_summary`（完整工作表数据）和 `dfs`（分割后的DataFrame列表）
  - ✅ 在日志中记录表格检测结果
  - ✅ 支持批量处理多个数据表格
- **位置**: `excel_processor/sheet_processor.py` 中的 `process_excel_sheet` 函数

### ✅ COMPLETED - Batch Processing Enhancement
- **功能**: 改进批处理逻辑和数据处理流程
- **实现状态**: 已完成
- **实现内容**:
  - ✅ 优化文件扫描和排序逻辑
  - ✅ 改进工作表分割算法
  - ✅ 增强数据验证和错误处理
  - ✅ 完善日志记录和进度跟踪

### ✅ COMPLETED - Data Loading Enhancement
- **功能**: 改进数据加载逻辑
- **实现状态**: 已完成
- **实现内容**:
  - ✅ 移除50行限制，加载所有数据
  - ✅ 支持完整数据加载到数据库
  - ✅ 优化数据处理流程

### ✅ COMPLETED - Database Integration Enhancement
- **功能**: 数据库集成和数据处理优化
- **实现状态**: 已完成
- **实现内容**:
  - ✅ 实现SQLite数据库集成，创建`payroll_details`表
  - ✅ 添加`文件名`和`sheet名`列用于数据源跟踪
  - ✅ 实现`load_df_to_db`函数，支持数据加载到数据库
  - ✅ 创建`load_log`表记录被丢弃的列信息
  - ✅ 实现自动数据处理流程
  - ✅ 添加数据处理状态跟踪和日志记录
- **位置**: `excel_processor/sheet_processor.py`

### NEW - Data Formatting Enhancement
- **功能**: 格式化DataFrame中的金额、定额、日期等字段
- **当前状态**: 待实现
- **未来需求**: 需要实现逻辑来识别和格式化特定类型的字段
- **实现思路**:
  - 识别包含"金额"、"定额"、"日期"等关键词的列
  - 应用适当的格式化（货币格式、数字格式、日期格式）
  - 在显示时保持数据可读性
- **位置**: `excel_processor/sheet_processor.py` 中的数据处理部分

### NEW - Advanced DataFrame Processing
- **功能**: 添加更多逻辑来处理DataFrame中的数据
- **当前状态**: 待实现
- **未来需求**: 需要实现更复杂的数据处理逻辑
- **实现思路**:
  - 数据清洗和验证
  - 自动识别数据模式
  - 智能数据转换
  - 数据质量检查
- **位置**: `excel_processor/sheet_processor.py` 中的数据处理部分

### ✅ COMPLETED - Database Schema Enhancement
- **功能**: 增强数据库架构和数据处理逻辑
- **实现状态**: 已完成
- **实现内容**:
  - ✅ 在`load_log`表中添加`discarded_cols_num`列，记录被丢弃列的数量
  - ✅ 改进`load_df_to_db`函数中的列清理逻辑，移除空列和带`_1`、`_2`等后缀的列
  - ✅ 仅在存在被丢弃列时记录到`load_log`表，减少不必要的日志记录
  - ✅ 在终端输出警告消息当列被丢弃时
  - ✅ 清理数据库表`payroll_details`和`load_log`中的所有记录
- **位置**: `excel_processor/sheet_processor.py` 中的 `load_df_to_db` 函数

### ✅ COMPLETED - Batch Processing Logic Enhancement
- **功能**: 改进批处理逻辑和自动化流程
- **实现状态**: 已完成
- **实现内容**:
  - ✅ 修复文件处理逻辑，解决文件跳过的问题
  - ✅ 实现自动文件处理流程，无需手动干预
  - ✅ 优化处理状态管理，防止运行时错误
  - ✅ 实现自动工作表切换，跳过包含"汇总"的工作表
  - ✅ 当所有工作表处理完成后自动处理下一个文件
- **位置**: `batch_process.py` 中的批处理逻辑

### ✅ COMPLETED - Sheet Content Filtering and Log Output Enhancement
- **功能**: 工作表内容过滤和日志输出增强
- **实现状态**: 已完成
- **实现内容**:
  - ✅ 添加工作表内容检查逻辑，过滤掉空工作表
  - ✅ 在日志输出空工作表丢弃信息：`工作表 '{sheet_name}' 被丢弃 - 空工作表`
  - ✅ 添加文件分隔符：在处理新文件前打印 `---------------------`
  - ✅ 添加文件处理状态消息：`开始处理文件: {file_name}` 和 `文件处理完成: {file_name}`
  - ✅ 处理工作表错误时也显示相应消息：`工作表 '{sheet_name}' 被丢弃 - 处理错误: {error_message}`
  - ✅ 当文件没有包含数据的工作表时显示：`文件 '{file_name}' 中没有包含数据的工作表`
- **位置**: `batch_process.py` 中的处理逻辑

### ✅ COMPLETED - Automatic Data Processing
- **功能**: 完全自动化的数据处理
- **实现状态**: 已完成
- **实现内容**:
  - ✅ 实现自动文件处理流程
  - ✅ 自动处理所有表格、工作表、文件，无需手动干预
  - ✅ 实现自动工作表切换和文件切换
  - ✅ 添加数据处理状态跟踪和日志记录
- **位置**: `batch_process.py` 中的自动处理逻辑

### ✅ COMPLETED - Special Logic Processing for Data Loading
- **功能**: 特殊逻辑预处理功能
- **实现状态**: 已完成
- **实现内容**:
  - ✅ 创建 `special_logic.py` 模块，包含 `special_logic_preprocess_df` 函数
  - ✅ 实现7种特殊逻辑规则：
    - **逻辑1**: "喷漆装配"工作表中，将"前装"/"中装"/"后装"/"刘雷"替换为"职员全名"
    - **逻辑2**: "喷漆装配"工作表中，将"后装曾大军"替换为"职员全名"并设置所有值为"曾大军"
    - **逻辑3**: 任意工作表中，将"姓名"替换为"职员全名"
    - **逻辑4**: "绕嵌排"工作表中，将"型号"列后的"嵌线"/"排线"替换为"工序全名"
    - **逻辑5**: "绕嵌排"工作表中，将"型号"列后的"工序名称"替换为"工序全名"
    - **逻辑6**: 如果存在列名为'数量'且同时存在列名为'职工全名'，则将'数量'改为'计件数量'
    - **逻辑7**: 如果存在列名为'加工型号'，则将'加工型号'改为'型号'
  - ✅ 在 `load_df_to_db` 函数开始时移除列名和工作表名中的空格
  - ✅ 在移除空列和后缀列后调用特殊逻辑预处理函数
  - ✅ 实现日志记录功能，记录所有应用的特殊逻辑到 `special_logic_applied.log` 文件
  - ✅ 日志包含时间戳、文件名、工作表名、表索引和逻辑描述
- **位置**: `excel_processor/special_logic.py` 和 `excel_processor/sheet_processor.py`

### ✅ COMPLETED - Code Refactoring and Function Separation
- **功能**: 代码重构和函数分离
- **实现状态**: 已完成
- **实现内容**:
  - ✅ 创建 `get_all_data_from_sheet` 函数，专门用于从Excel工作表提取原始数据
  - ✅ 创建 `split_raw_sheet_contents` 函数，专门用于将原始数据分割成多个DataFrame
  - ✅ 重构 `process_excel_sheet` 函数，使用新的分离函数
  - ✅ 更新 `sheet_gen.py` 使用 `get_all_data_from_sheet` 函数
  - ✅ 更新 `df_gen.py` 使用 `split_raw_sheet_contents` 函数
  - ✅ 创建 `config.py` 配置文件，包含共享的 `expected_columns` 和 `COMMON_COL_COUNT`
  - ✅ 实现智能列验证逻辑：
    - 检查DataFrame中期望列的数量
    - 如果期望列数量不足，自动搜索更好的表头行
    - 支持自动表头检测和替换
    - 丢弃不符合要求的数据框并输出详细日志
  - ✅ 增强日志功能，输出被丢弃数据框的完整内容
- **位置**: `excel_processor/sheet_processor.py`, `excel_processor/sheet_gen.py`, `excel_processor/df_gen.py`, `excel_processor/config.py`

## 特殊逻辑处理规则

系统实现了18种特殊逻辑规则，用于自动处理Excel文件中的列名不一致问题：

### 当前特殊逻辑规则列表：

1. **逻辑1**: "喷漆装配"工作表中，将"前装"/"中装"/"后装"/"刘雷"/"装配"替换为"职员全名"
2. **逻辑2**: "喷漆装配"工作表中，将"后装曾大军"替换为"职员全名"并设置所有值为"曾大军"
3. **逻辑3**: 任意工作表中，将"姓名"替换为"职员全名"
4. **逻辑4**: "绕嵌排"工作表中，将"型号"列后的"嵌线"/"排线"替换为"工序全名"
5. **逻辑5**: "绕嵌排"工作表中，将"型号"列后的"工序名称"替换为"工序全名"
6. **逻辑6**: 如果存在列名为'数量'且同时存在列名为'职工全名'，则将'数量'改为'计件数量'
7. **逻辑7**: 如果存在列名为'加工型号'，则将'加工型号'改为'型号'
8. **逻辑8**: 当'计件数量'包含在列名中时，将该列替换为'计件数量'
9. **逻辑9**: 将'单位工资'替换为'定额'
10. **逻辑10**: 将'合计金额'替换为'金额'
11. **逻辑11**: 将'规格'替换为'型号'
12. **逻辑12**: 当'定额'列后存在'合计'列时，将'合计'替换为'金额'
13. **逻辑13**: 将'任务名称'替换为'客户名称'
14. **逻辑14**: 当'职员全名'是'前装'或'前装人员'时，将记录拆分为2行（黄志梅和陈会清）
15. **逻辑15**: 当'职员全名'是'中装'或'中装人员'时，将值改为'李兆军'
16. **逻辑16**: 当'职员全名'是'后装'或'后装人员'时，将值改为'汤雅林'
17. **逻辑17**: 当'职员全名'列的值为空、空格（或中文空格）、或None时，从数据框中丢弃该行
18. **逻辑18**: 当'职员全名'列包含特定中文短语时，丢弃对应的行（短语：['下料', '铣底脚：', '铣：', '校平衡', '车转子', '压：', '磨：']）

**特殊逻辑总数**: 系统现在包含 **18个特殊逻辑规则**

### 日志记录
所有特殊逻辑应用都会记录到 `special_logic_applied.log` 文件，包含时间戳、文件名、工作表名、表索引和逻辑描述。

## 近期更新

### 2025年11月3日更新

#### 1. 工作表名称映射功能增强
- **功能**: 实现工作表名称标准化映射，统一处理历史文件中的不同命名
- **映射规则**:
  - `14年6月精加工` → `精加工`
  - `14年6月装配 喷漆` → `装配喷漆`
  - `14年6月绕嵌排` → `绕嵌排`
  - `装配 喷漆` → `装配喷漆`
  - `喷漆装配` → `装配喷漆`
  - `金加工` → `精加工`
- **实现**: 在 `excel_processor/special_logic.py` 中添加工作表名称映射逻辑
- **结果**: 历史文件中的工作表名称自动标准化，便于数据统一管理

#### 2. 特殊逻辑架构重构
- **功能**: 重构特殊逻辑处理函数返回值和数据流
- **实现**: 
  - 修改 `special_logic_preprocess_df` 函数返回 `(processed_df, updated_sheet_name, updated_file_name)`
  - 在特殊逻辑函数中添加文件名和工作表名列到DataFrame
  - 更新 `load_df_to_db` 函数使用返回的更新值
  - 移除冗余的文件名和工作表名设置代码
- **结果**: 映射后的工作表名称正确存储到数据库，解决映射不生效的问题

#### 3. 前装/中装/后装人员处理增强
- **功能**: 扩展特殊逻辑处理范围，支持"前装人员"、"中装人员"、"后装人员"
- **实现**:
  - 将 `row['职员全名'] == '前装'` 改为 `row['职员全名'].startswith('前装')`
  - 将 `df['职员全名'] == '中装'` 改为 `df['职员全名'].str.startswith('中装', na=False)`
  - 将 `df['职员全名'] == '后装'` 改为 `df['职员全名'].str.startswith('后装', na=False)`
- **结果**: 现在同时处理"前装"和"前装人员"、"中装"和"中装人员"、"后装"和"后装人员"

#### 4. 重复文件识别
- **功能**: 识别 new_payroll 和 old_payroll 文件夹中的重复文件
- **发现**: 12个重复文件（202001.xls 到 202012.xlsx）
- **结果**: moved to old_folder/placehoder folder (including 202010_2.xls)

### 2025年10月30日更新

#### 1. 小数转换错误修复
- **问题**: `decimal.ConversionSyntax` 错误，处理202507.xls文件时出现无效计件数量值 '1*2'
- **解决方案**: 增强特殊逻辑预处理，使用 `pd.to_numeric()` 和 `errors='coerce'` 进行安全转换
- **实现**: 在 `excel_processor/special_logic.py` 中添加错误处理逻辑
- **结果**: 无效数值自动替换为默认值0，处理继续正常进行

#### 2. 日期列类型转换优化
- **问题**: 日期列值存储为带小数点的格式（如"2.0", "3.0"）
- **解决方案**: 增强数据类型转换逻辑，移除整数日期值的'.0'后缀
- **实现**: 在 `excel_processor/sheet_processor.py` 中优化数据格式化
- **结果**: 日期列正确显示为整数格式（如"2", "3"）

#### 3. 增强日志记录功能
- **问题**: 无效计件数量日志缺少上下文信息
- **解决方案**: 在日志中增加定额和金额值信息
- **实现**: 更新 `excel_processor/special_logic.py` 中的日志格式
- **结果**: 日志格式从 `无效的计件数量值 '1*2' 在行 65，使用默认值0` 改进为 `无效的计件数量值 '1*2' 在行 65，使用默认值0, (定额 的值是：11, 金额的值是：2222)`

#### 4. 数据库路径配置修复
- **问题**: DATABASE_PATH 配置指向错误位置
- **解决方案**: 更新为绝对路径 `/home/richard/shared/jianglei/payroll/payroll_database.db`
- **实现**: 修改 `excel_processor/config.py` 中的 DATABASE_PATH 配置
- **结果**: 数据库连接正确指向父目录的数据库文件

### 2025年10月26日更新

#### 1. 特殊逻辑日志优化
- **问题**: 特殊逻辑日志文件过大，包含过多重复条目
- **解决方案**: 从记录每个单独操作改为记录汇总信息
- **实现**: 添加操作计数器，跟踪"前装拆分"、"中装替换"、"后装替换"操作次数
- **结果**: 日志文件从数百行减少到6行，显示汇总信息如："将'前装'记录拆分为2行: 黄志梅 和 陈会清 共119次"

#### 2. 数据类型优化
- **问题**: "计件数量"字段应为浮点数，但数据库存储为整数
- **解决方案**: 将数据库中的"计件数量"字段从INT改为FLOAT
- **实现**: 更新 `excel_processor/sheet_processor.py` 中的 `expected_columns` 定义
- **结果**: "计件数量"正确存储为浮点数（如20.0, 10.0, 1.0等）

#### 3. Excel公式处理说明
- **问题**: `.xls` 文件中的公式显示值与读取值不一致
- **发现**: `xlrd` 库在处理 `.xls` 文件时会自动评估所有公式，无法保留公式显示值
- **解决方案**: 在 `get_all_data_from_sheet` 函数中添加警告信息
- **结果**: 处理 `.xls` 文件时显示："Processing .xls file '{filename}': Formulas will be evaluated and calculated values returned. Formula display values (like #VALUE! errors) cannot be preserved."

## 待定问题

在当前实现中，现有Excel文件中的以下列被忽略：

- **工序编号** - 工序编号信息
- **乘系数定额** - 乘以系数后的定额金额
- **乘系数合计金额** - 乘以系数后的合计金额

这些列目前不在预期的列结构范围内，因此在数据处理过程中会被自动过滤掉。如果需要处理这些列，可以考虑：

1. 扩展数据库架构以包含这些字段
2. 添加相应的特殊逻辑规则来处理这些列
3. 更新`expected_columns`配置以包含这些字段

## 许可证

本项目仅供内部使用。

## 技术支持

如有问题，请联系开发团队。
