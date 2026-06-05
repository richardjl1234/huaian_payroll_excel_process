# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

工资数据批处理工具：从 `new_payroll/` 和 `old_payroll/` 文件夹读取 Excel 文件（.xls/.xlsx），将非结构化的工作表分割为表格、规范化列名、加载到 SQLite 数据库，并提供 8 步数据清洗流程。

## 环境配置

### 数据库路径
- **环境变量**: `SQLITE_DB_PATH`（必须设置，`batch_process.py` 和核心模块使用）
- **清洗脚本路径**: 多数 `cleansing_*.py` 脚本使用父目录的 `payroll_database.db`（即 `../payroll_database.db`）
- 数据库文件本身被 `.gitignore` 排除

## 常用命令

### 1. 批处理主程序
```bash
# 处理 new_payroll/ 和 old_payroll/ 中的所有 Excel 文件
export SQLITE_DB_PATH=/home/richard/shared/jianglei/payroll/payroll_database.db
python batch_process.py

# 单文件模式（不清理数据库）
python batch_process.py 201406.xls
```

### 1.5 端到端刷新脚本（推荐）
```bash
./sqlite_payroll_details_refresh.sh
```
该脚本组合了 `batch_process.py` 和全部 9 个清洗步骤（Step 0–9），从空的 SQLite 数据库开始，完整地生成 `payroll_details` 表。
- 脚本会从父目录 `source payroll_local_sqlite.sh` 来设置环境变量
- 中间任何步骤出错需排查后再重跑整个脚本（脚本不会自动恢复）

### 2. 数据检查工具
```bash
python check_payroll_database.py     # 工资数据交互式查询
python check_load_log_table.py       # 加载日志检查
python data_overview_analysis.py     # 综合数据分析 + 图表输出
```

### 3. Web 文件查看器
```bash
python one_time_pgms/excel_viewer.py
# 访问 http://localhost:5000
```

### 4. 数据清洗流程（重要：必须按顺序执行）
所有清洗脚本都支持 `--dry-run` 预览模式。

```bash
# Step 0: 全角转半角 (1,640 条)
python cleansing_data_dbcs_handling_step0.py [--dry-run]

# Step 1: 删除异常日期记录
python cleansing_outliers_step1.py [--dry-run]

# Step 2: 前向填充空日期
python cleansing_outliers_step2.py [--dry-run]

# Step 3: 删除"合计"行
python cleansing_outliers_step3.py [--dry-run]

# Step 4: 月份日期→工作日（用 holidays 模块识别中国法定节假日）
python cleansing_outliers_step4.py [--dry-run]

# Step 5: 统一为逗号分隔格式 (39,223 条)
python cleansing_date_handling_step5.py [--dry-run]

# Step 6: 复杂混合日期模式 (2,741 条)
python cleansing_date_handling_step6.py [--dry-run]

# Step 7: 展开波浪号范围 (7,757 条)
python cleansing_date_handling_step7.py [--dry-run]

# Step 8: 杂项清理（删除非日期 + 空格/休息日/& 替换）(24 删除 + 273 更新)
python cleansing_misc_step8.py [--dry-run]

# Step 9: 清理残留的 yy,m / m 前缀（如 `14,6,3`→`3`、`6,1`→`1`）(975 更新 + 19 错误保留)
python cleansing_date_handling_step9.py [--dry-run]
```

### 5. 验证脚本
```bash
# 验证日期列数据质量
python validate_date_column.py [--show-errors] [--export-html]

# 查找 Excel 文件中的 #VALUE! 错误（最新添加的一次性程序）
python one_time_pgms/find_excel_errors.py  # 详见 todo.md
```

完整的清洗流程文档见 `cleansing_outliers_process.md`。

## 架构

### 数据处理管道
```
Excel 文件
  → sheet_gen()         [excel_processor/sheet_gen.py]
      生成 SheetContents (file_name, sheet_name, raw_sheet_contents)
  → df_gen()            [excel_processor/df_gen.py]
      生成 SplitDataFrame (file_name, sheet_name, table_index, split_df)
  → load_df_to_db()     [excel_processor/sheet_processor.py]
      应用 special_logic + 插入 payroll_details / load_log 表
  → SQLite 数据库
```

### 核心模块 (`excel_processor/`)
- **`config.py`**: `expected_columns` 列表（12 列）+ 全局日志配置
- **`sheet_gen.py`**: `get_excel_files()` 扫描两个文件夹；`sheet_gen()` 是生成器，跳过 `汇总`/`统计`/`deleted` 工作表
- **`sheet_processor.py`**: 关键函数
  - `_process_cell_value()`: 将整数如 `132.0` → `132`（型号/日期列必需）
  - `get_all_data_from_sheet()`: 用 openpyxl (xlsx) / xlrd (xls) 读取
  - `split_raw_sheet_contents()`: 按空行分割为多个表
  - `_validate_and_fix_dataframe_columns()`: 表头检测，< `COMMON_COL_COUNT=4` 时扫描行内寻找更好的表头
  - `load_df_to_db()`: 应用 18 条特殊逻辑 + 数值列 `Decimal` 精度处理 + 写入数据库
- **`df_gen.py`**: `df_gen()` 生成器，包装 `split_raw_sheet_contents()` 为带 `table_index` 的 namedtuple
- **`special_logic.py`**: `special_logic_preprocess_df()` 包含 18 条规则，工作表名映射（`14年6月精加工` → `精加工` 等），`前装/中装/后装` 人员拆分

### 数据库表结构（详见 README.md）
- **`payroll_details`**: 14 列，NUMERIC(10,2) 精度（金额、系数等）
- **`load_log`**: 记录被丢弃的列名，用于数据质量监控

### 特殊逻辑（18 条规则，详见 README.md）
关键规则摘要：
- L1: 喷漆装配表"前装/中装/后装/刘雷/装配"→"职员全名"
- L14-16: 前装/中装/后装人员的行展开
- L17: 职员全名为空→删除行
- L18: 特定中文短语（下料/铣底脚/校平衡等）→删除行

## 重要注意事项

1. **.xls 文件的 xlrd 限制**: xlrd 不保留 `#VALUE!` 等公式显示错误（会计算为值）。详见 `todo.md` 中待修复的 2,400+ 行 #VALUE! 问题
2. **环境变量必需**: 不设置 `SQLITE_DB_PATH` 时 `batch_process.py` 会因 `sqlite3.connect(None)` 失败
3. **文件名约定**: 格式为 `YYYYMM.xls`（如 `202506.xls`），清洗脚本用此验证日期是否在当月有效范围内
4. **清洗步骤顺序敏感**: 必须 Step 0 → 1 → 2 → ... → 9，每步假设上一步已完成
5. **执行前先 dry-run**: 所有清洗脚本都支持 `--dry-run`，会生成 `*_output.html` 预览报告
6. **日志文件**:
   - `log_batch.txt`: 主处理日志（被 .gitignore 排除）
   - `special_logic_applied.log`: 特殊逻辑应用记录
7. **当前分支**: `process_date_col`，主线 `main`；近期工作集中在日期列数据质量
