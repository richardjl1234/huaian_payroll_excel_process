# 待修复 - Excel 公式错误（#VALUE! / #REF!）

> 扫描方式：使用 LibreOffice 重新计算公式后检测
> 扫描日期：2026-05-11（最近一次）
> 修复日期：2026-06-13（已修复第 G 列 #VALUE!，含 G 列 cached value 写入）
> 注意：todo.md 中的错误源自源 Excel 文件（LibreOffice 重新计算后能看到），
>     但 batch_process.py 用 xlrd 读 .xls 时会**吞掉 #VALUE! 变成 0/空**，
>     所以 DB 中相应行的 `金额` 列被记为 0（audit finding "金额=0 但计件数量≠0"）。
>     修复源文件是根治办法, 重新跑 `sqlite_payroll_details_refresh.sh` 后 DB 金额会变正确。

---

## ✅ 已修复 (2026-06-13): 第 G 列（金额列）— #VALUE!

**修复工具**: `one_time_pgms/fix_f_column_value_error.py`
**修改内容**: F 列（定额）单元格 → 0.8 + 黄底红字；同时触发 LibreOffice 重算让 G 列 cached value 写入
**修复总数**: 1,678 个 F 单元格（G 列 #VALUE! 全部清零）
**影响文件**: 12 个 new_payroll 文件 (.xls → .xlsx)
**备份位置**: `/home/richard/shared/jianglei/payroll/{new,old}_payroll_backup_20260612/`
**DB 验证**: 修复前 12 个文件 金额=0 但计件数量≠0 共约 450+ 行，修复后 0 行（仅 202012.xlsx 非目标 sheet 残留 2 行）

### 修复中遇到的关键坑 (2026-06-13)

G 列 `=F*E` 公式的 cached value 在修复后必须正确写入，否则 batch_process.py 会读 G=None → 金额=0。三处必须满足：

1. **`wb.calculation.fullCalcOnLoad = False`** — workbook.xml 的 `<calcPr fullCalcOnLoad="1"/>` 会让 LibreOffice 重算时故意不写 `<v>` 标签（让 Excel 打开时再算）。openpyxl 改 F 后必须关掉这个标志。
2. **显式 filter `--convert-to "xlsx:Calc Office Open XML"`** — 隐式 `--convert-to xlsx` 留下空 `<v></v>` 标签。
3. **`--outdir` 与输入路径必须不同** — 否则 LibreOffice 写保存失败 (`SfxBaseModel::impl_store failed 0x4c0c`, 同路径 in==out)。脚本用 `temp_dir/_recalc_tmp/` 中转。

幂等性：`build_work_list` 现在兼容 .xls 和 .xlsx 两种后缀（首次跑改 .xls，二次跑发现 .xls 不存在会改 .xlsx）。

### 装配喷漆 / 喷漆装配 表

| 文件 | 工作表 | 错误行数 | 修复后 |
|------|--------|----------|--------|
| new_payroll/202011.xls | 装配喷漆 | 303 | ✅ → 202011.xlsx |
| new_payroll/202010.xls | 装配喷漆 | 283 | ✅ → 202010.xlsx |
| new_payroll/202009.xls | 装配喷漆 | 159 | ✅ → 202009.xlsx |
| new_payroll/202006.xls | 喷漆装配 | 134 | ✅ → 202006.xlsx |
| new_payroll/202007.xls | 喷漆装配 | 66 | ✅ → 202007.xlsx |
| new_payroll/202012.xlsx | 喷漆装配 | 92 | ✅ (原本就 .xlsx) |
| new_payroll/202101.xls | 喷漆装配 | 120 | ✅ → 202101.xlsx |
| new_payroll/202102.xls | 喷漆装配 | 60 | ✅ → 202102.xlsx |
| new_payroll/202105.xls | 喷漆装配 | 38 | ✅ → 202105.xlsx |
| new_payroll/202108.xls | 装配喷漆 | 39 | ✅ → 202108.xlsx |

### 精加工 / 金加工 表

| 文件 | 工作表 | 错误行数 | 修复后 |
|------|--------|----------|--------|
| new_payroll/202006.xls | 精加工 | 168 | ✅ → 202006.xlsx |
| new_payroll/202106.xls | 金加工 | 133 | ✅ → 202106.xlsx |
| new_payroll/202101.xls | 金加工 | 23 | ✅ → 202101.xlsx |
| new_payroll/202108.xls | 金加工 | 23 | ✅ → 202108.xlsx |
| new_payroll/202105.xls | 金加工 | 16 | ✅ → 202105.xlsx |
| new_payroll/202110.xls | 金加工 | 21 | ✅ → 202110.xlsx |

---

## 仍未处理 (out of scope - 暂不修)

### 第 L 列（备注列）— #VALUE!

| 文件 | 工作表 | 错误行数 |
|------|--------|----------|
| new_payroll/202108.xls | 装配喷漆 | 61 |
| new_payroll/202106.xls | 喷漆装配 | 35 |
| new_payroll/202105.xls | 喷漆装配 | 1 |

### 第 R 列 — #VALUE!

| 文件 | 工作表 | 错误行数 |
|------|--------|----------|
| old_payroll/201711.xls | 装配喷漆 | 2 |

### 第 C/E/F 列（汇总表）— #REF!

| 文件 | 工作表 | 错误列 | 错误行数 |
|------|--------|--------|----------|
| new_payroll/202504.xls | 汇总 | C、E、F | 3 |

### 文件损坏（无法读取）

| 文件 | 说明 |
|------|------|
| new_payroll/202003.xls | XML解析错误 |
| new_payroll/202109.xls | XML解析错误 |

---

## 待办：自动化扫描工具

- **`one_time_pgms/take_excel_screenshot.py`**: 当前**未运行**. 脚本中 `FILES_WITH_ISSUES` 是 2026-05 手工 hardcoded 清单, 不会扫描新错误. 需要:
  1. 跑一次 `python one_time_pgms/take_excel_screenshot.py` 生成当前错误位置的 HTML 报告到 `screenshot/` 目录
  2. 用脚本里 `find_all_errors_in_column()` 函数 (基于 openpyxl `data_only=True`) 做一次**全文件全列扫描**, 找到所有 #VALUE!/#REF! 行 (不依赖 hardcoded 清单)
  3. 把扫描结果写回 todo.md 顶部表格, 替换过期数据

---

**总计影响：约 2,400+ 行数据**（G 列 #VALUE! 已修 1,678 行; L/R/汇总 列 105 行未修; 损坏文件 2 个）
**主要问题：** 金额列（第G列）公式 `=计件数量×定额` 因 F 列 VLOOKUP 引用断掉导致 #VALUE!
