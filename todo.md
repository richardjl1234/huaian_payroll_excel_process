# 待修复 - Excel 公式错误（#VALUE! / #REF!）

> 扫描方式：使用 LibreOffice 重新计算公式后检测
> 扫描日期：2026-05-11（最近一次）
> 注意：todo.md 中的错误源自源 Excel 文件（LibreOffice 重新计算后能看到），
>     但 batch_process.py 用 xlrd 读 .xls 时会**吞掉 #VALUE! 变成 0/空**，
>     所以 DB 中相应行的 `金额` 列被记为 0（audit finding "金额=0 但计件数量≠0"）。
>     修复源文件是根治办法, 重新跑 `sqlite_payroll_details_refresh.sh` 后 DB 金额会变正确。

---

## 第 G 列（金额列）— #VALUE!

### 装配喷漆 / 喷漆装配 表

| 文件 | 工作表 | 错误行数 |
|------|--------|----------|
| new_payroll/202011.xls | 装配喷漆 | 303 |
| new_payroll/202010.xls | 装配喷漆 | 283 |
| new_payroll/202009.xls | 装配喷漆 | 159 |
| new_payroll/202006.xls | 喷漆装配 | 134 |
| new_payroll/202007.xls | 喷漆装配 | 66 |
| new_payroll/202012.xlsx | 喷漆装配 | 92 |
| new_payroll/202101.xls | 喷漆装配 | 120 |
| new_payroll/202102.xls | 喷漆装配 | 60 |
| new_payroll/202105.xls | 喷漆装配 | 38 |
| new_payroll/202108.xls | 装配喷漆 | 39 |

### 精加工 表

| 文件 | 错误行数 |
|------|----------|
| new_payroll/202006.xls | 168 |
| new_payroll/202106.xls | 133 |
| new_payroll/202101.xls | 23 |
| new_payroll/202108.xls | 23 |
| new_payroll/202105.xls | 16 |
| new_payroll/202110.xls | 21 |

---

## 第 L 列（备注列）— #VALUE!

| 文件 | 工作表 | 错误行数 |
|------|--------|----------|
| new_payroll/202108.xls | 装配喷漆 | 61 |
| new_payroll/202106.xls | 喷漆装配 | 35 |
| new_payroll/202105.xls | 喷漆装配 | 1 |

---

## 第 C/E/F 列（汇总表）— #REF!

| 文件 | 工作表 | 错误列 | 错误行数 |
|------|--------|--------|----------|
| new_payroll/202504.xls | 汇总 | C、E、F | 3 |

---

## 其他（主目录）

| 文件 | 说明 |
|------|------|
| old_payroll/201711.xls 装配喷漆 | 第R列（超范围），#VALUE!，2行 |

---

## 文件损坏（无法读取）

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

**总计影响：约 2,400+ 行数据**（其中 DB 端表现为"金额=0 但计件数量≠0" ~21,486 行, 含历史 #VALUE! 残留）
**主要问题：** 金额列（第G列）公式 `=计件数量×定额` 因引用断掉导致 #VALUE!
