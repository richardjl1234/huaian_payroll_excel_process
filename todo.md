# 待修复 - Excel 公式错误（#VALUE! / #REF!）

> 扫描方式：使用 LibreOffice 重新计算公式后检测
> 扫描日期：2026-05-11

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

### placeholder 目录（旧版备份）

| 文件 | 工作表 | 错误行数 |
|------|--------|----------|
| old_payroll/placeholder/202006.xls | 喷漆装配 | 136 |
| old_payroll/placeholder/202006.xls | 精加工 | 172 |
| old_payroll/placeholder/202007.xls | 喷漆装配 | 66 |
| old_payroll/placeholder/202009.xls | 装配喷漆 | 254 |
| old_payroll/placeholder/202010.xls | 装配喷漆 | 283 |
| old_payroll/placeholder/202011.xls | 装配喷漆 | 300 |

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

## 其他

| 文件 | 说明 |
|------|------|
| old_payroll/201711.xls 装配喷漆 | 第R列（超范围），#VALUE!，2行 |
| old_payroll/placeholder/202003.xls 喷漆装配 | 第E列（计件数量），#VALUE!，1行 |
| old_payroll/placeholder/202010_2.xls | 第G列（装配喷漆），#VALUE!，17行 |

---

## 文件损坏（无法读取）

| 文件 | 说明 |
|------|------|
| new_payroll/202003.xls | XML解析错误 |
| new_payroll/202109.xls | XML解析错误 |

---

## 已删除的旧文件

| 文件 | 说明 |
|------|------|
| deleted/202010.xls | 装配喷漆，第G列，#VALUE!，192行 |

---

**总计影响：约 2,400+ 行数据**
**主要问题：** 金额列（第G列）公式 `=计件数量×定额` 因引用断掉导致 #VALUE!
