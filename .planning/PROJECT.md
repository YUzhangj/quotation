# 报价管理系统 (Vendor Quotation System)

## What This Is

基于浏览器的报价管理系统，导入本厂报价明细 Excel，交互式编辑成本数据，导出 TOMY 格式的 Vendor Quotation Excel。支持注塑产品和毛绒公仔两种报价格式。

## Core Value

准确高效地将内部报价明细转换为客户报价单，消除手工填写的错误和重复劳动。

## Requirements

### Validated

- ✓ Excel 文件导入解析（SheetJS/ExcelJS） — existing
- ✓ 注塑产品报价明细解析（47712 格式） — existing
- ✓ 产品/版本 CRUD 管理 — existing
- ✓ 11 个 tab 的交互式编辑界面 — existing
- ✓ 参数面板（汇率、加价率等）自动重算 — existing
- ✓ 成本计算引擎 — existing
- ✓ TOMY 模板驱动的 Excel 导出 — existing
- ✓ Docker 部署 — existing
- ✓ SQLite 数据持久化 — existing

### Active

- [ ] 毛绒公仔报价格式支持（L21014 格式：3K报价主 sheet + 搪胶/车缝明细等子 sheet）
- [ ] Raw Material 自动从 MoldPart 提取（两种格式）
- [ ] 导入时正确识别产品编号（从主报价 sheet 的 B1）
- [ ] 车缝明细 sheet 解析（布料名称、部位、用量、物料价等）
- [ ] 搪胶件 sheet 解析（搪胶脸/脚等部件）
- [ ] 多子 sheet 数据解析（五金、吊咭、贴纸、PE袋等）
- [ ] 未提交的 21 个文件整理提交

### Out of Scope

- 后端服务器认证 — 内部工具无需登录
- 在线部署 — 本地工具
- 修改 TOMY 模板结构

## Context

- 已有完整的注塑产品报价系统（47712），约 2600 行代码
- 毛绒公仔报价 Excel 结构完全不同：主 sheet 为 `3K报价-地区-YYMMDD`，含搪胶件表(R20-R22)、成本明细(R25+)，另有车缝明细等多个子 sheet
- 两种格式的 MoldPart 区域（R16/R17）结构相同，但数据量差异大
- 21 个未提交文件包含 header-info 面板、tab 优化等增强功能

## Constraints

- **Tech stack**: Node.js + Express + SQLite + ExcelJS + vanilla HTML/CSS/JS
- **Architecture**: 现有 client/server 结构不变
- **Template**: TOMY 模板格式不可修改

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Raw Material 从 MoldPart.unit_price_hkd_g 取价格 | MoldPart 自带对应料价，比从 MaterialPrice 匹配更准确 | 已实施 |
| Raw Material weight 不乘 sets_per_toy | sets_per_toy 是啤工计算用，原料重量应为单件克重累加 | 已实施 |
| 产品编号从主报价 sheet B1 提取 | 车缝明细等子 sheet 的 B1 不是货号 | 已实施 |

## Evolution

This document evolves at phase transitions and milestone boundaries.

---
*Last updated: 2026-03-28 after initialization*
