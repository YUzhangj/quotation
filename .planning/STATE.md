# Project State

## Project Reference

See: .planning/PROJECT.md (updated 2026-03-28)

**Core value:** 准确高效地将内部报价明细转换为客户报价单
**Current focus:** Phase 1 — 基础修复与代码整理

## Current Position

Phase: 1 of 3 (基础修复与代码整理)
Plan: 0 of ? in current phase
Status: Ready to plan
Last activity: 2026-03-28 — Project initialized, FIX-01 and FIX-02 partially implemented

Progress: [░░░░░░░░░░] 0%

## Accumulated Context

### Decisions

- Raw Material 从 MoldPart.unit_price_hkd_g 取价格（不从 MaterialPrice 匹配）
- Raw Material weight 不乘 sets_per_toy（单件克重直接累加）
- 产品编号从含"报价"关键字的 sheet 的 B1 提取
- 两种报价格式：注塑（报价明细-YYMMDD）和毛绒公仔（3K报价-地区-YYMMDD）

### Pending Todos

None yet.

### Blockers/Concerns

- FIX-01 和 FIX-02 的代码已改但尚未 commit
- 21 个未提交文件需要先整理提交

## Session Continuity

Last session: 2026-03-28
Stopped at: GSD 初始化完成，准备 plan Phase 1
Resume file: None
