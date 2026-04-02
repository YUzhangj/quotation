---
status: complete
phase: 03-plush-frontend
source: [01-SUMMARY.md]
started: 2026-04-02
updated: 2026-04-02
---

## Current Test

[testing complete]

## Tests

### 1. F. Sewing Detail Tab 出现在 Body Cost Breakdown
expected: 打开一个毛绒公仔报价，在 Body Cost Breakdown 区域能看到新增的 "F. Sewing Detail"（车缝明细）tab。点击后显示包含以下9列的表格：布料名称、部位、裁片数、用量、物料价、价钱、码点、总价。
result: skipped
reason: 功能已变更 — F. Sewing Detail tab 已取消，车缝明细移至其他位置

### 2. Sewing Detail 数据可编辑
expected: 在 F. Sewing Detail tab 中，点击任意单元格可以编辑数据，修改后数值正确保存和显示。
result: skipped
reason: F. Sewing Detail tab 已取消，此测试不适用

### 3. G. Rotocast Items Tab 出现在 Body Cost Breakdown
expected: 在 Body Cost Breakdown 区域能看到新增的 "G. Rotocast Items"（搪胶件）tab。点击后显示包含以下8列的表格：模号、名称、出数、用量、单价、合计、备注。
result: skipped
reason: 功能已变更 — 搪胶件内容移入 Body B.2，不再作为独立 tab

### 4. Rotocast Items 数据可编辑
expected: 在 G. Rotocast Items tab 中，点击任意单元格可以编辑数据，修改后数值正确保存和显示。
result: skipped
reason: G. Rotocast Items tab 已取消，此测试不适用

### 5. 注塑产品不受影响
expected: 打开一个注塑产品报价（非毛绒公仔），Body Cost Breakdown 中不出现 F. Sewing Detail 和 G. Rotocast Items tab，原有 tab 正常显示。
result: pass

## Summary

total: 5
passed: 1
issues: 0
pending: 0
skipped: 4

## Gaps

[none yet]
