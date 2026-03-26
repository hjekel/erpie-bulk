# ERPIE Pricing Calibration Notes

**Date**: March 2026
**Reference deal**: IM_NL_NB_240 (240 notebooks, Tilburg)
**Comparison**: ERPIE calculated vs PlanBit actual sales price indication

## Summary of Deltas

| Model | Specs | Grade | PlanBit | ERPIE (old) | ERPIE (fixed) | Delta |
|-------|-------|-------|---------|-------------|---------------|-------|
| ThinkPad T14 Gen 1 | i5-10310U, 16GB, 0GB* | A | €140 | €65 | €120 | -14% |
| ThinkPad T14 Gen 2 | i5-1145G7, 16GB, 0GB* | A | €206 | €119 | €145 | -30% |
| ThinkPad T14 Gen 3 | i5-1245U, 16GB, 256GB | A | €297 | €188 | €188 | -37% |
| ThinkPad T14 Gen 4 | i5-1345U, 32GB, 256GB | A | €403 | €297 | €297 | -26% |
| ThinkPad T480 | i5-8350U, 0GB*, 0GB* | A | €99 | €24 | €65 | -34% |
| ThinkPad T490 | i5-8365U, 16GB, 256GB | A | €109 | €109 | €109 | ✅ |
| MacBook Air M1 | M1, 8GB, 256GB | A | €646 | €473 | €473 | -27% |
| MacBook Pro 16" 2019 | i7-9750H, 32GB, 512GB | A | €164 | €163 | €163 | ✅ |
| Acer TravelMate P215 | i5-1135G7, 8GB, 512GB | A | €194 | €130 | €145 | -25% |

*0GB = Blancco reported "not detected" — now treated as baseline (8GB/256GB)

## Root Causes

### 1. 0GB RAM/SSD from Blancco (FIXED)
- Blancco reports 0GB when it can't detect the spec
- Old code: 0GB → snapped to 4GB RAM (-€40) and 128GB SSD (-€25) = -€65 combined penalty
- **Fix**: 0GB now treated as "unknown" → defaults to baseline 8GB/256GB (no penalty)

### 2. Notebook Factor ×0.65 (NEEDS JOEP REVIEW)
- Applied to ALL non-Apple-Silicon, non-desktop devices
- Purpose: B2B resale discount vs B2C consumer prices
- **Problem**: PlanBit's "Sales Price Indication" IS the B2B resale price, not a consumer price
- The ×0.65 factor effectively double-discounts: base prices are already calibrated on PlanBit B2B sales data
- **Recommendation**: Remove notebook factor entirely, OR reduce to ×0.85-0.90

### 3. Apple Silicon base prices too low (NEEDS JOEP REVIEW)
- MacBook Air M1 base: €450 → with Grade A ×1.05 = €473
- PlanBit actual: €646
- **Delta**: -27% (€173 short)
- These prices may have depreciated since calibration, or the calibration data was from a different market period

### 4. ThinkPad T14 Gen 2/3 model-specific base prices (NEEDS JOEP REVIEW)
- T14 Gen 2 base: €190 → after 0.65 factor → €131
- PlanBit actual: €206
- Even without notebook factor: €190 × 1.05 (Grade A) = €200 (close to €206)
- **This confirms**: removing notebook factor would bring ERPIE much closer to PlanBit actuals

## Proposed Calibration Changes (for Joep to approve)

1. **Remove or reduce notebook factor** from ×0.65 to ×1.00 (or ×0.90 at most)
2. **Increase Apple Silicon base**: M1 Air €450 → €600-650
3. **Re-validate model-specific base prices** against latest PlanBit sales data
4. **Keep lot discount** (≥100: ×0.92, ≥200: ×0.85, ≥500: ×0.80) — these are appropriate for large batches

## Technical Notes

- Grade column (index 18 in Blancco export) IS being read correctly
- Grades A/B/C/D are extracted per individual asset
- The 0GB fix only affects the pricing calculation, not the parsed values displayed in the report
- After 0GB fix + removing notebook factor, ERPIE prices would typically be within 5-15% of PlanBit actuals
