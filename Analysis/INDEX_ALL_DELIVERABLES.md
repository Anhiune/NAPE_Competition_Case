# 2026 SMU NAPE CASE COMPETITION - ANALYSIS DELIVERABLES
## Complete Index of All Files and Findings

**Analysis Date:** November 19, 2025  
**Competition:** 2026 SMU NAPE Energy Case Competition  
**Analyst:** Ben Nguyen EHK Team

---

## 📋 EXECUTIVE DOCUMENTS

### 1. Executive Summary (START HERE)
**File:** `EXECUTIVE_SUMMARY_NAPE_CASE_ANALYSIS.md`

**Contents:**
- Strategic recommendations with hybrid approach
- All 6 prediction targets summarized
- Real co-located deal analysis (Talen-AWS $18B nuclear PPA)
- Impact on company growth targets (30%+ Adj FCF/share)
- Sensitivity analysis and risk scenarios
- Strategic value creation framework

**Key Finding:** Hybrid approach (Nuclear + Existing Gas) achieves 65% EBITDA growth and positions company for AI-related market multiple expansion.

---

### 2. Detailed Calculations & Methodology
**File:** `DETAILED_CALCULATIONS_ALL_TARGETS.md`

**Contents:**
- Complete formulas for all 6 prediction targets
- Step-by-step calculations for each scenario
- Input features and data sources documented
- Year-by-year cash flow methodology
- NPV, IRR, LCOE derivations
- Comparative analysis across all metrics

**Use This For:** Understanding exactly how every number was calculated, verifying methodology, replicating analysis.

---

## 📊 COMPREHENSIVE VISUALIZATIONS

### 3. All 6 Targets Dashboard (18 Charts)
**File:** `comprehensive_analysis_all_6_targets.png`

**Visualization Grid (6 rows × 3 charts each):**

**Row 1 - Target 1: Physical Generation & Utilization**
- Chart 1a: Installed Capacity (MW) comparison
- Chart 1b: Capacity Factor (%) with baseload threshold
- Chart 1c: Annual Energy Output (TWh/year)

**Row 2 - Target 2: Operating Costs & Unit Economics**
- Chart 2a: Operating cost components (stacked bar: fuel, O&M, carbon)
- Chart 2b: Total operating cost per MWh
- Chart 2c: Annual fixed O&M costs

**Row 3 - Target 3: Capital Costs & Project Finance**
- Chart 3a: Total capital investment ($M)
- Chart 3b: Capital cost per kW ($/kW)
- Chart 3c: Capital structure (debt vs equity stacked)

**Row 4 - Target 4: Carbon & Policy Exposure**
- Chart 4a: Carbon emissions intensity (lb/MWh)
- Chart 4b: Annual carbon emissions (kt/year)
- Chart 4c: Annual carbon cost ($M)

**Row 5 - Target 5: LCOE & Comparative Economics**
- Chart 5a: LCOE components stacked (capital, fuel, O&M, carbon)
- Chart 5b: Total LCOE comparison with market price line
- Chart 5c: Energy margin vs market price

**Row 6 - Target 6: Year-by-Year Forecasts**
- Chart 6a: Net Present Value ($M) with positive/negative indicators
- Chart 6b: Average annual EBITDA ($M/year)
- Chart 6c: Cumulative free cash flow over project life ($M)

**Key Insight from Charts:** Nuclear dominates on generation, operating efficiency, and carbon; gas wins on capital efficiency and traditional NPV.

---

### 4. Year-Over-Year Projections (6 Time-Series Charts)
**File:** `year_over_year_projections.png`

**Time-Series Comparisons:**
1. Annual Revenue Projection (3 scenario lines)
2. Annual EBITDA Projection (3 scenario lines)
3. Annual Free Cash Flow Projection (3 scenario lines)
4. Operating Costs with Escalation (3 scenario lines)
5. Cumulative Free Cash Flow (3 scenario lines)
6. Annual Net Income (3 scenario lines)

**Shows:** How each scenario performs year-by-year over full project life (15-20 years), including O&M escalation effects, depreciation impacts, and cash flow trajectories.

---

## 📈 DETAILED DATA FILES

### 5. Yearly Financial Projections - Nuclear
**File:** `yearly_projections_scenario_A.csv`

**Columns (17 years of data):**
- year, mwh_generated, energy_revenue, capacity_revenue, total_revenue
- fuel_cost, variable_om, fixed_om, carbon_cost, total_operating_cost
- ebitda, depreciation, ebit, taxes, net_income
- operating_cash_flow, interest_expense, fcff
- discount_factor, pv_fcff

**Use For:** Building your own models, sensitivity analysis, detailed financial statement preparation.

---

### 6. Yearly Financial Projections - Build Gas
**File:** `yearly_projections_scenario_B.csv`

**Same structure as above, 20 years of data**

---

### 7. Yearly Financial Projections - Acquire Gas
**File:** `yearly_projections_scenario_C.csv`

**Same structure as above, 15 years of data**

---

## 🔢 RAW DATA & CONFIGURATION

### 8. Scenario Definitions
**File:** `scenario_definitions.json`

**Contains:** All input parameters for three scenarios including:
- Capacity specifications (MW, capacity factor)
- Cost structures (fuel, O&M, capital)
- Financial assumptions (WACC, debt/equity, tax rate)
- Technical parameters (heat rate, emissions, economic life)
- Strategic attributes (deal type, partnership model)

---

### 9. Targets 1-5 Analysis Results
**File:** `analysis_results_targets_1_5.json`

**Structured JSON with:**
- Generation & utilization metrics (Target 1)
- Operating cost breakdowns (Target 2)
- Capital cost calculations (Target 3)
- Carbon exposure analysis (Target 4)
- LCOE component breakdown (Target 5)

---

### 10. Target 6 Forecast Results
**File:** `forecast_results_target_6.json`

**Contains:**
- NPV calculations
- IRR results
- Average annual EBITDA
- Average annual free cash flow
- Cumulative cash flows
- Total revenue over project life

---

## 🎯 THREE STRATEGIC SCENARIOS ANALYZED

### Scenario A: Co-located Nuclear Partnership (Talen-AWS Model)

**Real-World Deal Basis:**
- **Source:** Talen Energy - Amazon Web Services
- **Value:** $18 billion over 17 years
- **Capacity:** 800 MW nuclear (from 2,500 MW Susquehanna plant)
- **Structure:** Front-of-meter, grid-connected PPA
- **Location:** PJM market, Pennsylvania

**Key Metrics:**
- Capacity: 800 MW at 92% capacity factor
- Annual Output: 6.45 TWh/year (baseload)
- Operating Cost: $15.48/MWh (lowest)
- Total LCOE: $65.41/MWh
- Carbon: Zero emissions
- Capex: $3,000M acquisition
- NPV: -$1,467M (but strategic value not captured)
- Avg EBITDA: $392M/year (highest)

**Strategic Value:**
- Perfect data center alignment (24/7 carbon-free)
- Zero carbon policy risk
- Platform for future partnerships
- Potential for "AI-related" market multiple

---

### Scenario B: Build New CCGT Natural Gas

**Key Metrics:**
- Capacity: 550 MW at 70% capacity factor
- Annual Output: 3.37 TWh/year (mid-merit)
- Operating Cost: $35.38/MWh
- Total LCOE: $55.65/MWh (middle)
- Carbon: 720 lb/MWh (1.2M tons/year)
- Capex: $687.5M new build
- Construction: 24 months
- NPV: -$134M
- Avg EBITDA: $115M/year

**Advantages:**
- Lower capital than nuclear
- Flexible, dispatchable generation
- Can be positioned near data centers
- Modern, efficient technology

---

### Scenario C: Acquire Existing CCGT Natural Gas

**Key Metrics:**
- Capacity: 550 MW at 60% capacity factor
- Annual Output: 2.89 TWh/year (mid-merit)
- Operating Cost: $36.03/MWh
- Total LCOE: $53.86/MWh (lowest)
- Carbon: 875 lb/MWh (1.3M tons/year)
- Capex: $450M acquisition
- NPV: +$64M (POSITIVE - best traditional return)
- Avg EBITDA: $108M/year
- Payback: 4-5 years (fastest)

**Advantages:**
- Positive NPV on DCF basis
- Lowest capital requirement
- Immediate cash flow
- Fast payback period

---

## 📊 KEY COMPARATIVE FINDINGS

### Financial Metrics Winner Matrix

| Metric | Nuclear | Build Gas | Buy Gas | Winner |
|--------|:-------:|:---------:|:-------:|:------:|
| **NPV** | -$1,467M | -$134M | +$64M | ✓ Gas Buy |
| **IRR** | <7.68% | ~7.5% | ~8.5% | ✓ Gas Buy |
| **Payback** | 8-9 yrs | 7-8 yrs | 4-5 yrs | ✓ Gas Buy |
| **LCOE** | $65.41 | $55.65 | $53.86 | ✓ Gas Buy |
| **Avg EBITDA** | $392M | $115M | $108M | ✓ Nuclear |
| **Annual MWh** | 6.45M | 3.37M | 2.89M | ✓ Nuclear |
| **Operating Cost** | $15.48 | $35.38 | $36.03 | ✓ Nuclear |
| **Carbon Cost** | $0 | $24M | $25M | ✓ Nuclear |
| **Capacity Factor** | 92% | 70% | 60% | ✓ Nuclear |

### Strategic Value Winner Matrix

| Factor | Nuclear | Build Gas | Buy Gas | Winner |
|--------|:-------:|:---------:|:-------:|:------:|
| **Data Center Alignment** | Perfect | Good | Good | ✓ Nuclear |
| **Carbon Policy Risk** | Zero | High | High | ✓ Nuclear |
| **Market Multiple Potential** | High | Low | Low | ✓ Nuclear |
| **Growth Platform** | Excellent | Limited | Limited | ✓ Nuclear |
| **Contracted Revenue** | Yes | No | No | ✓ Nuclear |
| **Capital Efficiency** | Low | Medium | High | ✓ Gas Buy |
| **Near-term Cash Flow** | Medium | Medium | High | ✓ Gas Buy |
| **Balance Sheet Impact** | High | Medium | Low | ✓ Gas Buy |

---

## 💡 STRATEGIC RECOMMENDATION SUMMARY

### Hybrid Approach: Nuclear + Existing Gas = $3.45B Investment

**Why This Combination?**

1. **Nuclear ($3.0B)** provides:
   - Strategic positioning as "AI power provider"
   - $392M annual EBITDA
   - Zero carbon for data center partnerships
   - Growth platform for future deals
   - Potential market multiple expansion

2. **Existing Gas ($450M)** provides:
   - Positive NPV (+$64M)
   - $108M annual EBITDA
   - Near-term cash flow support
   - 4-5 year payback
   - Balance sheet cushion

**Combined Impact:**
- **Total EBITDA:** ~$500M/year (65% increase)
- **Adj FCF/Share Growth:** 42%+ (exceeds 30% target)
- **Market Multiple:** 35-40x from "AI positioning" (vs 30x current)
- **Market Cap Potential:** $44B (122% increase from $20B)
- **Strategic Position:** Major player in $517B data center market

---

## 📞 NEXT STEPS FOR COMPETITION

### For Presentation Development:

1. **Start with:** `EXECUTIVE_SUMMARY_NAPE_CASE_ANALYSIS.md`
   - Contains full strategic narrative
   - All key findings summarized
   - Recommendation clearly stated

2. **Use for slides:** `comprehensive_analysis_all_6_targets.png`
   - 18 charts covering all prediction targets
   - Can extract individual charts for PowerPoint
   - Clear visual comparisons across scenarios

3. **Support with data:** CSV files for detailed backup
   - Show year-by-year projections
   - Demonstrate rigorous financial modeling
   - Available for Q&A on any metric

4. **Reference methodology:** `DETAILED_CALCULATIONS_ALL_TARGETS.md`
   - All formulas documented
   - Calculations shown step-by-step
   - Data sources cited (Talen-AWS deal)

---

## 🔍 HOW TO USE THESE FILES

### For PowerPoint Presentation (Max 15 slides):

**Suggested Slide Structure:**
1. Title Slide
2. Executive Summary (Key recommendation)
3. Context: IPP + Data Center Growth Opportunity
4. Scenario Overview (3 options analyzed)
5-10. One slide per prediction target (6 slides) with key charts
11. Financial Comparison Matrix
12. Strategic Value Analysis
13. Recommended Hybrid Approach
14. Impact on Company Targets
15. Conclusion & Next Steps

### For Video Presentation (15 minutes):

**Suggested Flow (2-3 minutes per section):**
1. Opening: Problem setup & company context (2 min)
2. Real-world deal: Talen-AWS nuclear partnership (2 min)
3. Three scenarios & methodology (2 min)
4. Target findings highlights: Walk through key charts (4 min)
5. Comparative analysis: What each scenario offers (2 min)
6. Recommendation: Why hybrid approach wins (2 min)
7. Closing: Impact on growth targets & strategic value (1 min)

### For Supporting Materials:

- **Appendix:** Include all CSV files and JSON data
- **Backup slides:** Individual target charts from PNG files
- **Financial model:** Excel workbook built from CSV projections
- **Citations:** Talen-AWS deal sources documented in summary

---

## 📚 DATA SOURCES CITED

### Real Co-located Deal:
- **Power Magazine:** "Talen, Amazon Launch $18B Nuclear PPA" (June 13, 2025)
- **Utility Dive:** "Talen to sell Amazon 1.9 GW from Susquehanna nuclear plant" (June 11, 2025)
- **World Nuclear News:** "New supply agreement expands Talen-Amazon partnership" (June 12, 2025)

### Technical Data:
- **Case Data Excel File:** All plant specifications, cost parameters, financial assumptions
- **PJM Market Data:** Capacity pricing, wholesale energy prices
- **Industry Standards:** MACRS depreciation, nuclear/gas performance benchmarks

---

## ✅ ANALYSIS COMPLETION CHECKLIST

- [x] All 6 prediction targets calculated
- [x] Real co-located deal researched (Talen-AWS)
- [x] Three scenarios fully modeled
- [x] Year-by-year forecasts (15-20 years)
- [x] Comprehensive visualizations created
- [x] NPV, IRR, LCOE computed
- [x] Strategic recommendation developed
- [x] All formulas documented
- [x] Data sources cited
- [x] Files organized and indexed

---

## 📧 QUESTIONS OR CLARIFICATIONS?

All calculations are transparent and documented in the detailed files. For any questions about:
- **Methodology:** See DETAILED_CALCULATIONS_ALL_TARGETS.md
- **Results:** See EXECUTIVE_SUMMARY_NAPE_CASE_ANALYSIS.md  
- **Raw data:** See CSV and JSON files
- **Visualizations:** See PNG chart files

---

**Good luck with the 2026 SMU NAPE Case Competition!**

The analysis demonstrates rigorous financial modeling combined with strategic thinking about the intersection of AI, data centers, and power generation. The hybrid approach balances traditional financial returns with forward-looking strategic positioning in the emerging AI infrastructure market.

**Go compete with confidence!** 🎯📊💪
