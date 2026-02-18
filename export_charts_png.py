"""
Export all charts from CASE_TWIST_ANALYSIS.qmd as high-res PNGs for PPTX.
Extracts Python code blocks, runs them, saves PNGs in Output/ folder.
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
import numpy as np
import os

os.makedirs("Output", exist_ok=True)

# === QMD Chart Setup ===
plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "grid.alpha": 0.30,
    "grid.linestyle": "--",
    "figure.dpi": 200,
    "axes.facecolor": "#FDFAF3",
    "figure.facecolor": "#FFFFFF",
})

C_NUCLEAR = "#3A4E42"
C_BUILD   = "#5E9A66"
C_BUY     = "#6AC758"
C_JV      = "#3A4E42"
C_NEUTRAL = "#B0A599"
C_UP      = "#6AC758"
C_DOWN    = "#C0392B"
C_SELL    = "#323232"
C_DC      = "#AFD480"

# ──────────────────────────────────────────────────
# 1. Talen Stock Chart
# ──────────────────────────────────────────────────
months = np.arange(0, 25)
prices = [25, 30, 38, 42, 55, 62, 68, 72,
          80, 95, 120, 160, 180, 210, 240,
          270, 300, 320, 340, 360, 380, 395, 410, 420, 389]

fig, ax = plt.subplots(figsize=(10, 4))
ax.fill_between(months, prices, alpha=0.15, color=C_NUCLEAR)
ax.plot(months, prices, color=C_NUCLEAR, linewidth=2.5, marker="o", markersize=3)
ax.annotate("AWS/Cumulus\nDeal Announced\n$650M",
            xy=(10, 120), xytext=(4, 280),
            fontsize=9, fontweight="bold", color=C_NUCLEAR,
            arrowprops=dict(arrowstyle="->", color=C_NUCLEAR, lw=1.5), ha="center")
ax.axvline(x=10, color=C_NUCLEAR, linestyle=":", alpha=0.5)
ax.set_xticks(range(0, 25, 4))
ax.set_xticklabels(["May'23", "Sep'23", "Jan'24", "May'24", "Sep'24", "Jan'25", "Feb'25"][:7])
ax.set_ylabel("Stock Price ($)", fontsize=10)
ax.set_title("Talen Energy (TLN): From Bankruptcy to $389/share via Data Center Strategy",
             fontweight="bold", fontsize=11)
ax.set_ylim(0, 460)
ax.annotate("Current: $389", xy=(24, 389), xytext=(20, 440),
            fontsize=9, fontweight="bold", color=C_DOWN,
            arrowprops=dict(arrowstyle="->", color=C_DOWN, lw=1.5), ha="center")
plt.tight_layout()
plt.savefig("Output/fig_talen_stock.png", bbox_inches="tight", dpi=200)
plt.close()
print("1/13 fig_talen_stock.png")

# ──────────────────────────────────────────────────
# 2. Option 1 Waterfall
# ──────────────────────────────────────────────────
categories = ["Current\nMkt Cap", "Premium\nPaid", "Synergy\n(Bull)", "Dilution\nDrag", "Integration\nCost", "Net\nValue"]
values      = [20, -2.5, 6, -3.5, -1.5, 18.5]
colors      = [C_NUCLEAR, C_DOWN, C_UP, C_DOWN, C_DOWN, C_NEUTRAL]

fig, ax = plt.subplots(figsize=(10, 4))
cumulative = 0
bottoms = []
for v in values[:-1]:
    bottoms.append(cumulative)
    cumulative += v
bottoms.append(0)

for i, (cat, val, col) in enumerate(zip(categories, values, colors)):
    ax.bar(i, val if i < len(values)-1 else cumulative, bottom=bottoms[i],
           color=col, edgecolor="white", width=0.6)
    y_pos = bottoms[i] + (val/2 if i < len(values)-1 else cumulative/2)
    ax.text(i, y_pos, f"${val:+.1f}B" if i > 0 and i < len(values)-1 else f"${abs(val):.1f}B",
            ha="center", va="center", fontweight="bold", fontsize=9, color="white")

ax.set_xticks(range(len(categories)))
ax.set_xticklabels(categories, fontsize=9)
ax.set_ylabel("Market Cap ($B)", fontsize=10)
ax.set_title("Option 1 (Acquire DC) — Shareholder Value Waterfall ($B)", fontweight="bold", fontsize=11)
ax.set_ylim(0, 30)
plt.tight_layout()
plt.savefig("Output/fig_option1_waterfall.png", bbox_inches="tight", dpi=200)
plt.close()
print("2/13 fig_option1_waterfall.png")

# ──────────────────────────────────────────────────
# 3. Option 2 Premium Analysis
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 4))
premiums = [10, 15, 20, 25, 30, 35, 40]
per_share = [444*p/100 for p in premiums]
total = [444*45e6*p/100/1e9 for p in premiums]

ax2 = ax.twinx()
bars = ax.bar(range(len(premiums)), per_share, color=C_SELL, alpha=0.7, width=0.5)
line = ax2.plot(range(len(premiums)), total, color=C_DOWN, linewidth=2.5, marker="D", markersize=7)

ax.set_xticks(range(len(premiums)))
ax.set_xticklabels([f"{p}%" for p in premiums], fontsize=9)
ax.set_xlabel("Acquisition Premium", fontsize=10)
ax.set_ylabel("Per-Share Premium ($)", fontsize=10, color=C_SELL)
ax2.set_ylabel("Total Deal Size ($B)", fontsize=10, color=C_DOWN)
ax.set_title("Option 2 (Sell to Oil Major) — Premium Analysis", fontweight="bold", fontsize=11)

for i, (ps, t) in enumerate(zip(per_share, total)):
    ax.text(i, ps + 5, f"${ps:.0f}", ha="center", fontsize=8, fontweight="bold", color=C_SELL)

plt.tight_layout()
plt.savefig("Output/fig_option2_premium.png", bbox_inches="tight", dpi=200)
plt.close()
print("3/13 fig_option2_premium.png")

# ──────────────────────────────────────────────────
# 4. Option 3 Revenue Impact
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 4))

years = ["2024\n(Current)", "2025\n(Year 1)", "2026\n(Year 2)", "2027\n(Year 3)", "2028\n(Year 4)", "2029\n(Year 5)"]
merchant = [637, 200, 200, 200, 200, 200]
ppa_rev  = [0,   950, 969, 988, 1008, 1028]

x = np.arange(len(years))
w = 0.35
ax.bar(x - w/2, merchant, w, label="Merchant Revenue", color=C_NEUTRAL, edgecolor="white")
ax.bar(x + w/2, ppa_rev,  w, label="PPA Revenue ($95/MWh)", color=C_NUCLEAR, edgecolor="white")

for i in range(len(years)):
    if merchant[i] > 0:
        ax.text(i - w/2, merchant[i] + 15, f"${merchant[i]}M", ha="center", fontsize=8, fontweight="bold")
    if ppa_rev[i] > 0:
        ax.text(i + w/2, ppa_rev[i] + 15, f"${ppa_rev[i]}M", ha="center", fontsize=8, fontweight="bold", color=C_NUCLEAR)

ax.set_xticks(x)
ax.set_xticklabels(years, fontsize=9)
ax.set_ylabel("Nuclear Revenue ($M)", fontsize=10)
ax.set_title("Option 3 — Nuclear Revenue: Merchant vs PPA", fontweight="bold", fontsize=11)
ax.legend(fontsize=9, loc="upper left")
ax.set_ylim(0, 1200)
plt.tight_layout()
plt.savefig("Output/fig_option3_revenue.png", bbox_inches="tight", dpi=200)
plt.close()
print("4/13 fig_option3_revenue.png")

# ──────────────────────────────────────────────────
# 5. Option 3 FCF Impact
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 4))

years_fcf = ["Current", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"]
fcf_base  = [10.20, 10.20, 10.20, 10.20, 10.20, 10.20]
fcf_option3 = [10.20, 16.87, 17.78, 18.73, 19.72, 20.75]

x = np.arange(len(years_fcf))
ax.plot(x, fcf_base, color=C_NEUTRAL, linewidth=2, linestyle="--", marker="s",
        markersize=6, label="Base (No Action)")
ax.plot(x, fcf_option3, color=C_NUCLEAR, linewidth=3, marker="o",
        markersize=8, label="Option 3 (JV/PPA)")
ax.fill_between(x, fcf_base, fcf_option3, alpha=0.15, color=C_NUCLEAR)

for i in range(len(years_fcf)):
    ax.text(i, fcf_option3[i] + 0.4, f"${fcf_option3[i]:.2f}", ha="center",
            fontsize=9, fontweight="bold", color=C_NUCLEAR)

target_line = 10.20 * 1.30
ax.axhline(y=target_line, color=C_DOWN, linestyle=":", linewidth=1.5, alpha=0.7)
ax.text(5.1, target_line + 0.2, "Dissident Target\n(+30% = $13.26)", fontsize=8,
        color=C_DOWN, fontweight="bold", va="bottom")

ax.set_xticks(x)
ax.set_xticklabels(years_fcf, fontsize=9)
ax.set_ylabel("Adj FCF / Share ($)", fontsize=10)
ax.set_title("Option 3 — Adj FCF/Share Trajectory", fontweight="bold", fontsize=11)
ax.legend(fontsize=9, loc="upper left")
plt.tight_layout()
plt.savefig("Output/fig_option3_fcf.png", bbox_inches="tight", dpi=200)
plt.close()
print("5/13 fig_option3_fcf.png")

# ──────────────────────────────────────────────────
# 6. Tornado / Sensitivity Analysis
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))

params = [
    "PPA Price ($/MWh)",
    "EV/EBITDA Multiple",
    "Nuclear Capacity Factor",
    "WACC",
    "PPA Contract Term",
    "Natural Gas Curve",
]
low_vals  = [-6.2, -4.5, -3.1, 2.4, -2.0, 1.2]
high_vals = [4.8, 3.2, 1.8, -1.8, 1.4, -1.5]
base_npv = 12.0

sorted_indices = sorted(range(len(params)), key=lambda i: abs(high_vals[i] - low_vals[i]))
params = [params[i] for i in sorted_indices]
low_vals = [low_vals[i] for i in sorted_indices]
high_vals = [high_vals[i] for i in sorted_indices]

y_pos = np.arange(len(params))
ax.barh(y_pos, low_vals, height=0.5, color=C_DOWN, alpha=0.8, label="Low Case (Downside)")
ax.barh(y_pos, high_vals, height=0.5, color=C_UP, alpha=0.8, label="High Case (Upside)")

for i in range(len(params)):
    if low_vals[i] != 0:
        ax.text(low_vals[i] - 0.15, i, f"${base_npv+low_vals[i]:.1f}B",
                ha="right", va="center", fontsize=8, fontweight="bold", color=C_DOWN)
    if high_vals[i] != 0:
        ax.text(high_vals[i] + 0.15, i, f"${base_npv+high_vals[i]:.1f}B",
                ha="left", va="center", fontsize=8, fontweight="bold", color="#2E7D32")

ax.axvline(x=0, color=C_NUCLEAR, linewidth=1.5)
ax.set_yticks(y_pos)
ax.set_yticklabels(params, fontsize=9)
ax.set_xlabel("NPV Impact vs Base Case ($B)", fontsize=10)
ax.set_title(f"Tornado Sensitivity — Option 3 NPV (Base = ${base_npv:.1f}B)", fontweight="bold", fontsize=11)
ax.legend(fontsize=9, loc="lower right")
plt.tight_layout()
plt.savefig("Output/fig_tornado_opt3.png", bbox_inches="tight", dpi=200)
plt.close()
print("6/13 fig_tornado_opt3.png")

# ──────────────────────────────────────────────────
# 7. Regulatory Dual-Track
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 4))

tasks = [
    ("Track A: Front-of-Meter PPA", 0, 3, C_UP, "Bilateral PPA — No FERC approval"),
    ("  → PPA Revenue Begins", 3, 20, C_NUCLEAR, "Revenue flows for 20 years"),
    ("Track B: Behind-the-Meter (Backup)", 2, 8, C_NEUTRAL, "NRC 10 CFR 50.90 amendment"),
    ("  → BTM Revenue (if approved)", 8, 20, C_BUILD, "+$10–15/MWh transmission savings"),
    ("DC Campus Development", 3, 12, C_DC, "100–200 MW co-located campus"),
    ("  → Campus Sale to AWS", 12, 14, C_BUY, "$750–900M sale proceeds"),
]

for i, (task, start, end, color, note) in enumerate(tasks):
    ax.barh(i, end - start, left=start, height=0.6, color=color, alpha=0.85, edgecolor="white")
    ax.text(start + (end - start) / 2, i, note, ha="center", va="center", fontsize=7, fontweight="bold", color="white")

ax.set_yticks(range(len(tasks)))
ax.set_yticklabels([t[0] for t in tasks], fontsize=9)
ax.set_xlabel("Months from Deal Signing", fontsize=10)
ax.set_title("Regulatory Dual-Track Strategy — Front-of-Meter (Primary) vs Behind-the-Meter (Backup)",
             fontweight="bold", fontsize=11)
ax.invert_yaxis()
ax.set_xlim(0, 22)
plt.tight_layout()
plt.savefig("Output/fig_regulatory_dual_track.png", bbox_inches="tight", dpi=200)
plt.close()
print("7/13 fig_regulatory_dual_track.png")

# ──────────────────────────────────────────────────
# 8. CCUS Economics
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 4))

phases = ["Phase 1\n(Pilot)\n800 MW", "Phase 2\n(Scale)\n2,000 MW", "Phase 3\n(Full Fleet)\n3,500 MW"]
capex = [400, 900, 1500]
annual_45q = [85, 210, 350]
payback = [4.7, 4.3, 4.3]

x = np.arange(len(phases))
w = 0.3
bars1 = ax.bar(x - w/2, capex, w, label="Capex ($M)", color=C_DOWN, alpha=0.7)
bars2 = ax.bar(x + w/2, annual_45q, w, label="Annual 45Q Credits ($M/yr)", color=C_UP)

for i in range(len(phases)):
    ax.text(i - w/2, capex[i] + 20, f"${capex[i]}M", ha="center", fontsize=9, fontweight="bold", color=C_DOWN)
    ax.text(i + w/2, annual_45q[i] + 20, f"${annual_45q[i]}M/yr", ha="center", fontsize=9, fontweight="bold", color="#2E7D32")

ax.set_xticks(x)
ax.set_xticklabels(phases, fontsize=9)
ax.set_ylabel("Value ($M)", fontsize=10)
ax.set_title("CCUS Phased Deployment — Investment vs Annual Tax Credit Revenue", fontweight="bold", fontsize=11)
ax.legend(fontsize=9)
ax.set_ylim(0, 1800)
plt.tight_layout()
plt.savefig("Output/fig_ccus_economics.png", bbox_inches="tight", dpi=200)
plt.close()
print("8/13 fig_ccus_economics.png")

# ──────────────────────────────────────────────────
# 9. ESG Comparison
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 4))

esg_cats = ["Carbon\nIntensity", "Regulatory\nRisk", "Hyperscaler\nAlignment", "SEC Climate\nReadiness", "Community\nImpact", "Long-term\nSustainability"]
opt1_esg = [3, 2, 4, 3, 3, 3]
opt2_esg = [1, 1, 1, 2, 1, 1]
opt3_esg = [5, 4, 5, 5, 5, 5]

x = np.arange(len(esg_cats))
w = 0.25
ax.bar(x - w, opt1_esg, w, label="Opt 1: Acquire DC", color=C_DC, alpha=0.8)
ax.bar(x, opt2_esg, w, label="Opt 2: Sell to Oil Major", color=C_SELL, alpha=0.6)
ax.bar(x + w, opt3_esg, w, label="Opt 3: JV/PPA with AWS", color=C_NUCLEAR)

ax.set_xticks(x)
ax.set_xticklabels(esg_cats, fontsize=9)
ax.set_ylabel("Score (1–5)", fontsize=10)
ax.set_title("ESG Scorecard — All Three Options", fontweight="bold", fontsize=11)
ax.legend(fontsize=9, loc="upper right")
ax.set_ylim(0, 6)
plt.tight_layout()
plt.savefig("Output/fig_esg_comparison.png", bbox_inches="tight", dpi=200)
plt.close()
print("9/13 fig_esg_comparison.png")

# ──────────────────────────────────────────────────
# 10. Three Options Comparison
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))

criteria = ["FCF/Share\nGrowth\n(25%)", "Multiple\nExpansion\n(20%)", "Execution\nRisk\n(20%)",
            "Shareholder\nValue\n(15%)", "Independence\n(10%)", "Satisfies\nDissident\n(10%)"]
opt1 = [2, 3, 2, 2, 5, 3]
opt2 = [0, 0, 3, 4, 0, 5]
opt3 = [5, 5, 4, 5, 5, 5]

x = np.arange(len(criteria))
w = 0.25
ax.bar(x - w, opt1, w, label="Opt 1: Acquire DC (2.45/5)", color=C_DC, alpha=0.8)
ax.bar(x, opt2, w, label="Opt 2: Sell to Oil (2.15/5)", color=C_SELL, alpha=0.6)
ax.bar(x + w, opt3, w, label="Opt 3: JV/PPA (4.80/5) ★", color=C_NUCLEAR)

ax.set_xticks(x)
ax.set_xticklabels(criteria, fontsize=8)
ax.set_ylabel("Score (0–5)", fontsize=10)
ax.set_title("Weighted Decision Matrix — All Three Options", fontweight="bold", fontsize=11)
ax.legend(fontsize=9, loc="upper right")
ax.set_ylim(0, 6)
plt.tight_layout()
plt.savefig("Output/fig_three_options.png", bbox_inches="tight", dpi=200)
plt.close()
print("10/13 fig_three_options.png")

# ──────────────────────────────────────────────────
# 11. Radar Chart
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 7), subplot_kw=dict(polar=True))

categories_r = ["FCF Growth", "Multiple\nExpansion", "Execution\nFeasibility",
                "Shareholder\nValue", "Independence", "Satisfies\nDissident"]
N = len(categories_r)
angles = [n / float(N) * 2 * np.pi for n in range(N)]
angles += angles[:1]

opt1_r = [2, 3, 2, 2, 5, 3] + [2]
opt2_r = [0, 0, 3, 4, 0, 5] + [0]
opt3_r = [5, 5, 4, 5, 5, 5] + [5]

ax.plot(angles, opt1_r, "o-", linewidth=2, label="Opt 1", color=C_DC)
ax.fill(angles, opt1_r, alpha=0.1, color=C_DC)
ax.plot(angles, opt2_r, "s-", linewidth=2, label="Opt 2", color=C_SELL)
ax.fill(angles, opt2_r, alpha=0.1, color=C_SELL)
ax.plot(angles, opt3_r, "D-", linewidth=2, label="Opt 3 ★", color=C_NUCLEAR)
ax.fill(angles, opt3_r, alpha=0.2, color=C_NUCLEAR)

ax.set_xticks(angles[:-1])
ax.set_xticklabels(categories_r, fontsize=9)
ax.set_ylim(0, 5.5)
ax.set_title("Strategic Fit — Radar Comparison", fontweight="bold", fontsize=12, pad=20)
ax.legend(fontsize=9, loc="upper right", bbox_to_anchor=(1.3, 1.1))
plt.tight_layout()
plt.savefig("Output/fig_radar.png", bbox_inches="tight", dpi=200)
plt.close()
print("11/13 fig_radar.png")

# ──────────────────────────────────────────────────
# 12. Shareholder Value Trajectory
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 4))

years_sv = ["Current", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"]
opt1_sv  = [20.0, 18.0, 20.0, 24.0, 28.0, 32.0]
opt2_sv  = [24.0, 24.0, 24.0, 24.0, 24.0, 24.0]
opt3_sv  = [20.0, 32.0, 38.0, 44.0, 48.0, 52.0]

x = np.arange(len(years_sv))
ax.plot(x, opt1_sv, "o-", linewidth=2.5, label="Opt 1: Acquire DC", color=C_DC, markersize=7)
ax.plot(x, opt2_sv, "s--", linewidth=2, label="Opt 2: Sell (exit at $24B)", color=C_SELL, markersize=7)
ax.plot(x, opt3_sv, "D-", linewidth=3, label="Opt 3: JV/PPA with AWS ★", color=C_NUCLEAR, markersize=8)

for i in range(len(years_sv)):
    ax.text(i, opt3_sv[i] + 1.0, f"${opt3_sv[i]:.0f}B", ha="center", fontsize=9,
            fontweight="bold", color=C_NUCLEAR)

ax.axhline(y=20, color=C_NEUTRAL, linestyle=":", linewidth=1, alpha=0.5)
ax.text(5.1, 20.5, "Current: $20B", fontsize=8, color=C_NEUTRAL)

ax.set_xticks(x)
ax.set_xticklabels(years_sv, fontsize=9)
ax.set_ylabel("Market Cap ($B)", fontsize=10)
ax.set_title("5-Year Shareholder Value Trajectory — All Options", fontweight="bold", fontsize=11)
ax.legend(fontsize=9, loc="upper left")
ax.set_ylim(10, 60)
plt.tight_layout()
plt.savefig("Output/fig_shareholder_value.png", bbox_inches="tight", dpi=200)
plt.close()
print("12/13 fig_shareholder_value.png")

# ──────────────────────────────────────────────────
# 13. IPP Re-Rate
# ──────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 4))

companies = ["VST\n(Vistra)", "CEG\n(Constellation)", "TLN\n(Talen)", "NRG\n(NRG Energy)", "Our Co\n(Current)", "Our Co\n(Post-Deal)"]
ev_ebitda = [18, 22, 28, 8, 30, 38]
colors_bar = [C_NEUTRAL, C_BUILD, C_UP, C_NEUTRAL, C_DOWN, C_NUCLEAR]

bars = ax.bar(range(len(companies)), ev_ebitda, color=colors_bar, edgecolor="white", width=0.6)
for i, v in enumerate(ev_ebitda):
    ax.text(i, v + 0.5, f"{v}x", ha="center", fontsize=11, fontweight="bold",
            color=colors_bar[i])

ax.set_xticks(range(len(companies)))
ax.set_xticklabels(companies, fontsize=9)
ax.set_ylabel("EV / EBITDA", fontsize=10)
ax.set_title("IPP Sector Re-Rating — EV/EBITDA Multiples", fontweight="bold", fontsize=11)
ax.set_ylim(0, 45)
ax.axhline(y=30, color=C_NEUTRAL, linestyle=":", alpha=0.5)
plt.tight_layout()
plt.savefig("Output/fig_ipp_rerate.png", bbox_inches="tight", dpi=200)
plt.close()
print("13/13 fig_ipp_rerate.png")

print("\n✓ All 13 charts exported as PNGs to Output/")
