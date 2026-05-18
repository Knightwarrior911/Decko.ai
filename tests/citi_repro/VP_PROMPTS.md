# Final VP prompts — Citi Q1'26 slide recreation

These are the one-paragraph natural-language prompts a banker types into Decko.
Each maps to the exact action JSON listed (driven via `ExecuteFromString`).
Source data: Citi Q1'26 earnings PDF, pages 6 & 7.

---

## Slide 1 — "Financial results overview"

> Build a 16:9 slide titled "Financial results overview" in Citi blue with a
> rule under it. On the left, a "Financial Results" table ($ in MM, except EPS)
> with a blue banner header and columns 1Q26 / %ΔQoQ / %ΔYoY for: Net Interest
> Income 15,741 / - / 12%; Non-Interest Revenue 8,892 / 111% / 17%; **Total
> Revenues 24,633 / 24% / 14%**; Expenses 14,311 / 3% / 7%; NCLs 2,208 / 1% /
> (10)%; ACL Build and Other(1) 597 / NM / 126%; Provision for Credit Losses
> 2,805 / 26% / 3%; **EBT 7,517 / 97% / 38%**; Income Taxes 1,578 / 23% / 18%;
> **Net Income 5,785 / 134% / 42%**; **Net Income to Common(2) 5,442 / 151% /
> 44%**; **Diluted EPS $3.06 / 157% / 56%**; Efficiency Ratio 58.1% / (1,150) /
> (410); ROE 11.5%; RoTCE(d) 13.1% / 800 / 400; CET1 Capital Ratio(c) 12.7%;
> then a Memo: NII ex-Markets 12,944 / - / 7%; NIR ex-Markets 4,443 / 88% / 29%.
> Bold the subtotal rows with a light-blue band; shrink the font so all rows
> fit; faint row rules. Top-right: a "1Q26 Financial Overview Highlights"
> heading with bullets on Revenues +14% YoY (NII +12%, NII ex-Markets +7%, NIR
> +17%, NIR ex-Markets +29%), Expenses +7% YoY, Provision $2.8B incl. $597MM ACL
> build, RoTCE 13.1%. Bottom-right: a "Revenue by Segment" blue banner over a
> real stacked column chart ($ in B) for 1Q25/4Q25/1Q26 with Services
> 5.2/6.3/6.1, Markets 6.1/4.6/7.2, Banking 2.8/2.9/3.1, Wealth 1.5/1.8/1.8,
> U.S. Consumer Cards 4.6/4.6/4.8, All Other -? (1.5/-0.2/1.7), totals
> 21.6/19.9/24.6 labeled above the bars, legend at the bottom, value axis
> hidden.

Action JSON: `tests/citi_repro/slide1.actions.json` (436 actions, 0 non-ok).
Builder: `tests/citi_repro/build_slide1.py`.

---

## Slide 2 — "Quarterly expense trend and year-over-year expense drivers"

> Build a 16:9 slide titled "Quarterly expense trend and year-over-year expense
> drivers" in Citi blue. Left side: an "Expense Overview" blue banner over a
> real combo chart ($ in B) for 1Q25–1Q26 — stacked columns for Compensation
> and Benefits & Restructuring 7.5/7.6/7.5/7.1/8.4, Transactional and Product
> Servicing 1.1/1.2/1.1/1.2/1.2, Technology/Communication 2.4/2.3/2.3/2.4/2.3,
> Other Expenses ex-notable item 2.5/2.5/2.7/3.2/2.4, Goodwill Impairment Charge
> only 0.7 in 3Q25 — with Reported Expenses totals $13.4/$13.6/$14.3/$13.8/$14.3
> labeled above each bar, and a Reported Efficiency Ratio line on a secondary
> axis at 62.2%/62.7%/64.7%/69.6%/58.1%. Beneath it a small table: Direct Staff
> (thousands) 229/230/227/226/224 and Severance(3) ($B) 0.1/0.4/0.2/0.1/0.5.
> Right side: a "1Q26 Expense Drivers" heading and four rounded callout cards —
> Transactional and Product Servicing (Up 11% YoY): higher volumes in Markets,
> Services, U.S. Consumer Cards and Banking; Technology/Communication (Down (2)%
> YoY): fewer technology contractors, largely offset by technology charges and
> continued investment; Other Expenses ex-notable item (Down (5)% YoY): lower
> legal and professional fees, partially offset by higher tax and deposit
> insurance costs; Compensation and Benefits & Restructuring (Up 12% YoY):
> higher severance, comp from Banking/Services investment and performance,
> partially offset by productivity savings and lower transformation expense.

Action JSON: `tests/citi_repro/slide2.actions.json` (72 actions, 0 non-ok).
Builder: `tests/citi_repro/build_slide2.py`.

---

## Reproducibility

Combined 2-slide deck regenerated from the above:
`python tests/citi_repro/build_combined.py` →
`python tests/citi_repro/run_citi.py tests/citi_repro/citi_final.actions.json --deck citi_final`
→ **508/508 actions, 0 non-ok**, slide 1 + slide 2 each = 1 real ChartObject +
1 real table, zero autoshape-as-chart.
