# GA Dashboard PPT Change Tracker

Source PPT: `/Users/vbond/Documents/GA_대시보드_1차수정.pptx`
Updated: 2026-04-01

## Slide-by-slide

1. Slide 1 - Key Insights / Current Month Status terminology
- Status: Partially applied
- Applied:
  - Main KPI labels in Key Insights now keep GA terms in English: `TOTAL SESSIONS`, `TOTAL CLICKS`, `AVG PLP CONV.`, `AVG PURCHASE INTENT CONV.`
  - AI labels also keep GA terms in English.
- Pending:
  - Additional no-translation cleanup outside the main hero if needed.

2. Slide 2 - Remove standalone MOM CHANGE and show MoM under each KPI
- Status: Applied
- Applied:
  - Removed standalone `MoM Change` hero KPI.
  - Added per-KPI MoM under each main KPI card in Key Insights.
  - Replaced purchase KPI in Key Insights with `AVG PURCHASE INTENT CONV.`
- Pending:
  - `Pages Tracked` definition still needs business confirmation.

3. Slide 3 - Country Status Classification should support rising pages
- Status: Applied
- Applied:
  - Growing / Stable / Declining groups already render and remain active.

4. Slide 4 - Country Cards metric order and duration rounding
- Status: Applied
- Applied:
  - Country card metric order aligned around sessions, PLP, duration, purchase.
  - Duration display now floors decimals to whole seconds across key card views.

5. Slide 5 - Fold/unfold exposure rules in Country Cards
- Status: Partially applied
- Applied:
  - Detailed product breakdown remains in expanded area only.
- Pending:
  - Confirm whether any additional top-area content should be hidden until fold open.

6. Slide 6 - Show page links beside page title with link icon
- Status: Applied
- Applied:
  - Product/category headers in Country Cards now show inline link icons.
  - Separate link context is now attached to the relevant category title instead of only separate chips.

7. Slide 7 - `Top` category logic in Country Cards
- Status: Applied
- Applied:
  - `Top` now uses the highest-session category when `All Categories` is selected.
  - When a specific category is selected, the selected category is shown.

8. Slide 8 - Category Compare should show all countries
- Status: Applied
- Applied:
  - Category ranking tables now compute from all countries, not a country-filtered subset.
  - Canada EN / FR remain separate entries in rankings.

9. Slide 9 - Selecting a country should preserve actual ranking
- Status: Applied
- Applied:
  - Category Compare rankings are now computed from the global country set.
  - Selected country is highlighted instead of being artificially moved to rank 1.

10. Slide 10 - Raw Data wording should change by content type
- Status: Applied
- Applied:
  - Overview KPI labels now change between `Lineup Guide` and `Feature Library`.
  - Sessions-by-country chart title also changes by selected content type.

11. Slide 11 - Hover tooltip on Raw Data / Overview line chart
- Status: Applied
- Applied:
  - Added explicit chart interaction/tooltip defaults for line charts.

12. Slide 12 - Show percentage inside Traffic Source Breakdown
- Status: Applied
- Applied:
  - Added doughnut percentage labels inside the chart slices.

13. Slide 13 - Exit rate / Session Trend mapping
- Status: Pending review
- Notes:
  - Current implementation renders both charts, but workbook mapping still needs business-side validation with live data.

14. Slide 14 - Internal vs External by Country should respect category selection
- Status: Pending verification
- Notes:
  - Current render path already respects selected category through filtered sheets.
  - Needs visual QA against expected workbook output for Monitor-only cases.

## Validation completed

- `node --check` on inline dashboard script
- `python3 -m py_compile` on data build scripts

## Next recommended pass

- Finish slide 5 exact fold/unfold rule after visual confirmation.
- Validate slide 13 and 14 with workbook examples and screenshots.
- Apply any remaining no-translation text cleanup from the PPT reference images.
