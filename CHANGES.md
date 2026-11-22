# ExcelInsight: fixes and UI refresh (22 Nov 2026)

Everything else works as before. Deploy the frontend (Vercel) and the backend (Render) to get all fixes live.
The Gemini fixes only take effect after the backend in `server/` is redeployed.

## Bugs fixed

### Calculations
- **Averages with blank cells.** `AVERAGE`, Category Summary, Chart Studio, Pivot and the Outliers mean now skip empty cells instead of counting them as 0. Example: average profit on the test file was 3,796.58, now 4,141.73.
- **Formula bar** rewritten. Supports `SUM, AVERAGE, MEDIAN, MIN, MAX, COUNT, COUNTA, COUNTBLANK, STDEV` and arithmetic between them (`=SUM(Profit)/SUM(Sales)*100`). Unknown columns, unknown functions and bare column names now give a clear error instead of `0` or `1`.
- **KPI cards** pick the meaningful column (Sales, Revenue, Profit...) instead of the first numeric one.
- Blank cells show as a light dash instead of the word `null`; blanks sort last.

### Sheet Diff
- Rows are matched on a **key column** (auto-detected, e.g. OrderID) instead of row position. Deleting one row no longer marks every row below it as modified. Repeated keys pair in order. "Show only changes" toggle added.

### Data Hygiene
- **Standardize text** now merges case variants ("saket", " Saket ") into the most common spelling, not just trims.
- Health score, missing, duplicate and text-issue counts update after every action.
- **Smart pipeline** order fixed: drop empty rows → fix text → dedupe → fill gaps (ID columns left alone, so no fake records) → cap only extreme values (3 × IQR). The toast reports exactly what it did.

### Outliers
- Blank cells are skipped (they were flagged as `0`).
- Isolation Forest is now seeded (same score every run) and the contamination setting actually changes the cut-off.
- Ensemble flags a value when 2 of 3 methods agree. Z-score threshold is 2.5 everywhere.

### AutoML Studio (rewritten engine in `app.js`)
- Real models: OLS, Ridge, Random Forest, Gradient Boosted Trees; Logistic Regression (softmax), Random Forest, k-NN; Naive, Moving Average, Linear Trend, Holt for forecasting. Each task has a baseline.
- **Evaluation:** 5-fold cross-validation under 400 rows, 80/20 holdout above. Rows with no target value are skipped.
- **Ranking metric** you choose is now the one used for ranking, the KPI and the chart.
- **Metrics** follow scikit-learn definitions (macro F1 with zero-division = 0, R² can be negative). The impossible baseline scores are gone.
- **Prediction panel** names the model that actually made the prediction. Classification shows a real class probability; regression shows a range from cross-validated RMSE, warns on inputs outside the training range and never shows negative values for a target that is never negative.
- ID and date columns start unticked as features.
- Forecast mode hides the feature inputs, shows the next period, and the CSV export comes from the winning model.
- Feature importance is permutation importance on the winning model, grouped by original column.
- Audit JSON includes the full leaderboard, metrics, evaluation method and seed.

### AI (Gemini) — `server/routes/ai.js`
- Answers were cut off because `maxOutputTokens` (800/1024) was used up by the model's internal thinking. Now 8192 (set `GEMINI_MAX_TOKENS` to change), and all text parts are joined.
- Chat asks Gemini for JSON mode and recovers the answer even if JSON is incomplete; no more raw `{ "answer": ...` in the chat.
- The browser sends exact group totals for all rows, so "which outlet sold most" uses real totals, not a 15-row sample.
- Insights and chat render markdown properly. User questions are HTML-escaped.

### Other
- PDF report works from any tab (chart is drawn off-screen) and includes column summary, AutoML results and the AI summary.
- JSON export renamed from `SheetPulse_...` to `ExcelInsight_...`.
- Admin tab can only be opened by admin accounts.
- CDN libraries pinned to exact versions; missing icon warning fixed.
- `/api/ml/train` stores the browser's results for signed-in users instead of retraining with a different engine.

## UI refresh
White canvas, black headings, charcoal body text and one blue accent (#0d44e8). Google Sans Flex for UI text (Inter fallback), 14px controls, 20px cards, lighter header, calmer AI and AutoML panels, cleaner chat drawer and KPI tiles.

## Files changed
`index.html`, `app.js`, `style.css`, `server/routes/ai.js`, `server/routes/ml.js`
