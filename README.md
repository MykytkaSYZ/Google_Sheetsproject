# Google Sheet: Financial Dashboard | [Table](https://docs.google.com/spreadsheets/d/1QhSVeRENaQNoAtI3IXSgDy67G3ATw5NwYSuCex7a0y0/edit?gid=759494513#gid=759494513)

This project is an interactive financial analytics dashboard built entirely in Google Sheets, based on real-world data from 12 major global companies over the period 2009–2023.

# 📦 File Structure (Google Sheets)

- Financial Statements — Raw financial data
- Pivot Tables - multiple tables
- Dashboard — Visual summaries, slicers, and charts
- Insights — Automatically generated statements

# :file_folder: Financial Statements Data Overview 

The dataset includes the following :
- Market Cap 
- Revenue, Net Income, Gross Profit, EBITDA
- Key Ratios: ROE, ROI, ROA, Net Profit Margin, Debt/Equity, Current Ratio
- Cash Flows: Operating, Investing, Financial
- Earnings Per Share, Free Cash Flow per Share
- Number of Employees, Inflation Rate

# :bar_chart: Features of the Dashboard 

## 1. Interactive Slicers

Filter visualizations and metrics by Company, Year, Category and Quality.

## 2. Pivot Tables

- Revenue trends by company/year
- ROE, ROI, Net Profit Margin per company over time
- Average financial ratios per industry
- Quality by category
- Earning Per Share by company/year
  
## 3. Visualizations

- Area chart: Cash Flow
- Bar charts: Market Cap, Number of Employees
- Scatter plot: Revenue
- Column chart: Comparison of key metrics

## 4. Custom Quality Score

Implemented custom logic with nested IF() formulas based on: ROE, Net Profit Margin, Debt/Equity, Positive Free Cash Flow

This scoring classifies companies as: ✅ Excellent 👍 Good ⚠️ Average ❌ Poor

`IF(AND(S2>15,T2>0,P2>30,O2<1.5),"✅ Excellent",IF(AND(S2>10,T2>0,P2>20,O2<2),"👍 Good",IF(AND(S2>5,T2>=0,P2>10,O2<3),"⚠️ Average","❌ Poor")))`

# 🤖 Google Apps Script: "Insights"

A script generates automated insights with a click of a button:

- Highlights the top company for a selected metric and year
- Compares its value to the market average

Example: "AAPL (IT) had a EBITDA (in M USD) of 81801.00 in 2018 — 4.1x the market average (20058.97)."

```js
function generateInsights() {
  const sheetName = "Insights";
  const dataSheetName = "Financial Statements";  // замените на имя вашего основного листа с данными

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const dataSheet = ss.getSheetByName(dataSheetName);

  const metric = sheet.getRange("B4").getValue();
  const year = sheet.getRange("B5").getValue();

  const headers = dataSheet.getDataRange().getValues()[0];
  const data = dataSheet.getDataRange().getValues().slice(1);

  const metricIndex = headers.indexOf(metric);
  const yearIndex = headers.indexOf("Year");
  const companyIndex = headers.indexOf("Company");
  const categoryIndex = headers.indexOf("Category");

  if (metricIndex === -1 || yearIndex === -1 || companyIndex === -1) {
    sheet.getRange("A8").setValue("❌ Metric or Year column not found.");
    return;
  }

  // Filter by year and clean metric
  const filtered = data.filter(row => row[yearIndex] == year && typeof row[metricIndex] === "number" && !isNaN(row[metricIndex]));

  if (filtered.length === 0) {
    sheet.getRange("A8").setValue("⚠️ No data available for this year and metric.");
    return;
  }

  // Calculate average
  const total = filtered.reduce((sum, row) => sum + row[metricIndex], 0);
  const avg = total / filtered.length;

  // Sort and take top 3
  const sorted = filtered.slice().sort((a, b) => b[metricIndex] - a[metricIndex]);
  const top3 = sorted.slice(0, 3);

  // Generate insights
  const insights = top3.map(row => {
    const company = row[companyIndex];
    const value = row[metricIndex];
    const ratio = (value / avg).toFixed(1);
    return `📊 ${company} (${row[categoryIndex]}) had a ${metric} of ${value.toFixed(2)} in ${year} — ${ratio}x the market average (${avg.toFixed(2)}).`;
  });

  // Output
  sheet.getRange("A8:A").clearContent();
  sheet.getRange(8, 1, insights.length).setValues(insights.map(i => [i]));
}
```
## 🚀 Try It Yourself

[Table(view only)](https://docs.google.com/spreadsheets/d/1QhSVeRENaQNoAtI3IXSgDy67G3ATw5NwYSuCex7a0y0/edit?gid=759494513#gid=759494513)

Or clone the repository and test the Apps Script locally.

# :facepunch: Conclusion

Thank you for reviewing this project!
I hope this dashboard clearly demonstrates how financial data can be translated into valuable business insights through effective analysis and visualization.
Feedback is welcome, and I'm always open to improving the work further.




