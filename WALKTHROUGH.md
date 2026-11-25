# Walkthrough: from data to dashboard

This is how the tracker was built, in the order you would do it yourself.

## Step 1. Prepare the data
- One row per transmittal. One fact per column: Trans. No, Doc. No, Title, Discipline, Date Sent, Rev, Review days, Date Reply, Reply status.
- Dates must be real Excel dates, not text. Do not trust the alignment: test a cell with `=ISNUMBER(E2)` (TRUE for a real date), or set the cell format to General and see whether a number such as 46000 appears. The Register also has a date validation that rejects text.
- Keep lists (disciplines, reply statuses) on a Settings sheet and use Data > Data Validation > List so people pick values instead of typing them. This keeps counts correct.
- Freeze the header row (View > Freeze Panes) and add a filter (Data > Filter).

## Step 2. Add the calculated columns
1. Review days (G): `=Settings!$B$2`. The period lives in one cell, so you change it once.
2. Days waiting (J): `=IF(E2="","",IF(H2="",TODAY(),H2)-E2)`. If there is no reply date, count to today. Otherwise count to the reply date.
3. Flag (K): `=IF(E2="","",IF(NOT(ISNUMBER(J2)),"Check date",IF(J2<0,"Check date",IF(H2<>"","Closed",IF(J2>G2,"Overdue","Open")))))`. The first two tests catch text dates and a reply dated before the send date.
4. Highlight overdue rows: select the table, Home > Conditional Formatting > New Rule > Use a formula: `=$K2="Overdue"`, choose a red fill.

## Step 3. Build the metrics on the Dashboard sheet
The cell references below (B46, D58, F5, D5) are the ones in this workbook. Always refer to the Register ranges (for example `Register!$K$2:$K$500`) so new rows are counted.
| Metric | Formula idea |
|---|---|
| Total transmittals (rows with a Date sent) | `=SUMPRODUCT(--(Register!$E$2:$E$500<>""))` |
| Overdue | `=COUNTIF(Register!$K$2:$K$500,"Overdue")` |
| Waiting for reply | `=COUNTIF(Register!$K$2:$K$500,"Open")+COUNTIF(Register!$K$2:$K$500,"Overdue")` |
| Overdue % of waiting | `=IFERROR(F5/D5,0)` (overdue divided by waiting) |
| Average days to reply | `=IFERROR(ROUND(AVERAGEIF(Register!$K$2:$K$500,"Closed",Register!$J$2:$J$500),1),0)` |
| Longest wait still open (Open and Overdue rows) | `=IFERROR(MAX(MAXIFS(Register!$J$2:$J$500,Register!$K$2:$K$500,"Open"),MAXIFS(Register!$J$2:$J$500,Register!$K$2:$K$500,"Overdue")),0)` (Excel 2019 / 365) |
| Status by discipline | `=COUNTIFS(Register!$D$2:$D$500,$B46,Register!$K$2:$K$500,"Open")` (same for Closed and Overdue) |
| Average reply days by discipline | `=IFERROR(ROUND(AVERAGEIFS(Register!$J$2:$J$500,Register!$D$2:$D$500,$B46,Register!$K$2:$K$500,"Closed"),1),0)` |
| Ageing of open items (8 to 14 days) | `=COUNTIFS(Register!$H$2:$H$500,"",Register!$E$2:$E$500,"<>",Register!$K$2:$K$500,"<>Check date",Register!$J$2:$J$500,">=8",Register!$J$2:$J$500,"<=14")` |
| Sent per month | `=COUNTIFS(Register!$E$2:$E$500,">="&D58,Register!$E$2:$E$500,"<"&EDATE(D58,1))`, with the month start in D58 as `=EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),-3)` |

## Step 4. Make the charts
1. Select a small table (for example Discipline with Open, Overdue, Closed).
2. Insert > Chart > Stacked Column.
3. Give it a clear title. Remove the legend when there is only one series.
4. Place charts in a grid with equal sizes, and use the same few colours (here navy, gold, slate, red).
Charts read the dashboard tables, so they update when the Register changes.

## Step 5. Check your work
- Pick three rows and count the days by hand. Compare with column J.
- Total = Open + Overdue + Closed + Check date. A row with no Date sent is not counted at all (for example a row with only a title).
- A date typed as text makes Days waiting show `#VALUE!` and the Flag show `Check date`. Fix the date, do not hide the message.
- Add a test row with no reply date and an old sent date: it must show Overdue.
- Clear the dummy data (columns A to F, H and I only, never G, J, K). Every dashboard number must show 0 and no errors.
- Make every row open and recent: Overdue must be 0 and Longest open wait must still show the oldest open row.

## Limits
Dummy data. `TODAY()` changes the numbers daily. Rows after 500 are not counted until you extend the ranges. Add a bad date on purpose and confirm the row shows "Check date".
