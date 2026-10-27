Worked for 3m 38s

I’ve inspected your actual Excel file and checked the email matches against the employee list. The dashboard should use EMA Assurance Employee List as the master employee list, use trimmed email addresses to match attendees, and calculate each person's duration separately for Option 1, Option 2, and both meetings combined.

### 1\. Attendance rules for your dashboard

* Fully attended: 1 hour 30 minutes or more (90 minutes).

* Partially attended: More than 0 minutes, but less than 90 minutes.

* Did not attend: The employee does not appear in that meeting’s Participants section with a matching email, or has zero attendance.

* Combined: Add Option 1 and Option 2 durations together, then apply the same 90-minute threshold to the total.

One important distinction: someone can be classed as fully attended in the combined results without having fully attended either meeting individually. The combined classification measures their total time across the two meetings.

I’ll use the seven requested countries—Kenya (KE), Uganda (UG), Mauritius (MU), Tanzania (TZ), Rwanda (RW), South Africa (ZA), and Zambia (ZM)—and exclude the employee whose country code is `#N/A`.

### 2\. Download your completed Excel dashboard

EMA Assurance Attendance Dashboard.xlsx

Your original workbook plus the dashboard, employee analysis, country summaries and matching audit.

Download the completed Excel dashboard

The workbook includes your three original sheets and six new sheets:

| Sheet               | What it does                                                              |
| ------------------- | ------------------------------------------------------------------------- |
| Dashboard           | KPI cards, country breakdown and comparison charts                        |
| Employee Attendance | Each employee’s three durations and attendance statuses; filterable table |
| Country Summary     | Country-by-country totals for all three meeting views                     |
| Pivot Data          | Prepared data for PivotTables, slicers and interactive charts             |
| Data Quality        | Matching checks and a list of unmatched report participants               |
| Read Me             | Definitions, methodology and notes for using the workbook                 |

Important: This is a checked analysis of the file you uploaded. The KPI cards and charts are based on that snapshot; changes to your original sheets will not automatically recalculate the new analysis sheets. The instructions below explain how to build a repeatable Excel version and add interactive slicers.

### 3\. Check the results first

I found 446 employees in your master list. One has a country code of `#N/A`, so 445 employees are included in the seven-country dashboard.

Here are the verified overall counts:

| Meeting  | Fully attended | Partially attended | Did not attend | Total |
| -------- | -------------- | ------------------ | -------------- | ----- |
| Option 1 | 151            | 56                 | 238            | 445   |
| Option 2 | 188            | 68                 | 189            | 445   |
| Combined | 330            | 74                 | 41             | 445   |

The full country breakdown is already in the Dashboard sheet.

A few data-quality findings are worth knowing:

* Option 1: 207 employees matched the master list by email. There is one attendee named Wilson Judah (Unverified) whose email is blank, so that record cannot be matched using your email-primary-key rule.

* Option 2: 256 employees matched the master list by email.

* Some participants in both meeting reports are not present in your EMA Assurance Employee List. They are listed on the Data Quality sheet but are not counted in the country dashboard.

This avoids incorrectly counting everyone in a meeting report as an EMA Assurance employee.

### 4\. Step-by-step: How to create the dashboard in Excel

The following instructions assume you are using desktop Excel. The formulas use modern Excel functions, including `XLOOKUP` and `TEXTSPLIT`.

### Step 1: Create a safe working copy

* Open the original workbook.

* Click File → Save As.

* Save a copy with a name such as `EMA Assurance Attendance Dashboard Working.xlsx`.

* Keep your original workbook unchanged, so you always have the raw reports available.

We will use only the Participants section in each attendance report, not the In-Meeting Activities or Meeting Engagement sections.

### Step 2: Turn the Participants sections into Excel Tables

Tables make formulas easier to use and provide filter arrows automatically.

### Option 1

* Open `Attendance Report - Option 1`.

* Select cells `A10:O315`. Row 10 contains the column headings, and rows 11–315 contain the participant records.

* Press Ctrl + T.

* Make sure My table has headers is ticked.

* Click OK.

* Click any cell in the new table.

* Open Table Design on the ribbon.

* Find the Table Name box and rename the table to `tblOpt1`.

### Option 2

Repeat the same steps on `Attendance Report - Option 2`, selecting `A10:O412`, and name the table `tblOpt2`.

Be careful not to include the In-Meeting Activities section below either table.

These are the starting ranges for the report you supplied. If a future report has more or fewer participants, resize each table so it ends at the final participant row immediately before Section 3.

### Step 3: Clean the emails and convert durations into seconds

This step is essential. It prevents extra spaces in email addresses from causing false non-attendance results, and makes the different duration formats comparable.

### 3A. Add two helper columns to each Participants table

In `Attendance Report - Option 1`, add two columns directly to the right of the table:

* Column P: `Clean Email`

* Column Q: `Duration Seconds`

Do the same in Option 2.

If Excel does not automatically extend the table, use Table Design → Resize Table to include the two new columns.

### 3B. Clean the email addresses

In the first data cell under `Clean Email`, enter:

For example, in P11 on the Option 1 sheet:

=LOWER(TRIM(\[@Email\]))

Copy formula

Press Enter. Excel should fill the formula down the whole table.

This formula removes leading and trailing spaces and converts emails to lowercase. Repeat it in the Option 2 table.

### 3C. Convert the duration into seconds

Your attendance data contains values such as `2h 22m 42s`, `9m 37s`, `2h 3m` and `6s`. You must convert them into one comparable number.

In the first data cell under `Duration Seconds`, enter this formula:

=LET(p,TEXTSPLIT(TRIM(\[@\[In-Meeting Duration\]\])," "),SUM(IFERROR(--TEXTBEFORE(p,"h"),0))\*3600+SUM(IFERROR(--TEXTBEFORE(p,"m"),0))\*60+SUM(IFERROR(--TEXTBEFORE(p,"s"),0)))

Copy formula

Press Enter and let Excel fill it down the table. Repeat for Option 2.

The formula recognizes the hour, minute and second components independently. For example, `2h 22m 42s` becomes 8,562 seconds, while `9m 37s` becomes 577 seconds.

Why seconds? Comparing numbers is much more reliable than comparing text such as `1h 30m` and `90m`.

### Step 4: Create the employee attendance sheet

* Click the + at the bottom of Excel to add a new worksheet.

* Rename it `Employee Attendance`.

* Create these headings in row 1:

| Column | Heading           |
| ------ | ----------------- |
| A      | Employee Name     |
| B      | Country Code      |
| C      | Email             |
| D      | Normalized Email  |
| E      | Country           |
| F      | Option 1 Seconds  |
| G      | Option 1 Duration |
| H      | Option 1 Status   |
| I      | Option 2 Seconds  |
| J      | Option 2 Duration |
| K      | Option 2 Status   |
| L      | Combined Seconds  |
| M      | Combined Duration |
| N      | Combined Status   |

* Copy the employee names, country codes and email addresses from `EMA Assurance Employee List` into columns A–C, starting in row 2.

* Select the populated range and press Ctrl + T.

* Tick My table has headers and click OK.

* Name the table `tblAttendance` using Table Design.

This sheet will become the central list from which your dashboard is built.

### Step 5: Add formulas to classify the employees

Enter these formulas into the first data row of their respective table columns. Excel Tables should fill the formulas down automatically.

### A. Normalize the employee email

In `D2`:

=LOWER(TRIM(\[@Email\]))

Copy

### B. Convert country codes into country names

In `E2`:

=IFERROR(SWITCH(UPPER(TRIM(\[@\[Country Code\]\])),"KE","Kenya","UG","Uganda","MU","Mauritius","TZ","Tanzania","RW","Rwanda","ZA","South Africa","ZM","Zambia",""),"")

Copy

The formula returns a blank for `#N/A` or an unsupported country code. This means those rows can be excluded from the country dashboard.

### C. Retrieve Option 1 duration

In `F2`:

=IFERROR(XLOOKUP(\[@\[Normalized Email\]\],tblOpt1\[Clean Email\],tblOpt1\[Duration Seconds\],0),0)

Copy

This searches for the employee's normalized email in the Option 1 Participants table. If there is no matching email, it returns zero.

In `H2`, classify the duration:

=IF(F2=0,"Did not attend",IF(F2>=5400,"Fully attended","Partially attended"))

Copy

Remember, 90 minutes × 60 seconds = 5,400 seconds.

### D. Retrieve Option 2 duration

In `I2`:

=IFERROR(XLOOKUP(\[@\[Normalized Email\]\],tblOpt2\[Clean Email\],tblOpt2\[Duration Seconds\],0),0)

Copy

In `K2`:

=IF(I2=0,"Did not attend",IF(I2>=5400,"Fully attended","Partially attended"))

Copy

### E. Calculate the combined duration

In `L2`:

=F2+I2

Copy

In `N2`:

=IF(L2=0,"Did not attend",IF(L2>=5400,"Fully attended","Partially attended"))

Copy

The combined classification applies the 90-minute threshold to the sum of the two meetings, not separately to each one.

### F. Display durations in readable time format

In `G2`, enter:

`=F2/86400`

In `J2`, enter:

`=I2/86400`

In `M2`, enter:

`=L2/86400`

Select columns G, J and M. Press Ctrl + 1 → Number → Custom, and enter this format:

`[h]"h "m"m "s"s"`

Excel will then display durations as hours, minutes and seconds instead of fractions of a day.

Important for future updates: The formulas can fill down when rows are added to `tblAttendance`. When the master employee list grows, add the new employees to this analysis table as well. When the attendance report changes, extend the Option 1 or Option 2 table to cover the new Participants rows, without including Section 3.

### Step 6: Create the country summary using PivotTables

A PivotTable is an Excel tool that groups your data and calculates counts without needing to write dozens of formulas.

For the most flexible version, use the Pivot Data sheet in the downloaded workbook. It already has one row per employee per meeting view, making it suitable for interactive filters.

* Open `Pivot Data`.

* Click any cell inside the table.

* On the ribbon, select Insert → PivotTable.

* Choose New Worksheet, then click OK.

* In the PivotTable Fields panel, drag the fields as follows:

Rows

Country, then Meeting Type

This groups employees by country and displays Option 1, Option 2 and Combined under each country.

Columns

Attendance Status

This creates separate counts for fully attended, partially attended and did not attend.

Values

Employee Name

Change Value Field Settings to Count, not Sum.

You should now see a country summary with three attendance classifications for each meeting view.

### Step 7: Add real slicers and interactive charts

Slicers are clickable buttons that filter a PivotTable. The workbook already has filter arrows on its tables; use these steps to add actual slicers.

* Click inside the PivotTable you created.

* Go to PivotTable Analyze → Insert Slicer.

* Select `Country`, `Meeting Type` and `Attendance Status`.

* Click OK.

* Move the slicers onto your Dashboard sheet and resize them to fit neatly beside the charts.

* Click the PivotTable, then choose PivotTable Analyze → PivotChart.

* Choose a Stacked Bar or Stacked Column chart.

* Move the chart onto the Dashboard sheet.

* Test the slicers by clicking a country or meeting type. The PivotTable and its chart should respond.

The chart you create from this PivotTable will be interactive. The prebuilt dashboard charts at the top of the downloaded workbook remain snapshot charts; the new PivotChart will respond to your slicers.

### Step 8: Create KPI cards for a board presentation

Your downloaded workbook already has styled KPI cards. To make cards that update when your slicers change, link their values to the PivotTable.

For example, to show the total number of fully attended employees for Option 1:

* Click a blank cell where you want the KPI number.

* Type `=`.

* Click the relevant fully attended total inside the PivotTable.

* Press Enter.

Excel will usually generate a `GETPIVOTDATA` formula that retrieves the relevant total. Use the same approach for Option 2, Combined, partial attendance and non-attendance.

Then format the cards:

* Use a large, bold font for the number.

* Use a short heading such as `OPTION 1 — FULLY ATTENDED`.

* Show percentages below the values where useful.

* Keep the same colors throughout the dashboard.

* Use the slicers and PivotChart together for interactive country and meeting comparisons.

### Step 9: Validate your dashboard before presenting it

Before using the workbook for a board meeting, check that:

* Option 1 totals add up to 445.

* Option 2 totals add up to 445.

* Combined totals add up to 445.

* No employee appears twice in the master list because of an email-matching issue.

* The `#N/A` country row is excluded from the country totals.

* The employee with a missing attendance-report email is visible in the Data Quality sheet.

* Changing the country slicer changes the PivotTable and its linked chart as expected.

When the source data is updated, resize the attendance tables if their participant ranges have changed, add any new roster employees to the analysis table, and use Data → Refresh All to refresh your PivotTables.

My recommendation: Open the downloaded workbook first and review the Dashboard and Employee Attendance sheets. You already have the checked totals and country breakdowns. Then use `Pivot Data` to create the interactive PivotTable and slicers, so you can present the static board-level overview and also drill into individual countries when questions arise.

