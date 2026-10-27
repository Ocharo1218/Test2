You can do this using Excel formulas, without manually copying names or removing duplicates one by one. I will walk you through it step by step as a beginner.

We will create a new sheet called Combined Attendance containing four columns:

| Full Original Entry | Name           | Country Code | Email                  |
| ------------------- | -------------- | ------------ | ---------------------- |
| Andrew Chibuye (ZM) | Andrew Chibuye | ZM           | andrew.chibuye@pwc.com |

The new sheet will combine both attendance lists, remove duplicate entries, and automatically separate each person's name, country code, and email address.

These instructions are for Microsoft 365 or a newer version of Excel that supports dynamic array formulas.

### Step 1: Create a new sheet

* Open your Excel workbook.

* At the bottom of Excel, find the sheet tabs named Attendance List 1 and Attendance List 2.

* Click the \+ (plus) button beside the sheet tabs to create a new sheet.

* Right-click the new sheet tab.

* Select Rename.

* Type `Combined Attendance` and press Enter.

You now have an empty sheet where your combined results will appear.

### Step 2: Create the four column headings

Go to the new Combined Attendance sheet.

Enter the following headings in row 1:

| Cell | Heading             |
| ---- | ------------------- |
| A1   | Full Original Entry |
| B1   | Name                |
| C1   | Country Code        |
| D1   | Email               |

To do this, click each cell and type the heading.

Your worksheet should now look like this:

A

Full Original Entry

B

Name

C

Country Code

D

Email

Original entry

Person's name

ZM

Email address

### Step 3: Combine both lists and remove duplicates

This is the most important step. We will tell Excel to take all the entries from both sheets, combine them into one list, and keep only unique entries.

* Click cell A2 on the `Combined Attendance` sheet.

* Copy the formula below.

* Paste it into the formula bar or directly into cell A2.

* Press Enter.

Formula for cell A2

=LET(all,VSTACK('Attendance List 1'!A4:A269,'Attendance List 2'!A4:A385),SORT(UNIQUE(FILTER(all,all<>""))))

Copy formula

What this formula does:

* Reads entries from `Attendance List 1`, cells A4 to A269.

* Reads entries from `Attendance List 2`, cells A4 to A385.

* Combines both lists.

* Removes duplicate entries and ignores blank cells.

* Sorts the resulting list alphabetically.

You do not need to drag the formula down. Excel will automatically fill the results into the cells below A2.

Important: The cells below A2 in column A must be empty so Excel can display the full list. If you see a `#SPILL!` error, clear any existing content blocking the results.

### Step 4: Extract each person's name

Now we will separate the person's actual name from the original entry.

For example:

`Andrew Chibuye (ZM) <andrew.chibuye@pwc.com>`

We want to extract just:

`Andrew Chibuye`

Follow these steps:

* Click cell B2.

* Copy and paste this formula.

* Press Enter.

Formula for cell B2

=TEXTBEFORE(A2#," (")

Copy formula

This formula takes everything before the space and opening bracket ` (`.

Excel will automatically extract the names for all the entries in column A.

For example:

| Full Original Entry | Name           |
| ------------------- | -------------- |
| Andrew Chibuye (ZM) | Andrew Chibuye |
| Jane Mwansa (ZM)    | Jane Mwansa    |

### Step 5: Extract the country code

Your country code is the text between the opening and closing brackets.

For example, in `Andrew Chibuye (ZM)`, the country code is `ZM`.

Follow these steps:

* Click cell C2.

* Paste the following formula.

* Press Enter.

Formula for cell C2

=TEXTBEFORE(TEXTAFTER(A2#,"("),")")

Copy formula

The formula finds the opening bracket `(`, takes the text after it, and returns everything before the closing bracket `)`.

The result will look like this:

| Full Original Entry | Name           | Country Code |
| ------------------- | -------------- | ------------ |
| Andrew Chibuye (ZM) | Andrew Chibuye | ZM           |
| Jane Mwansa (KE)    | Jane Mwansa    | KE           |

The country codes above are examples; Excel will extract whichever codes appear in your source data.

### Step 6: Extract the email address

The email address is the text between the angle brackets `<` and `>`.

Follow these steps:

* Click cell D2.

* Paste this formula.

* Press Enter.

Formula for cell D2

=TEXTBEFORE(TEXTAFTER(A2#,"<"),">")

Copy formula

This formula extracts the email address without the `<` and `>` symbols.

Excel should automatically populate the email addresses for all the entries in the combined list.

### Step 7: Check that your results are correct

Your new sheet should now look similar to the example below.

Full Original Entry

Name

Country Code

Email

Andrew Chibuye (ZM) <andrew.chibuye@pwc.com>

Andrew Chibuye

ZM

andrew.chibuye@pwc.com

Jane Mwansa (KE) <jane.mwansa@pwc.com>

Jane Mwansa

KE

jane.mwansa@pwc.com

Illustration of the expected output. Your actual entries and country codes will come from your workbook.

Check the following:

* Column A: Contains the original complete entry.

* Column B: Contains only the person's name.

* Column C: Contains only the country code.

* Column D: Contains only the email address.

* A person whose exact original entry appears on both attendance sheets should appear only once in the combined list.

### Step 8: Make the sheet easier to read

Once the formulas are working:

* Select cells A1:D1 and click Bold on the Home tab.

* Select columns A to D by clicking their column letters.

* Go to Home → Format → AutoFit Column Width (or double-click the boundary between the column letters to resize each column).

* If necessary, go to View → Freeze Panes → Freeze Top Row so the headings remain visible while scrolling.

### One final thing to remember

The formulas in this guide use modern Excel functions, including `VSTACK`, `UNIQUE`, and `TEXTBEFORE`. If Excel returns `#NAME?` when you enter them, your version may not support one or more of these functions; a Power Query method would be a suitable alternative.

Also, the deduplication in Step 3 is based on the entire original entry. If a person's name appears with different country codes or email addresses, Excel will treat those entries as different.
