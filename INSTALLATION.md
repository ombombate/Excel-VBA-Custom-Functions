Installation Guide – Excel VBA Custom Functions (SUMCOLOR & COUNTIFCOLOR)

Introduction

Thank you for downloading these custom Excel VBA functions.

This guide will help you install and use the following formulas in Excel:

- "SUMCOLOR()" – Calculate a score based on cell colors.
- "COUNTIFCOLOR()" – Count cells of a specific color.

No prior VBA experience is required. Simply follow the steps below.

---

What is VBA?

VBA (Visual Basic for Applications) is Microsoft's programming language built into Excel and other Office applications.

VBA allows you to:

- Automate repetitive tasks
- Create custom buttons and macros
- Build custom worksheet functions
- Extend Excel beyond its built-in capabilities

Examples of built-in Excel formulas:

=SUM(A1:A10)
=COUNTIF(A1:A10,">0")
=VLOOKUP(...)

Examples of custom VBA formulas:

=SUMCOLOR(A1:A10)
=COUNTIFCOLOR(A1:A10,D1)

---

What is PERSONAL.XLSB?

PERSONAL.XLSB is Excel's Personal Macro Workbook.

Think of it as your personal toolbox.

Any VBA code stored in PERSONAL.XLSB becomes available in:

- New workbooks
- Existing workbooks
- Any Excel file opened on your computer

This means you only need to install the code once.

---

Step 1 – Create the Personal Macro Workbook

If you already have PERSONAL.XLSB, skip to Step 4.

Open Excel

Create a blank workbook.

Open the Macro Recorder

Navigate to:

View → Macros → Record Macro

Configure the Recording

For:

Store Macro In:

Select:

Personal Macro Workbook

Click:

OK

Stop Recording

Navigate to:

View → Macros → Stop Recording

Excel has now created PERSONAL.XLSB.

---

Step 2 – Open the VBA Editor

Press:

ALT + F11

The VBA Editor will open.

You should see a Project Explorer window on the left.

If you do not see it:

CTRL + R

to display it.

---

Step 3 – Locate PERSONAL.XLSB

In Project Explorer find:

VBAProject (PERSONAL.XLSB)

Expand it.

You should see:

Microsoft Excel Objects
Modules

---

Step 4 – Create a New Module

Right-click:

VBAProject (PERSONAL.XLSB)

Select:

Insert → Module

A new module will be created.

Typically it will appear as:

Module1

---

Step 5 – Paste the VBA Code

Open the provided VBA code file.

Copy the entire code.

Paste it into the newly created module.

Example:

Function SumColor(rng As Range) As Double

    Dim cell As Range
    Dim cellValue As Double

    For Each cell In rng

        Select Case cell.Interior.ColorIndex

            Case 6
                cellValue = 5

            Case 3
                cellValue = 0

            Case 43
                cellValue = 10

            Case 56
                cellValue = -5

            Case 33
                cellValue = 15

            Case Else
                cellValue = 0

        End Select

        SumColor = SumColor + cellValue

    Next cell

End Function

Repeat the process for any additional functions such as:

COUNTIFCOLOR()

---

Step 6 – Save PERSONAL.XLSB

Press:

CTRL + S

Or choose:

File → Save PERSONAL.XLSB

Close the VBA Editor.

---

Step 7 – Close Excel Properly

Close Excel completely.

Excel will display:

Do you want to save changes to PERSONAL.XLSB?

Click:

Save

This step is critical.

If you click "Don't Save", your VBA code will be lost.

---

Step 8 – Reopen Excel

Close Excel.

Open Excel again.

Your custom formulas should now be available in every workbook on your computer.

---

How to Use SUMCOLOR()

Color Mapping

Color| Value
Light Blue| 15
Light Green| 10
Yellow| 5
Red| 0
Black| -5
Other Colors| 0

Formula

=SUMCOLOR(A1:A20)

Example

Suppose:

- 2 Light Blue cells
- 3 Light Green cells
- 1 Yellow cell

Calculation:

(2 × 15) + (3 × 10) + (1 × 5)

= 65

Result:

65

---

How to Use COUNTIFCOLOR()

This function counts cells having the same color as a reference cell.

Formula

=COUNTIFCOLOR(A1:A20,D1)

Where:

- A1:A20 = Range to count
- D1 = Cell containing the target color

Example

If D1 is Light Green:

=COUNTIFCOLOR(A1:A20,D1)

The formula will return the number of Light Green cells in A1:A20.

---

Important Notes

Formula Not Working?

Press:

ALT + F11

and confirm that:

PERSONAL.XLSB

contains the VBA code.

---

Getting #NAME?

This means Excel cannot find the custom function.

Check that:

- Code is stored in PERSONAL.XLSB
- Macros are enabled
- Excel was restarted after installation

---

Color Changed But Formula Didn't Update?

Excel does not automatically recalculate when cell colors change.

To refresh:

- Press F9
- Or edit the formula and press Enter

---

Congratulations!

You have successfully installed custom VBA-powered Excel functions.

You can now use:

=SUMCOLOR()

and

=COUNTIFCOLOR()

in any workbook on your computer.

Happy Exceling!
