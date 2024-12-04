---
layout: post
title:  "Introduction to Excel Macros"
date:   2024-11-20 09:00:00 -0800
categories: tutorial
---

**Table of Contents**

* TOC
{:toc}

## Introduction

Microsoft's Excel is a powerful spreadsheet tool to record and analyze data. However, raw data is often messy and unstructured. Many people spend hours repeating the same data-cleaning tasks. Fortunately, repetitive tasks can be automated with macros.

Macros are like a set of instructions. Each macro can tell Excel to do a couple or hundreds of things at once. Programmers write macros in the Visual Basic for Applications (VBA) language, but you do not need to know programming or VBA to get started.

### Audience

This tutorial is for anyone new to Excel macros. You will learn how to set up Excel, record macros, and explore the VBA code. You may find macros especially useful if your role includes:

- Data entry
- Data analysis
- Bookkeeping
- Timekeeping

### Prerequisites

You will need:

- A copy of Excel
- A basic understanding of Excel

This tutorial applies to:

- Windows computers
- Excel for Microsoft 365 
- Excel 2024
- Excel 2019
- Excel 2016

> Note: Screenshots and animations use Excel 2016.

---

## Setup

By default, Excel hides the Developer tab which provides access to macros. To unhide the Developer tab:

1. In the Excel menu bar, click File. 

   ![menu bar file]({{ site.baseurl }}/assets/images/excel_menu_bar_file.png)

2. Towards the bottom of the side panel menu, click Options.

   ![menu options]({{ site.baseurl }}/assets/images/excel_menu_options.png)

3. In the Excel Options side panel menu, click Customize Ribbon, select the Developer checkbox, and then click OK.

   ![customize ribbon]({{ site.baseurl }}/assets/images/excel_customize_ribbon.png)

The Developer tab is now visible in your Excel menu bar.

![developer tab]({{ site.baseurl }}/assets/images/excel_menu_bar_developer_tab.png)

---

## Record Macro

The Record Macro tool is the simplest way to create macros. Excel will record and save your instructions into a macro. You can then run ("play back") the recorded macro as many times as needed.

### Format texts

Let's begin with a formatting macro:

1. Enter some text into cells A1, B1, and C1.

    ![enter text]({{ site.baseurl }}/assets/images/excel_random_text.png)

2. Click the Developer tab and then Record Macro.

    ![record macro]({{ site.baseurl }}/assets/images/excel_developer_tab_record_macro.png)

3. In the Record Macro window, leave the defaults as-is and click OK to continue (i.e., Macro name: Macro1).

    ![macro window]({{ site.baseurl }}/assets/images/excel_macro_window.png)

    > Note: By default, recorded macros are stored in "This Workbook". You need to save your current Workbook to save the macro too. To save and access macros across different Workbooks, see this [guide on saving macros](https://support.microsoft.com/en-us/office/create-and-save-all-your-macros-in-a-single-workbook-66c97ab3-11c2-44db-b021-ae005a9bc790).

4. Bold A1, highlight B2, and change C1 text color to orange.

    ![record formatting]({{ site.baseurl }}/assets/gifs/record_formatting.gif)

5. In the Developer Tab, click Stop Recording.

   ![stop recording]({{ site.baseurl }}/assets/images/excel_developer_tab_stop_record.png)

6. Select row 1, right-click the selection, and click Delete.

   ![delete row]({{ site.baseurl }}/assets/images/excel_delete_row.png)

7. Enter new text into cells A1, B1, and C1.
8. In the Developer tab, click Macros, and then Run.

    ![run formatting]({{ site.baseurl }}/assets/gifs/run_formatting.gif)

Congrats! You just recorded and played back your first macro.

### Time sheet calculation

The following example is a time sheet with employee clock in/clock out times:

![time sheet]({{ site.baseurl }}/assets/images/excel_timesheet.png)

You would normally calculate hours worked with a custom formula:

1. `=(C2*24)-(B2*24)` in cell D2.
2. Format the result as Number.

Instead, you can record a macro with a shortcut key:

1. In the Developer tab, click Use Relative References, and then Record Macro.

    ![record macro relative references]({{ site.baseurl }}/assets/images/excel_developer_tab_macros_relative_ref.png)

    > Note: Use Relative References ensures our macro will play back *relative* to future cells. To learn more about references, see this [guide on relative, absolute, and mixed references](https://support.microsoft.com/en-us/office/switch-between-relative-absolute-and-mixed-references-dfec08cd-ae65-4f56-839e-5f0d8d0baca9).

2. Rename the macro and enter a shortcut key. A description can help you remember what the macro does. Click OK.
  
    ![custom macro info]({{ site.baseurl }}/assets/images/excel_macro_custom_info.png)

3. In cell D2, enter the formula `=(C2*24)-(B2*24)` and change its format to Number.

    ![record custom formula]({{ site.baseurl }}/assets/gifs/record_custom_formula.gif)

4. In the Developer Tab, click Stop Recording.

You can play back your macro in any selected cell by pressing your custom shortcut key.

![run custom formula]({{ site.baseurl }}/assets/gifs/run_custom_formula.gif)

---

## Explore VBA

The last step is to explore a macro's VBA code. This exploration is meant as an overview - the explanations use generic words and avoid programming terminology.

The following example uses the recorded macro from [Time sheet calculation](#time-sheet-calculation).

### VBA Editor

The VBA Editor is where you view, organize, and edit your macro's code.

To open the VBA Editor:

1. In the Developer tab, click Macros.
2. Make sure the correct macro is selected.
3. Click Edit.

    ![edit macro]({{ site.baseurl }}/assets/images/excel_developer_tab_edit_macro.png)

The VBA Editor window opens to the selected macro.

![vba window]({{ site.baseurl }}/assets/images/excel_vba_window.png)

### Layout

Recorded macros are saved as `Module1` within the Modules folder. You can see the macro's name (highlighted below) after the `Sub` statement and in the Procedure pane.

![module name]({{ site.baseurl }}/assets/images/excel_macro_module_name.png)

Similar to reading a book, VBA code is read from top to bottom, left to right. All of the recorded code is in between the `Sub` and `End Sub` statements. 

![code area]({{ site.baseurl }}/assets/images/excel_macro_code_area.png)

### VBA code

The following sections explain the code line-by-line. Blank spaces do nothing.

![lines]({{ site.baseurl }}/assets/images/vba_code_lines.png)

#### Line 1

Blank space.

#### Line 2

The `Sub` statement tells VBA this is the beginning.

#### Lines 3-5

Apostrophes indicate comments. VBA ignores and skips over any text in the same line after an apostrophe.

![vba comments]({{ site.baseurl }}/assets/images/excel_vba_comments.png)

#### Line 6

Blank space.

#### Line 7

See [Lines 3-5](#lines-3-5).

#### Line 8

The first line of code that runs is `ActiveCell.FormulaR1C1 = "=(RC[-1]*24)-(RC[-2]*24)"`.

![first line]({{ site.baseurl }}/assets/images/vba_code_first_line.png)

Broken down:

- `ActiveCell`: The selected cell in the spreadsheet.
- `.FormulaR1C1`: Enter a formula using relative references.
- `=`: To the right is the formula.
- `"=(RC[-1]*24)-(RC[-2]*24)"`: Multiply one cell to the left of `ActiveCell` by 24. Multiply two cells to the left of `ActiveCell` by 24. Subtract the second result from the first result.

#### Line 9

The second line of code that runs is `ActiveCell.Select`.

![second line]({{ site.baseurl }}/assets/images/vba_code_second_line.png)

Broken down:

- `ActiveCell`: The selected cell in the spreadsheet.
- `.Select`: Select it.

#### Line 10

The third line of code that runs is `Selection.NumberFormat = "0.00"`.

![third line]({{ site.baseurl }}/assets/images/vba_code_third_line.png)

Broken down:

- `Selection`: The selected cell.
- `.NumberFormat`: Enter a value in a specific format.
- `=`: To the right is the format.
- `"0.00"`: Number format.

#### Line 11

The `End Sub` statement tells VBA this is the end.

---

## Conclusion

This introductory tutorial is a launching point into the world of Excel macros. You can now record your own macros and automate repetitive tasks!

### Next steps

- To explore VBA further, watch [this video tutorial](https://www.youtube.com/watch?v=IJQHMFLXk_c).
- To review official documentation, see [Microsoft's article](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office).

---

## References

- [Create and save all your macros in a single workbook](https://support.microsoft.com/en-us/office/create-and-save-all-your-macros-in-a-single-workbook-66c97ab3-11c2-44db-b021-ae005a9bc790)
- [Switch between relative, absolute, and mixed references](https://support.microsoft.com/en-us/office/switch-between-relative-absolute-and-mixed-references-dfec08cd-ae65-4f56-839e-5f0d8d0baca9)
- [The Complete Guide to the VBA Sub](https://excelmacromastery.com/excel-vba-sub/)