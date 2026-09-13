# Excel Macro Toolkit

VBA macros for Excel formatting, numeric scaling, copying values, and pasting financial statement data using keyboard shortcuts.

## Keyboard Shortcuts

| Shortcut | Action |
| --- | --- |
| Alt + Shift + . | Divide selected values by 10 |
| Alt + Shift + , | Multiply selected values by 10 |
| Ctrl + Alt + A | Cycle font colors: black → blue → green → black |
| Ctrl + Alt + S | Reverse positive and negative signs |
| Ctrl + Alt + 4 | Apply currency formatting |
| Ctrl + Alt + 5 | Apply percentage formatting |
| Ctrl + Alt + 6 | Apply multiples formatting, such as 5.0x |
| Ctrl + Alt + 7 | Apply whole-number formatting with comma separators |
| Ctrl + Alt + [ | Decrease displayed decimal places |
| Ctrl + Alt + ] | Increase displayed decimal places |
| Ctrl + Alt + 8 | Display zeros as a dash while preserving the existing number format |
| Ctrl + Alt + C | Not assigned |
| Alt + Shift + C | Copy selected cell values to the clipboard and collect numeric values for summing |
| Alt + Shift + V | Insert a SUM formula using collected numbers, then clear the list |
| Alt + Shift + X | Clear collected numbers without inserting a formula |
| Ctrl + Alt + Q | Paste clipboard text down rows, skipping completely blank lines |
| Alt + Shift + Z | Paste nonblank clipboard text lines across columns |

## Copy and Collect Values

**Alt + Shift + C** copies calculated cell values rather than formulas. Tabs separate columns, and line breaks separate rows.

The same shortcut appends numeric values to an internal VBA list. Copy repeatedly to collect values from multiple selections.

Select a destination cell and press **Alt + Shift + V** to insert a formula such as:

    =SUM(100,-25,50)

The sum shortcut uses the internal list, not the current Windows clipboard. The list clears after inserting the formula or pressing **Alt + Shift + X**.

The collected list is temporary and does not persist after closing Excel or resetting VBA.

## Paste Down Rows

**Ctrl + Alt + Q** removes completely blank lines from clipboard text before pasting.

- Starts at the active cell.
- Preserves tab-separated columns and empty fields between tabs.
- Converts recognized numbers into numeric values.
- Keeps unrecognized entries as text.
- Highlights suspicious numeric text with a red fill.
- Does not delete worksheet cells or shift existing data upward.

## Paste Across Columns

**Alt + Shift + Z** places each nonblank clipboard text line into the next column, starting at the active cell.

- Splits on line breaks, not spaces.
- Keeps labels such as “Total Revenues” together.
- Removes dollar signs and skips entries containing only a dollar sign.
- Converts recognized numbers into numeric values.
- Highlights suspicious numeric text with a red fill.

Copy one source table row at a time when using this shortcut.

## Number Validation and Paste Limitations

Both paste macros check comma placement before converting numbers. For example, `10,731` becomes a number, while `10,73` remains text and is flagged for review.

- Commas are treated as thousands separators and periods as decimal points.
- A standalone dash is preserved as text.
- Formats not currently recognized, such as `(1,234)` and `12%`, may also be flagged.
- Red highlighting is a direct cell fill, not conditional formatting. Clear it manually after reviewing or correcting an entry.
- Missing digits cannot be detected if the remaining value still looks valid.
- Intentional blank lines are skipped along with unwanted blank lines.
- Pasting overwrites cells within the destination area.
- These macros paste values from clipboard text, not source formatting or formulas.

## Formatting Notes

- Scaling shortcuts change numeric values; decimal-place formatting only changes their display.
- Scaling formulas uses powers of 10.
- Mixed font colors reset to black before cycling.
- The zero-as-dash shortcut requires an existing number format with multiple sections. Apply a number format first; General is unchanged.

## Installation

1. Open Excel and press **Alt + F11** to open the VBA Editor.
2. Import `excel_macro_toolkit.bas` into `PERSONAL.XLSB` or a macro-enabled workbook.
3. Run `SetCustomShortcuts` to activate the keyboard shortcuts.
4. Save `PERSONAL.XLSB`, or save the workbook as `.xlsm`.
5. Run `SetCustomShortcuts` again after restarting Excel unless you have configured it to run automatically.

Using `PERSONAL.XLSB` makes the macros available while working in other workbooks.

## Files

- `excel_macro_toolkit.bas` — VBA macro library
- `Excel_Macro_Cheat_Sheet.pdf` — Keyboard shortcut reference

## Requirements

- Microsoft Excel for Windows with VBA7 support
- Macros enabled
- The workbook containing the macros must be open

## License

MIT License