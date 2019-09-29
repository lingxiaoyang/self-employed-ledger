To make it work with your bank, you will need to edit the VBA code by double-clicking the module `Converted Macro- DoImportCSV` (certainly, I'll clean this name later).

Press Ctrl-F and find with `If acName`. The cursor will highlight the condition branches reading the bank that you're selecting. If you don't see the bank that you are using, go ahead and add yours into the `If` block. What you will need to change here are:

  - The condition (`acName like ...` or `acName = ...`); and
  - The name of function (`ImportXXXXCSV`).

Then, go to the bottom of the code, copy the last `Private Function` block (from `Private Function` until `End Function`), and duplicate it.

*First*, adapt the name of function (the same as what you just changed above, `ImportXXXXCSV`). You will need to adapt the 1st line (`Private Function ImportXXXXCSV(.....`) and the last line above `End Function` (`Set ImportXXXXCSV = DataLines`).

*Second*, open your CSV file exported from your bank (if your bank only provides PDF statements, see "FAQ" section below), and specify from which column you can extract following information:

  - Date of transaction: easier to be all-numerical, like `2019-06-01` or `6/1/2019` or so. A difficult one would be `Jun 1 2019` which needs more difficult VBA code to parse, depending on your coding level.
  - Description line 1
  - Description line 2: this could be empty, or could be assembled from multiple fields, as long as it will help you later on.
  - Amount: some banks separate debits and credits into two columns. Keep in mind, what you earn will be entered positive, and what you spend will be negative.

*Finally*, adapt the private function `ImportXXXXCSV`. In preset examples, you will find:

  - How to skip the header row;
  - How to parse a `YYYY-mm-dd` formatted date and build VBA date object using `CDate()`;
    - Example 1: `CDate("09/30/2019")`
    - Example 2: `CDate("Sep 30 2019")`
  - How to build a decimal object for the amount.

Keep in mind:

  - Look for `dict("tranDate")`, `dict("desc1")`, `dict("desc2")`, and `dict("amount")`.
  - `Items(1)` corresponds to the value of *first* column of the CSV file.

Good luck!
