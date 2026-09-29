# 🧪 Complete Tool Test Questions for `sales.xlsx`

> **File:** `sales.xlsx`  
> **Sheets:** `Orders` (501 rows, 15 cols) · `SalesReps` (12 rows, 7 cols) · `Returns` (36 rows, 6 cols)

Upload `sales.xlsx` first, then copy-paste each question below into the chat. Each question targets specific MCP tools.

---

## 1️⃣ get_workbook_metadata
> What sheets are in this workbook and how many rows and columns does each sheet have?

---

## 2️⃣ read_data_from_excel
> Read the first 5 rows of the Orders sheet including headers and show me the data.

---

## 3️⃣ validate_excel_range
> Validate if the range A1:O501 exists in the Orders sheet.

---

## 4️⃣ write_data_to_excel
> In the Orders sheet, write "HIGH PRIORITY" in cell P1 and write "Yes" in cell P2.

---

## 5️⃣ apply_formula
> In cell P3 of the Orders sheet, apply the formula =L3*N3 to calculate Sales times Profit.

---

## 6️⃣ validate_formula_syntax
> Check if the formula =SUMIF(D2:D501,"North America",L2:L501) is valid.

---

## 7️⃣ format_range
> Format the header row A1:O1 in the Orders sheet with bold text, white font color, and a dark blue background.

---

## 8️⃣ merge_cells
> Merge cells P1:Q1 in the Orders sheet.

---

## 9️⃣ get_merged_cells
> Show me all merged cells in the Orders sheet.

---

## 🔟 unmerge_cells
> Unmerge cells P1:Q1 in the Orders sheet.

---

## 1️⃣1️⃣ create_worksheet
> Create a new worksheet called "Q4_Summary" in the workbook.

---

## 1️⃣2️⃣ copy_worksheet
> Copy the SalesReps sheet to create a new sheet.

---

## 1️⃣3️⃣ rename_worksheet
> Rename the "Q4_Summary" sheet to "Q4_Analysis".

---

## 1️⃣4️⃣ delete_worksheet
> Delete the "Q4_Analysis" sheet from the workbook.

---

## 1️⃣5️⃣ insert_rows
> Insert 2 blank rows at row 2 in the Returns sheet.

---

## 1️⃣6️⃣ insert_columns
> Insert 1 new column at column C in the SalesReps sheet.

---

## 1️⃣7️⃣ delete_sheet_rows
> Delete row 2 from the Returns sheet.

---

## 1️⃣8️⃣ delete_sheet_columns
> Delete column C from the SalesReps sheet.

---

## 1️⃣9️⃣ copy_range
> Copy the range A1:G1 from the SalesReps sheet to A1 of a new sheet called "Rep_Backup".

---

## 2️⃣0️⃣ delete_range
> Delete the range P1:P5 in the Orders sheet and shift cells up.

---

## 2️⃣1️⃣ get_data_validation_info
> Show me all data validation rules in the Orders sheet.

---

## 2️⃣2️⃣ create_table
> Convert the data range A1:O501 in the Orders sheet into a native Excel table called "OrdersTable".

---

## 2️⃣3️⃣ create_pivot_table
> Create a pivot table in a new sheet that summarizes total Sales and Profit grouped by Region from the Orders sheet.

---

## 2️⃣4️⃣ create_chart
> Create a bar chart in the Orders sheet showing total Sales by Region.

---

## 2️⃣5️⃣ create_workbook
> Create a brand new workbook called "Monthly_Report.xlsx" with a sheet named "January" and write the headers "Metric", "Value", and "Notes" in row 1.

---

## 🔥 Bonus: Multi-Tool Chain Questions

These questions will force the agent to call **multiple tools in sequence**:

### Chain Test A (read → write → formula → format)
> Read the SalesReps sheet, then add a new column H with header "Annual_Target" and in H2 apply a formula that sums columns C through F. Format the header row with bold and a green background.

### Chain Test B (create sheet → copy → pivot → chart)
> Create a new sheet called "Dashboard". Then create a pivot table summarizing total Profit by Category from Orders. Add a pie chart to visualize it.

### Chain Test C (read → insert → write → format)
> In the Returns sheet, insert a new column at column G with header "Resolution_Days". Then calculate the number of days between OrderDate and ReturnDate for each row. Format the header in bold red.

---

> [!TIP]
> After running write/edit tests, click the **⬇️ download button** next to the filename in the sidebar to download the modified file and verify changes in Excel!
