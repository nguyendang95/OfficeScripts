function main(workbook: ExcelScript.Workbook)
{
  let selectedRange = workbook.getSelectedRange();
  let selectedRangeValues = selectedRange.getValues();
  let rowCount: number;
  let columnCount: number;
  let row: number;
  let column: number;
  rowCount = selectedRange.getRowCount();
  columnCount = selectedRange.getColumnCount();
  for (row = 0; row < rowCount; row++)
  {
    for (column = 0; column < columnCount; column++)
    {
      let cell = selectedRange.getCell(row, column);
      if (cell.getFormula)
      {
        cell.setValue(selectedRangeValues[row][column]);
      }
    }
  }
}