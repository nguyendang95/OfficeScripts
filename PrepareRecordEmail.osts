function main(workbook: ExcelScript.Workbook)
{
  let emailsSheet = workbook.addWorksheet("EmailWorksheet");
  let tableRange = emailsSheet.getRange("A1:D1").setValues([["Tiêu đề", "Người gửi", "Địa chỉ email", "Ngày, giờ nhận"]]);
  emailsSheet.getRange("A:D").getFormat().autofitColumns();
  let emailsTable = emailsSheet.addTable(emailsSheet.getRange("A1:D2"), true);
  emailsTable.setName("EmailTable");
  let pivotTableWorksheet = workbook.addWorksheet("EmailPivotTable");
  let pivotTableRange = pivotTableWorksheet.getRange("A3:D20");
  let emailPivotTable = pivotTableWorksheet.addPivotTable("EmailPivotTable", "EmailTable", pivotTableRange);
  emailPivotTable.addRowHierarchy(emailPivotTable.getHierarchy("Người gửi"));
  emailPivotTable.addRowHierarchy(emailPivotTable.getHierarchy("Địa chỉ email"));
  emailPivotTable.addRowHierarchy(emailPivotTable.getHierarchy("Ngày, giờ nhận"));
  emailPivotTable.addDataHierarchy(emailPivotTable.getHierarchy("Tiêu đề"));
  }