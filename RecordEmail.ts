function main(workbook: ExcelScript.Workbook, from: string, dateTimeReceived: string, subject: string, emailAddress: string)
{
  let emailWorksheet = workbook.getWorksheet("EmailWorksheet");
  let emailTable = emailWorksheet.getTable("EmailTable");
  let subjectText = subject.replace("Re: ", "");
  subjectText = subjectText.replace("Re: ","");
  emailTable.addRow(-1, [subject, from, emailAddress, dateTimeReceived]);
  emailWorksheet.getRange("A:D").getFormat().autofitColumns;
  let emailPivotTable = workbook.getWorksheet("EmailPivotTable").getPivotTable("EmailPivotTable");
  emailPivotTable.refresh();
}