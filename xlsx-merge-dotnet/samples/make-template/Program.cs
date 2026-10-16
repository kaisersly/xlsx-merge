using ClosedXML.Excel;

// Generates the demo template used by the sample CLI invocations and the README.
// Usage: dotnet run --project samples/make-template -- <output-path>

var output = args.Length > 0 ? args[0] : "samples/attendance-template.xlsx";

var workbook = new XLWorkbook();
var sheet = workbook.Worksheets.Add("Attendance");

sheet.Cell(1, 1).Value = "#COURSE# — attendance";
sheet.Cell(2, 1).Value = "Room: #ROOM#";
sheet.Range("A1:D1").Merge();

sheet.Cell(4, 1).Value = "#EMPLOYEE_ID#";
sheet.Cell(4, 2).Value = "#NAME#";
sheet.Cell(4, 3).Value = "#SCORE# / 20";
sheet.Cell(4, 4).Value = "Signed";
sheet.Range("E4:F4").Merge();
sheet.Cell(4, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#DDEBF7");
sheet.Row(4).Height = 18;
sheet.Cell(5, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFFFF");

sheet.Cell(7, 1).Value = "Generated on #DATE# (#SOMEWHERE#)";

sheet.PageSetup.SetDifferentFirstPageOnHF(true);
sheet.PageSetup.Header.Left.AddText("#COURSE#", XLHFOccurrence.OddPages);
sheet.PageSetup.Header.Left.AddText("#COURSE# — cover", XLHFOccurrence.FirstPage);
sheet.PageSetup.Footer.Center.AddText("#DATE#", XLHFOccurrence.OddPages);

Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(output))!);
workbook.SaveAs(output);
Console.WriteLine($"Sample template written to {output}");
