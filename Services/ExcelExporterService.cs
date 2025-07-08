using ExcelTest1.Models;
using OfficeOpenXml;

namespace ExcelTest1.Services;

public class ExcelExporterService
{
    public byte[] ExportStudentsToExcel(List<Student> students)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Students");

        worksheet.Cells[1, 1].Value = "ID";
        worksheet.Cells[1, 2].Value = "Name";
        worksheet.Cells[1, 3].Value = "DOB";
        worksheet.Cells[1, 4].Value = "Email";
        worksheet.Cells[1, 5].Value = "Mobile";

        for (int i = 0; i < students.Count; i++)
        {
            var s = students[i];
            worksheet.Cells[i + 2, 1].Value = s.Id;
            worksheet.Cells[i + 2, 2].Value = s.Name;
            worksheet.Cells[i + 2, 3].Value = s.DOB.ToShortDateString();
            worksheet.Cells[i + 2, 4].Value = s.Email;
            worksheet.Cells[i + 2, 5].Value = s.Mob;
        }

        return package.GetAsByteArray();
    }
}