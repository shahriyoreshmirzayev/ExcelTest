using ExcelTest1.Models;
using ExcelTest1.Repositories;
using OfficeOpenXml;

namespace ExcelTest1.Services;

public class ExcelService : IExcelService
{
    private readonly IStudentRepository _studentRepository;
    private const string ExcelContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public ExcelService(IStudentRepository studentRepository)
    {
        _studentRepository = studentRepository;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public byte[] ExportStudentsToExcel(IEnumerable<Student> students)
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Students");

        // Header
        var headers = new[] { "ID", "Name", "Date of Birth", "Email", "Mobile" };
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cells[1, i + 1].Value = headers[i];
            worksheet.Cells[1, i + 1].Style.Font.Bold = true;
        }

        // Data
        int row = 2;
        foreach (var student in students)
        {
            worksheet.Cells[row, 1].Value = student.Id;
            worksheet.Cells[row, 2].Value = student.Name;
            worksheet.Cells[row, 3].Value = student.DOB;
            worksheet.Cells[row, 3].Style.Numberformat.Format = "yyyy-mm-dd";
            worksheet.Cells[row, 4].Value = student.Email;
            worksheet.Cells[row, 5].Value = student.Mob;
            row++;
        }

        // Auto-fit columns
        worksheet.Cells.AutoFitColumns();

        return package.GetAsByteArray();
    }

    public async Task<ImportResult> ImportStudentsFromExcelAsync(IFormFile file)
    {
        var result = new ImportResult();

        if (!IsValidExcelFile(file))
        {
            result.ErrorMessage = "Invalid file format. Only .xlsx files are supported.";
            return result;
        }

        try
        {
            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);

            var students = ParseExcelFile(stream.ToArray(), result);

            if (students.Any())
            {
                await _studentRepository.CreateBulkAsync(students);
                result.ProcessedRows = students.Count;
                result.IsSuccess = true;
            }
            else if (result.Errors.Count == 0)
            {
                result.ErrorMessage = "No valid data found to import.";
            }
        }
        catch (Exception ex)
        {
            result.ErrorMessage = $"Error processing Excel file: {ex.Message}";
        }

        return result;
    }

    //public byte[] GenerateImportSample()
    //{
    //    var sampleStudents = new List<Student>
    //    {
    //        new Student { Id = 1, Name = "Sample Student 1", DOB = DateTime.Now.AddYears(-20), Email = "sample1@email.com", Mob = "+998901234567" },
    //        new Student { Id = 2, Name = "Sample Student 2", DOB = DateTime.Now.AddYears(-22), Email = "sample2@email.com", Mob = "+998912345678" },
    //        new Student { Id = 3, Name = "Sample Student 3", DOB = DateTime.Now.AddYears(-21), Email = "sample3@email.com", Mob = "+998923456789" }
    //    };

    //    return ExportStudentsToExcel(sampleStudents);
    //}

    private static bool IsValidExcelFile(IFormFile file)
    {
        return file != null &&
               file.Length > 0 &&
               file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
    }

    private static List<Student> ParseExcelFile(byte[] excelData, ImportResult result)
    {
        var students = new List<Student>();

        using var package = new ExcelPackage(new MemoryStream(excelData));
        var worksheet = package.Workbook.Worksheets.FirstOrDefault();

        if (worksheet == null)
        {
            result.ErrorMessage = "No worksheet found in Excel file.";
            return students;
        }

        var rowCount = worksheet.Dimension?.Rows ?? 0;

        for (int row = 2; row <= rowCount; row++) // Skip header row
        {
            try
            {
                var student = ParseStudentFromRow(worksheet, row);

                if (IsValidStudentData(student))
                {
                    students.Add(student);
                }
                else
                {
                    result.SkippedRows++;
                    result.Errors.Add($"Row {row}: Invalid or incomplete data");
                }
            }
            catch (Exception ex)
            {
                result.SkippedRows++;
                result.Errors.Add($"Row {row}: {ex.Message}");
            }
        }

        return students;
    }

    private static Student ParseStudentFromRow(ExcelWorksheet worksheet, int row)
    {
        return new Student
        {
            Name = worksheet.Cells[row, 2].Value?.ToString()?.Trim() ?? string.Empty,
            DOB = ParseDate(worksheet.Cells[row, 3].Value?.ToString()),
            Email = worksheet.Cells[row, 4].Value?.ToString()?.Trim() ?? string.Empty,
            Mob = worksheet.Cells[row, 5].Value?.ToString()?.Trim() ?? string.Empty
        };
    }

    private static DateTime ParseDate(string? dateString)
    {
        if (string.IsNullOrWhiteSpace(dateString))
            return DateTime.Now.AddYears(-18); // Default age

        if (DateTime.TryParse(dateString, out var date))
            return date;

        return DateTime.Now.AddYears(-18);
    }

    private static bool IsValidStudentData(Student student)
    {
        return !string.IsNullOrWhiteSpace(student.Name) &&
               !string.IsNullOrWhiteSpace(student.Email) &&
               student.DOB <= DateTime.Now;
    }
}
