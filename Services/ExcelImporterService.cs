using ExcelTest1.Models;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.Data;

namespace ExcelTest1.Services;

public class ExcelImporterService
{
    private readonly string _connectionString;

    public ExcelImporterService(string connectionString)
    {
        _connectionString = connectionString;
    }

    public async Task<ImportResult> ImportStudentsFromExcel(byte[] excelData)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var result = new ImportResult();
        var students = new List<Student>();

        try
        {
            using var package = new ExcelPackage(new MemoryStream(excelData));
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();

            if (worksheet == null)
            {
                result.IsSuccess = false;
                result.ErrorMessage = "Excel faylida worksheet topilmadi";
                return result;
            }

            // Header tekshirish (2-qatordan boshlanadi, 1-qator header)
            var rowCount = worksheet.Dimension?.Rows ?? 0;

            for (int row = 2; row <= rowCount; row++)
            {
                try
                {
                    var student = new Student
                    {
                        Name = worksheet.Cells[row, 2].Value?.ToString() ?? string.Empty,
                        DOB = DateTime.Parse(worksheet.Cells[row, 3].Value?.ToString() ?? DateTime.Now.ToString()),
                        Email = worksheet.Cells[row, 4].Value?.ToString() ?? string.Empty,
                        Mob = worksheet.Cells[row, 5].Value?.ToString() ?? string.Empty
                    };

                    // Validatsiya
                    if (string.IsNullOrWhiteSpace(student.Name) ||
                        string.IsNullOrWhiteSpace(student.Email))
                    {
                        result.SkippedRows++;
                        continue;
                    }

                    students.Add(student);
                }
                catch (Exception ex)
                {
                    result.SkippedRows++;
                    result.Errors.Add($"Qator {row}: {ex.Message}");
                }
            }

            // Ma'lumotlarni database'ga saqlash
            if (students.Any())
            {
                await SaveStudentsToDatabase(students);
                result.ProcessedRows = students.Count;
                result.IsSuccess = true;
            }
            else
            {
                result.IsSuccess = false;
                result.ErrorMessage = "Import qilish uchun yaroqli ma'lumotlar topilmadi";
            }
        }
        catch (Exception ex)
        {
            result.IsSuccess = false;
            result.ErrorMessage = $"Excel faylini o'qishda xatolik: {ex.Message}";
        }

        return result;
    }

    public async Task<ImportResult> ImportStudentsFromExcelFile(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return new ImportResult
            {
                IsSuccess = false,
                ErrorMessage = "Fayl tanlanmagan yoki bo'sh"
            };
        }

        if (!file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return new ImportResult
            {
                IsSuccess = false,
                ErrorMessage = "Faqat .xlsx formatdagi fayllar qo'llab-quvvatlanadi"
            };
        }

        using var memoryStream = new MemoryStream();
        await file.CopyToAsync(memoryStream);
        return await ImportStudentsFromExcel(memoryStream.ToArray());
    }

    private async Task SaveStudentsToDatabase(List<Student> students)
    {
        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        using var transaction = connection.BeginTransaction();

        try
        {
            const string insertQuery = @"
                INSERT INTO [TestExcel].[dbo].[Students] ([Name], [DOB], [Email], [Mob])
                VALUES (@Name, @DOB, @Email, @Mob)";

            foreach (var student in students)
            {
                using var command = new SqlCommand(insertQuery, connection, transaction);
                command.Parameters.AddWithValue("@Name", student.Name);
                command.Parameters.AddWithValue("@DOB", student.DOB);
                command.Parameters.AddWithValue("@Email", student.Email);
                command.Parameters.AddWithValue("@Mob", student.Mob);

                await command.ExecuteNonQueryAsync();
            }

            transaction.Commit();
        }
        catch
        {
            transaction.Rollback();
            throw;
        }
    }

    // Bulk insert uchun optimallashtirilgan versiya
    private async Task SaveStudentsToDatabaseBulk(List<Student> students)
    {
        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        var dataTable = new DataTable();
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("DOB", typeof(DateTime));
        dataTable.Columns.Add("Email", typeof(string));
        dataTable.Columns.Add("Mob", typeof(string));

        foreach (var student in students)
        {
            dataTable.Rows.Add(student.Name, student.DOB, student.Email, student.Mob);
        }

        using var bulkCopy = new SqlBulkCopy(connection);
        bulkCopy.DestinationTableName = "[TestExcel].[dbo].[Students]";
        bulkCopy.ColumnMappings.Add("Name", "Name");
        bulkCopy.ColumnMappings.Add("DOB", "DOB");
        bulkCopy.ColumnMappings.Add("Email", "Email");
        bulkCopy.ColumnMappings.Add("Mob", "Mob");

        await bulkCopy.WriteToServerAsync(dataTable);
    }
}
//public class ImportResult
//{
//    public bool IsSuccess { get; set; }
//    public int ProcessedRows { get; set; }
//    public int SkippedRows { get; set; }
//    public string ErrorMessage { get; set; } = string.Empty;
//    public List<string> Errors { get; set; } = new List<string>();
//}
