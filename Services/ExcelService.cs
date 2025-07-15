using ExcelTest1.Models;
using ExcelTest1.Repositories;
using OfficeOpenXml;
using QRCoder;
using System.Drawing;
using System.Drawing.Imaging;

namespace ExcelTest1.Services;

public class ExcelService : IExcelService
{
    private readonly IStudentRepository _studentRepository;
    private readonly IWebHostEnvironment _webHostEnvironment;
    private const string ExcelContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    public ExcelService(IStudentRepository studentRepository, IWebHostEnvironment webHostEnvironment)
    {
        _studentRepository = studentRepository;
        _webHostEnvironment = webHostEnvironment;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public byte[] ExportStudentsToExcel(IEnumerable<Student> students)
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Students");

        // Header styling
        var headers = new[] { "ID", "Name", "Date of Birth", "Email", "Mobile" };
        for (int i = 0; i < headers.Length; i++)
        {
            worksheet.Cells[1, i + 1].Value = headers[i];
            worksheet.Cells[1, i + 1].Style.Font.Bold = true;
            worksheet.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            worksheet.Cells[1, i + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
        }

        // Data rows
        int row = 2;
        foreach (var student in students)
        {
            worksheet.Cells[row, 1].Value = student.Id;
            worksheet.Cells[row, 2].Value = student.Name;
            worksheet.Cells[row, 3].Value = student.DOB;
            worksheet.Cells[row, 3].Style.Numberformat.Format = "yyyy-mm-dd";
            worksheet.Cells[row, 4].Value = student.Email;
            worksheet.Cells[row, 5].Value = student.Mob;

            // Add borders to data rows
            for (int col = 1; col <= 5; col++)
            {
                worksheet.Cells[row, col].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            }

            row++;
        }

        // Auto-fit columns
        worksheet.Cells.AutoFitColumns();

        return package.GetAsByteArray();
    }

    public async Task<ExportResult> ExportStudentsToExcelWithQRAsync(IEnumerable<Student> students, string baseUrl)
    {
        var result = new ExportResult();

        try
        {
            // Excel faylini yaratish
            var excelData = ExportStudentsToExcel(students);

            // Fayl nomini yaratish
            var fileName = $"students_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            var downloadsPath = Path.Combine(_webHostEnvironment.WebRootPath, "downloads");

            // Downloads papkasini yaratish
            if (!Directory.Exists(downloadsPath))
            {
                Directory.CreateDirectory(downloadsPath);
            }

            var filePath = Path.Combine(downloadsPath, fileName);

            // Faylni saqlash
            await File.WriteAllBytesAsync(filePath, excelData);

            // Fayl URL'ini yaratish
            var fileUrl = $"{baseUrl}/downloads/{fileName}";

            // QR kod yaratish
            var qrCodeBytes = GenerateQRCode(fileUrl);

            // QR kod faylini saqlash
            var qrFileName = $"qr_{DateTime.Now:yyyyMMdd_HHmmss}.png";
            var qrCodesPath = Path.Combine(_webHostEnvironment.WebRootPath, "qrcodes");

            if (!Directory.Exists(qrCodesPath))
            {
                Directory.CreateDirectory(qrCodesPath);
            }

            var qrFilePath = Path.Combine(qrCodesPath, qrFileName);
            await File.WriteAllBytesAsync(qrFilePath, qrCodeBytes);

            result.ExcelData = excelData;
            result.QRCodeData = qrCodeBytes;
            result.FileUrl = fileUrl;
            result.FileName = fileName;
            result.QRCodeUrl = $"{baseUrl}/qrcodes/{qrFileName}";
            result.ProcessedRows = students.Count();
            result.IsSuccess = true;
        }
        catch (Exception ex)
        {
            result.ErrorMessage = $"Export xatoligi: {ex.Message}";
            result.IsSuccess = false;
        }

        return result;
    }

    public async Task<ImportResult> ImportStudentsFromExcelAsync(IFormFile file)
    {
        var result = new ImportResult();

        if (!IsValidExcelFile(file))
        {
            result.ErrorMessage = "Noto'g'ri fayl formati. Faqat .xlsx fayllar qo'llab-quvvatlanadi.";
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
                result.ErrorMessage = "Import qilish uchun hech qanday to'g'ri ma'lumot topilmadi.";
            }
        }
        catch (Exception ex)
        {
            result.ErrorMessage = $"Excel faylini qayta ishlashda xatolik: {ex.Message}";
        }

        return result;
    }

    public byte[] GenerateImportSample()
    {
        var sampleStudents = new List<Student>
        {
            new Student { Id = 1, Name = "Ali Valiyev", DOB = DateTime.Now.AddYears(-20), Email = "ali@example.com", Mob = "+998901234567" },
            new Student { Id = 2, Name = "Madina Karimova", DOB = DateTime.Now.AddYears(-22), Email = "madina@example.com", Mob = "+998912345678" },
            new Student { Id = 3, Name = "Jasur Toshmatov", DOB = DateTime.Now.AddYears(-21), Email = "jasur@example.com", Mob = "+998923456789" },
            new Student { Id = 4, Name = "Nigora Ahmadova", DOB = DateTime.Now.AddYears(-19), Email = "nigora@example.com", Mob = "+998934567890" },
            new Student { Id = 5, Name = "Sardor Umarov", DOB = DateTime.Now.AddYears(-23), Email = "sardor@example.com", Mob = "+998945678901" }
        };

        return ExportStudentsToExcel(sampleStudents);
    }

    public byte[] GenerateQRCode(string url, int size = 300)
    {

        try
        {
            using var qrGenerator = new QRCodeGenerator();
            using var qrCodeData = qrGenerator.CreateQrCode(url, QRCodeGenerator.ECCLevel.Q);

            // ⚠️ Bu yerda QRCoder.QRCode dan foydalaning:
            using var qrCode = new QRCoder.QRCode(qrCodeData);
            using var qrCodeImage = qrCode.GetGraphic(20, Color.Black, Color.White, true);

            using var stream = new MemoryStream();
            qrCodeImage.Save(stream, ImageFormat.Png);
            return stream.ToArray();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"QR kod yaratishda xatolik: {ex.Message}");
        }
        /*try
        {
            using var qrGenerator = new QRCodeGenerator();
            using var qrCodeData = qrGenerator.CreateQrCode(url, QRCodeGenerator.ECCLevel.Q);
            using var qrCode = new QRCode(qrCodeData);
            using var qrCodeImage = qrCode.GetGraphic(20, Color.Black, Color.White, true);

            using var stream = new MemoryStream();
            qrCodeImage.Save(stream, ImageFormat.Png);
            return stream.ToArray();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"QR kod yaratishda xatolik: {ex.Message}");
        }*/
    }

    public async Task<QRCodeResult> GenerateQRCodeForFileAsync(string fileUrl)
    {
        var result = new QRCodeResult();

        try
        {
            var qrCodeBytes = GenerateQRCode(fileUrl);

            var qrFileName = $"qr_{DateTime.Now:yyyyMMdd_HHmmss}.png";
            var qrCodesPath = Path.Combine(_webHostEnvironment.WebRootPath, "qrcodes");

            if (!Directory.Exists(qrCodesPath))
            {
                Directory.CreateDirectory(qrCodesPath);
            }

            var qrFilePath = Path.Combine(qrCodesPath, qrFileName);
            await File.WriteAllBytesAsync(qrFilePath, qrCodeBytes);

            result.QRCodeData = qrCodeBytes;
            result.QRCodeUrl = $"/qrcodes/{qrFileName}";
            result.TargetUrl = fileUrl;
            result.IsSuccess = true;
        }
        catch (Exception ex)
        {
            result.ErrorMessage = $"QR kod yaratishda xatolik: {ex.Message}";
            result.IsSuccess = false;
        }

        return result;
    }

    private static bool IsValidExcelFile(IFormFile file)
    {
        return file != null &&
               file.Length > 0 &&
               file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
    }

    private static List<Student> ParseExcelFile(byte[] excelData, ImportResult result)
    {
        var students = new List<Student>();

        try
        {
            using var package = new ExcelPackage(new MemoryStream(excelData));
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();

            if (worksheet == null)
            {
                result.ErrorMessage = "Excel faylida hech qanday worksheet topilmadi.";
                return students;
            }

            var rowCount = worksheet.Dimension?.Rows ?? 0;

            if (rowCount < 2)
            {
                result.ErrorMessage = "Excel faylida ma'lumotlar topilmadi yoki faqat header mavjud.";
                return students;
            }

            for (int row = 2; row <= rowCount; row++)
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
                        result.Errors.Add($"Qator {row}: Noto'g'ri yoki to'liq bo'lmagan ma'lumot");
                    }
                }
                catch (Exception ex)
                {
                    result.SkippedRows++;
                    result.Errors.Add($"Qator {row}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            result.ErrorMessage = $"Excel faylini o'qishda xatolik: {ex.Message}";
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
            return DateTime.Now.AddYears(-18);

        // Try parsing different date formats
        var formats = new[] { "yyyy-MM-dd", "dd/MM/yyyy", "MM/dd/yyyy", "dd-MM-yyyy", "yyyy/MM/dd" };

        foreach (var format in formats)
        {
            if (DateTime.TryParseExact(dateString, format, null, System.Globalization.DateTimeStyles.None, out var date))
                return date;
        }

        if (DateTime.TryParse(dateString, out var parsedDate))
            return parsedDate;

        return DateTime.Now.AddYears(-18);
    }

    private static bool IsValidStudentData(Student student)
    {
        return !string.IsNullOrWhiteSpace(student.Name) &&
               !string.IsNullOrWhiteSpace(student.Email) &&
               IsValidEmail(student.Email) &&
               student.DOB <= DateTime.Now &&
               student.DOB >= DateTime.Now.AddYears(-100);
    }

    private static bool IsValidEmail(string email)
    {
        try
        {
            var addr = new System.Net.Mail.MailAddress(email);
            return addr.Address == email;
        }
        catch
        {
            return false;
        }
    }






    /*private readonly IStudentRepository _studentRepository;
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
    }*/
}
