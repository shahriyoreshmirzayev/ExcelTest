using ExcelTest1.Models;

namespace ExcelTest1.Services;

public interface IExcelService
{
    byte[] ExportStudentsToExcel(IEnumerable<Student> students);
    Task<ImportResult> ImportStudentsFromExcelAsync(IFormFile file);
    Task<List<Student>> GetAllAsync();
    Task<ExcelExportWithQRResult> ExportStudentsToExcelWithQRAsync(IEnumerable<Student> students, string baseUrl);
    byte[] GenerateImportSample();
    Task<QRCodeResult> GenerateQRCodeForFileAsync(string url);
}
