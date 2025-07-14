using ExcelTest1.Models;

namespace ExcelTest1.Services;

public interface IExcelService
{
    byte[] ExportStudentsToExcel(IEnumerable<Student> students);
    Task<ImportResult> ImportStudentsFromExcelAsync(IFormFile file);
    //byte[] GenerateImportSample();
}
