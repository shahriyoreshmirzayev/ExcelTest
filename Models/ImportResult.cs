namespace ExcelTest1.Models;

public class ImportResult
{
    public bool IsSuccess { get; set; }
    public int ProcessedRows { get; set; }
    public int SkippedRows { get; set; }
    public string ErrorMessage { get; set; } = string.Empty;
    public List<string> Errors { get; set; } = new();

    public DateTime ImportedAt { get; set; } = DateTime.Now;
    public string? ImportedBy { get; set; }
    public List<Student> ImportedStudents { get; set; } = new(); // Qo'shilgan

}
