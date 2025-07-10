namespace ExcelTest1.Models;

public class ImportResult
{
    public bool IsSuccess { get; set; }
    public int ProcessedRows { get; set; }
    public int SkippedRows { get; set; }
    public string ErrorMessage { get; set; } = string.Empty;
    public List<string> Errors { get; set; } = new();
}
