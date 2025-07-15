namespace ExcelTest1.Models;

public class ExportResult
{
    public bool IsSuccess { get; set; }
    public string? ErrorMessage { get; set; }
    public byte[]? ExcelData { get; set; }
    public byte[]? QRCodeData { get; set; }
    public string? FileUrl { get; set; }
    public string? FileName { get; set; }
    public string? QRCodeUrl { get; set; }
    public int ProcessedRows { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}
