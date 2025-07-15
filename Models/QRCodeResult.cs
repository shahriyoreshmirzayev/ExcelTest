namespace ExcelTest1.Models
{
    public class QRCodeResult
    {
        public bool IsSuccess { get; set; }
        public string? ErrorMessage { get; set; }
        public byte[]? QRCodeData { get; set; }
        public string? QRCodeUrl { get; set; }
        public string? TargetUrl { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.Now;
    }
}
