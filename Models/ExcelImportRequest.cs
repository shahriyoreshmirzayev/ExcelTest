namespace ExcelTest1.Models
{
    public class ExcelImportRequest
    {
        public IFormFile File { get; set; } = null!;
        public bool SkipErrors { get; set; } = true;
        public bool UpdateExisting { get; set; } = false;
    }
}
