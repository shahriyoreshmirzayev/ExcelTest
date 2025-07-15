namespace ExcelTest1.Models
{
    public class ExcelExportRequest
    {
        public string? TableName { get; set; }
        public bool IncludeQR { get; set; } = false;
        public bool StyleHeader { get; set; } = true;
        public string? FilterBy { get; set; }
    }
}
