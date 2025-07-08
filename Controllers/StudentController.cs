using ExcelTest1.Data;
using ExcelTest1.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
namespace ExcelTest1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class StudentController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly ExcelExporterService _excelExporter;

        public StudentController(ApplicationDbContext context)
        {
            _context = context;
            _excelExporter = new ExcelExporterService(); // yoki DI orqali uzating
        }

        [HttpGet("export")]
        public async Task<IActionResult> ExportToExcel()
        {
            var students = await _context.Students.ToListAsync();

            if (!students.Any())
                return NotFound("Bazadan studentlar topilmadi.");

            var fileContents = _excelExporter.ExportStudentsToExcel(students);
            return File(fileContents,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "students.xlsx");
        }
    }
}
