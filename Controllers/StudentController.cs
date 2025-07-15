using ExcelTest1.Models;
using ExcelTest1.Services;
using Microsoft.AspNetCore.Mvc;
namespace ExcelTest1.Controllers;

[Route("api/[controller]")]
[ApiController]
public class StudentController : ControllerBase
{
    private readonly IStudentService _studentService;
    private readonly IExcelService _excelService;

    public StudentController(IStudentService studentService, IExcelService excelService)
    {
        _studentService = studentService;
        _excelService = excelService;
    }

    // Mavjud CRUD operatsiyalaringiz...

    [HttpGet("export-excel")]
    public async Task<IActionResult> ExportStudentsToExcel()
    {
        try
        {
            var students = await _studentService.GetAllAsync();
            var excelData = _excelService.ExportStudentsToExcel(students);

            var fileName = $"students_export_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

            return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }

    [HttpGet("export-excel-with-qr")]
    public async Task<IActionResult> ExportStudentsToExcelWithQR()
    {
        try
        {
            var students = await _studentService.GetAllAsync();
            var baseUrl = $"{Request.Scheme}://{Request.Host}";

            var result = await _excelService.ExportStudentsToExcelWithQRAsync(students, baseUrl);

            if (!result.IsSuccess)
            {
                return BadRequest(new { error = result.ErrorMessage });
            }

            return Ok(new
            {
                success = true,
                message = "Excel fayl va QR kod muvaffaqiyatli yaratildi",
                data = new
                {
                    fileUrl = result.FileUrl,
                    fileName = result.FileName,
                    qrCodeUrl = result.QRCodeUrl,
                    qrCodeBase64 = Convert.ToBase64String(result.QRCodeData ?? Array.Empty<byte>()),
                    processedRows = result.ProcessedRows,
                    createdAt = result.CreatedAt
                }
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }

    [HttpPost("import-excel")]
    public async Task<IActionResult> ImportStudentsFromExcel([FromForm] ExcelImportRequest request)
    {
        try
        {
            if (request.File == null || request.File.Length == 0)
            {
                return BadRequest(new { error = "Fayl yuklanmadi" });
            }

            var result = await _excelService.ImportStudentsFromExcelAsync(request.File);

            if (!result.IsSuccess)
            {
                return BadRequest(new
                {
                    error = result.ErrorMessage,
                    details = result.Errors.Take(10).ToArray() // Faqat birinchi 10 ta xatoni ko'rsatish
                });
            }

            return Ok(new
            {
                success = true,
                message = "Ma'lumotlar muvaffaqiyatli import qilindi",
                data = new
                {
                    processedRows = result.ProcessedRows,
                    skippedRows = result.SkippedRows,
                    totalErrors = result.Errors.Count,
                    errors = result.Errors.Take(5).ToArray(), // Birinchi 5 ta xato
                    importedAt = result.ImportedAt
                }
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }

    [HttpGet("download-import-sample")]
    public IActionResult DownloadImportSample()
    {
        try
        {
            var sampleData = _excelService.GenerateImportSample();
            var fileName = $"student_import_sample_{DateTime.Now:yyyyMMdd}.xlsx";

            return File(sampleData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }

    [HttpPost("generate-qr")]
    public async Task<IActionResult> GenerateQRCode([FromBody] QRCodeRequest request)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(request.Url))
            {
                return BadRequest(new { error = "URL bo'sh bo'lishi mumkin emas" });
            }

            var result = await _excelService.GenerateQRCodeForFileAsync(request.Url);

            if (!result.IsSuccess)
            {
                return BadRequest(new { error = result.ErrorMessage });
            }

            return Ok(new
            {
                success = true,
                message = "QR kod muvaffaqiyatli yaratildi",
                data = new
                {
                    qrCodeBase64 = Convert.ToBase64String(result.QRCodeData ?? Array.Empty<byte>()),
                    qrCodeUrl = result.QRCodeUrl,
                    targetUrl = result.TargetUrl,
                    createdAt = result.CreatedAt
                }
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }

    [HttpGet("qr-code/{fileName}")]
    public IActionResult GetQRCode(string fileName)
    {
        try
        {
            var qrCodesPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "qrcodes");
            var filePath = Path.Combine(qrCodesPath, fileName);

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound(new { error = "QR kod fayli topilmadi" });
            }

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "image/png");
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }

    [HttpGet("download/{fileName}")]
    public IActionResult DownloadFile(string fileName)
    {
        try
        {
            var downloadsPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "downloads");
            var filePath = Path.Combine(downloadsPath, fileName);

            if (!System.IO.File.Exists(filePath))
            {
                return NotFound(new { error = "Fayl topilmadi" });
            }

            var fileBytes = System.IO.File.ReadAllBytes(filePath);
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }

    [HttpDelete("cleanup-files")]
    public IActionResult CleanupOldFiles()
    {
        try
        {
            var deletedFiles = 0;
            var cutoffDate = DateTime.Now.AddHours(-24); // 24 soat eski fayllarni o'chirish

            // Downloads papkasini tozalash
            var downloadsPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "downloads");
            if (Directory.Exists(downloadsPath))
            {
                var files = Directory.GetFiles(downloadsPath);
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    if (fileInfo.CreationTime < cutoffDate)
                    {
                        System.IO.File.Delete(file);
                        deletedFiles++;
                    }
                }
            }

            // QR codes papkasini tozalash
            var qrCodesPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "qrcodes");
            if (Directory.Exists(qrCodesPath))
            {
                var files = Directory.GetFiles(qrCodesPath);
                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    if (fileInfo.CreationTime < cutoffDate)
                    {
                        System.IO.File.Delete(file);
                        deletedFiles++;
                    }
                }
            }

            return Ok(new
            {
                success = true,
                message = $"{deletedFiles} ta eski fayl o'chirildi",
                deletedFiles = deletedFiles
            });
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }









    /*
    private readonly IStudentService _studentService;
    private readonly IExcelService _excelService;

    public StudentController(IStudentService studentService, IExcelService excelService)
    {
        _studentService = studentService;
        _excelService = excelService;
    }

    [HttpGet("[action]")]
    [Authorize(Roles = "Admin")]
    public async Task<ActionResult<IEnumerable<Student>>> GetAllStudents()
    {
        var students = await _studentService.GetAllStudentsAsync();
        return Ok(students);
    }

    [HttpGet("[action]/{id}")]
    public async Task<ActionResult<Student>> GetStudent(int id)
    {
        var student = await _studentService.GetStudentByIdAsync(id);
        if (student == null)
            return NotFound();

        return Ok(student);
    }

    [HttpPost("[action]")]
    [Authorize(Roles = "Admin,User")]
    public async Task<ActionResult<Student>> CreateStudent(Student student)
    {
        if (!ModelState.IsValid)
            return BadRequest(ModelState);

        var createdStudent = await _studentService.CreateStudentAsync(student);
        return CreatedAtAction(nameof(GetStudent), new { id = createdStudent.Id }, createdStudent);
    }

    [HttpPut("[action]/{id}")]
    [Authorize(Roles = "Admin,User")]
    public async Task<IActionResult> UpdateStudent(int id, Student student)
    {
        if (!ModelState.IsValid)
            return BadRequest(ModelState);

        var updated = await _studentService.UpdateStudentAsync(id, student);
        if (!updated)
            return NotFound();

        return NoContent();
    }

    [HttpDelete("[action]/{id}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> DeleteStudent(int id)
    {
        var deleted = await _studentService.DeleteStudentAsync(id);
        if (!deleted)
            return NotFound();

        return NoContent();
    }

    [HttpPost("[action]")]
    [Authorize(Roles = "Admin,User")]
    public async Task<IActionResult> ExportToExcel()
    {
        var students = await _studentService.GetAllStudentsAsync();

        if (!students.Any())
            return BadRequest("Ma'lumotlar topilmadi");

        var excelData = _excelService.ExportStudentsToExcel(students);
        var fileName = $"Students_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

        return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    [HttpPost("[action]")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> ImportFromExcel(IFormFile file)
    {
        if (file == null || file.Length == 0)
            return BadRequest("Fayl tanlanmagan");

        var result = await _excelService.ImportStudentsFromExcelAsync(file);

        if (result.IsSuccess)
        {
            return Ok(new
            {
                message = "Ma'lumotlar muvaffaqiyatli import qilindi",
                processedRows = result.ProcessedRows,
                skippedRows = result.SkippedRows,
                errors = result.Errors
            });
        }

        return BadRequest(new
        {
            message = result.ErrorMessage,
            errors = result.Errors
        });
    }

    //[HttpGet("import/sample")]
    //[Authorize(Roles = "Admin,User")]
    //public IActionResult GenerateImportSample()
    //{
    //    var excelData = _excelService.GenerateImportSample();
    //    return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Import_Sample.xlsx");
    //}

    [HttpGet("[action]")]
    [Authorize]
    public IActionResult GetUserInfo()
    {
        var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
        var username = User.FindFirst(ClaimTypes.Name)?.Value;

        return Ok(new { UserId = userId, Username = username });
    }*/
}




/*[Route("api/[controller]")]
[ApiController]
public class StudentController : ControllerBase
{
    private readonly ExcelExporterService _excelExporter;
    private readonly ExcelImporterService _excelImporter;
    private readonly string _connectionString;

    public StudentController(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DefaultConnection");
        _excelExporter = new ExcelExporterService();
        _excelImporter = new ExcelImporterService(_connectionString);
    }

    // Mavjud metodlaringiz (CRUD operatsiyalari)
    [HttpGet("[action]")]
    public async Task<ActionResult<List<Student>>> GetAllStudents()     // GetStudents
    {
        var userId = User.FindFirst(ClaimTypes.NameIdentifier)?.Value;
        var username = User.FindFirst(ClaimTypes.Name)?.Value;

        var students = new List<Student>();

        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        var query = "SELECT [Id], [Name], [DOB], [Email], [Mob] FROM [TestExcel].[dbo].[Students]";
        using var command = new SqlCommand(query, connection);
        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            students.Add(new Student
            {
                Id = reader.GetInt32("Id"),
                Name = reader.GetString("Name"),
                DOB = reader.GetDateTime("DOB"),
                Email = reader.GetString("Email"),
                Mob = reader.GetString("Mob")
            });
        }

        return Ok(students);
    }

    [HttpGet("[action]")]
    public async Task<ActionResult<Student>> GetStudentById(int id)
    {
        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        var query = "SELECT [Id], [Name], [DOB], [Email], [Mob] FROM [TestExcel].[dbo].[Students] WHERE Id = @Id";
        using var command = new SqlCommand(query, connection);
        command.Parameters.AddWithValue("@Id", id);

        using var reader = await command.ExecuteReaderAsync();

        if (await reader.ReadAsync())
        {
            var student = new Student
            {
                Id = reader.GetInt32("Id"),
                Name = reader.GetString("Name"),
                DOB = reader.GetDateTime("DOB"),
                Email = reader.GetString("Email"),
                Mob = reader.GetString("Mob")
            };
            return Ok(student);
        }
        return NotFound();
    }

    [HttpPost("[action]")]
    [Authorize(Roles = "Admin,User")]
    public async Task<ActionResult<Student>> CreateStudent(Student student)
    {
        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        var query = @"INSERT INTO [TestExcel].[dbo].[Students] ([Name], [DOB], [Email], [Mob]) 
                     OUTPUT INSERTED.Id VALUES (@Name, @DOB, @Email, @Mob)";

        using var command = new SqlCommand(query, connection);
        command.Parameters.AddWithValue("@Name", student.Name);
        command.Parameters.AddWithValue("@DOB", student.DOB);
        command.Parameters.AddWithValue("@Email", student.Email);
        command.Parameters.AddWithValue("@Mob", student.Mob);

        var newId = (int)await command.ExecuteScalarAsync();
        student.Id = newId;

        return CreatedAtAction(nameof(GetStudentById), new { id = student.Id }, student);
    }

    [HttpPut("[action]")]
    [Authorize(Roles = "Admin,User")]
    public async Task<IActionResult> UpdateStudent(int id, Student student)
    {
        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        var query = @"UPDATE [TestExcel].[dbo].[Students] 
                     SET [Name] = @Name, [DOB] = @DOB, [Email] = @Email, [Mob] = @Mob 
                     WHERE [Id] = @Id";

        using var command = new SqlCommand(query, connection);
        command.Parameters.AddWithValue("@Id", id);
        command.Parameters.AddWithValue("@Name", student.Name);
        command.Parameters.AddWithValue("@DOB", student.DOB);
        command.Parameters.AddWithValue("@Email", student.Email);
        command.Parameters.AddWithValue("@Mob", student.Mob);

        var rowsAffected = await command.ExecuteNonQueryAsync();

        if (rowsAffected == 0)
        {
            return NotFound();
        }

        return Ok(student);
    }

    [HttpDelete("[action]")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> DeleteStudent(int id)
    {
        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        var query = "DELETE FROM [TestExcel].[dbo].[Students] WHERE [Id] = @Id";
        using var command = new SqlCommand(query, connection);
        command.Parameters.AddWithValue("@Id", id);

        var rowsAffected = await command.ExecuteNonQueryAsync();

        if (rowsAffected == 0)
        {
            return NotFound();
        }

        return NoContent();
    }

    // YANGI: Excel Export/Import metodlari
    [HttpPost("[action]")]
    [Authorize(Roles = "Admin,User")]
    public async Task<IActionResult> ExportToExcel()
    {
        var students = await GetStudentsFromDatabase();

        if (!students.Any())
        {
            return BadRequest("Ma'lumotlar topilmadi");
        }

        var excelData = _excelExporter.ExportStudentsToExcel(students);

        return File(excelData,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            $"Students_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
    }

    [HttpPost("[action]")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> ImportFromExcel(IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest("Fayl tanlanmagan");
        }

        var result = await _excelImporter.ImportStudentsFromExcelFile(file);

        if (result.IsSuccess)
        {
            return Ok(new
            {
                message = "Ma'lumotlar muvaffaqiyatli import qilindi",
                processedRows = result.ProcessedRows,
                skippedRows = result.SkippedRows,
                errors = result.Errors
            });
        }
        else
        {
            return BadRequest(new
            {
                message = result.ErrorMessage,
                errors = result.Errors
            });
        }
    }

    [HttpGet("[action]")]
    [Authorize(Roles = "Admin,User")]
    public async Task<IActionResult> GenerateImportSample()
    {
        var sampleStudents = new List<Student>
        {
        new Student { Name = "Namuna Talaba 1", DOB = DateTime.Now.AddYears(-20), Email = "namuna1@email.com", Mob = "+998901234567" },
        new Student { Name = "Namuna Talaba 2", DOB = DateTime.Now.AddYears(-22), Email = "namuna2@email.com", Mob = "+998912345678" },
        new Student { Name = "Namuna Talaba 3", DOB = DateTime.Now.AddYears(-21), Email = "namuna3@email.com", Mob = "+998923456789" }
        };

        var excelData = _excelExporter.ExportStudentsToExcel(sampleStudents);

        return File(excelData, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Import_Sample.xlsx");
    }

    // Helper metod
    private async Task<List<Student>> GetStudentsFromDatabase()
    {
        var students = new List<Student>();

        using var connection = new SqlConnection(_connectionString);
        await connection.OpenAsync();

        var query = "SELECT [Id], [Name], [DOB], [Email], [Mob] FROM [TestExcel].[dbo].[Students]";
        using var command = new SqlCommand(query, connection);
        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            students.Add(new Student
            {
                Id = reader.GetInt32("Id"),
                Name = reader.GetString("Name"),
                DOB = reader.GetDateTime("DOB"),
                Email = reader.GetString("Email"),
                Mob = reader.GetString("Mob")
            });
        }

        return students;
    }
}*/
