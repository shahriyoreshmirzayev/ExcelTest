using DocumentFormat.OpenXml.InkML;
using ExcelTest1.Data;
using ExcelTest1.Models;
using ExcelTest1.Repositories;

namespace ExcelTest1.Services;

public class StudentService : IStudentService
{
    private readonly IStudentRepository _studentRepository;
    private readonly ApplicationDbContext _context;
    public StudentService(IStudentRepository studentRepository, ApplicationDbContext context)
    {
        _studentRepository = studentRepository;
        _context = context;
    }

    public async Task<IEnumerable<Student>> GetAllStudentsAsync()
    {
        return await _studentRepository.GetAllAsync();
    }

    public async Task<Student?> GetStudentByIdAsync(int id)
    {
        if (id <= 0)
            return null;

        return await _studentRepository.GetByIdAsync(id);
    }

    public async Task<Student> CreateStudentAsync(Student student)
    {
        ValidateStudent(student);
        return await _studentRepository.CreateAsync(student);
    }

    public async Task<bool> UpdateStudentAsync(int id, Student student)
    {
        if (id <= 0)
            return false;

        ValidateStudent(student);
        student.Id = id;
        return await _studentRepository.UpdateAsync(student);
    }

    public async Task<bool> DeleteStudentAsync(int id)
    {
        if (id <= 0)
            return false;

        return await _studentRepository.DeleteAsync(id);
    }

    private static void ValidateStudent(Student student)
    {
        if (student == null)
            throw new ArgumentNullException(nameof(student));

        if (string.IsNullOrWhiteSpace(student.Name))
            throw new ArgumentException("Student name is required", nameof(student.Name));

        if (string.IsNullOrWhiteSpace(student.Email))
            throw new ArgumentException("Student email is required", nameof(student.Email));

        if (!IsValidEmail(student.Email))
            throw new ArgumentException("Invalid email format", nameof(student.Email));

        if (student.DOB > DateTime.Now)
            throw new ArgumentException("Birth date cannot be in the future", nameof(student.DOB));
    }

    private static bool IsValidEmail(string email)
    {
        try
        {
            var addr = new System.Net.Mail.MailAddress(email);
            return addr.Address == email;
        }
        catch
        {
            return false;
        }
    }
}
