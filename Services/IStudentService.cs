using ExcelTest1.Models;

namespace ExcelTest1.Services;

public interface IStudentService
{
    Task<IEnumerable<Student>> GetAllStudentsAsync();
    Task<Student?> GetStudentByIdAsync(int id);
    Task<Student> CreateStudentAsync(Student student);
    Task<bool> UpdateStudentAsync(int id, Student student);
    Task<bool> DeleteStudentAsync(int id);
}
