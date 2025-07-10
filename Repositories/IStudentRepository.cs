using ExcelTest1.Models;

namespace ExcelTest1.Repositories;

public interface IStudentRepository
{
    Task<IEnumerable<Student>> GetAllAsync();
    Task<Student?> GetByIdAsync(int id);
    Task<Student> CreateAsync(Student student);
    Task<bool> UpdateAsync(Student student);
    Task<bool> DeleteAsync(int id);
    Task<int> CreateBulkAsync(IEnumerable<Student> students);
}
