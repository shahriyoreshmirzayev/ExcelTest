using ExcelTest1.Models;
using Microsoft.Data.SqlClient;
using System.Data;

namespace ExcelTest1.Repositories
{
    public class StudentRepository : IStudentRepository
    {
        private readonly string _connectionString;
        private const string TableName = "[TestExcel].[dbo].[Students]";

        public StudentRepository(string connectionString)
        {
            _connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
        }

        public async Task<IEnumerable<Student>> GetAllAsync()
        {
            var students = new List<Student>();
            const string query = $"SELECT [Id], [Name], [DOB], [Email], [Mob] FROM {TableName}";

            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            using var command = new SqlCommand(query, connection);
            using var reader = await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                students.Add(MapReaderToStudent(reader));
            }

            return students;
        }

        public async Task<Student?> GetByIdAsync(int id)
        {
            const string query = $"SELECT [Id], [Name], [DOB], [Email], [Mob] FROM {TableName} WHERE [Id] = @Id";

            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Id", id);

            using var reader = await command.ExecuteReaderAsync();

            if (await reader.ReadAsync())
            {
                return MapReaderToStudent(reader);
            }

            return null;
        }

        public async Task<Student> CreateAsync(Student student)
        {
            const string query = $@"
                INSERT INTO {TableName} ([Name], [DOB], [Email], [Mob]) 
                OUTPUT INSERTED.Id 
                VALUES (@Name, @DOB, @Email, @Mob)";

            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            using var command = new SqlCommand(query, connection);
            AddStudentParameters(command, student);

            var newId = (int)await command.ExecuteScalarAsync();
            student.Id = newId;

            return student;
        }

        public async Task<bool> UpdateAsync(Student student)
        {
            const string query = $@"
                UPDATE {TableName} 
                SET [Name] = @Name, [DOB] = @DOB, [Email] = @Email, [Mob] = @Mob 
                WHERE [Id] = @Id";

            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            using var command = new SqlCommand(query, connection);
            AddStudentParameters(command, student);
            command.Parameters.AddWithValue("@Id", student.Id);

            var rowsAffected = await command.ExecuteNonQueryAsync();
            return rowsAffected > 0;
        }

        public async Task<bool> DeleteAsync(int id)
        {
            const string query = $"DELETE FROM {TableName} WHERE [Id] = @Id";

            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Id", id);

            var rowsAffected = await command.ExecuteNonQueryAsync();
            return rowsAffected > 0;
        }

        public async Task<int> CreateBulkAsync(IEnumerable<Student> students)
        {
            var studentList = students.ToList();
            if (!studentList.Any())
                return 0;

            var dataTable = CreateStudentDataTable(studentList);

            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            using var bulkCopy = new SqlBulkCopy(connection)
            {
                DestinationTableName = TableName,
                BatchSize = 1000
            };

            bulkCopy.ColumnMappings.Add("Name", "Name");
            bulkCopy.ColumnMappings.Add("DOB", "DOB");
            bulkCopy.ColumnMappings.Add("Email", "Email");
            bulkCopy.ColumnMappings.Add("Mob", "Mob");

            await bulkCopy.WriteToServerAsync(dataTable);
            return studentList.Count;
        }

        private static Student MapReaderToStudent(SqlDataReader reader)
        {
            return new Student
            {
                Id = reader.GetInt32("Id"),
                Name = reader.GetString("Name"),
                DOB = reader.GetDateTime("DOB"),
                Email = reader.GetString("Email"),
                Mob = reader.GetString("Mob")
            };
        }

        private static void AddStudentParameters(SqlCommand command, Student student)
        {
            command.Parameters.AddWithValue("@Name", student.Name);
            command.Parameters.AddWithValue("@DOB", student.DOB);
            command.Parameters.AddWithValue("@Email", student.Email);
            command.Parameters.AddWithValue("@Mob", student.Mob);
        }

        private static DataTable CreateStudentDataTable(IEnumerable<Student> students)
        {
            var dataTable = new DataTable();
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("DOB", typeof(DateTime));
            dataTable.Columns.Add("Email", typeof(string));
            dataTable.Columns.Add("Mob", typeof(string));

            foreach (var student in students)
            {
                dataTable.Rows.Add(student.Name, student.DOB, student.Email, student.Mob);
            }

            return dataTable;
        }
    }
}
