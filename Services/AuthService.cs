using ExcelTest1.DTOs;
using ExcelTest1.Models;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Security.Cryptography;
using System.Text;

namespace ExcelTest1.Services
{

    public class AuthService
    {
        private readonly string _connectionString;
        private readonly JwtTokenService _jwtService;

        public AuthService(string connectionString, JwtTokenService jwtService)
        {
            _connectionString = connectionString;
            _jwtService = jwtService;
        }

        public async Task<AuthResponseDto?> LoginAsync(LoginDto loginDto)
        {
            var user = await GetUserByUsernameAsync(loginDto.Username);

            if (user == null || !VerifyPassword(loginDto.Password, user.PasswordHash))
            {
                return null;
            }

            if (!user.IsActive)
            {
                return null;
            }

            var token = _jwtService.GenerateToken(user);

            return new AuthResponseDto
            {
                Token = token,
                Username = user.Username,
                Email = user.Email,
                Role = user.Role,
                Expires = DateTime.UtcNow.AddMinutes(60)
            };
        }

        public async Task<AuthResponseDto?> RegisterAsync(RegisterDto registerDto)
        {
            // Foydalanuvchi mavjudligini tekshirish
            if (await IsUsernameExistsAsync(registerDto.Username))
            {
                throw new InvalidOperationException("Username allaqachon mavjud");
            }

            if (await IsEmailExistsAsync(registerDto.Email))
            {
                throw new InvalidOperationException("Email allaqachon mavjud");
            }

            // Yangi foydalanuvchi yaratish
            var user = new User
            {
                Username = registerDto.Username,
                Email = registerDto.Email,
                PasswordHash = HashPassword(registerDto.Password),
                Role = "User",
                CreatedAt = DateTime.Now,
                IsActive = true
            };

            var userId = await CreateUserAsync(user);
            user.Id = userId;

            var token = _jwtService.GenerateToken(user);

            return new AuthResponseDto
            {
                Token = token,
                Username = user.Username,
                Email = user.Email,
                Role = user.Role,
                Expires = DateTime.UtcNow.AddMinutes(60)
            };
        }

        private async Task<User?> GetUserByUsernameAsync(string username)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            var query = "SELECT * FROM [TestExcel].[dbo].[Users] WHERE [Username] = @Username";
            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Username", username);

            using var reader = await command.ExecuteReaderAsync();

            if (await reader.ReadAsync())
            {
                return new User
                {
                    Id = reader.GetInt32("Id"),
                    Username = reader.GetString("Username"),
                    Email = reader.GetString("Email"),
                    PasswordHash = reader.GetString("PasswordHash"),
                    Role = reader.GetString("Role"),
                    CreatedAt = reader.GetDateTime("CreatedAt"),
                    IsActive = reader.GetBoolean("IsActive")
                };
            }

            return null;
        }

        private async Task<bool> IsUsernameExistsAsync(string username)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            var query = "SELECT COUNT(*) FROM [TestExcel].[dbo].[Users] WHERE [Username] = @Username";
            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Username", username);

            var count = (int)await command.ExecuteScalarAsync();
            return count > 0;
        }

        private async Task<bool> IsEmailExistsAsync(string email)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            var query = "SELECT COUNT(*) FROM [TestExcel].[dbo].[Users] WHERE [Email] = @Email";
            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Email", email);

            var count = (int)await command.ExecuteScalarAsync();
            return count > 0;
        }

        private async Task<int> CreateUserAsync(User user)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();

            var query = @"
            INSERT INTO [TestExcel].[dbo].[Users] ([Username], [Email], [PasswordHash], [Role], [CreatedAt], [IsActive])
            OUTPUT INSERTED.Id
            VALUES (@Username, @Email, @PasswordHash, @Role, @CreatedAt, @IsActive)";

            using var command = new SqlCommand(query, connection);
            command.Parameters.AddWithValue("@Username", user.Username);
            command.Parameters.AddWithValue("@Email", user.Email);
            command.Parameters.AddWithValue("@PasswordHash", user.PasswordHash);
            command.Parameters.AddWithValue("@Role", user.Role);
            command.Parameters.AddWithValue("@CreatedAt", user.CreatedAt);
            command.Parameters.AddWithValue("@IsActive", user.IsActive);

            return (int)await command.ExecuteScalarAsync();
        }

        private string HashPassword(string password)
        {
            using var sha256 = SHA256.Create();
            var hashedBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
            return Convert.ToBase64String(hashedBytes);
        }

        private bool VerifyPassword(string password, string hashedPassword)
        {
            var hashedInput = HashPassword(password);
            return hashedInput == hashedPassword;
        }
    }
}
