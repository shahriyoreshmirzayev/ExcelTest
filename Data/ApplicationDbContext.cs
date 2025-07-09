using ExcelTest1.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelTest1.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {
        }

        public DbSet<Student> Students { get; set; }
        public DbSet<User> Users { get; set; }
    }
}
