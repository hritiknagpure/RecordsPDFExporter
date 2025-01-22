using Microsoft.EntityFrameworkCore;
using WebApiWithPdf.Models;

namespace WebApiWithPdf.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options) { }

        public DbSet<Record> Records { get; set; }
    }
}