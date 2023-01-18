using Microsoft.EntityFrameworkCore;

namespace excel_generator.Models
{
    public class StoreContext : DbContext
    {
        public StoreContext(DbContextOptions<StoreContext> options)
            : base(options)
        {

        }

        public DbSet<Student> Students { get; set; }

    }
}