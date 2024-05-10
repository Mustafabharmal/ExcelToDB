using System.Data.Entity;

namespace xltodb.Models
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext() : base("Data Source=LAPTOP-FKPSL27T;Initial Catalog=SpclDB;Integrated Security=True")
        {
        }

        public DbSet<Employee> Employees { get; set; }
    }
}

