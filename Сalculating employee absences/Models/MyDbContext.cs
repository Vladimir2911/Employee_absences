using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace Сalculating_employee_absences.Models
{
    internal class MyDbContext : DbContext
    {
        public MyDbContext()
        {
            Database.EnsureCreated();
        }
        public DbSet<Employee> Employees { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlite("Data Source=EmployeeDatabase.mdf");
        }
       
    }
}
