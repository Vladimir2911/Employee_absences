﻿using Microsoft.EntityFrameworkCore;
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
          //  Database.SetCommandTimeout(600);
        }
        public DbSet<Employee> Employees { get; set; }
        public DbSet<Period> Periods { get; set; }

        //protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        //{
        //    optionsBuilder.UseSqlite("Data Source=EmployeeDatabase.mdf");
        //}
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("Server=(localdb)\\MSSQLLocalDB;Database=AnnaDatabase;Trusted_Connection=True");
        }
    }
}
