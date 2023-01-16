using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Сalculating_employee_absences.Models
{
    internal class Employee
    {
        public int Id { get; set; }
        public string Name { get; set; }

        public Department EmployeeDepartment { get; set; }
        public Dictionary<DateOnly, string> AbsentDay = new Dictionary<DateOnly, string>();
        
    }
}
