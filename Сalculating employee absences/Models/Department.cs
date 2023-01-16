using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Сalculating_employee_absences.Models
{
    internal class Department
    {
        [Key]
        public int Id { get; set; }

        public string DepartmentName { get; set; }
        public List<Employee> Employees { get; set; }
    }
}
