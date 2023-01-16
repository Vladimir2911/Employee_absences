using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Сalculating_employee_absences.Models
{
    internal class Employee
    {
        public int Id { get; set; }
        [Required]
        public string Name { get; set; }
        [Required]
        public string Department { get; set; }

        public Dictionary<DateOnly, string> AbsentDay = new Dictionary<DateOnly, string>();
        public override string ToString()
        {
            return this.Name;
        }

    }
}
