using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Сalculating_employee_absences.Models
{
    public class Employee
    {
        public Employee()
        {
            Absence = new List<Period>();
        }
        public int Id { get; set; }
        [Required]
        public string Name { get; set; }
        [Required]
        public string Department { get; set; }
        public int PeriodId { get; set; }
        
        public List<Period> Absence { get; set; }       

        public override string ToString()
        {
            return this.Name;
        }

    }
}
