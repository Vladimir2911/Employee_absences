using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Windows.Documents;

namespace Сalculating_employee_absences.Models
{
    public class Period
    {
        public int Id { get; set; }
        [Required]
        public string Reason { get; set; }
        public DateTime FirstDay { get; set; }       
        public int DaysCount { get; set; }
        public string? DateNote { get; set; }

        public int EmployeeId { get; set; }
        public Employee? Employee { get; set; }
    }
}
