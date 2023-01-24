using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data;
using System.Linq;
using System.Printing.IndexedProperties;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

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

        public int? EmployeeId { get; set; }        
        public Employee Employee { get; set; }

    }
}
