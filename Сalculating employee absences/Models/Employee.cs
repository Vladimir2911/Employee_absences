using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Сalculating_employee_absences.Models
{
    public class Employee : INotifyPropertyChanged
    {
        public Employee()
        {
            Periods = new List<Period>();
        }
        private string _name;
        private string _department;
        public int Id { get; set; }
        [Required]
        public string Name { get { return _name; } set { _name = value; OnPropertyChanged("Name"); } }


        [Required]
        public string Department { get { return _department; } set{_department=value; OnPropertyChanged("Department"); } }

        public int PeriodId { get; set; }

        public List<Period> Periods { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }

        public override string ToString()
        {            
            return Name;
        }

    }
}
