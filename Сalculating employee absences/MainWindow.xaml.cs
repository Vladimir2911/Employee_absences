using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Сalculating_employee_absences.Models;

namespace Сalculating_employee_absences
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            SetDefaultDates();
        }

        private void SetDefaultDates()
        {
            var year = DateTime.Now.Year;
            CalendarJanuary.DisplayDate = Convert.ToDateTime("01/01/" + year);
            CalendarFabruary.DisplayDate = Convert.ToDateTime("01/02/" + year);
            CalendarMarch.DisplayDate = Convert.ToDateTime("01/03/" + year);
            CalendarApril.DisplayDate = Convert.ToDateTime("01/04/" + year);
            CalendarMay.DisplayDate = Convert.ToDateTime("01/05/" + year);
            CalendarJune.DisplayDate = Convert.ToDateTime("01/06/" + year);
            CalendarJuly.DisplayDate = Convert.ToDateTime("01/07/" + year);
            CalendarAugust.DisplayDate = Convert.ToDateTime("01/08/" + year);
            CalendarSeptember.DisplayDate = Convert.ToDateTime("01/09/" + year);
            CalendarOctober.DisplayDate = Convert.ToDateTime("01/10/" + year);
            CalendarNovember.DisplayDate = Convert.ToDateTime("01/11/" + year);
            CalendarDesember.DisplayDate = Convert.ToDateTime("01/12/" + year);

        }

        private void AddEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            AddEmployeDialog addEmployeDialog = new AddEmployeDialog();
            
            addEmployeDialog.Show();
            LoadData();
        }
    
        public void LoadData()
        {
            using (MyDbContext myDb = new MyDbContext())
            {              
                ListBoxEmployee.ItemsSource = myDb.Employees.ToList().OrderBy(x=>x.Name);              
            }
        }

        private void RemuveEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            LoadData();
        }
    }
}
