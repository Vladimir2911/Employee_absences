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
using System.Windows.Shapes;
using Сalculating_employee_absences.Models;

namespace Сalculating_employee_absences
{
    /// <summary>
    /// Логика взаимодействия для AddEmployeDialog.xaml
    /// </summary>
    public partial class AddEmployeDialog : Window
    {
        public AddEmployeDialog()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        { 
            Employee employee = new Employee();
            employee.Name=AddEmployeTextBox.Text;
            MessageBox.Show(employee.Name);
            using (MyDbContext myDb = new MyDbContext())
            {
                myDb.Employees.Add(employee);
               
            }
            MessageBox.Show("Done!!!");
           
        }
    }
}
