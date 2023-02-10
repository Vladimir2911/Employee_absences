using Microsoft.IdentityModel.Tokens;
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
        bool editMode = false;
        Employee employee;
        public AddEmployeDialog()
        {
            editMode = false;
            InitializeComponent();
            ComboBox1.ItemsSource = StaticResourses.Departments;
        }
        public AddEmployeDialog(Employee employee)
        {
            InitializeComponent(); 
            ComboBox1.ItemsSource = StaticResourses.Departments;
            this.employee = employee;
            editMode = true;
            AddEmployeTextBox.Text = employee.Name;
            ComboBox1.SelectedItem = employee.Department;           
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        public async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (!editMode)
            {
                if (AddEmployeTextBox.Text.IsNullOrEmpty())
                {
                    MessageBox.Show("Введите имя сотрудника");
                    return;
                }
                else if (ComboBox1.Text == StaticResourses.Departments[0])
                {
                    MessageBox.Show("не выбран отдел!!");
                    return;
                }
                else
                {
                    Click();
                    this.Close();
                }
            }
            else
            {
                EditEmployee();
                this.Close();
            }
        }

        private void EditEmployee()
        {
            Button1.Content = "Изменить";
            using (MyDbContext myDbContext = new MyDbContext())
            {
                Employee emp = myDbContext.Employees.First(n => n.Name == employee.Name);
                emp.Name = AddEmployeTextBox.Text;
                emp.Department = ComboBox1.Text;
                myDbContext.SaveChanges();
            }
        }

        public async void Click()
        {
            Employee employee = new Employee();
            employee.Name = AddEmployeTextBox.Text;
            employee.Department = ComboBox1.Text;
            MessageBox.Show("Добавлен " + employee.Name + " \n" + employee.Department);

            using (MyDbContext myDb = new MyDbContext())
            {
                myDb.Employees.Add(employee);
                await myDb.SaveChangesAsync();
            }

        }

    }
}
