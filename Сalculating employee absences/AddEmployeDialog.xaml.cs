using Microsoft.IdentityModel.Tokens;
using System.Linq;
using System.Windows;
using Сalculating_employee_absences.Models;


namespace Сalculating_employee_absences
{
    /// <summary>
    /// Логика взаимодействия для AddEmployeDialog.xaml
    /// </summary>
    public partial class AddEmployeDialog : System.Windows.Window
    {
        MainWindow _window { get; set; }
        readonly bool _editMode;

        private Employee _employee;

        public AddEmployeDialog(MainWindow window)
        {
            _window = window;
            _editMode = false;
            InitializeComponent();
            ComboBox1.ItemsSource = StaticResourses.Departments;
        }

        public AddEmployeDialog(MainWindow window, Employee employee)
        {
            InitializeComponent();
            ComboBox1.ItemsSource = StaticResourses.Departments;
            _window = window;
            _employee = employee;
            _editMode = true;
            AddEmployeTextBox.Text = employee.Name;
            ComboBox1.SelectedItem = employee.Department;

        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
            _window.Close();
        }

        public void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_editMode)
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
                    AddNewEmployee();
                    Close();
                }
            }
            else
            {
                EditEmployee();
                Close();
            }
        }

        private void EditEmployee()
        {
            using (MyDbContext myDbContext = new MyDbContext())
            {
                Employee emp = myDbContext.Employees.First(n => n.Name == _employee.Name);     
                _window.Employees.FirstOrDefault(_employee).Name = AddEmployeTextBox.Text;
                _window.Employees.FirstOrDefault(_employee).Department = ComboBox1.Text;
                emp.Name = AddEmployeTextBox.Text;
                emp.Department = ComboBox1.Text;
                myDbContext.SaveChanges();
                _window.ListBoxEmployee.Items.Refresh(); 
            }
        }
        public void AddNewEmployee()
        {
            using (MyDbContext myDb = new MyDbContext())
            {
                Employee employee = new Employee();
                employee.Name = AddEmployeTextBox.Text;
                employee.Department = ComboBox1.Text;
                _window.Employees.Add(employee);
                MessageBox.Show("Добавлен " + employee.Name + " \n" + employee.Department);
                myDb.Employees.Add(employee);
                myDb.SaveChanges();
                _window.ListBoxEmployee.Items.Refresh();
            }
        }
    }
}
