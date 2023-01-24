using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Conventions;
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
            RefreshWindows();
            LoadData();
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
                ListBoxEmployee.ItemsSource = myDb.Employees.ToList().OrderBy(x => x.Name);
            }
        }

        private void RemuveEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            using (MyDbContext myDb = new MyDbContext())
            {
                if (ListBoxEmployee.SelectedItem != null)
                {
                    var itemToDelete = ListBoxEmployee.SelectedItem.ToString();
                    Employee employee = myDb.Employees.FirstOrDefault(x => x.Name == itemToDelete);
                    if (employee != null)
                    {
                        myDb.Employees.Remove(employee);
                        myDb.SaveChanges();
                        MessageBox.Show("Завпись удалена!");
                    }
                    else
                    {
                        MessageBox.Show("Ошибка");
                    }
                }
            }
            LoadData();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            LoadData();
        }
        private void MenuHealthReason_Click(object sender, RoutedEventArgs e)
        {
            InsertDataToDb("HealthReason");
            MessageBox.Show("Health reason");
        }
        private void MenuFamilyReason_Click(object sender, RoutedEventArgs e)
        {
            InsertDataToDb("FamilyReason");
            MessageBox.Show("Family reason");
        }
        private void MenuVacation_Click(object sender, RoutedEventArgs e)
        {
            InsertDataToDb("Vacation");
            MessageBox.Show("Vacation");
        }
        private void MenuUnknownReason_Click(object sender, RoutedEventArgs e)
        {
            InsertDataToDb("UnknownReason");
            MessageBox.Show("Unknown reason");
        }
        private void MenuDeleteRecord_Click(object sender, RoutedEventArgs e)
        {
            DeleteDateFromDb();
            MessageBox.Show("Delete records");
        }

        private void DeleteDateFromDb()
        {
            using (MyDbContext myDb = new MyDbContext())
            {
                if (ListBoxEmployee.SelectedItem != null)
                {
                    var itemToUse = (Employee)ListBoxEmployee.SelectedItem;
                    var employee = myDb.Employees.Include(x => x.Absence).FirstOrDefault(x => x.Name == itemToUse.Name);
                    if (employee != null)
                    {
                        List<DateTime>? AbsencePeriod = new List<DateTime>();

                        AbsencePeriod.AddRange(CalendarJanuary.SelectedDates);
                        AbsencePeriod.AddRange(CalendarFabruary.SelectedDates);
                        AbsencePeriod.AddRange(CalendarMarch.SelectedDates);
                        AbsencePeriod.AddRange(CalendarApril.SelectedDates);
                        AbsencePeriod.AddRange(CalendarMay.SelectedDates);
                        AbsencePeriod.AddRange(CalendarJune.SelectedDates);
                        AbsencePeriod.AddRange(CalendarJuly.SelectedDates);
                        AbsencePeriod.AddRange(CalendarAugust.SelectedDates);
                        AbsencePeriod.AddRange(CalendarSeptember.SelectedDates);
                        AbsencePeriod.AddRange(CalendarOctober.SelectedDates);
                        AbsencePeriod.AddRange(CalendarNovember.SelectedDates);
                        AbsencePeriod.AddRange(CalendarDesember.SelectedDates);

                        var emp = myDb.Employees.Include(x => x.Absence).FirstOrDefault(employee => employee.Name == itemToUse.Name);
                        for (int i = 0; i < emp.Absence.Count; i++)
                        {
                            if (emp.Absence[i].FirstDay == AbsencePeriod[0] && emp.Absence[i].DaysCount == AbsencePeriod.Count)
                            {
                                employee.Absence.Remove(emp.Absence[i]);
                                break;
                            }
                        }
                        myDb.SaveChanges();

                        MessageBox.Show($"Запись удалена.");
                    }
                    else
                    {
                        MessageBox.Show("Ошибка");
                    }
                }
                RefreshWindows();
            }

        }

        private void InsertDataToDb(string reason)
        {
            using (MyDbContext myDb = new MyDbContext())
            {
                if (ListBoxEmployee.SelectedItem != null)
                {
                    var itemToUse = (Employee)ListBoxEmployee.SelectedItem;
                    var employee = myDb.Employees.FirstOrDefault(x => x.Name == itemToUse.Name);
                    if (employee != null)
                    {
                        List<DateTime>? AbsencePeriod = new List<DateTime>();

                        AbsencePeriod.AddRange(CalendarJanuary.SelectedDates);
                        AbsencePeriod.AddRange(CalendarFabruary.SelectedDates);
                        AbsencePeriod.AddRange(CalendarMarch.SelectedDates);
                        AbsencePeriod.AddRange(CalendarApril.SelectedDates);
                        AbsencePeriod.AddRange(CalendarMay.SelectedDates);
                        AbsencePeriod.AddRange(CalendarJune.SelectedDates);
                        AbsencePeriod.AddRange(CalendarJuly.SelectedDates);
                        AbsencePeriod.AddRange(CalendarAugust.SelectedDates);
                        AbsencePeriod.AddRange(CalendarSeptember.SelectedDates);
                        AbsencePeriod.AddRange(CalendarOctober.SelectedDates);
                        AbsencePeriod.AddRange(CalendarNovember.SelectedDates);
                        AbsencePeriod.AddRange(CalendarDesember.SelectedDates);


                        Period period = new Period();

                        period.Reason = reason;
                        period.DateNote = TextBoxNote.Text;
                        period.FirstDay = AbsencePeriod[0];
                        period.DaysCount = AbsencePeriod.Count;

                        employee.Absence.Add(period);

                        MessageBox.Show(employee.Absence.Count().ToString());

                        myDb.SaveChanges();
                        MessageBox.Show($"Запись добавлена. Причина отсутствия {reason}");
                    }
                    else
                    {
                        MessageBox.Show("Ошибка");
                    }
                }
                RefreshWindows();
            }
        }

        private void RefreshWindows()
        {
            CalendarJanuary.SelectedDates.Clear();
            CalendarFabruary.SelectedDates.Clear();
            CalendarMarch.SelectedDates.Clear();
            CalendarApril.SelectedDates.Clear();
            CalendarMay.SelectedDates.Clear();
            CalendarJune.SelectedDates.Clear();
            CalendarJuly.SelectedDates.Clear();
            CalendarAugust.SelectedDates.Clear();
            CalendarSeptember.SelectedDates.Clear();
            CalendarOctober.SelectedDates.Clear();
            CalendarNovember.SelectedDates.Clear();
            CalendarDesember.SelectedDates.Clear();
        }

        private void StatButon_Click(object sender, RoutedEventArgs e)
        {
            using (MyDbContext myDb = new MyDbContext())
            {
                Employee selectedItem = (Employee)ListBoxEmployee.SelectedItem;
                StringBuilder outputString = new StringBuilder();
                var instance = myDb.Employees.Include(x => x.Absence).FirstOrDefault(x => x.Name == selectedItem.Name);

                outputString.AppendLine(instance.Department);
                outputString.AppendLine(instance.Name);
                foreach (var absPeriod in instance.Absence)
                {
                    outputString.Append($"Сотрудник отсутствовал на работе с {absPeriod.FirstDay.ToShortDateString()} по {absPeriod.FirstDay.AddDays(absPeriod.DaysCount - 1).ToShortDateString()}" +
                        $" в течении {absPeriod.DaysCount} днея(й). {absPeriod.Reason}");
                    if (absPeriod.DateNote != "")
                        outputString.AppendLine($"примечание: {absPeriod.DateNote}");
                    else outputString.AppendLine();
                }
                TextBoxStatistic.Text = outputString.ToString();              
            }
        }

        private void ListBoxEmployee_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            using (MyDbContext myDb = new MyDbContext())
            {
                Employee selectedItem = (Employee)ListBoxEmployee.SelectedItem;
                StringBuilder outputString = new StringBuilder();
                var instance = myDb.Employees.Include(x => x.Absence).FirstOrDefault(x => x.Name == selectedItem.Name);

                outputString.AppendLine(instance.Department);
                outputString.AppendLine(instance.Name);
                foreach (var absPeriod in instance.Absence)
                {
                    outputString.Append($"Сотрудник отсутствовал на работе с {absPeriod.FirstDay.ToShortDateString()} по {absPeriod.FirstDay.AddDays(absPeriod.DaysCount - 1).ToShortDateString()}" +
                        $" в течении {absPeriod.DaysCount} днея(й). {absPeriod.Reason}");
                    if (absPeriod.DateNote != "")
                        outputString.AppendLine($"примечание: {absPeriod.DateNote}");
                    else outputString.AppendLine();
                }
                TextBoxStatistic.Text = outputString.ToString();               
            }

        }
    }
}
