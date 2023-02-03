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
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.IO;
using System.Collections;

namespace Сalculating_employee_absences
{    
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        MyDbContext myDb = new MyDbContext();
        ObservableCollection<Employee> employees;
       

        public MainWindow()
        {
            employees = new ObservableCollection<Employee>();
            InitializeComponent();
            InitializeCombobox();
            SetDefaultValues();
            
        }

        private void InitializeCombobox()
        {                       
            YearSelectionComboBox.ItemsSource = StaticResourses.Years;
            YearSelectionComboBox.SelectedItem=DateTime.Now.Year;
            DepartmentCombobox.ItemsSource = StaticResourses.Departments;            
        }

        private void SetDefaultValues()
        {
            RefreshCalendarDates();
            LoadData();
        }

        private void AddEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            AddEmployeDialog addEmployeDialog = new AddEmployeDialog();
            addEmployeDialog.Show();

        }

        public async void LoadData()
        {
            List<Employee> result;
            if (DepartmentCombobox.SelectedItem == null)
            {
                result = myDb.Employees.Include(p => p.Periods).OrderBy(x => x.Name).ToList();
            }
            else
            {
                result = myDb.Employees.Include(p => p.Periods).OrderBy(x => x.Name).Where(d=>d.Department == DepartmentCombobox.SelectedItem).ToList();
            }
            ListBoxEmployee.ItemsSource = result;
        }

        private void RemuveEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            if (ListBoxEmployee.SelectedItem != null)
            {
                var itemToDelete = (Employee)ListBoxEmployee.SelectedItem;
                var employee = myDb.Employees.Include(p => p.Periods).FirstOrDefault(x => x.Name == itemToDelete.Name);
                if (employee != null)
                {
                    myDb.Employees.Remove(employee);
                    myDb.SaveChanges();
                    MessageBox.Show("Запись удалена!");
                }
            }
            else
            {
                MessageBox.Show("Не выбран сотрудник");
            }
            ClearSelectedDates();
            RefreshCalendarDates();
            try
            {
                ListBoxEmployee.SelectedIndex=0;
            }
            catch
            {
                MessageBox.Show("Error");
            }
            LoadData();


        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SetDefaultValues();
        }
        #region Context menu
        private void MenuHealthReason_Click(object sender, RoutedEventArgs e)
        {
            InsertDataToDb("HealthReason");           
        }
        private void MenuFamilyReason_Click(object sender, RoutedEventArgs e)
        {
            InsertDataToDb("FamilyReason");          
        }
        private void MenuVacation_Click(object sender, RoutedEventArgs e)
        {
            InsertDataToDb("Vacation");           
        }
        private void MenuUnknownReason_Click(object sender, RoutedEventArgs e)
        {
            InsertDataToDb("UnknownReason");          
        }
        private void MenuDeleteRecord_Click(object sender, RoutedEventArgs e)
        {
            DeleteDateFromDb();           
        }
        #endregion

        private void DeleteDateFromDb()
        {
            if (ListBoxEmployee.SelectedItem != null)
            {
                var itemToUse = (Employee)ListBoxEmployee.SelectedItem;
                var employee = myDb.Employees.Include(x => x.Periods).FirstOrDefault(x => x.Name == itemToUse.Name);
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
                    bool found = false;
                    var emp = myDb.Employees.Include(x => x.Periods).FirstOrDefault(employee => employee.Name == itemToUse.Name);
                    for (int i = 0; i < emp.Periods.Count; i++)
                    {
                        if (emp.Periods[i].FirstDay == AbsencePeriod[0] && emp.Periods[i].DaysCount == AbsencePeriod.Count)
                        {
                            found = true;
                            employee.Periods.Remove(emp.Periods[i]);
                            myDb.SaveChanges();
                            break;
                        }
                    }
                    if (found == true)
                    {
                        MessageBox.Show($"Запись удалена.");
                    }
                    else
                        MessageBox.Show($"Запись не удалена.");
                }
                else
                {
                    MessageBox.Show("Ошибка");
                }
            }
            else
                MessageBox.Show("Не выбран сотрудник");
            ClearSelectedDates();
            RefreshCalendarDates();
        }

        private void InsertDataToDb(string reason)
        {
            string _reason = "";
            switch (reason)
            {
                case "UnknownReason":
                    _reason = "по не выясненых причинах";
                    break;

                case "Vacation":
                    _reason = "отпуск";
                    break;

                case "FamilyReason":
                    _reason = "по семейным обстоятельствам";
                    break;
                case "HealthReason":
                    _reason = "по состоянию здоровья";
                    break;

                default:
                    break;
            }

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

                    period.Reason = _reason;
                    period.DateNote = TextBoxNote.Text;
                    period.FirstDay = AbsencePeriod[0];
                    period.DaysCount = AbsencePeriod.Count;

                    employee.Periods.Add(period);

                    MessageBox.Show(employee.Periods.Count().ToString());

                    myDb.SaveChanges();
                    MessageBox.Show($"Запись добавлена. Причина отсутствия {reason}");
                }
                else
                {
                    MessageBox.Show("Ошибка");
                }
            }
            ClearSelectedDates();
            RefreshCalendarDates();
        }

        private void RefreshCalendarDates()
        {
            var year = DateTime.Now.Year;
            CalendarJanuary.DisplayDate = Convert.ToDateTime("01/01/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarFabruary.DisplayDate = Convert.ToDateTime("01/02/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarMarch.DisplayDate = Convert.ToDateTime("01/03/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarApril.DisplayDate = Convert.ToDateTime("01/04/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarMay.DisplayDate = Convert.ToDateTime("01/05/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarJune.DisplayDate = Convert.ToDateTime("01/06/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarJuly.DisplayDate = Convert.ToDateTime("01/07/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarAugust.DisplayDate = Convert.ToDateTime("01/08/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarSeptember.DisplayDate = Convert.ToDateTime("01/09/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarOctober.DisplayDate = Convert.ToDateTime("01/10/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarNovember.DisplayDate = Convert.ToDateTime("01/11/" + YearSelectionComboBox.SelectedItem.ToString());
            CalendarDesember.DisplayDate = Convert.ToDateTime("01/12/" + YearSelectionComboBox.SelectedItem.ToString());
        }

        private void ClearSelectedDates()
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
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            #region TableHead
            xlWorkSheet.Cells[1, 1] = "П/П";
            xlWorkSheet.Cells[1, 2] = "Ф.И.О";
            xlWorkSheet.Cells[1, 3] = "Январь";
            xlWorkSheet.Cells[1, 4] = "Февраль";
            xlWorkSheet.Cells[1, 5] = "Март";
            xlWorkSheet.Cells[1, 6] = "Апрель";
            xlWorkSheet.Cells[1, 7] = "Май";
            xlWorkSheet.Cells[1, 8] = "Июнь";
            xlWorkSheet.Cells[1, 9] = "Июль";
            xlWorkSheet.Cells[1, 10] = "Август";
            xlWorkSheet.Cells[1, 11] = "Сентябрь";
            xlWorkSheet.Cells[1, 12] = "Октябрь";
            xlWorkSheet.Cells[1, 13] = "Ноябрь";
            xlWorkSheet.Cells[1, 14] = "Декабрь";
            xlWorkSheet.Cells[1, 15] = "Общее количество отпускных дней (к.д.)";
            xlWorkSheet.Cells[1, 16] = "Общее количество дней болезни (Б к.д)";
            xlWorkSheet.Cells[1, 17] = "Общее количество отсутствия по семейным обстоятельствам (СО к.д.)";
            xlWorkSheet.Cells[1, 18] = "Отсутствие по невыясненным причинам (НВ к.д.)";
            #endregion
            var employee = myDb.Employees.Include(p => p.Periods).OrderBy(d => d.Department).ThenBy(d => d.Name).ToList();

            for (int i = 0; i < employee.Count; i++)
            {
                int dayVacation = 0, dayFamily = 0, dayHeals = 0, dayUnknown = 0;
                xlWorkSheet.Cells[i + 2, 1] = i + 1;
                xlWorkSheet.Cells[i + 2, 2] = employee[i].Name;
                StringBuilder sb = new StringBuilder();
                foreach (Period period in employee[i].Periods)
                {
                    string shortReason = "";
                    switch (period.Reason)
                    {
                        case "по не выясненых причинах":
                            dayUnknown +=period.DaysCount;
                            shortReason = "НВ";
                            break;

                        case "отпуск":
                            dayVacation+=period.DaysCount;
                            shortReason = "О";
                            break;

                        case "по семейным обстоятельствам":
                            dayFamily+=period.DaysCount;
                            shortReason = "СО";
                            break;
                        case "по состоянию здоровья":
                            dayHeals+=period.DaysCount;
                            shortReason = "Б";
                            break;

                        default:
                            break;
                    }

                    if (xlWorkSheet.Cells[i + 2, period.FirstDay.Month + 2].Value == null || xlWorkSheet.Cells[i + 2, period.FirstDay.Month + 2].Value == null)
                    {
                        sb.Clear();
                        sb.Append($"{period.FirstDay.ToShortDateString()}-{period.FirstDay.AddDays(period.DaysCount - 1).ToShortDateString()} ({period.DaysCount})-{shortReason} {period.DateNote} ");
                        xlWorkSheet.Cells[i + 2, period.FirstDay.Month + 2] = sb.ToString();
                    }
                    else
                    {
                        sb.Append($"{period.FirstDay.ToShortDateString()}-{period.FirstDay.AddDays(period.DaysCount - 1).ToShortDateString()} ({period.DaysCount})-{shortReason}{period.DateNote} ");
                        xlWorkSheet.Cells[i + 2, period.FirstDay.Month + 2] = sb.ToString();
                    }
                    xlWorkSheet.Cells[i + 2, 15] = dayVacation;
                    xlWorkSheet.Cells[i + 2, 16] = dayHeals;
                    xlWorkSheet.Cells[i + 2, 17] = dayFamily;
                    xlWorkSheet.Cells[i + 2, 18] = dayUnknown;
                }
            }

            xlWorkBook.SaveAs($"{Directory.GetCurrentDirectory()}\\отпуск.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Файл создан, ві можете его найти в текущей папке. файл отпуск.xls");

            /*            Employee selectedItem = (Employee)ListBoxEmployee.SelectedItem;
                        StringBuilder outputString = new StringBuilder();
                        var instance = myDb.Employees.Include(x => x.Periods).FirstOrDefault(x => x.Name == selectedItem.Name);

                        outputString.AppendLine(instance.Department);
                        outputString.AppendLine(instance.Name);
                        foreach (var absPeriod in instance.Periods)
                        {
                            outputString.Append($"Сотрудник отсутствовал на работе с {absPeriod.FirstDay.ToShortDateString()} по {absPeriod.FirstDay.AddDays(absPeriod.DaysCount - 1).ToShortDateString()}" +
                                $" в течении {absPeriod.DaysCount} днея(й). {absPeriod.Reason}");
                            if (absPeriod.DateNote != "")
                                outputString.AppendLine($"примечание: {absPeriod.DateNote}");
                            else outputString.AppendLine();
                        }
                        TextBoxStatistic.Text = outputString.ToString();*/
        }

        private void ListBoxEmployee_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Employee selectedItem = (Employee)ListBoxEmployee.SelectedItem;
            StringBuilder outputString = new StringBuilder();
            var instance = myDb.Employees.Include(x => x.Periods).FirstOrDefault(x => x.Name == selectedItem.Name);

            outputString.AppendLine(instance.Department);
            outputString.AppendLine(instance.Name);
            foreach (var absPeriod in instance.Periods)
            {
                outputString.Append($"Сотрудник отсутствовал на работе с {absPeriod.FirstDay.ToShortDateString()} по" +
                    $" {absPeriod.FirstDay.AddDays(absPeriod.DaysCount - 1).ToShortDateString()} в течении {absPeriod.DaysCount} днея(й). {absPeriod.Reason}");
                if (absPeriod.DateNote != "")
                    outputString.AppendLine($"примечание: {absPeriod.DateNote}");
                else outputString.AppendLine();
            }
            TextBoxStatistic.Text = outputString.ToString();
        }
      
        private void DepartmentCombobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadData();
        }

        private void YearSelectionComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ClearSelectedDates();
            RefreshCalendarDates();
        }
    }
}
