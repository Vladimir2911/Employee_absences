using Microsoft.EntityFrameworkCore;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Сalculating_employee_absences.Models;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Сalculating_employee_absences
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        MyDbContext myDb = new MyDbContext();
        public ObservableCollection<Employee> Employees { get; set; }


        public MainWindow()
        {
            Employees = new ObservableCollection<Employee>();
            InitializeComponent();
            InitializeCombobox();
            SetDefaultValues();
            // ListBoxEmployee.ItemsSource = Employees;
        }

        private void InitializeCombobox()
        {
            YearSelectionComboBox.ItemsSource = StaticResourses.Years;
            YearSelectionComboBox.SelectedItem = DateTime.Now.Year;
            DepartmentCombobox.ItemsSource = StaticResourses.Departments;
        }

        private void SetDefaultValues()
        {
            RefreshCalendarDates();
            LoadDataFromCollection();
        }

        private void LoadDataFromCollection()
        {
            ClearCollection();

            var loadResult = LoadDataToList();
            foreach (var item in loadResult)
            {
                if (DepartmentCombobox.SelectedItem == null || (string)DepartmentCombobox.SelectedValue == StaticResourses.Departments[0])
                    Employees.Add(item);
                else if (item.Department == DepartmentCombobox.SelectedItem.ToString())
                    Employees.Add(item);
            }
        }

        private void ClearCollection()
        {
            Employees.Clear();
        }

        private void AddEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            AddEmployeDialog addEmployeDialog = new AddEmployeDialog(this);
            addEmployeDialog.Show();
        }

        public List<Employee> LoadDataToList()
        {
            return myDb.Employees.Include(p => p.Periods).OrderBy(x => x.Name).ToList();
        }

        private void RemuveEmployeeButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Удалить сотрудника?", "Внимание!!!",
                MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                if (ListBoxEmployee.SelectedItem != null)
                {
                    var itemToDelete = ListBoxEmployee.SelectedItem.ToString();
                    var employee = myDb.Employees.Include(p => p.Periods)
                        .FirstOrDefault(x => x.Name == itemToDelete);
                    if (employee != null)
                    {
                        Employees.Remove(employee);
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
                SetDefaultValues();
            }
            else return;
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
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
            DisplayEmployeeInfo();
        }

        private void InsertDataToDb(string reason)
        {
            string _reason = string.Empty;
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

                    if (CheckAvalibleDate(employee, AbsencePeriod))
                    {
                        Period period = new Period();
                        period.Reason = _reason;
                        period.DateNote = TextBoxNote.Text;
                        period.FirstDay = AbsencePeriod[0];
                        period.DaysCount = AbsencePeriod.Count;

                        employee.Periods.Add(period);
                        myDb.SaveChanges();
                        MessageBox.Show($"Запись добавлена. Причина отсутствия {_reason}");
                    }
                    else MessageBox.Show("На выделенные даты уже существует запись у этого сотрудника.");
                }
                else
                {
                    MessageBox.Show("Ошибка");
                }
            }
            ClearSelectedDates();
            TextBoxNote.Text = string.Empty;
            RefreshCalendarDates();
            DisplayEmployeeInfo();
        }

        private bool CheckAvalibleDate(Employee employee, List<DateTime> absencePeriod)
        {
            if (employee.Periods.Count == 0) return true;

            foreach (var period in employee.Periods)
            {
                for (int i = 0; i < period.DaysCount; i++)
                {
                    foreach (DateTime date in absencePeriod)
                    {
                        if (period.FirstDay.AddDays(i) == date) return false;
                    }
                }
            }
            return true;
        }

        private void RefreshCalendarDates()
        {
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
        //выгрузка в ексель
        private void UploadToExcelButon_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();


            if (xlApp == null)
            {
                MessageBox.Show("Excel не установлен!!");
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
            xlWorkSheet.Cells[1, 16] = "Общее количество дней по болезни (Б к.д)";
            xlWorkSheet.Cells[1, 17] = "Общее количество отсутствия по семейным обстоятельствам (СО к.д.)";
            xlWorkSheet.Cells[1, 18] = "Отсутствие по невыясненным причинам (НВ к.д.)";
            #endregion
            var employee = myDb.Employees.Include(p => p.Periods).OrderBy(x => x.Department).ThenBy(n => n.Name).ToList();

            for (int i = 0; i < employee.Count; i++)
            {
                int dayVacation = 0, dayFamily = 0, dayHealth = 0, dayUnknown = 0;
                xlWorkSheet.Cells[i + 2, 1] = i + 1;
                xlWorkSheet.Cells[i + 2, 2] = employee[i].Name;
                StringBuilder sb = new StringBuilder();
                foreach (Period period in employee[i].Periods.OrderBy(x=>x.FirstDay))
                {
                    string shortReason = string.Empty;
                    switch (period.Reason)
                    {
                        case "по не выясненых причинах":
                            dayUnknown += period.DaysCount;
                            shortReason = "НВ";
                            break;

                        case "отпуск":
                            dayVacation += period.DaysCount;
                            shortReason = "О";
                            break;

                        case "по семейным обстоятельствам":
                            dayFamily += period.DaysCount;
                            shortReason = "СО";
                            break;
                        case "по состоянию здоровья":
                            dayHealth += period.DaysCount;
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
                    if (dayVacation == 28) xlWorkSheet.Cells[i + 2, 15].Interior.Color = Color.MediumAquamarine;
                    if (dayVacation > 28) xlWorkSheet.Cells[i + 2, 15].Interior.Color = Color.Coral;

                    xlWorkSheet.Cells[i + 2, 16] = dayHealth;
                    if (dayHealth > 14) xlWorkSheet.Cells[i + 2, 16].Interior.Color = Color.Coral;
                    xlWorkSheet.Cells[i + 2, 17] = dayFamily;
                    if (dayFamily > 7) xlWorkSheet.Cells[i + 2, 17].Interior.Color = Color.Coral;
                    xlWorkSheet.Cells[i + 2, 18] = dayUnknown;
                    if (dayUnknown > 5) xlWorkSheet.Cells[i + 2, 18].Interior.Color = Color.Coral;
                    if (dayUnknown + dayFamily + dayHealth > 16)
                    {
                        xlWorkSheet.Cells[i + 2, 19] = dayUnknown + dayFamily + dayHealth + " к.д. отсутствий";
                        xlWorkSheet.Cells[i + 2, 19].Interior.Color = Color.Coral;
                    }
                }
            }

            Range rng = xlWorkSheet.get_Range("B:B", System.Type.Missing);
            rng.EntireColumn.ColumnWidth = 23;
            rng = xlWorkSheet.get_Range("A:A", System.Type.Missing);
            rng.EntireColumn.ColumnWidth = 4;
            rng = xlWorkSheet.get_Range("C:N", System.Type.Missing);
            rng.EntireColumn.ColumnWidth = 11;
            rng.WrapText = true;
            rng = xlWorkSheet.get_Range("O:S", System.Type.Missing);
            rng.EntireColumn.ColumnWidth = 15.8;
            rng.WrapText = true;
            rng = xlWorkSheet.get_Range("A:R", System.Type.Missing);

            rng.get_Range("A1", "R" + (employee.Count + 1));
            rng.Borders.ColorIndex = 0;
            rng.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            rng.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         



            xlWorkBook.SaveAs($"{Directory.GetCurrentDirectory()}\\График_{YearSelectionComboBox.SelectedValue}.xls",
                                 Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                                 Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();


            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            if (MessageBox.Show($"Файл График_{YearSelectionComboBox.SelectedValue}.xls создан в текущей папке.\nОткрыть файл???", "!!!",
               MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                OpenExcellFile($"{Directory.GetCurrentDirectory()}\\График_{YearSelectionComboBox.SelectedValue}.xls");
            }

        }

        private void OpenExcellFile(string path)
        {
            Excel.Application xlApp = new Excel.Application();
            Workbook wb = xlApp.Workbooks.Open(path);
            xlApp.Visible = true;
        }

        private void ListBoxEmployee_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DisplayEmployeeInfo();
        }

        private void DisplayEmployeeInfo()
        {
            if (ListBoxEmployee.SelectedIndex == -1 || ListBoxEmployee.SelectedItems == null)
            {
                TextBoxStatistic.Text = "Статистика";
            }
            else
            {
                StringBuilder outputString = new StringBuilder();
                var instance = myDb.Employees.Include(e => e.Periods).FirstOrDefault(x => x.Name == ListBoxEmployee.SelectedItem.ToString());
                if (instance == null)
                {
                    TextBoxStatistic.Text = "...";
                    return;
                }
                outputString.AppendLine(instance.Department);
                outputString.AppendLine(instance.Name);
                foreach (var absPeriod in instance.Periods.OrderBy(x=>x.FirstDay))
                {
                    if (absPeriod.FirstDay.Year == (int)YearSelectionComboBox.SelectedItem)
                    {
                        outputString.Append($"Сотрудник отсутствовал на работе с {absPeriod.FirstDay.ToShortDateString()} по" +
                            $" {absPeriod.FirstDay.AddDays(absPeriod.DaysCount - 1).ToShortDateString()} " +
                            $"в течение {absPeriod.DaysCount} к.д. {absPeriod.Reason}.");
                        if (absPeriod.DateNote != string.Empty)
                            outputString.AppendLine($" примечание: {absPeriod.DateNote}");
                        else outputString.AppendLine();
                    }
                }
                TextBoxStatistic.Text = outputString.ToString();
            }
        }

        private void DepartmentCombobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadDataFromCollection();
        }

        private void YearSelectionComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ClearSelectedDates();
            RefreshCalendarDates();
        }

        private void MenuRenameEmployee_Click(object sender, RoutedEventArgs e)
        {
            EmployeeEdit();
            LoadDataFromCollection();
        }

        private void MenuRefreshList_Click(object sender, RoutedEventArgs e)
        {
            SetDefaultValues();
        }

        private void EmployeeEdit()
        {
            var employeeToEdit = myDb.Employees.Include(e => e.Periods).FirstOrDefault(x => x.Name == ((Employee)ListBoxEmployee.SelectedItem).Name);
            AddEmployeDialog edit = new AddEmployeDialog(this, employeeToEdit);
            edit.Show();
        }
    }
}
