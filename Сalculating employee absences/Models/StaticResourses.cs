using System.Collections.Generic;
using System.Linq;

namespace Сalculating_employee_absences.Models
{
    internal static class StaticResourses
    {
        public static List<int> Years = Enumerable.Range(2022, 12).ToList();
        public static string[] Departments = {"Выберите отдел", "1. Бухгалтерия", "2. Маркетинг", "3. Логисты", "4. Склад", "5. Менеджеры", "6. IТ" };       
    }
}
