using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Расчет_отпускных
{
    public class TableLayout
    {

        public string Month { get; set; }
        public int DayInMonth { get; set; }
        public int SickDays { get; set; }
        public int VacationDays { get; set; }
        public int TotalDays { get; set; }
        public int Wages { get; set; }
        public int PaymentSick { get; set; }
        public int PaymentVacation { get; set; }
        public int TotalWages { get; set; }
        public double DaysCalculate { get; set; }
    }
}
