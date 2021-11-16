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
        public string DayInMonth { get; set; }
        public string SickDays { get; set; }
        public string VacationDays { get; set; }
        public string TotalDays { get; set; }
        public string Wages { get; set; }
        public string PaymentSick { get; set; }
        public string PaymentVacation { get; set; }
        public string TotalWages { get; set; }
        public string DaysCalculate { get; set; }
    }
}
