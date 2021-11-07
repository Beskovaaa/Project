using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Расчет_отпускных
{
    public class Model
    {
        public string month { get; set; }
        public int number1 { get; set; }
        public int number2 { get; set; }
        public int sum { get; set; }
        public string add_inform { get; set; }
        //public ExcelRangeBase Formula {get; set;}
    }
}
