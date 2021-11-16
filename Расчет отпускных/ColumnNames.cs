using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Расчет_отпускных
{
    class ColumnNames
    {
        public string A1 { get {return "месяц"; } set { } }
        public string B1 { get {return "дней в месяце"; } set { } }
        public string C1 { get {return "кол-во дней больничного"; } set { } }
        public string D1 { get {return "кол-во дней отпуска"; } set { } }
        public string E1 { get {return "итого дней"; } set { } }
        public string F1 { get {return "ЗП"; } set { } }
        public string G1 { get {return "выплаты за больничный"; } set { } }
        public string H1 { get {return "выплаты за отпуск"; } set { } }
        public string I1 { get {return "итого ЗП без выплат"; } set { } }
        public string J1 { get {return "дни для расчета"; } set { } }
    }
}
