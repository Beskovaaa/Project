using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Xml.Linq;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;


namespace Расчет_отпускных
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private FileInfo FileExcel = new FileInfo(@"person.xlsx");
        private FileInfo FileXML= new FileInfo(@"Employees.xml");
        private string LastSelected = "";
        public bool FixFlag = false;
        public List<TableLayout> GlobalList = new List<TableLayout>();
        public MainWindow()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            mainGrid.ItemsSource = GlobalList;
            UpdateCalcData();
            AddEmployees();

            if (!FileExcel.Exists)
            {
                using (var package = new ExcelPackage(FileExcel))
                {
                    //package.Workbook.CalcMode = ExcelCalcMode.Manual;
                    var ws = package.Workbook.Worksheets.Add("Hidden_9753194576945164766");
                    List<ColumnNames> columnNames = new List<ColumnNames>
                    {
                        new ColumnNames() {}
                    };
                    ws.Cells["A1"].LoadFromCollection(columnNames, false);
                    package.Save();
                }
            }
        }

        

        private List<TableLayout> GetTable(FileInfo file, string nameEmpl) //
        {
            List<TableLayout> getTable = new List<TableLayout>();
            if (file.Exists) getTable = OpenTable(file, nameEmpl); //
            else getTable = GenerateAndOpenTable(file, nameEmpl); //
            return getTable;
        }

        private List<TableLayout> GenerateAndOpenTable(FileInfo file, string nameEmpl) //
        {
            SaveNewTable(file, GenerateTable(), nameEmpl); //
            return OpenTable(file, nameEmpl); //
        }

        private List<TableLayout> OpenTable(FileInfo file, string nameEmpl) //
        {
            List<TableLayout> getList = new List<TableLayout>();
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets[nameEmpl];
                for (int i = 2; i <= 13; i++)
                {
                    getList.Add(new TableLayout()
                    {
                        Month = ws.Cells[i, 1].Value.ToString(),
                        DayInMonth = int.Parse(ws.Cells[i, 2].Value.ToString()),
                        SickDays = int.Parse(ws.Cells[i, 3].Value.ToString()),
                        VacationDays = int.Parse(ws.Cells[i, 4].Value.ToString()),
                        TotalDays = int.Parse(ws.Cells[i, 5].Value.ToString()),
                        Wages = int.Parse(ws.Cells[i, 6].Value.ToString()),
                        PaymentSick = int.Parse(ws.Cells[i, 7].Value.ToString()),
                        PaymentVacation = int.Parse(ws.Cells[i, 8].Value.ToString()),
                        TotalWages = int.Parse(ws.Cells[i, 9].Value.ToString()),
                        DaysCalculate = double.Parse(ws.Cells[i, 10].Value.ToString())
                    });
                }
            }
            return getList;
        }

        private void DeleteList(FileInfo file, string nameEmpl)
        {
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.SingleOrDefault(x => x.Name == nameEmpl);
                package.Workbook.Worksheets.Delete(ws);
                package.Save();
            }
        }

        private void SaveNewTable(FileInfo file, List<TableLayout> data, string nameEmpl) //
        {
            using (var package = new ExcelPackage(file))
            {
                //package.Workbook.CalcMode = ExcelCalcMode.Manual;
                var ws = package.Workbook.Worksheets.Add(nameEmpl);
                
                //var range = ws.Cells["A2"].LoadFromCollection(data, false);
                ws.Cells["A2"].LoadFromCollection(data, false);
                //range.AutoFitColumns();
                List<ColumnNames> columnNames = new List<ColumnNames>
                {
                    new ColumnNames() {}
                };
                //range = ws.Cells["A1"].LoadFromCollection(columnNames, false);
                ws.Cells["A1"].LoadFromCollection(columnNames, false);
                //range.AutoFitColumns();
                package.Save();
            }
        }

        private void SaveTable(FileInfo file, List<TableLayout> data)
        {
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets[0];
                var range = ws.Cells["A2"].LoadFromCollection(data, false);
                range.AutoFitColumns();
                package.Save();
            }
        }

        private List<TableLayout> getListTableLayout()
        {
            var date = DateTime.Today.Year;
            if (date % 4 == 0) date = 29;
            else date = 28;
            List<TableLayout> outList = new List<TableLayout>()
            {
                new TableLayout { Month = "январь", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "февраль", DayInMonth = date, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "март", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "апрель", DayInMonth = 30, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "май", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "июнь", DayInMonth = 30, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "июль", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "август", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "сентябрь", DayInMonth = 30, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "октябрь", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "ноябрь", DayInMonth = 30, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "декабрь", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 }
            };
            return outList;
        }

        private List<TableLayout> GenerateTable() //
        {
            var date = DateTime.Today.Year;
            if (date % 4 == 0) date = 29;
            else date = 28;
            List<TableLayout> outList = new List<TableLayout>()
            {
                new TableLayout { Month = "январь", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "февраль", DayInMonth = date, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "март", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "апрель", DayInMonth = 30, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "май", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "июнь", DayInMonth = 30, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "июль", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "август", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "сентябрь", DayInMonth = 30, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "октябрь", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "ноябрь", DayInMonth = 30, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 },
                new TableLayout { Month = "декабрь", DayInMonth = 31, DaysCalculate = 0, PaymentSick = 0, PaymentVacation = 0, SickDays = 0, TotalDays = 0, TotalWages = 0, VacationDays = 0, Wages = 0 }
            };
            List<TableLayout> newOutList = new List<TableLayout>();
            var date2 = DateTime.Today.Month;
            for (int i = date2; i < 12; i++)
            {
                newOutList.Add(outList[i]);
            }
            for (int i = 0; i < date2; i++)
            {
                newOutList.Add(outList[i]);
            }
            return newOutList;
        }

        private void SaveExcel(List<TableLayout> data, FileInfo file, string nameEmpl)
        {
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets[nameEmpl];
                var range = ws.Cells["A2"].LoadFromCollection(data, false);
                range.AutoFitColumns();
                package.Save();
            }
        }
        private List<TableLayout> GetDataFromGrid()
        {
            List<TableLayout> getList = new List<TableLayout>();
            for (int i = 0; i < 12; i++)
            {
                TableLayout tl = (TableLayout)mainGrid.Items[i];
                getList.Add(new TableLayout()
                {
                    Month = tl.Month,
                    DayInMonth = tl.DayInMonth,
                    SickDays = tl.SickDays,
                    VacationDays = tl.VacationDays,
                    TotalDays = tl.TotalDays,
                    Wages = tl.Wages,
                    PaymentSick = tl.PaymentSick,
                    PaymentVacation = tl.PaymentVacation,
                    TotalWages = tl.TotalWages,
                    DaysCalculate = tl.DaysCalculate
                });
            }
            return getList;
        }

        private void UpdateCalcData()
        {
            double sumDay = 0;
            double sumWages = 0;
            double sumCalcDay = 0;
            foreach (TableLayout item in GlobalList)
            {
                if (item.VacationDays != 0) sumDay += item.VacationDays;
                if (item.TotalWages != 0) sumWages += item.TotalWages;
                if (item.DaysCalculate != 0) sumCalcDay += item.DaysCalculate;
            }

            tb1.Text = string.Format("{0:f0}", sumDay);
            tb2.Text = string.Format("{0:f2}", sumWages / sumCalcDay);
            tb3.Text = string.Format("{0:f2}", sumWages / sumCalcDay * sumDay);
        }

        private void MainGrid_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (!FixFlag)
            {
                var TableLayoutObj = e.Row.Item as TableLayout;
                int numRow = e.Row.GetIndex();
                UpdateAfterEdit(TableLayoutObj, numRow);
                mainGrid.Items.Refresh();
                UpdateCalcData();
                //MessageBox.Show($"{TableLayoutObj.Month} {TableLayoutObj.DayInMonth} {numRow}");
            }
        }

        private void UpdateAfterEdit(TableLayout tableLayoutObj, int numRow)
        {
            tableLayoutObj.TotalDays =
                tableLayoutObj.DayInMonth - tableLayoutObj.SickDays - tableLayoutObj.VacationDays;
            tableLayoutObj.TotalWages =
                tableLayoutObj.Wages - tableLayoutObj.PaymentSick - tableLayoutObj.PaymentVacation;
            if (tableLayoutObj.DayInMonth != 0)
            {
                tableLayoutObj.DaysCalculate = double.Parse(tableLayoutObj.TotalDays.ToString()) /
                    double.Parse(tableLayoutObj.DayInMonth.ToString()) * 29.3;
            }
            else
            {
                tableLayoutObj.DaysCalculate = 0;
            }

            GlobalList.RemoveAt(numRow);
            GlobalList.Insert(numRow, tableLayoutObj);
            FixFlag = true;
            mainGrid.CancelEdit();
            mainGrid.CancelEdit();
            FixFlag = false;
            //mainGrid.CommitEdit();
            //mainGrid.Items.Refresh();
        }

        private void ButtonPrint_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog Printdlg = new PrintDialog();
            if ((bool)Printdlg.ShowDialog().GetValueOrDefault())
            {
                Size pageSize = new Size(Printdlg.PrintableAreaWidth, Printdlg.PrintableAreaHeight);
                mainGrid.Measure(pageSize);
                mainGrid.Arrange(new Rect(5, 5, pageSize.Width + 100.0, pageSize.Height));
                Printdlg.PrintVisual(mainGrid, Title);
            }
        }
        private void WriteXML(List<string> itemsList)
            => new XDocument(new XElement("itemsList", itemsList.ConvertAll(item => new XElement("item", item)))).Save(@"Employees.xml");
        private List<string> ReadXML()
          => XDocument.Load(@"Employees.xml").Element("itemsList")?.Elements("item").Select(item => item.Value).ToList();

        private void AddEmployees() //добавление на форму при запуске
        {
            List<string> NameEmployees = new List<string>();
            var file = new FileInfo(@"Employees.xml");
            if (!file.Exists) WriteXML(NameEmployees);
            NameEmployees = ReadXML();
            foreach (var nameEmployee in NameEmployees)
            {
                ListEmployees.Items.Add(nameEmployee);
            }
        }

        private void ButtonAddEmployee_Click(object sender, RoutedEventArgs e)
        {
            WindowAddEmployee addEmployee = new WindowAddEmployee();
            addEmployee.ShowDialog();
            List<string> NameEmployees = new List<string>();
            var file = new FileInfo(@"Employees.xml");
            if (!file.Exists) WriteXML(NameEmployees);
            NameEmployees = ReadXML();
            if (!NameEmployees.Contains(WindowAddEmployee.TextBox) && WindowAddEmployee.TextBox != "" && WindowAddEmployee.TextBox != null)
            {
                NameEmployees.Add(WindowAddEmployee.TextBox);
                WriteXML(NameEmployees);
                ListEmployees.Items.Clear();
                for (int i = 0; i < NameEmployees.Count; i++)
                {
                    ListEmployees.Items.Add(NameEmployees[i]);
                }
                //+++добавить листочек в книгу
                SaveNewTable(FileExcel, GenerateTable(), NameEmployees[NameEmployees.Count - 1]);
            }

        }
        
        private void ButtonDeleteEmployee_Click(object sender, RoutedEventArgs e)
        {
            if( ListEmployees.SelectedIndex!=-1)
            {
                if (MessageBox.Show("Вы уверены что хотите удалить сотрудника?","Удаление сотрудника",MessageBoxButton.YesNo,MessageBoxImage.Warning ) == MessageBoxResult.Yes)
                {
                    List<string> NameEmployees = ReadXML();
                    NameEmployees.RemoveAt(ListEmployees.SelectedIndex);
                    //++++удаление странички сотруднника
                    LastSelected = "";
                    GlobalList.Clear();
                    mainGrid.Items.Refresh();
                    UpdateCalcData();
                    string nameEmpl = ListEmployees.SelectedItem.ToString();
                    DeleteList(FileExcel, nameEmpl);
                    WriteXML(NameEmployees);
                    ListEmployees.Items.Clear();
                    for (int i = 0; i < NameEmployees.Count; i++)
                    {
                        ListEmployees.Items.Add(NameEmployees[i]);
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите сотрудника в списке!","Внимание",MessageBoxButton.OKCancel,MessageBoxImage.Error);
            }
        }

        private void UpdateLine()
        {
            List<string> monthList = new List<string>()
            {
                "январь",
                "февраль",
                "март", 
                "апрель",
                "май",
                "июнь", 
                "июль",
                "август", 
                "сентябрь", 
                "октябрь",
                "ноябрь",
                "декабрь"
            };
            string LastMonth = GlobalList[GlobalList.Count - 1].Month;
            int LastNumMonth = monthList.IndexOf(LastMonth) + 1;
            var date = DateTime.Today.Month;
            if (date < LastNumMonth) LastNumMonth -= 12;
            int difference = date - LastNumMonth;
            if (difference != 0)
            {
                List<TableLayout> tl = getListTableLayout();
                for (int i = 0; i < difference; i++) GlobalList.RemoveAt(0);
                for (int i = 0; i < difference; i++) GlobalList.Add(tl[date - difference + i]);
            }
        }

        private void ListEmployees_PreviewMouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (LastSelected != "") SaveExcel(GlobalList, FileExcel, LastSelected);
            if (ListEmployees.SelectedItem != null)
            {
                string nameEmpl = ListEmployees.SelectedItem.ToString();
                LastSelected = nameEmpl;
                //MessageBox.Show(nameEmpl);
                GlobalList.Clear();
                var dataTable = OpenTable(FileExcel, nameEmpl);
                foreach (var item in dataTable)
                {
                    GlobalList.Add(item);
                }

                UpdateLine();
                mainGrid.Items.Refresh();
                UpdateCalcData();
            }
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (LastSelected != "") SaveExcel(GlobalList, FileExcel, LastSelected);
        }
    }
}
