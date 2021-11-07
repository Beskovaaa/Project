using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using OfficeOpenXml;

namespace Расчет_отпускных
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
            var file = new FileInfo(@"D:\C#\Расчет отпускных\bin\Debug\person.xlsx");
            var data = GetData();
            SaveExcel(data, file);
            MainGrid.ItemsSource = GetExcel(file);
        }

        private List<Model> GetExcel(FileInfo file)
        {
            List<Model> getList = new List<Model>();
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets[0];
                for (int i = 2; i <= 6; i++)
                {
                    getList.Add(new Model()
                    {
                        month = ws.Cells[i, 1].Value.ToString(),
                        number1 = int.Parse(ws.Cells[i, 2].Value.ToString()),
                        number2 = int.Parse(ws.Cells[i, 3].Value.ToString()),
                        sum = int.Parse(ws.Cells[i, 4].Value.ToString()),
                        add_inform = ws.Cells[i, 5].Value.ToString()
                    });
                }
            }
            return getList;
        }
        private void SaveExcel(List<Model> data, FileInfo file)
        {
            using (var package = new ExcelPackage(file))
            {
                if (!file.Exists)
                {
                    var ws = package.Workbook.Worksheets.Add("Main");
                    var range = ws.Cells["A1"].LoadFromCollection(data, true);
                    range.AutoFitColumns();
                    package.Save();
                }
            }
        }
        static List<Model> GetData()
        {
            List<Model> output = new List<Model>()
            {
                new Model() { month = "январь", number1 = 1, number2 = 3, sum = 4, add_inform = "=B1+C1" },
                new Model() { month = "февраль", number1 = 2, number2 = 2, sum = 4, add_inform = "=B2+C2" },
                new Model() { month = "март", number1 = 5, number2 = 5, sum = 10, add_inform = "=B3+C3" },
                new Model() { month = "апрель", number1 = 6, number2 = 4, sum = 10, add_inform = "=B4+C4" },
                new Model() { month = "май", number1 = 8, number2 = 9, sum = 17, add_inform = "=B5+C5" }
            };
            return output;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }
    }
}
