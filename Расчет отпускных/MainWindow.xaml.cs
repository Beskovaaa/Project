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
using System.Drawing;
using System.IO;
using Syncfusion.XlsIO;

namespace Расчет_отпускных
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                application.DefaultVersion = ExcelVersion.Excel2016;

                //Create a workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding a picture
                FileStream imageStream = new FileStream("../../Image/DQS-Logo.jpg", FileMode.Open, FileAccess.Read);
                IPictureShape shape = worksheet.Pictures.AddPicture(1, 1, imageStream,15,10);
                worksheet.Range["A1:A2"].Merge();
                worksheet.Range["A1:C1"].Merge();

                //Disable gridlines in the worksheet
                worksheet.IsGridLinesVisible = false;

                //Enter values to the cells from A3 to A5
                worksheet.Range["A6"].Text = "150003, г.Ярославль";
                worksheet.Range["A7"].Text = "dqs@dqs-russia.ru";
                worksheet.Range["A8"].Text = "Phone: +7 4852 69 50 21";

                //Make the text bold
                worksheet.Range["A6:A8"].CellStyle.Font.Bold = true;

                //Merge cells
                worksheet.Range["F1:N1"].Merge();

                //Enter text to the cell D1 and apply formatting.
                worksheet.Range["F1"].Text = "ОРГАН ПО СЕРТИФИКАЦИИ ООО ССУ ДЭКУЭС";
                worksheet.Range["F1"].CellStyle.Font.Bold = true;
                worksheet.Range["F1"].CellStyle.Font.RGBColor = System.Drawing.Color.FromArgb(42, 118, 189);
                worksheet.Range["F1"].CellStyle.Font.Size = 18;


                //Create table with the data in given range
                IListObject table = worksheet.ListObjects.Create("Table", worksheet["E6:O19"]);

                //Create data
                worksheet.Range["E6"].Text = "№";                
                worksheet.Range["F6"].Text = "Месяц";
                worksheet.Range["G6"].Text = "Кол-во дней больничного";
                worksheet.Range["H6"].Text = "Кол-во дней отпуска";
                worksheet.Range["I6"].Text = "Итого дней";
                worksheet.Range["J6"].Text = "Сумма ЗП";
                worksheet.Range["K6"].Text = "Сумма больничных";
                worksheet.Range["L6"].Text = "Сумма отпускных";
                worksheet.Range["M6"].Text = "ЗП в месяц без больгичных и отпускных";
                worksheet.Range["N6"].Text = "Дни для расчета ";



                //Save the Excel document
                workbook.SaveAs("Расчет.xlsx");

                System.Diagnostics.Process.Start("Расчет.xlsx");
            }
        }
    }
}
