using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Shapes;

namespace Расчет_отпускных
{
    /// <summary>
    /// Логика взаимодействия для WindowAddEmployee.xaml
    /// </summary>
    public partial class WindowAddEmployee : Window
    {
        public WindowAddEmployee()
        {
            InitializeComponent();
        }

        private void ButtonAdd_Click(object sender, RoutedEventArgs e)
        {
            string writePath = @"Employees.txt";
            using (StreamWriter sw = new StreamWriter(writePath))
            {
                if (TextBoxData.Text != "")
                    sw.WriteLine(TextBoxData.Text);
                else
                    MessageBox.Show("Введите данные сотрудника!");
            }
            TextBoxData.Text = "";
        }

    }
}
