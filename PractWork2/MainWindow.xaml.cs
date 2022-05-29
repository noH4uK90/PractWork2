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
using Excel = Microsoft.Office.Interop.Excel;

namespace PractWork2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void MultiplyTableButton_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application application = new();
            Excel.Workbooks workbooks = application.Workbooks;
            Excel.Workbook workbook = workbooks.Add();
            Excel.Sheets worksheets = workbook.Worksheets;
            Excel.Worksheet worksheet = worksheets.Item[1];
            Excel.Range title = (Excel.Range) worksheet.get_Range("D9", "L9");
            Excel.Range table = (Excel.Range) worksheet.get_Range("D10", "l18");
            Excel.Range firstMyltiplier = (Excel.Range)worksheet.get_Range("E10", "l10");
            Excel.Range secondMyltiplier = (Excel.Range)worksheet.get_Range("D11", "D18");
            worksheet.Name = "Умножение";

            title.Merge(Type.Missing);
            title.Value = "Таблица умножения";
            title.Font.Italic = true;
            title.Font.Bold = true;
            title.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            title.Font.Size = 20;

            table.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            table.Font.Size = 15;
            table.BorderAround2();

            firstMyltiplier.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(158, 253, 250));
            firstMyltiplier.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(46, 106, 133));
            firstMyltiplier.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            secondMyltiplier.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(158, 253, 250));
            secondMyltiplier.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(46, 106, 133));
            secondMyltiplier.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            for (int i = 2; i < 10; i++)
            {
                table = worksheet.Cells[i + 9, 4];
                table.Value = i;
            }
            for (int i = 2; i < 10; i++)
            {
                table = worksheet.Cells[10, i + 3];
                table.Value = i;
                for (int j = 2; j < 10; j++)
                {
                    table = worksheet.Cells[i + 9, j + 3];
                    table.Value = i * j;
                }
            }
            application.Visible = true;
        }
    }
}
