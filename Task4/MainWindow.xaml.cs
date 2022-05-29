using System;
using System.IO;
using System.Windows;
using Spire.Xls;
using Excel = Microsoft.Office.Interop.Excel;

namespace Task4
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

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            var path = pathTextBox.Text;

            if (Directory.Exists(path))
            {
                DirectoryInfo directory = new DirectoryInfo(path);
                DirectoryInfo[] subDirectoryes = directory.GetDirectories("*", SearchOption.AllDirectories);
                FileInfo[] files = directory.GetFiles();

                Excel.Application application = new();
                Excel.Workbooks workbooks = application.Workbooks;
                Excel.Workbook workbook = workbooks.Add();
                Excel.Sheets worksheets = workbook.Worksheets;
                var directoryInfo = (Excel.Worksheet)worksheets.Add(Type.Missing, worksheets[1], Type.Missing, Type.Missing);
                Excel.Worksheet filesInfo = worksheets.Item[1];
                Excel.Range filesRange = (Excel.Range)filesInfo.get_Range("A1", "C" + files.Length + 1);
                Excel.Range directoryRange = (Excel.Range)directoryInfo.get_Range("A1", "B" + subDirectoryes.Length + 1);
                filesInfo.Name = "Файлы";
                filesInfo.Cells[1, "A"].Value = "Номер";
                filesInfo.Cells[1, "B"].Value = "Имя файла";
                filesInfo.Cells[1, "C"].Value = "Размер в КБ";
                directoryInfo.Name = "Папка";
                directoryInfo.Cells[1, "A"].Value = "Номер";
                directoryInfo.Cells[1, "B"].Value = "Имя папки";

                for (int i = 1; i <= files.Length; i++)
                {
                    filesInfo.Cells[i + 1, "A"].Value = i;
                    filesInfo.Cells[i + 1, "B"].Value = files[i - 1].Name;
                    filesInfo.Cells[i + 1, "C"].Value = files[i - 1].Length / 1024;
                }
                filesInfo.Cells[files.Length + 2, "C"].Formula = $"=SUM(C2:C{files.Length + 1}) / 1024";
                for (int i = 1; i <= subDirectoryes.Length; i++)
                {
                    directoryInfo.Cells[i + 1, "A"].Value = i;
                    directoryInfo.Cells[i + 1, "B"].Value = subDirectoryes[i - 1].Name;
                }
                filesInfo.Columns.AutoFit();
                directoryInfo.Columns.AutoFit();
                application.Visible = true;
            }
            else
            {
                MessageBox.Show("Такой папки не существует");
            }
        }
    }
}
