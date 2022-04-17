using Microsoft.Win32;
using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static OpenFileDialog openfile;
        public static bool? browsefile;
        public static Excel.Application excelApp;
        public static Excel.Workbook excelBook;
        public static Excel.Worksheet excelSheet;
        public static Excel.Range excelRange;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            openfile = new OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "(.xlsx)|*.xlsx"
            };
            browsefile = openfile.ShowDialog();
            txtFilePath.Text = openfile.FileName;
            excelApp = new Excel.Application();
            excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(1);
            excelRange = excelSheet.UsedRange;
            MessageBox.Show("Загрузка файла прошла успешно!");
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void dtGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Short_OutP_Click(object sender, RoutedEventArgs e)
        {
            if (browsefile == true)
            {
                string strCellData = "";
                double douCellData;
                int rowCnt = 0;
                int colCnt = 0;

                DataTable dt = new DataTable();
                for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                {
                    if (colCnt >2)
                    {
                        continue;
                    }
                    string strColumn = "";
                    strColumn = (string)(excelRange.Cells[1, colCnt] as Excel.Range).Value2;
                    dt.Columns.Add(strColumn, typeof(string));
                }
                for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                {
                    string strData = "";
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        try
                        {
                            if (colCnt <= 2)
                            {
                                    strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                                    strData = "УБИ." + strData + strCellData + "|";
                            }
                        }
                        catch
                        {
                            douCellData = (excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += douCellData.ToString() + "|";
                        }
                    }
                    if (rowCnt == 2)
                    {
                        strData = strData.Remove(0, 8);
                    }
                    strData = strData.Remove(strData.Length - 1, 1);
                    dt.Rows.Add(strData.Split('|'));
                }
                dtGrid.ItemsSource = dt.DefaultView;

                excelBook.Close(true, null, null);
                excelApp.Quit();
            }
        }

        private void Full_OutP_Click(object sender, RoutedEventArgs e)
        {
            if (browsefile == true)
            {
                string strCellData;
                double douCellData;
                int rowCnt;
                int colCnt;

                DataTable dt = new DataTable();
                for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                {
                    string strColumn = "";
                    strColumn = (string)(excelRange.Cells[1, colCnt] as Excel.Range).Value2;
                    dt.Columns.Add(strColumn, typeof(string));
                }
                for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                {
                    string strData = "";
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        try
                        {
                            strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += strCellData + "|";
                        }
                        catch
                        {
                            douCellData = (excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += douCellData.ToString() + "|";
                        }
                    }
                    strData = strData.Remove(strData.Length - 1, 1);
                    dt.Rows.Add(strData.Split('|'));
                }

                dtGrid.ItemsSource = dt.DefaultView;

                excelBook.Close(true, null, null);
                excelApp.Quit();
            }
        }

        private void btnSaveFile_Click(object sender, RoutedEventArgs e)
        {
                dtGrid.SelectAllCells();
                dtGrid.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
                ApplicationCommands.Copy.Execute(null, dtGrid);
                String resultat = (string)Clipboard.GetData(DataFormats.CommaSeparatedValue);
                String result = (string)Clipboard.GetData(DataFormats.Text);
                dtGrid.UnselectAllCells();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel file (*.xls)|*.xls|Text file (*.txt)|*.txt";
            if (saveFileDialog.ShowDialog() == true)
            {
                StreamWriter file = new StreamWriter(saveFileDialog.FileName, true, Encoding.UTF8);
                file.WriteLine(result.Replace(',', ' '));
                file.Close();
            }
            MessageBox.Show(@" Exporting DataGrid data to Excel file: "+ saveFileDialog.FileName);
        }

        private void btnCheckUpdate_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfile1;
            bool? browsefile1;
            Excel.Application excelApp1;
            Excel.Workbook excelBook1;
            Excel.Worksheet excelSheet1;
            Excel.Range excelRange1;
            openfile1 = new OpenFileDialog
            {
                DefaultExt = ".xlsx",
                Filter = "(.xlsx)|*.xlsx"
            };
            browsefile1 = openfile1.ShowDialog();
            txtFilePath.Text = openfile1.FileName;
            excelApp1 = new Excel.Application();
            excelBook1 = excelApp1.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            excelSheet1 = (Excel.Worksheet)excelBook1.Worksheets.get_Item(1);
            excelRange1 = excelSheet1.UsedRange;
            MessageBox.Show("Загрузка файла прошла успешно!");
            if (browsefile == true)
            {
                string strCellData;
                double douCellData;
                int rowCnt;
                int colCnt;
                string strCellData1;
                double douCellData1;
                int colCnt1;
                DataTable dt = new DataTable();
                for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                {
                    string strColumn = "";
                    strColumn = (string)(excelRange.Cells[1, colCnt] as Excel.Range).Value2;
                    dt.Columns.Add(strColumn, typeof(string));
                }
                for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                {
                    string strData = "";
                    for (colCnt = 1; colCnt <= excelRange.Columns.Count; colCnt++)
                    {
                        try
                        {
                            strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += strCellData + "|";
                        }
                        catch
                        {
                            douCellData = (excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += douCellData.ToString() + "|";
                        }
                    }
                    strData = strData.Remove(strData.Length - 1, 1);
                    //сверху парсинг первого файла
                    //снизу парсинг второго файла
                    string strData1 = "";
                    for (colCnt1 = 1; colCnt1 <= excelRange1.Columns.Count; colCnt1++)
                    {
                        try
                        {
                            strCellData1 = (string)(excelRange1.Cells[rowCnt, colCnt1] as Excel.Range).Value2;
                            strData1 += strCellData1 + "|";
                        }
                        catch
                        {
                            douCellData1 = (excelRange1.Cells[rowCnt, colCnt1] as Excel.Range).Value2;
                            strData1 += douCellData1.ToString() + "|";
                        }
                    }
                    strData1 = strData1.Remove(strData1.Length - 1, 1);
                   if (strData!=strData1)
                    {
                        strData = "Было:" + strData;
                        strData1 = "Стало:" + strData1;
                        dt.Rows.Add(strData.Split('|'));
                        dt.Rows.Add(strData1.Split('|'));
                    }
                }
                dtGrid.ItemsSource = dt.DefaultView;
                excelBook.Close(true, null, null);
                excelApp.Quit();
                excelBook1.Close(true, null, null);
                excelApp1.Quit();
                MessageBox.Show("Сравнение файлов прошло успешно!");
            }
        }
    }
}
