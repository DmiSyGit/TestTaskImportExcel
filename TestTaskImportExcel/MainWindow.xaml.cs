using Microsoft.Win32;
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
using ExcelDataReader;
using System.Data;
using System.IO;

namespace TestTaskImportExcel
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

        private void ChooseExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EXCEL Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() != true)
                return;

            ExcelGrig.ItemsSource = ReadFile(openFileDialog.FileName);

        }
        IExcelDataReader ExcelDataReader;
        DataSet dataSet;
        private DataView ReadFile(string fileNames)
        {
            var extension = fileNames.Substring(fileNames.LastIndexOf('.'));
            // Создание потока для чтения.
            FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
            // В зависимости от расширения файла, создаётся тот или иной читатель.
            if (extension == ".xlsx")
                ExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            else if (extension == ".xls")
                ExcelDataReader = ExcelReaderFactory.CreateBinaryReader(stream);
            else if (extension == ".csv")
                ExcelDataReader = ExcelReaderFactory.CreateCsvReader(stream);

            // Создание конфигурации, что превая строка в файле это заголовки
            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };
            // Читаем, получаем DataView
            dataSet = ExcelDataReader.AsDataSet(conf);
            DataView dataView = dataSet.Tables[0].AsDataView();

            // После завершения чтения освобождаем ресурсы.
            ExcelDataReader.Close();
            return dataView;
        }

        private void SaveExcel_Click(object sender, RoutedEventArgs e)
        {
            if (dataSet != null && dataSet.Tables.Count > 0)
            {
                DataTable dataTable = dataSet.Tables[0];// Таблица с данными
                Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();//Приложение Excel
                Microsoft.Office.Interop.Excel.Workbook workbook = appExcel.Workbooks.Add(Type.Missing);// Файл Excel
                Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.ActiveSheet;// Страница Excel
                try
                {
                    // Построчное копирование данных на страницу Exel
                    for (int i = 1; i < dataTable.Columns.Count + 1; i++)
                    {
                        worksheet.Cells[1, i] = ExcelGrig.Columns[i - 1].Header;
                    }
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            worksheet.Cells[i + 2, j + 1] = dataTable.Rows[i].ItemArray[j].ToString();
                        }
                    }
                    // Сохранение файла
                    var saveFileDialoge = new SaveFileDialog();
                    saveFileDialoge.FileName = "FileName";
                    saveFileDialoge.DefaultExt = ".xlsx";
                    if (saveFileDialoge.ShowDialog() == true)
                    {
                        workbook.SaveAs(saveFileDialoge.FileName);
                        MessageBox.Show("Сохранение успешно!");
                    }
                    appExcel.Quit();// Завершение работы приложения Excel
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    appExcel.Quit();
                }
            }
            else
            {
                MessageBox.Show("Данных для сохранения не найдено!");
            }
        }

       // private
        private void ExcelGrig_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (dataSet != null && dataSet.Tables.Count > 0)
                {
                    DataTable dataTable = dataSet.Tables[0];// Таблица с данными
                    InformationBox.Content = "";
                    for (int i = 1; i < dataTable.Columns.Count + 1; i++)
                    {
                        InformationBox.Content += ExcelGrig.Columns[i - 1].Header
                            + ": "
                            + ((DataRowView)((DataGrid)sender).SelectedItem).Row.ItemArray[i - 1]
                            + "\n";
                    }
                }
            }
            catch
            {
                // TODO
            }
        }
    }
}
