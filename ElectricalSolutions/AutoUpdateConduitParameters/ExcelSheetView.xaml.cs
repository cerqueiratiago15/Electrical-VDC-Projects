using Autodesk.Revit.DB;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
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

namespace ElectricalSolutions
{
    /// <summary>
    /// Interaction logic for ExcelSheetView.xaml
    /// </summary>
    public partial class ExcelSheetView : Window
    {
        private DataTable importedData;
        public string StringData { get; set; }
        public Document Doc { get; set; }


        public ExcelSheetView(Document doc)
        {
            InitializeComponent();
            this.Doc = doc;
           this.importedData = Utils.ReadDataTableFromExtensibleStorage(doc);
            if (this.importedData !=null)
            {
                dataGrid.ItemsSource = importedData.DefaultView; 
            }
        }

        private void btnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                importedData = Utils.ReadExcelFile(filePath); // Assuming you have this method
                dataGrid.ItemsSource = importedData.DefaultView;
            }
        }

        private void btnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            if (importedData == null)
            {
                MessageBox.Show("No data to export!");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                FileName = "ExportedData.xlsx"
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;
                Utils.WriteExcelFile(importedData, filePath);
                
            }
        }

        private void btnConvertToString_Click(object sender, RoutedEventArgs e)
        {
            if (importedData == null)
            {
                MessageBox.Show("No data to convert!");
                return;
            }

            this.StringData = Utils.DataTableToString(importedData);
            Utils.SaveStringToExtensibleStorage(this.Doc, this.StringData);
            Utils.conduitDataTable = this.importedData;
        }
    }
}
