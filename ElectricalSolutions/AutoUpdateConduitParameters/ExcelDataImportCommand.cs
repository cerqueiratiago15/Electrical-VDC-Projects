using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace ElectricalSolutions
{
    [Transaction(TransactionMode.Manual)]
    public class ExcelDataImportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Get the active document
                Document doc = commandData.Application.ActiveUIDocument.Document;

                Utils.ImportExcelData(doc);

                #region old code
                //// Prompt the user to select an Excel file
                //var fileDialog = new System.Windows.Forms.OpenFileDialog();
                //fileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                //fileDialog.Title = "Select an Excel File";
                //fileDialog.Multiselect = false;

                //if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                //{
                //    string filePath = fileDialog.FileName;

                //    // Call the ExcelDataImporter to import data into extensible storage
                //    Utils.ImportExcelData(doc, filePath);

                //    TaskDialog.Show("Excel Data Import", "Excel data has been imported successfully!");
                //} 
                #endregion

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}





