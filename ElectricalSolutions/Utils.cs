using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using DataTable = System.Data.DataTable;
using Parameter = Autodesk.Revit.DB.Parameter;

namespace ElectricalSolutions
{
    public class Utils
    {
        public static string conduitIDParameter = "Conduit ID";
        public static DataTable conduitDataTable;

        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new System.Drawing.Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        public static BitmapSource GetBitMapSourceFromImage(Image targetImage, int width, int height)
        {
            return Imaging.CreateBitmapSourceFromHBitmap(Utils.ResizeImage(targetImage, width, height).GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
        }

        public static List<Element> GetConnectedElements(Element electricalElement)
        {
            List<Element> connectedElements = new List<Element>();
            HashSet<int> visitedElementIds = new HashSet<int>();

            GetConnectedElementsRecursive(electricalElement, connectedElements, visitedElementIds);

            return connectedElements;
        }

        public static double GetNumericParameter(Element element, string parameterName)
        {

            Parameter parameter = element.LookupParameter(parameterName);

            if (parameter == null)
            {
                return 0;
            }


            if (parameter.StorageType == StorageType.Double)
            {
                return parameter.AsDouble();
            }
            else if (parameter.StorageType == StorageType.Integer)
            {
                return parameter.AsInteger();

            }
            else if (parameter.StorageType == StorageType.String)
            {
                string value = parameter.AsString();

                bool converted = double.TryParse(value, out double valueConverted);

                if (converted)
                {
                    return valueConverted;
                }
            }

            return 0;
        }

        private static void GetConnectedElementsRecursive(Element element, List<Element> connectedElements, HashSet<int> visitedElementIds)
        {
            if (element == null || visitedElementIds.Contains(element.Id.IntegerValue))
                return;

            visitedElementIds.Add(element.Id.IntegerValue);
            connectedElements.Add(element);

            List<Connector> connectors = (element is CableTrayConduitBase)
                ? (element as CableTrayConduitBase).ConnectorManager.Connectors.OfType<Connector>().ToList()
                : (element as FamilyInstance)?.MEPModel?.ConnectorManager.Connectors.OfType<Connector>().ToList();

            if (connectors == null)
                return;

            foreach (Connector connector in connectors)
            {
                IEnumerable<Connector> connectedConnectors = connector.AllRefs.OfType<Connector>().Where(x => x.Owner != null && x.Owner.Id.IntegerValue != element.Id.IntegerValue);

                foreach (Connector connectedConnector in connectedConnectors)
                {
                    GetConnectedElementsRecursive(connectedConnector.Owner, connectedElements, visitedElementIds);
                }
            }
        }

        public static void SetConduitSpacingDatatable()
        {
            if (conduitDataTable != null)
            {
                return;
            }

            DataTable tbl = new DataTable();

            tbl.Columns.Add("Size");
            tbl.Columns.Add("0.5");
            tbl.Columns.Add("0.75");
            tbl.Columns.Add("1");
            tbl.Columns.Add("1.25");
            tbl.Columns.Add("1.5");
            tbl.Columns.Add("2");
            tbl.Columns.Add("2.5");
            tbl.Columns.Add("3");
            tbl.Columns.Add("3.5");
            tbl.Columns.Add("4");
            tbl.Columns.Add("5");
            tbl.Columns.Add("6");

            DataRow r1 = tbl.NewRow();

            string[] r1Data = { "0.5", "1.375", "1.5", "1.75", "2", "2.125", "2.375", "2.625", "3", "3.375", "3.6875", "4.375", "5" };

            string[] r2Data = { "0.75", "1.5", "1.625", "1.875", "2.125", "2.25", "2.5", "2.75", "3.125", "3.5", "3.875", "4.5", "5.125" };

            string[] r3Data = { "1", "1.75", "1.875", "2", "2.25", "2.375", "2.75", "3", "3.375", "3.625", "4", "4.625", "5.25" };

            string[] r4Data = { "1.25", "2", "2.125", "2.25", "2.5", "2.625", "3", "3.25", "3.625", "3.875", "4.25", "4.875", "5.5" };

            string[] r5Data = { "1.5", "2.125", "2.25", "2.375", "2.625", "2.75", "3.125", "3.375", "3.75", "4", "4.375", "5", "5.625" };

            string[] r6Data = { "2", "2.375", "2.5", "2.75", "3", "3.125", "3.375", "3.625", "4", "4.375", "4.75", "5.375", "6" };

            string[] r7Data = { "2.5", "2.625", "2.75", "3", "3.25", "3.375", "3.625", "4", "4.375", "4.625", "5", "5.625", "6.25" };

            string[] r8Data = { "3", "3", "3.125", "3.375", "3.625", "3.75", "4", "4.375", "4.75", "5", "5.375", "6", "6.625" };

            string[] r9Data = { "3.5", "3.375", "3.5", "3.625", "3.875", "4", "4.375", "4.625", "5", "4.75", "5.625", "6.25", "7" };

            string[] r10Data = { "4", "3.6875", "3.875", "4", "4.25", "4.375", "4.75", "5", "5.375", "5.625", "6", "6.625", "7.25" };

            string[] r11Data = { "5", "4.375", "4.5", "4.625", "4.875", "5", "5.375", "5.625", "6", "6.25", "6.625", "7.25", "8" };

            string[] r12Data = { "6", "5", "5.125", "5.25", "5.5", "5.625", "6", "6.25", "6.625", "7", "7.25", "8", "8.625" };


            tbl.Rows.Add(r1Data);
            tbl.Rows.Add(r2Data);
            tbl.Rows.Add(r3Data);
            tbl.Rows.Add(r4Data);
            tbl.Rows.Add(r5Data);
            tbl.Rows.Add(r6Data);
            tbl.Rows.Add(r7Data);
            tbl.Rows.Add(r8Data);
            tbl.Rows.Add(r9Data);
            tbl.Rows.Add(r10Data);
            tbl.Rows.Add(r11Data);
            tbl.Rows.Add(r12Data);

            conduitDataTable = tbl;
        }

        public static string DataTableToString(DataTable dataTable)
        {
            if (dataTable == null || dataTable.Rows.Count == 0)
                return string.Empty;

            // Use a StringWriter to write the DataTable to an XML string
            using (StringWriter writer = new StringWriter())
            {
                dataTable.WriteXml(writer, XmlWriteMode.WriteSchema, false);
                return writer.ToString();
            }
        }

        public static DataTable StringToDataTable(string dataString)
        {
            if (string.IsNullOrEmpty(dataString))
                return null;
            
            // Use a StringReader to read the XML string and create a DataTable
            using (StringReader reader = new StringReader(dataString))
            {
                DataTable dataTable = new DataTable("ConduitData");
                dataTable.ReadXml(reader);
                return dataTable;
            }
        }


        public static bool IsConduitMagnetic(ElementId typeId, Document doc)
        {
            Element element = doc.GetElement(typeId);

            bool isMagnetic = element.Name.ToLower().Contains("emt") || element.Name.ToLower().Contains("steel");


            return isMagnetic;
        }

        public static string GetWireSize(Element conduit)
        {
           Parameter parameter= conduit.LookupParameter("Wire Size");

            if (parameter == null || parameter.HasValue==false)
            {
                return string.Empty;

            }

            return parameter.AsString();
        }

        public static string GetWireType(Element conduit)
        {
            Parameter parameter = conduit.LookupParameter("Wire Type");

            if (parameter == null || parameter.HasValue == false)
            {
                return string.Empty;

            }

            return parameter.AsString();
        }


        public static string GetPowerFactor(Element conduit)
        {
            Parameter parameter = conduit.LookupParameter("Power Factor");

            if (parameter == null || parameter.HasValue == false)
            {
                return string.Empty;

            }

            if (parameter.StorageType == StorageType.String)
            {
                return parameter.AsString();
            }
            else
            {
                return parameter.AsValueString();

            }
            
        }

        public static void CalculateVoltageDropOneConduit(Element electricalElement)
        {
            Parameter voltageDropParameter = electricalElement.LookupParameter("Voltage Drop");


            if (voltageDropParameter == null)
            {
                return;
            }

            List<Element> listOfConnectedElements =  GetConnectedElements(electricalElement);

            if (!listOfConnectedElements.Any())
            {
                return;
            }

            double totalLength = 0;
            foreach (Element element in listOfConnectedElements)
            {

                if (element is Conduit)
                {
                    totalLength += (element as Conduit).get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                }
                else
                {
                    Parameter fittingLength = element.LookupParameter("Fitting Length");

                    if (fittingLength == null)
                    {
                        continue;
                    }

                    totalLength += fittingLength.AsDouble();
                }
            }

            if (totalLength==0)
            {
                return;
            }

            Element conduitElement = listOfConnectedElements.FirstOrDefault(x => x is Conduit);

            if (conduitElement==null)
            {
                return;
            }

            bool isMagnetic = IsConduitMagnetic(conduitElement.GetTypeId(), conduitElement.Document);
            double current =   GetCurrentValue(conduitElement); 
            string wireSize = GetWireSize(conduitElement);
            string powerFactor = GetPowerFactor(conduitElement);
            string wireType = GetWireType(conduitElement);
            double sourceVoltage = GetSourceVoltage(conduitElement);

            bool validSize = IsValidWireSize(wireSize);
            int convertedPowerFactor = string.IsNullOrEmpty(powerFactor)?0:int.Parse(powerFactor);
            bool validPower = IsValidPowerFactor(convertedPowerFactor);
            bool validWireType = IsValidWireType(wireType);
            sourceVoltage = sourceVoltage == 0 ? 0.1 : sourceVoltage;

            convertedPowerFactor = validPower ? convertedPowerFactor : 100;
            wireSize = validSize ? wireSize : "14";
            wireType = validWireType ? wireType : "Copper";


            using (TransactionGroup group = new TransactionGroup(electricalElement.Document, "Set Voltage Drop and Length"))
            {
                group.Start();
                using (Transaction t = new Transaction(electricalElement.Document, "Set Length"))
                {
                    t.Start();
                    foreach (Element element1 in listOfConnectedElements)
                    {
                        Parameter totalRunLengthParameter = element1.LookupParameter("Total Run Length");

                        if (totalRunLengthParameter == null)
                        {
                            continue;
                        }

                        totalRunLengthParameter.Set(totalLength);


                    }
                    t.Commit();
                }

                double voltageDrop = CalculateVoltageDrop(current, totalLength, wireType, isMagnetic ? "Magnetic" : "NonMagnetic", convertedPowerFactor, wireSize);
                

                double dropPercentage = (voltageDrop / sourceVoltage)*100+0.3;

                using (Transaction t = new Transaction(electricalElement.Document, "Set Voltage Drop"))
                {
                    t.Start();
                    foreach (Element element1 in listOfConnectedElements)
                    {
                        Parameter vDropParameter = element1.LookupParameter("Voltage Drop");

                        if (vDropParameter == null)
                        {
                            continue;
                        }

                        vDropParameter.Set(dropPercentage);


                    }      
                    t.Commit();
                }
                group.Assimilate();
            }
        }

        public static double GetCurrentValue(Element conduitElement)
        {
            Parameter parameter = conduitElement.LookupParameter("Amperage");

            if (parameter == null || parameter.HasValue == false)
            {
                return 0;
            }

            if (parameter.StorageType == StorageType.Double)
            {
                return parameter.AsDouble();
            }
            else if (parameter.StorageType == StorageType.String)
            {
                string value = parameter.AsString();

                bool isNumber = double.TryParse(value, out double valueDouble);

                if (isNumber)
                {
                    return valueDouble;
                }
            }

            return 0;
        }

        public static double GetSourceVoltage(Element conduitElement)
        {
            Parameter parameter = conduitElement.LookupParameter("Voltage");

            if (parameter == null || parameter.HasValue == false)
            {
                return 0.0001;
            }

            if (parameter.StorageType == StorageType.Double)
            {

                if (parameter.Definition.GetDataType()== SpecTypeId.ElectricalPotential)
                {
                   return UnitUtils.ConvertFromInternalUnits(parameter.AsDouble(), UnitTypeId.Volts);
                }
                return parameter.AsDouble();
            }
            else if (parameter.StorageType == StorageType.String)
            {
                string value = parameter.AsString();

                bool isNumber = double.TryParse(value, out double valueDouble);

                if (isNumber)
                {
                    return valueDouble;
                }
            }

            return 0.0001;
        }

        public bool CheckParametersExist(Document doc)
        {
            List<string> requiredParameterNames = new List<string>
        {
            "Voltage",
            "Conduit Size",
            "Fitting Length",
            "Circuits",
            "Wire Size",
            "Origin",
            "Destination",
            "Wire Type",
            "Total Run Length",
            "Voltage Drop",
            "Level",
            "Amperage",
            "Conduit Fill",
            "Wire Count"
        };

            foreach (string paramName in requiredParameterNames)
            {
                if (!ParameterExists(paramName, doc))
                {
                    return false;
                }
            }

            return true;
        }

        private bool ParameterExists(string paramName, Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ParameterElement));

            ParameterElement parameter = collector
                .FirstOrDefault(p => p.Name == paramName) as ParameterElement;

            return parameter != null;
        }

        // Define a unique GUID for the schema (Replace with your own unique GUID)
        private static readonly Guid SchemaGUID = new Guid("A1654022-3109-4B7B-9F8D-D70BA6BE5A58");

        public static void SaveStringToExtensibleStorage(Document document, string data)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            if (string.IsNullOrEmpty(data))
                return;

            // Get or create the schema definition
            Schema schema = Schema.Lookup(SchemaGUID);

            if (schema == null)
            {
                // Create the schema if it doesn't exist
                SchemaBuilder schemaBuilder = new SchemaBuilder(SchemaGUID);

                // Define a field for the string data
                FieldBuilder dataFieldBuilder = schemaBuilder.AddSimpleField("B15ConduitData", typeof(string));
                schemaBuilder.SetSchemaName("B15");
                schema = schemaBuilder.Finish();
            }

            // Get or create the entity for the project document
            Entity entity = document.ProjectInformation.GetEntity(schema);

            // If the entity is not valid, start a transaction to make it valid
            using (Transaction transaction = new Transaction(document, "Make Entity Valid"))
            {
                transaction.Start();
                entity = new Entity(schema);
                entity.Set(schema.GetField("B15ConduitData"), data);
                document.ProjectInformation.SetEntity(entity);
                transaction.Commit();
            }

            // Set the string data in the entity


        }

        public static string ReadStringFromExtensibleStorage(Element element)
        {
            if (element == null)
                throw new ArgumentNullException(nameof(element));

            // Get the extensible storage schema
            Schema schema = Schema.Lookup(SchemaGUID);

            if (schema == null)
            {
                // Schema does not exist; the data hasn't been saved for this element.
                return string.Empty;
            }

            // Get the entity associated with the element
            Entity entity = element.GetEntity(schema);

            if (entity.IsValid())
            {
                // Read the data from the entity and return the string value
                Field dataField = schema.GetField("B15ConduitData");
                if (dataField != null && entity.IsValid())
                {
                    string data = entity.Get<string>(dataField);
                    return data;
                }
            }

            return string.Empty; // If the data field doesn't exist or is not valid, return an empty string.
        }

        public static DataTable ReadDataTableFromExtensibleStorage(Document document)
        {
            if (document == null)
                return null;

            // Get the extensible storage schema
            Schema schema = Schema.Lookup(SchemaGUID);

            if (schema == null)
            {
                // Schema does not exist; the data hasn't been saved for this element.
                return null;
            }

            // Get the entity associated with the project document
            Entity entity = document.ProjectInformation.GetEntity(schema);

            if (!entity.IsValid())
            {
                // Entity does not exist; the data hasn't been saved for this element.
                return null;
            }

            // Read the data from the entity and convert it back to a DataTable
            Field dataField = schema.GetField("B15ConduitData");
            if (dataField != null && entity.IsValid())
            {
                string data = entity.Get<string>(dataField);
                return StringToDataTable(data);
            }

            return null; // If the data field doesn't exist or is not valid, return null.
        }

        public static DataRow LookupLineInDataTable(Document document, string conduitId)
        {
            // Read the DataTable from extensible storage
            DataTable dataTable = ReadDataTableFromExtensibleStorage(document);

            if (dataTable == null || string.IsNullOrEmpty(conduitId))
                return null;

            // Perform the lookup based on "Conduit_ID" parameter
            DataRow[] foundRows = dataTable.Select("Conduit ID = '" + conduitId + "'");

            // Return the first matched row (assuming there is only one matching row)
            return foundRows.Length > 0 ? foundRows[0] : null;
        }

        public static void UpdateElementParametersFromDataRow(Element element, DataRow dataRow)
        {
            if (element == null)
                throw new ArgumentNullException(nameof(element));

            if (dataRow == null)
                throw new ArgumentNullException(nameof(dataRow));

            // Get the element's parameter bindings (definition and value)
            var parameterBindings = element.GetOrderedParameters()
                .Select(p => new { Definition = p.Definition, Value = p.AsValueString() })
                .ToList();

            // Loop through the DataRow columns (headers) and update the element's parameters
            foreach (DataColumn column in dataRow.Table.Columns)
            {
                var matchingParameterBinding = parameterBindings
                    .FirstOrDefault(binding => binding.Definition.Name.Equals(column.ColumnName));

                if (matchingParameterBinding != null)
                {
                    Parameter parameter = element.LookupParameter(matchingParameterBinding.Definition.Name);

                    if (parameter != null && !parameter.IsReadOnly)
                    {
                        // Convert the data from DataRow to the appropriate parameter type (handle different data types as needed)
                        string dataValue = dataRow[column].ToString();
                        if (parameter.StorageType == StorageType.Integer)
                        {
                            if (int.TryParse(dataValue, out int intValue))
                            {
                                parameter.Set(intValue);
                            }
                        }
                        else if (parameter.StorageType == StorageType.Double)
                        {
                            if (double.TryParse(dataValue, out double doubleValue))
                            {
                                parameter.Set(doubleValue);
                            }
                        }
                        else if (parameter.StorageType == StorageType.String)
                        {
                            parameter.Set(dataValue);
                        }
                        // Handle other parameter types as needed

                        // Optionally, you can remove the parameter from the list to speed up future lookups
                        parameterBindings.Remove(matchingParameterBinding);
                    }
                }
            }
        }

        public static string GetSpacingBasedOnConduitSize(string sizeInInchesConduit1, string sizeInInchesConduit2)
        {
            SetConduitSpacingDatatable();

            DataColumn column = conduitDataTable.Columns.OfType<DataColumn>().First(x => x.ColumnName == sizeInInchesConduit1);

            if (column == null)
            {
                return "";
            }

            DataColumn firstColumn = conduitDataTable.Columns[0];

            DataRow row = conduitDataTable.Rows.OfType<DataRow>().First(x => x[firstColumn].ToString() == sizeInInchesConduit2);


            if (row == null)
            {
                return "";
            }


            string value = row[column].ToString();

            return value;
        }

        public static void UpdateConnectedElementsParameters(Document document, Element electricalElement)
        {

            if (document == null)
                throw new ArgumentNullException(nameof(document));

            if (electricalElement == null)
                throw new ArgumentNullException(nameof(electricalElement));
            
            if (conduitDataTable==null)
            {
                conduitDataTable = ReadDataTableFromExtensibleStorage(document);
            }
            // Get the DataTable from the extensible storage
            DataTable dataTable = conduitDataTable;

            if (dataTable == null)
                return;

            // Find the "Conduit_ID" parameter in the element
            Parameter conduitIdParameter = electricalElement.LookupParameter(conduitIDParameter);

            if (conduitIdParameter == null)
                return;

            // Get the "Conduit_ID" value from the element
            string conduitIdValue = conduitIdParameter.AsString();

            // Find the DataRow with matching "Conduit_ID" in the DataTable
            DataRow foundRow = null;

            try
            {
                foundRow= dataTable.AsEnumerable().FirstOrDefault(row => row[conduitIDParameter].ToString() == conduitIdValue);
            }
            catch (Exception e)
            {

                TaskDialog.Show("Error",e.Message);
                return;
            }

            if (foundRow == null)
                return;

            // Update the "Conduit_ID" parameter in the entry element
            conduitIdParameter.Set(conduitIdValue);
            ConduitIdUpdater.updatingConnectedElements = true;
            // Update the other parameters in the connected elements
            List<Element> connectedElements = GetConnectedElements(electricalElement);
            foreach (Element connectedElement in connectedElements)
            {
                // Skip the electricalElement as it has already been updated
                if (connectedElement.Id == electricalElement.Id)
                    continue;

                // Update the parameters in the connected elements
                UpdateElementParametersFromDataRow(connectedElement, foundRow);

                // Set the "Conduit_ID" parameter in the connected elements
                Parameter connectedElementConduitIdParameter = connectedElement.LookupParameter(conduitIDParameter);
                if (connectedElementConduitIdParameter != null)
                {
                    connectedElementConduitIdParameter.Set(conduitIdValue);
                }
            }
            UpdateElementParametersFromDataRow(electricalElement, foundRow);

            ConduitIdUpdater.updatingConnectedElements = false;
        }

        public static DataTable ReadExcelFile(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbook = package.Workbook;
                if (workbook.Worksheets.Count == 0)
                {
                    throw new Exception("The Excel file does not contain any sheets.");
                }

                var worksheet = workbook.Worksheets[0];
                var dataTable = new DataTable("ConduitData");

                // Assume the first row contains headers
                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(firstRowCell.Text);
                }

                // Start from the second row to add data rows
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var newRow = dataTable.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    dataTable.Rows.Add(newRow);
                }

                return dataTable;
            }
        }

        public static void WriteExcelFile(DataTable dataTable, string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Add column headers
                for (int col = 1; col <= dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[1, col].Value = dataTable.Columns[col - 1].ColumnName;
                }

                // Add rows
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                    }
                }

                // Save to file
                File.WriteAllBytes(filePath, package.GetAsByteArray());
            }
        }

        public static void ImportExcelData(Document document)
        {
            ExcelSheetView excelView = new ExcelSheetView(document);
            excelView.ShowDialog();
           
        }

        private static Dictionary<string, Dictionary<int, double>> aluminumMagneticConduit = new Dictionary<string, Dictionary<int, double>>
        {
             {"12", new Dictionary<int, double>
        {
            {60, 0.3296}, {70, 0.3811}, {80, 0.4349}, {90, 0.4848}, {100, 0.533}
        }
    },
    {"10", new Dictionary<int, double>
        {
            {60, 0.2133}, {70, 0.2429}, {80, 0.2741}, {90, 0.318}, {100, 0.3363}
        }
    },
    {"8", new Dictionary<int, double>
        {
            {60, 0.1305}, {70, 0.1552}, {80, 0.1758}, {90, 0.1951}, {100, 0.2106}
        }
    },
    {"6", new Dictionary<int, double>
        {
            {60, 0.0898}, {70, 0.1018}, {80, 0.1142}, {90, 0.1254}, {100, 0.1349}
        }
    },
    {"4", new Dictionary<int, double>
        {
            {60, 0.0595}, {70, 0.066}, {80, 0.0747}, {90, 0.0809}, {100, 0.0862}
        }
    },
    {"2", new Dictionary<int, double>
        {
            {60, 0.0403}, {70, 0.0443}, {80, 0.0483}, {90, 0.0523}, {100, 0.0535}
        }
    },
    {"1", new Dictionary<int, double>
        {
            {60, 0.0332}, {70, 0.0357}, {80, 0.0396}, {90, 0.0423}, {100, 0.0428}
        }
    },
    {"1/0", new Dictionary<int, double>
        {
            {60, 0.0286}, {70, 0.0305}, {80, 0.0334}, {90, 0.035}, {100, 0.0341}
        }
    },
    {"2/0", new Dictionary<int, double>
        {
            {60, 0.0234}, {70, 0.0246}, {80, 0.0275}, {90, 0.0284}, {100, 0.0274}
        }
    },
    {"3/0", new Dictionary<int, double>
        {
            {60, 0.0209}, {70, 0.022}, {80, 0.0231}, {90, 0.0241}, {100, 0.0217}
        }
    },
    {"4/0", new Dictionary<int, double>
        {
            {60, 0.0172}, {70, 0.0174}, {80, 0.0179}, {90, 0.0177}, {100, 0.017}
        }
    },
    {"250", new Dictionary<int, double>
        {
            {60, 0.0158}, {70, 0.0163}, {80, 0.0162}, {90, 0.0159}, {100, 0.0145}
        }
    },
    {"300", new Dictionary<int, double>
        {
            {60, 0.0137}, {70, 0.0139}, {80, 0.0143}, {90, 0.0144}, {100, 0.0122}
        }
    },
    {"350", new Dictionary<int, double>
        {
            {60, 0.013}, {70, 0.0133}, {80, 0.0128}, {90, 0.0131}, {100, 0.01}
        }
    },
    {"500", new Dictionary<int, double>
        {
            {60, 0.0112}, {70, 0.0111}, {80, 0.0114}, {90, 0.0099}, {100, 0.0076}
        }
    },
    {"600", new Dictionary<int, double>
        {
            {60, 0.0101}, {70, 0.0106}, {80, 0.0097}, {90, 0.009}, {100, 0.0063}
        }
    },
    {"750", new Dictionary<int, double>
        {
            {60, 0.0095}, {70, 0.0094}, {80, 0.009}, {90, 0.0084}, {100, 0.0056}
        }
    },
    {"1000", new Dictionary<int, double>
        {
            {60, 0.0085}, {70, 0.0082}, {80, 0.0078}, {90, 0.0071}, {100, 0.0043}
        }
    }
        };
                
        private static Dictionary<string, Dictionary<int, double>> aluminumNonMagneticConduit = new Dictionary<string, Dictionary<int, double>>
        {
            {"12", new Dictionary<int, double>
        {
            {60, 0.3312}, {70, 0.3802}, {80, 0.4328}, {90, 0.4848}, {100, 0.5331}
        }
    },
    {"10", new Dictionary<int, double>
        {
            {60, 0.209}, {70, 0.241}, {80, 0.274}, {90, 0.3052}, {100, 0.3363}
        }
    },
    {"8", new Dictionary<int, double>
        {
            {60, 0.1286}, {70, 0.1534}, {80, 0.1745}, {90, 0.1933}, {100, 0.2115}
        }
    },
    {"6", new Dictionary<int, double>
        {
            {60, 0.0887}, {70, 0.1011}, {80, 0.1127}, {90, 0.1249}, {100, 0.1361}
        }
    },
    {"4", new Dictionary<int, double>
        {
            {60, 0.0583}, {70, 0.0654}, {80, 0.0719}, {90, 0.08}, {100, 0.0849}
        }
    },
    {"2", new Dictionary<int, double>
        {
            {60, 0.0389}, {70, 0.0435}, {80, 0.0473}, {90, 0.0514}, {100, 0.0544}
        }
    },
    {"1", new Dictionary<int, double>
        {
            {60, 0.0318}, {70, 0.0349}, {80, 0.0391}, {90, 0.0411}, {100, 0.0428}
        }
    },
    {"1/0", new Dictionary<int, double>
        {
            {60, 0.0263}, {70, 0.0287}, {80, 0.0322}, {90, 0.0337}, {100, 0.0339}
        }
    },
    {"2/0", new Dictionary<int, double>
        {
            {60, 0.0227}, {70, 0.0244}, {80, 0.0264}, {90, 0.0274}, {100, 0.0273}
        }
    },
    {"3/0", new Dictionary<int, double>
        {
            {60, 0.016}, {70, 0.0171}, {80, 0.0218}, {90, 0.0233}, {100, 0.0222}
        }
    },
    {"4/0", new Dictionary<int, double>
        {
            {60, 0.0152}, {70, 0.0159}, {80, 0.0171}, {90, 0.0179}, {100, 0.0172}
        }
    },
    {"250", new Dictionary<int, double>
        {
            {60, 0.0138}, {70, 0.0144}, {80, 0.0147}, {90, 0.0155}, {100, 0.0138}
        }
    },
    {"300", new Dictionary<int, double>
        {
            {60, 0.0126}, {70, 0.0128}, {80, 0.0133}, {90, 0.0132}, {100, 0.0125}
        }
    },
    {"350", new Dictionary<int, double>
        {
            {60, 0.0122}, {70, 0.0123}, {80, 0.0119}, {90, 0.012}, {100, 0.0101}
        }
    },
    {"500", new Dictionary<int, double>
        {
            {60, 0.0093}, {70, 0.0094}, {80, 0.0094}, {90, 0.0091}, {100, 0.0072}
        }
    },
    {"600", new Dictionary<int, double>
        {
            {60, 0.0084}, {70, 0.0085}, {80, 0.0085}, {90, 0.0081}, {100, 0.006}
        }
    },
    {"750", new Dictionary<int, double>
        {
            {60, 0.0081}, {70, 0.008}, {80, 0.0078}, {90, 0.0072}, {100, 0.0051}
        }
    },
    {"1000", new Dictionary<int, double>
        {
            {60, 0.0069}, {70, 0.0068}, {80, 0.0065}, {90, 0.0058}, {100, 0.0038}
        }
    }
        };
                
        private static Dictionary<string, Dictionary<int, double>> cooperMagneticConduit = new Dictionary<string, Dictionary<int, double>>
        {
            {"14", new Dictionary<int, double>
        {
            {60, 0.3390}, {70, 0.3910}, {80, 0.4430}, {90, 0.4940}, {100, 0.5410}
        }
    },
    {"12", new Dictionary<int, double>
        {
            {60, 0.2170}, {70, 0.2490}, {80, 0.2810}, {90, 0.3130}, {100, 0.3410}
        }
    },
    {"10", new Dictionary<int, double>
        {
            {60, 0.1390}, {70, 0.1590}, {80, 0.1790}, {90, 0.1980}, {100, 0.2150}
        }
    },
    {"8", new Dictionary<int, double>
        {
            {60, 0.0905}, {70, 0.1030}, {80, 0.1150}, {90, 0.1260}, {100, 0.1350}
        }
    },
    {"6", new Dictionary<int, double>
        {
            {60, 0.0595}, {70, 0.0670}, {80, 0.0742}, {90, 0.0809}, {100, 0.0850}
        }
    },
    {"4", new Dictionary<int, double>
        {
            {60, 0.0399}, {70, 0.0443}, {80, 0.0485}, {90, 0.0522}, {100, 0.0534}
        }
    },
    {"2", new Dictionary<int, double>
        {
            {60, 0.0275}, {70, 0.0300}, {80, 0.0323}, {90, 0.0342}, {100, 0.0336}
        }
    },
    {"1", new Dictionary<int, double>
        {
            {60, 0.0233}, {70, 0.0251}, {80, 0.0267}, {90, 0.0279}, {100, 0.0267}
        }
    },
    {"1/0", new Dictionary<int, double>
        {
            {60, 0.0198}, {70, 0.0211}, {80, 0.0222}, {90, 0.0229}, {100, 0.0213}
        }
    },
    {"2/0", new Dictionary<int, double>
        {
            {60, 0.0171}, {70, 0.0180}, {80, 0.0187}, {90, 0.0190}, {100, 0.0170}
        }
    },
    {"3/0", new Dictionary<int, double>
        {
            {60, 0.0148}, {70, 0.0154}, {80, 0.0158}, {90, 0.0158}, {100, 0.0136}
        }
    },
    {"4/0", new Dictionary<int, double>
        {
            {60, 0.0130}, {70, 0.0134}, {80, 0.0136}, {90, 0.0133}, {100, 0.0109}
        }
    },
    {"250", new Dictionary<int, double>
        {
            {60, 0.0122}, {70, 0.0124}, {80, 0.0124}, {90, 0.0120}, {100, 0.0094}
        }
    },
    {"300", new Dictionary<int, double>
        {
            {60, 0.0111}, {70, 0.0112}, {80, 0.0111}, {90, 0.0106}, {100, 0.0080}
        }
    },
    {"350", new Dictionary<int, double>
        {
            {60, 0.0104}, {70, 0.0104}, {80, 0.0102}, {90, 0.0096}, {100, 0.0069}
        }
    },
    {"500", new Dictionary<int, double>
        {
            {60, 0.0100}, {70, 0.0091}, {80, 0.0087}, {90, 0.0080}, {100, 0.0053}
        }
    },
    {"600", new Dictionary<int, double>
        {
            {60, 0.0088}, {70, 0.0086}, {80, 0.0082}, {90, 0.0074}, {100, 0.0046}
        }
    },
    {"750", new Dictionary<int, double>
        {
            {60, 0.0084}, {70, 0.0081}, {80, 0.0077}, {90, 0.0069}, {100, 0.0040}
        }
    },
    {"1000", new Dictionary<int, double>
        {
            {60, 0.0080}, {70, 0.0077}, {80, 0.0072}, {90, 0.0063}, {100, 0.0035}
        }
    }
        };
                
        private static Dictionary<string, Dictionary<int, double>> cooperNonMagneticConduit = new Dictionary<string, Dictionary<int, double>>
        {
            {"14", new Dictionary<int, double>
        {
            {60, 0.3370}, {70, 0.3900}, {80, 0.4410}, {90, 0.4930}, {100, 0.5410}
        }
    },
    {"12", new Dictionary<int, double>
        {
            {60, 0.2150}, {70, 0.2480}, {80, 0.2800}, {90, 0.3120}, {100, 0.3410}
        }
    },
    {"10", new Dictionary<int, double>
        {
            {60, 0.1370}, {70, 0.1580}, {80, 0.1780}, {90, 0.1970}, {100, 0.2150}
        }
    },
    {"8", new Dictionary<int, double>
        {
            {60, 0.0888}, {70, 0.1010}, {80, 0.1140}, {90, 0.1250}, {100, 0.1350}
        }
    },
    {"6", new Dictionary<int, double>
        {
            {60, 0.0579}, {70, 0.0656}, {80, 0.0730}, {90, 0.0800}, {100, 0.0849}
        }
    },
    {"4", new Dictionary<int, double>
        {
            {60, 0.0384}, {70, 0.0430}, {80, 0.0473}, {90, 0.0513}, {100, 0.0533}
        }
    },
    {"2", new Dictionary<int, double>
        {
            {60, 0.0260}, {70, 0.0287}, {80, 0.0312}, {90, 0.0333}, {100, 0.0335}
        }
    },
    {"1", new Dictionary<int, double>
        {
            {60, 0.0218}, {70, 0.0238}, {80, 0.0256}, {90, 0.0270}, {100, 0.0266}
        }
    },
    {"1/0", new Dictionary<int, double>
        {
            {60, 0.0183}, {70, 0.0198}, {80, 0.0211}, {90, 0.0220}, {100, 0.0211}
        }
    },
    {"2/0", new Dictionary<int, double>
        {
            {60, 0.0156}, {70, 0.0167}, {80, 0.0176}, {90, 0.0181}, {100, 0.0169}
        }
    },
    {"3/0", new Dictionary<int, double>
        {
            {60, 0.0134}, {70, 0.0141}, {80, 0.0147}, {90, 0.0149}, {100, 0.0134}
        }
    },
    {"4/0", new Dictionary<int, double>
        {
            {60, 0.0116}, {70, 0.0121}, {80, 0.0124}, {90, 0.0124}, {100, 0.0107}
        }
    },
    {"250", new Dictionary<int, double>
        {
            {60, 0.0107}, {70, 0.0111}, {80, 0.0112}, {90, 0.0110}, {100, 0.0091}
        }
    },
    {"300", new Dictionary<int, double>
        {
            {60, 0.0097}, {70, 0.0099}, {80, 0.0099}, {90, 0.0096}, {100, 0.0077}
        }
    },
    {"350", new Dictionary<int, double>
        {
            {60, 0.0090}, {70, 0.0091}, {80, 0.0091}, {90, 0.0087}, {100, 0.0066}
        }
    },
    {"500", new Dictionary<int, double>
        {
            {60, 0.0078}, {70, 0.0077}, {80, 0.0075}, {90, 0.0070}, {100, 0.0049}
        }
    },
    {"600", new Dictionary<int, double>
        {
            {60, 0.0074}, {70, 0.0072}, {80, 0.0070}, {90, 0.0064}, {100, 0.0042}
        }
    },
    {"750", new Dictionary<int, double>
        {
            {60, 0.0069}, {70, 0.0067}, {80, 0.0064}, {90, 0.0058}, {100, 0.0035}
        }
    },
    {"1000", new Dictionary<int, double>
        {
            {60, 0.0064}, {70, 0.0062}, {80, 0.0058}, {90, 0.0052}, {100, 0.0029}
        }
    }
        };
        
        public static bool IsValidWireSize(string wireSize)
        {
            string[] validWireSizes = { "14", "12", "10", "8", "6", "4", "2", "1", "1/0", "2/0", "3/0", "4/0", "250", "300", "350", "500", "600", "750", "1000" };
            return Array.Exists(validWireSizes, size => size.Equals(wireSize, StringComparison.OrdinalIgnoreCase));
        }
        public static bool IsValidWireType(string wireType)
        {
            string[] validWireTypes = { "copper", "aluminum" };
            return Array.Exists(validWireTypes, type => type.Equals(wireType, StringComparison.OrdinalIgnoreCase));
        }
        public static bool IsValidPowerFactor(int powerFactor)
        {
            int[] validPowerFactors = { 60, 70, 80, 90, 100 };
            return Array.Exists(validPowerFactors, factor => factor == powerFactor);
        }
        public static double CalculateVoltageDrop(double current, double length,  string wireType, string conduitType, int powerFactor, string wireSize)
        {

            if (!IsValidWireSize(wireSize))
            {
                return -1;
            }
            if (length<=0 || current<=0)
            {
                return -1;
            }
            if (!IsValidWireType(wireType))
            {
                return - 1;
            }
            if (!IsValidPowerFactor(powerFactor))
            {
                return -1;
            }

            Dictionary<string, Dictionary<int, double>> selectedTable = GetSelectedTable(wireType, conduitType);
            
            double resistance = selectedTable[wireSize][powerFactor];

            double voltageDrop = resistance * length * current / 100;

            return voltageDrop; 
        }

        private static Dictionary<string, Dictionary<int, double>> GetSelectedTable(string wireType, string conduitType)
        {
            if (wireType == "Aluminum" && conduitType == "Magnetic")
            {
                return aluminumMagneticConduit;
            }
            else if (wireType == "Aluminum" && conduitType == "NonMagnetic")
            {
                return aluminumNonMagneticConduit;
            }
            else if (wireType == "Copper" && conduitType == "Magnetic")
            {
                return cooperMagneticConduit;
            }
            else if (wireType == "Copper" && conduitType == "NonMagnetic")
            {
                return cooperNonMagneticConduit;
            }

            throw new ArgumentException("Invalid wire type or conduit type.");
        }

        public static double CalculateConduitFillPercentage(string wireSize, int wireQuantity, double conduitSizeInches)
        {
            // Create a dictionary of wire sizes and their corresponding areas (in square inches)
            Dictionary<string, double> wireAreas = new Dictionary<string, double>
        {
            { "14", 0.0026 },
            { "12", 0.0041 },
            { "10", 0.0065 },
            { "8", 0.0104 },
            { "6", 0.0165 },
            { "4", 0.0262 },
            { "2", 0.0418 },
            { "1", 0.0664 },
            { "1/0", 0.105 },
            { "2/0", 0.167 },
            { "3/0", 0.265 },
            { "4/0", 0.419 },
            { "250", 0.667 },
            { "300", 1.055 },
            { "350", 1.331 },
            // Continue adding more sizes as needed
        };

            // Validate input parameters
            if (!wireAreas.ContainsKey(wireSize))
            {
                throw new ArgumentException("Invalid wire size.");
            }

            if (wireQuantity <= 0)
            {
                throw new ArgumentException("Invalid wire quantity. Must be greater than zero.");
            }

            // Calculate total wire area
            double totalWireArea = wireAreas[wireSize] * wireQuantity;

            // Calculate conduit area based on diameter
            double conduitRadius = conduitSizeInches / 2.0;
            double conduitArea = Math.PI * Math.Pow(conduitRadius, 2);

            // Calculate conduit fill percentage
            double fillPercentage = (totalWireArea / conduitArea) * 100;
            return fillPercentage;
        }


    }

}





