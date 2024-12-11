using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using ElectricalSolutions.Properties;
using OfficeOpenXml;
using SJSSolutions.Properties;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ElectricalSolutions
{
    public class ElectricalApplication : IExternalApplication
    {
        public static string tabName = "Electrical App";
        public static string currentDll = string.Empty;
        public static string nameSpace = string.Empty;
        internal static AddInId AddinID;

        public Result OnShutdown(UIControlledApplication application)
        {

            return Result.Succeeded;
        }


       
        public Result OnStartup(UIControlledApplication application)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            currentDll = System.Reflection.Assembly.GetExecutingAssembly().Location;
            nameSpace = this.GetType().Assembly.GetName().Name + ".";
            AddinID = application.ActiveAddInId;

            try
            {
                application.CreateRibbonTab(tabName);
               
                RibbonPanel utilitiesPanel = AddRibbonPanel(application,
                                                              tabName,
                                                              "Utilities",
                                                              true);

                string copyParameterTooltip = "Copy parameters from a source element. First select the elements that you want to copy to and then select the element that contain the parameter values";
             
                //CreatePushButton(utilitiesPanel, "Copy Parameter", Resources.CopyParameters,nameof(CopyParameterCommand), copyParameterTooltip);
                CreatePushButton(utilitiesPanel, "Conduit Data\nExcel Sync", Resources.ImportExcel,nameof(ExcelDataImportCommand), "Sync the data from the conduits with an excel file through the Conduit ID parameter");
               // CreatePushButton(utilitiesPanel, "Voltage Drop", Resources.VoltageDrop,nameof(VoltageDropCommand), "Calculates the voltage drop");
            }
            catch 
            {

                
            }


            // Find the GUID of the shared parameter named "Conduit_ID"
            string parameterName = "Conduit_ID";

            // Get the current active document
            

        
            application.ControlledApplication.DocumentOpened += OnDocumentOpened;

            return Result.Succeeded;
        }

        private void ReadStringForConduitData(Document doc)
        {
           Utils.conduitDataTable = Utils.ReadDataTableFromExtensibleStorage(doc);
        }

        private ElementId FindSharedParameterId(Document doc, string parameterName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(SharedParameterElement));

            foreach (SharedParameterElement sharedParameterElement in collector)
            {
                if (sharedParameterElement.Name == parameterName)
                {
                    return sharedParameterElement.Id;
                }
            }

            return ElementId.InvalidElementId;
        }
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            // Set the active document in the ConduitIdUpdater whenever a new document is opened
            

            ElementId parameterId = FindSharedParameterId(e.Document, Utils.conduitIDParameter);

            // If the parameter is found, register the ConduitIdUpdater with the monitored parameter
            if (parameterId != ElementId.InvalidElementId)
            {
                // Register the ConduitIdUpdater
                ConduitIdUpdater updater= new ConduitIdUpdater();
                UpdaterRegistry.RegisterUpdater(updater);
                ElementMulticategoryFilter filter = new ElementMulticategoryFilter(new List<BuiltInCategory>() { BuiltInCategory.OST_Conduit, BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_CableTray, BuiltInCategory.OST_CableTrayFitting});
                UpdaterRegistry.AddTrigger(updater.GetUpdaterId(), filter, Element.GetChangeTypeParameter(parameterId));
                ReadStringForConduitData(e.Document);
            }
        }

        public static RibbonPanel AddRibbonPanel(UIControlledApplication application, string tabName, string panelName, bool addSeperator)
        {
            List<RibbonPanel> panels = application.GetRibbonPanels(tabName);
            RibbonPanel panel = panels.Where(x => x.Name == tabName).FirstOrDefault();
            if (panel == null)
            {
                panel = application.CreateRibbonPanel(tabName, panelName);
            }
            else if (addSeperator)
            {
                panel.AddSeparator();
            }
            return panel;
        }

        private PushButton CreatePushButton(RibbonPanel targetPanel, string targetName, Image targetImage, string targetCommand, string targetToolTip, string targetDescrip = "")
        {
            PushButton currentBtn = targetPanel.AddItem(new PushButtonData(targetCommand, targetName, currentDll, nameSpace + targetCommand)) as PushButton;
            try
            {
                BitmapSource currentImage32 = Utils.GetBitMapSourceFromImage(targetImage, 32, 32);
                BitmapSource currentImage16 = Utils.GetBitMapSourceFromImage(targetImage, 16, 16);
                currentBtn.LargeImage = currentImage32;
                currentBtn.Image = currentImage16;
            }
            catch { }
            currentBtn.ToolTip = targetToolTip;
            currentBtn.LongDescription = targetDescrip;

           

            return currentBtn;
        }
       
    }
}
