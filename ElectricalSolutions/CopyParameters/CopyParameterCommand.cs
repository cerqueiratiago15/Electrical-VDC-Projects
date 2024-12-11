using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using ElectricalSolutions.CopyParameters;


namespace ElectricalSolutions
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CopyParameterCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application?.ActiveUIDocument;
            if (uiDoc==null)
            {
                return Result.Cancelled;
            }
            Document doc = uiDoc.Document;

            if (doc.IsFamilyDocument)
            {
                return Result.Succeeded;
            }

            List<string> parameterValues = new List<string>();
            string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string filePath = Path.Combine(assemblyFolder, "CopyParameter.txt");

            if (!File.Exists(filePath))
            {
                TaskDialog.Show("Warning", "The file with parameters does not exist");
                return Result.Succeeded;
            }

            List<string> parameters = new List<string>();

            parameters.AddRange(File.ReadAllLines(filePath));

            if (parameters.Any()==false)
            {
                TaskDialog.Show("Warning", "The file does not contain parameters written");
                return Result.Succeeded;
            }

            ICollection<ElementId> seletedIds = uiDoc.Selection.GetElementIds();

            if (seletedIds.Any()==false)
            {
                TaskDialog.Show("Warning", "You must select elements before run the add-in!");

                return Result.Succeeded;
            }

            Reference sourceReference = null;
            try
            {
                sourceReference = uiDoc.Selection.PickObject(ObjectType.Element);
            }
            catch
            {
                TaskDialog.Show("Warning", "You must select the source to copy the parameters value!");

                return Result.Succeeded;

            }

            List<Element> selectedElements = seletedIds.Select(x => doc.GetElement(x)).ToList();


            Element source =  doc.GetElement(sourceReference);

            source.CopyParameters(selectedElements, parameters);


            return Result.Succeeded;
        }
    }
}
