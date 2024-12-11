using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ElectricalSolutions.TestCommand
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class TestCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ElementMulticategoryFilter multiCategory = new ElementMulticategoryFilter(new List<BuiltInCategory>() { BuiltInCategory.OST_Conduit, BuiltInCategory.OST_ConduitFitting, BuiltInCategory.OST_CableTray, BuiltInCategory.OST_CableTrayFitting });
            var elementsCollection = collector.WherePasses(multiCategory).ToElements();
            using (Transaction changeParameter = new Transaction(doc, "correcting parameter"))
            {
                changeParameter.Start();

                foreach (Element element in elementsCollection)
                {
                    Parameter wrongParameter = element.Parameters.OfType<Parameter>().FirstOrDefault(x => x.Id.IntegerValue > 0 && x.Definition.Name == "Service Type");

                    if (wrongParameter == null)
                    {
                        continue;
                    }

                    Parameter serviceTypeNative = element.GetParameter(ParameterTypeId.RbsCtcServiceType);
                    if (serviceTypeNative == null)
                    {
                        continue;
                    }

                    serviceTypeNative.Set(wrongParameter.AsString());
                }

                changeParameter.Commit();
            }

            return Result.Succeeded;
        }
    }
}
