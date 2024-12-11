using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ElectricalSolutions.DuctbankCreator
{
    
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class DuctbankCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application?.ActiveUIDocument;
            if (uiDoc == null)
            {
                return Result.Cancelled;
            }
            Document doc = uiDoc.Document;

            if (doc.IsFamilyDocument)
            {
                return Result.Succeeded;
            }
    


            return Result.Succeeded;
        }
    }
}
