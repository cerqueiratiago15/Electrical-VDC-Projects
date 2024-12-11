using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

namespace ElectricalSolutions
{
    [Transaction(TransactionMode.Manual)]
    public class VoltageDropCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;


                Document doc = uiDoc.Document;

                if (doc.IsFamilyDocument)
                {
                    return Result.Failed;
                }

                ICollection<ElementId> ids =  uiDoc.Selection.GetElementIds();


                if (ids.Count>0)
                {
                    foreach (ElementId id in ids)
                    {
                        Element element = doc.GetElement(id);

                        Utils.CalculateVoltageDropOneConduit(element);
                    }
                }

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





