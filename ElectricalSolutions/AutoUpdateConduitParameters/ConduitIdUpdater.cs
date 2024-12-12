using Autodesk.Revit.DB;
using System;

namespace ElectricalSolutions
{
    public class ConduitIdUpdater : IUpdater
    {
        internal static bool updatingConnectedElements = false;

        public static Document ActiveDocument { get; internal set; }

        public void Execute(UpdaterData data)
        {
            if (!data.GetDocument().IsFamilyDocument && data.GetModifiedElementIds().Count > 0)
            {
                Document document = data.GetDocument();
                // Retrieve the modified elements
                foreach (ElementId elementId in data.GetModifiedElementIds())
                {
                    Element element = document.GetElement(elementId);

                    // Check if the parameter "Conduit_ID" is modified
                    Parameter conduitIdParameter = element.LookupParameter(Utils.conduitIDParameter);
                    if (conduitIdParameter != null && !conduitIdParameter.IsReadOnly && !updatingConnectedElements)
                    {
                        Utils.UpdateConnectedElementsParameters(document, element);
                    }
                }
            }
        }

        public string GetAdditionalInformation() => "Updater to update connected elements when Conduit_ID is modified";

        public ChangePriority GetChangePriority() => ChangePriority.FloorsRoofsStructuralWalls;

        public UpdaterId GetUpdaterId() => new UpdaterId(ElectricalApplication.AddinID, new Guid("B8D7215D-8F8C-44F7-8B9B-9CBB965EF9F2"));

        public string GetUpdaterName() => "ConduitIdUpdater";
    }

}





