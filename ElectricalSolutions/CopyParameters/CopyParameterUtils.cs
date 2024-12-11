using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace ElectricalSolutions.CopyParameters
{
    public static class CopyParameterUtils
    {
        public static Parameter GetParameterFromString(this Element element, string parameterName)
        {
            Parameter parameter = element.Parameters.OfType<Parameter>().FirstOrDefault(x => x.Definition.Name == parameterName);

            if (parameter == null)
            {
                return null;
            }

            return parameter;
        }

        public static void CopyParameters(this Element source, List<Element> targetElements, List<string> parameterList)
        {
            using (Transaction copyParameters = new Transaction(source.Document, "Copy Parameters"))
            {
                copyParameters.Start();
                foreach (Element targetElement in targetElements)
                {
                    foreach (string parameterName in parameterList)
                    {
                        Parameter sourceParameter = source.GetParameterFromString(parameterName);

                        if (sourceParameter == null)
                        {
                            continue;
                        }

                        Parameter targetParameter = targetElement.GetParameterFromString(parameterName);

                        if (targetParameter == null)
                        {
                            continue;
                        }

                        string sourceValue = sourceParameter.GetParameterValue();

                        targetParameter.SetParameterValue(sourceValue);

                    }
                }
                copyParameters.Commit();
            }
        }

        public static string GetParameterValue(this Parameter parameter)
        {
            if (parameter == null)
            {
                return string.Empty;

            }
            if (parameter.StorageType == StorageType.String)
            {
                return parameter.AsString();
            }


            return parameter.AsValueString();

        }

        public static void SetParameterValue(this Parameter parameter, string value)
        {
            if (parameter == null)
            {
                return;

            }
            if (parameter.StorageType == StorageType.String)
            {
                parameter.Set(value);
                return ;
            }

            parameter.SetValueString(value);
           
        }
    }
}
