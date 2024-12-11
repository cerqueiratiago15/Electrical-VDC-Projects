using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;
using Autodesk.Revit.DB;
using Visibility = System.Windows.Visibility;

namespace ElectricalSolutions
{
    public class TrueToNotVisibleConverter : MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            Visibility visibility = System.Windows.Visibility.Visible;
            if (value is bool)
            {
                if ((bool)value)
                {
                    visibility = Visibility.Hidden;
                }
            }
            return visibility;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool isHidden = true;
            if (value is Visibility && (Visibility)value == Visibility.Visible)
            {

                isHidden = false;

            }
            return isHidden;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
