using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View.Theme;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace INV.Elearning.DesignControl.Converters
{
    public class SelectedThemesConverter : IValueConverter
    {      
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            //if (value != null)
            //{
            //    return value as EThemes;
            //}
            return null;
        }

      
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return Binding.DoNothing;
        }
    }
}
