using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace INV.Elearning.ImportPowerPoint.Converter
{
    public class ImageConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            BitmapImage image = new BitmapImage();
            image.BeginInit();
            //Hiển thị file ảnh mới được replace
            image.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
            //Replace file ảnh
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.UriSource = new Uri(value as string);
            image.EndInit();
            return image;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return Binding.DoNothing;
        }

        #endregion
    }
}
