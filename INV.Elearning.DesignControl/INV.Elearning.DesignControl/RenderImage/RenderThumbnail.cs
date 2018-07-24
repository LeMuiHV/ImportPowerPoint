using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace INV.Elearning.DesignControl.RenderImage
{
    public class RenderThumbnail : LayoutBase
    {

        public RenderThumbnail()
        {

        }

        public ELayoutMaster RenderData
        {
            get { return (ELayoutMaster)GetValue(RenderDataProperty); }
            set { SetValue(RenderDataProperty, value); }
        }

        // Using a DependencyProperty as the backing store for RenderData.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty RenderDataProperty =
            DependencyProperty.Register("RenderData", typeof(ELayoutMaster), typeof(RenderThumbnail), new PropertyMetadata(null, new PropertyChangedCallback(RenderDataPropertyChanged)));

        private static void RenderDataPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            double scaleX = (85 / 1024);
            double scaleY = (48 / 576);
            ScaleTransform scale = new ScaleTransform();
            scale.ScaleX = scaleX;
            scale.ScaleY = scaleY;
            ELayoutMaster layoutMaster = e.NewValue as ELayoutMaster;
            RenderThumbnail renderThumbnail = d as RenderThumbnail;
            renderThumbnail.Children.Clear();

            foreach (var item in layoutMaster.MainLayer.Children)
            {
                if (item is EStandardElement eStandardElement)
                {
                    StandardElement objectElement = new StandardElement();
                    objectElement.UpdateUI(eStandardElement);
                    objectElement.Width = objectElement.Width * 85 / 1024;
                    objectElement.Height = objectElement.Height * 48 / 576;
                    objectElement.Left = objectElement.Left * 85 / 1024;
                    objectElement.Top = objectElement.Top * 48 / 576;

                    //objectElement.RenderTransform = scale;
                    //objectElement.Left *= scaleX;
                    //objectElement.Top *= scaleY;
                    renderThumbnail.Children.Add(objectElement);
                }
            }
        }
    }
}
