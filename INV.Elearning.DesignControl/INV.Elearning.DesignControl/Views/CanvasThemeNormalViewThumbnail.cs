using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.DesignControl.Converters;
using INV.Elearning.ImageProcess.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Shapes;

namespace INV.Elearning.DesignControl.Views
{
    public class CanvasThemeNormalViewThumbnail : Canvas
    {

        #region Data

        /// <summary>
        /// Data của canvas Thumbnail
        /// </summary>
        public ETheme Data
        {
            get { return (ETheme)GetValue(DataProperty); }
            set { SetValue(DataProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Data.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DataProperty =
            DependencyProperty.Register("Data", typeof(ETheme), typeof(CanvasThemeNormalViewThumbnail), new PropertyMetadata(null, DataCallBack));

        private static void DataCallBack(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            Global.BeginInit();
            CanvasThemeNormalViewThumbnail _owner = d as CanvasThemeNormalViewThumbnail;
            _owner.Background = Brushes.Transparent;
            EColorManagment eColorManagment = new EColorManagment();
            //  (Application.Current as IAppGlobal).SelectedTheme.Colors = _owner.Data.Colors;
            _owner.Children.Clear();
            //Thêm các hình nền
            foreach (ObjectElementBase item in _owner.Data.SlideMasters[0].MainLayer.Children)
            {
                if (item.ToString() == (typeof(EStandardElement)).ToString() || item is EImageContent)
                {
                    var _standardElement = ObjectElementsHelper.LoadData(item);
                    if (_standardElement != null)
                    {
                        if ((_standardElement as StandardElement).Fill.ColorSpecialName == System.Windows.Application.Current.TryFindResource("ColorGallery_Accent1").ToString())
                        {
                            Binding biding = new Binding("Colors.Accent1.Color");
                            biding.Source = _owner.Data;
                            biding.Converter = new StringToColorBaseConverter();
                            (_standardElement as StandardElement).SetBinding(StandardElement.FillProperty, biding);
                        }
                        else if ((_standardElement as StandardElement).Fill.ColorSpecialName == System.Windows.Application.Current.TryFindResource("ColorGallery_Accent2").ToString())
                        {
                            Binding biding = new Binding("Colors.Accent2.Color");
                            biding.Source = _owner.Data;
                            biding.Converter = new StringToColorBaseConverter();
                            (_standardElement as StandardElement).SetBinding(StandardElement.FillProperty, biding);
                        }
                        else if ((_standardElement as StandardElement).Fill.ColorSpecialName == System.Windows.Application.Current.TryFindResource("ColorGallery_Accent3").ToString())
                        {
                            Binding biding = new Binding("Colors.Accent3.Color");
                            biding.Source = _owner.Data;
                            biding.Converter = new StringToColorBaseConverter();
                            (_standardElement as StandardElement).SetBinding(StandardElement.FillProperty, biding);
                        }
                        else if ((_standardElement as StandardElement).Fill.ColorSpecialName == System.Windows.Application.Current.TryFindResource("ColorGallery_Accent4").ToString())
                        {
                            Binding biding = new Binding("Colors.Accent4.Color");
                            biding.Source = _owner.Data;
                            biding.Converter = new StringToColorBaseConverter();
                            (_standardElement as StandardElement).SetBinding(StandardElement.FillProperty, biding);
                        }
                        else if ((_standardElement as StandardElement).Fill.ColorSpecialName == System.Windows.Application.Current.TryFindResource("ColorGallery_Accent5").ToString())
                        {
                            Binding biding = new Binding("Colors.Accent5.Color");
                            biding.Source = _owner.Data;
                            biding.Converter = new StringToColorBaseConverter();
                            (_standardElement as StandardElement).SetBinding(StandardElement.FillProperty, biding);
                        }
                        else if ((_standardElement as StandardElement).Fill.ColorSpecialName == System.Windows.Application.Current.TryFindResource("ColorGallery_Accent6").ToString())
                        {
                            Binding biding = new Binding("Colors.Accent6.Color");
                            biding.Source = _owner.Data;
                            biding.Converter = new StringToColorBaseConverter();
                            (_standardElement as StandardElement).SetBinding(StandardElement.FillProperty, biding);
                        }
                        _owner.Children.Add(_standardElement);
                    }
                }
            }

            //Thêm các rectangle hiển thị màu
            Rectangle rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 130);
            Canvas.SetTop(rectangle, 420);
            Binding binding = new Binding("Colors.Accent1.Color");
            binding.Source = _owner.Data;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 260);
            Canvas.SetTop(rectangle, 420);
            binding = new Binding("Colors.Accent2.Color");
            binding.Source = _owner.Data;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 390);
            Canvas.SetTop(rectangle, 420);
            binding = new Binding("Colors.Accent3.Color");
            binding.Source = _owner.Data;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 520);
            Canvas.SetTop(rectangle, 420);
            binding = new Binding("Colors.Accent4.Color");
            binding.Source = _owner.Data;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 650);
            Canvas.SetTop(rectangle, 420);
            binding = new Binding("Colors.Accent5.Color");
            binding.Source = _owner.Data;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 780);
            Canvas.SetTop(rectangle, 420);
            binding = new Binding("Colors.Accent6.Color");
            binding.Source = _owner.Data;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            //Thêm các text hiển thị font     
            TextBlock textBlock = new TextBlock();
            textBlock.Text = "A";
            textBlock.FontSize = 150;
            binding = new Binding("SelectedFont.MajorFont");
            binding.Source = _owner.Data;
            textBlock.SetBinding(TextBlock.FontFamilyProperty, binding);
            Canvas.SetLeft(textBlock, 180);
            Canvas.SetTop(textBlock, 250);
            Panel.SetZIndex(textBlock, 100);
            _owner.Children.Add(textBlock);

            textBlock = new TextBlock();
            textBlock.Text = "a";
            textBlock.FontSize = 150;
            binding = new Binding("SelectedFont.MinorFont");
            binding.Source = _owner.Data;
            textBlock.SetBinding(TextBlock.FontFamilyProperty, binding);
            Canvas.SetLeft(textBlock, 290);
            Canvas.SetTop(textBlock, 250);
            Panel.SetZIndex(textBlock, 100);
            _owner.Children.Add(textBlock);
            Global.EndInit();
        }

        #endregion

        #region SelectedColor

        /// <summary>
        /// Màu đang được chọn
        /// </summary>
        public EColorManagment SelectedColor
        {
            get { return (EColorManagment)GetValue(SelectedColorProperty); }
            set { SetValue(SelectedColorProperty, value); }
        }

        // Using a DependencyProperty as the backing store for SelectedColor.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SelectedColorProperty =
            DependencyProperty.Register("SelectedColor", typeof(EColorManagment), typeof(CanvasThemeNormalViewThumbnail), new PropertyMetadata(null, SelectedColorCallback));

        private static void SelectedColorCallback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasThemeNormalViewThumbnail _owner = d as CanvasThemeNormalViewThumbnail;
            foreach (var item in _owner.Children)
            {
                if (item is StandardElement standardElement)
                    standardElement.UpdateThemeColor();
            }
        }


        #endregion

        #region BackgroundStyle

        /// <summary>
        /// Cài đặt background cho thumbnail
        /// </summary>
        public BackgroundItem BackgroundStyle
        {
            get { return (BackgroundItem)GetValue(BackgroundStyleProperty); }
            set { SetValue(BackgroundStyleProperty, value); }
        }

        // Using a DependencyProperty as the backing store for BackgroundStyle.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty BackgroundStyleProperty =
            DependencyProperty.Register("BackgroundStyle", typeof(BackgroundItem), typeof(CanvasThemeNormalViewThumbnail), new PropertyMetadata(null, BackgroundStyleCallback));

        private static void BackgroundStyleCallback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasThemeNormalViewThumbnail _owner = d as CanvasThemeNormalViewThumbnail;
            BackgroundItem backgroundItem = e.NewValue as BackgroundItem;
            if (backgroundItem != null)
            {
                _owner.Background = backgroundItem.ColorStyle;
            }
        }


        #endregion

    }
}
