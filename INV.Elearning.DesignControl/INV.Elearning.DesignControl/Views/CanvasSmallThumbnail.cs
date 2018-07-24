using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.ImageProcess.Models;
using System;
using System.Collections.Generic;
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
    public class CanvasSmallThumbnail : Canvas
    {
        #region Data
        public ESlideMaster Data
        {
            get { return (ESlideMaster)GetValue(DataProperty); }
            set { SetValue(DataProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Data.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DataProperty =
            DependencyProperty.Register("Data", typeof(ESlideMaster), typeof(CanvasSmallThumbnail), new PropertyMetadata(null, DataCallback));

        private static void DataCallback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasSmallThumbnail _owner = d as CanvasSmallThumbnail;
            _owner.Background = ColorHelper.ConverFromColorData(_owner.Data.MainLayer.Background).Brush;
            _owner.Children.Clear();

            //Thêm các rectangle hiển thị màu
            Rectangle rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 130);
            Canvas.SetTop(rectangle, 450);
            Binding binding = new Binding("Colors.Accent1.Color");
            binding.Source = (Application.Current as IAppGlobal).SelectedTheme;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 260);
            Canvas.SetTop(rectangle, 450);
            binding = new Binding("Colors.Accent2.Color");
            binding.Source = (Application.Current as IAppGlobal).SelectedTheme;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 390);
            Canvas.SetTop(rectangle, 450);
            binding = new Binding("Colors.Accent3.Color");
            binding.Source = (Application.Current as IAppGlobal).SelectedTheme;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 520);
            Canvas.SetTop(rectangle, 450);
            binding = new Binding("Colors.Accent4.Color");
            binding.Source = (Application.Current as IAppGlobal).SelectedTheme;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 650);
            Canvas.SetTop(rectangle, 450);
            binding = new Binding("Colors.Accent5.Color");
            binding.Source = (Application.Current as IAppGlobal).SelectedTheme;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            rectangle = new Rectangle();
            rectangle.Width = 120;
            rectangle.Height = 50;
            Panel.SetZIndex(rectangle, 100);
            Canvas.SetLeft(rectangle, 780);
            Canvas.SetTop(rectangle, 450);
            binding = new Binding("Colors.Accent6.Color");
            binding.Source = (Application.Current as IAppGlobal).SelectedTheme;
            rectangle.SetBinding(Rectangle.FillProperty, binding);
            _owner.Children.Add(rectangle);

            //Thêm các text hiển thị font     
            TextBlock textBlock = new TextBlock();
            textBlock.Text = "A";
            textBlock.FontSize = 450;
            binding = new Binding("SelectedFont.MajorFont");
            binding.Source = (Application.Current as IAppGlobal).SelectedTheme;
            textBlock.SetBinding(TextBlock.FontFamilyProperty, binding);
            Canvas.SetLeft(textBlock, 180);
            Canvas.SetTop(textBlock, 0);
            Panel.SetZIndex(textBlock, 100);
            _owner.Children.Add(textBlock);

            textBlock = new TextBlock();
            textBlock.Text = "a";
            textBlock.FontSize = 400;
            binding = new Binding("SelectedFont.MinorFont");
            binding.Source = (Application.Current as IAppGlobal).SelectedTheme;
            textBlock.SetBinding(TextBlock.FontFamilyProperty, binding);
            Canvas.SetLeft(textBlock, 480);
            Canvas.SetTop(textBlock, 50);
            Panel.SetZIndex(textBlock, 100);
            _owner.Children.Add(textBlock);

            foreach (ObjectElementBase item in _owner.Data.MainLayer.Children)
            {
                if (item.ToString() == (typeof(EStandardElement)).ToString() || item is EImageContent)
                {
                    var _standardElement = ObjectElementsHelper.LoadData(item);
                    if (_standardElement != null) _owner.Children.Add(_standardElement);
                }
            }
        }
        #endregion

        #region SelectedColor

        /// <summary>
        /// Chọn màu color
        /// </summary>
        public EColorManagment SelectedColor
        {
            get { return (EColorManagment)GetValue(SelectedColorProperty); }
            set { SetValue(SelectedColorProperty, value); }
        }

        // Using a DependencyProperty as the backing store for SelectedColor.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SelectedColorProperty =
            DependencyProperty.Register("SelectedColor", typeof(EColorManagment), typeof(CanvasSmallThumbnail), new PropertyMetadata(null, SelectedColorCallback));

        private static void SelectedColorCallback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasSmallThumbnail _owner = d as CanvasSmallThumbnail;
            foreach (var item in _owner.Children)
            {
                if (item is StandardElement standardElement)
                    standardElement.UpdateThemeColor();
            }
        }


        #endregion

        #region SelectedFont

        /// <summary>
        /// Chọn màu font
        /// </summary>
        public EFontfamily SelectedFont
        {
            get { return (EFontfamily)GetValue(SelectedFontProperty); }
            set { SetValue(SelectedFontProperty, value); }
        }

        // Using a DependencyProperty as the backing store for SelectedFont.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SelectedFontProperty =
            DependencyProperty.Register("SelectedFont", typeof(EFontfamily), typeof(CanvasSmallThumbnail), new PropertyMetadata(null, SelectedFontCallback));

        private static void SelectedFontCallback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasSmallThumbnail _owner = d as CanvasSmallThumbnail;
            foreach (var item in _owner.Children)
            {
                if (item is StandardElement standardElement)
                    standardElement.UpdateThemeFont();
            }
        }


        #endregion

        #region BackgroundStyle

        /// <summary>
        /// Cài đặt background của Thumbnail
        /// </summary>
        public BackgroundItem BackgroundStyle
        {
            get { return (BackgroundItem)GetValue(BackgroundStyleProperty); }
            set { SetValue(BackgroundStyleProperty, value); }
        }

        // Using a DependencyProperty as the backing store for BackgroundStyle.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty BackgroundStyleProperty =
            DependencyProperty.Register("BackgroundStyle", typeof(BackgroundItem), typeof(CanvasSmallThumbnail), new PropertyMetadata(null, BackgroundStyleCallback));

        private static void BackgroundStyleCallback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasSmallThumbnail _owner = d as CanvasSmallThumbnail;
            BackgroundItem backgroundItem = e.NewValue as BackgroundItem;
            if (backgroundItem != null)
            {
                _owner.Background = backgroundItem.ColorStyle;
            }
        }


        #endregion
    }
}
