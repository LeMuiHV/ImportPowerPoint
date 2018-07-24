using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.DesignControl.Models;
using INV.Elearning.ImageProcess.Models;
using INV.Elearning.Text.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace INV.Elearning.DesignControl.Views
{
    public class CanvasLayoutThumbnail : Canvas
    {
        #region Data
        /// <summary>
        /// Data MainLayer
        /// </summary>
        public ELayoutMaster Data
        {
            get { return (ELayoutMaster)GetValue(DataProperty); }
            set { SetValue(DataProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Data.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DataProperty =
            DependencyProperty.Register("Data", typeof(ELayoutMaster), typeof(CanvasLayoutThumbnail), new PropertyMetadata(null, DataCallBack));

        private static void DataCallBack(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasLayoutThumbnail _owner = d as CanvasLayoutThumbnail;
            _owner.Children.Clear();
            foreach (ObjectElementBase item in _owner.Data.MainLayer.Children)
            {
                if (item.ToString() == (typeof(EStandardElement)).ToString() || item is EImageContent)
                {
                    var _standardElement = ObjectElementsHelper.LoadData(item);
                    if (_standardElement != null) _owner.Children.Add(_standardElement);
                }
                if (item.Visibility == Visibility.Visible)
                {
                    if (item is EPlaceHolder)
                    {
                        Rectangle rectangle = new Rectangle();
                        rectangle.Width = item.Width;
                        rectangle.Height = item.Height;
                        Canvas.SetTop(rectangle, item.Top);
                        Canvas.SetLeft(rectangle, item.Left);
                        rectangle.Stroke = Brushes.Black;
                        rectangle.StrokeThickness = 7;
                        rectangle.StrokeDashArray = new DoubleCollection { 3, 3 };
                        _owner.Children.Add(rectangle);
                    }
                    if (item is DataDocument dataDocument)
                    {
                        if (dataDocument.TypeText == Text.TypeText.Title || dataDocument.TypeText == Text.TypeText.Text)
                        {
                            Rectangle rectangle = new Rectangle();
                            rectangle.Width = dataDocument.Width;
                            rectangle.Height = dataDocument.Height;
                            Canvas.SetTop(rectangle, dataDocument.Top);
                            Canvas.SetLeft(rectangle, dataDocument.Left);
                            rectangle.Stroke = Brushes.Black;
                            rectangle.StrokeThickness = 7;
                            rectangle.StrokeDashArray = new DoubleCollection { 3, 3 };
                            _owner.Children.Add(rectangle);
                        }
                    }
                }
            }
            if (!_owner.Data.IsHideBackground && _owner.Data.Parent != null)
            {
                Canvas canvas = new Canvas();
                foreach (var item in _owner.Data.Parent.MainLayer.Children)
                {
                    if (item.ToString() == (typeof(EStandardElement)).ToString())
                    {
                        var _standardElement = ObjectElementsHelper.LoadData(item);
                        if (_standardElement != null) canvas.Children.Add(_standardElement);
                    }
                }
                for (int i = 0; i < _owner.Children.Count; i++)
                {
                    if (_owner.Children[i] is Canvas)
                    {
                        _owner.Children.RemoveAt(i);
                        break;
                    }
                }
                Panel.SetZIndex(canvas, -1);
                _owner.Children.Add(canvas);
            }

        }
        #endregion
        #region Background Thumbnail
        /// <summary>
        /// Background Thumbnail
        /// </summary>
        public ESlideMaster DataBackground
        {
            get { return (ESlideMaster)GetValue(DataBackgroundProperty); }
            set { SetValue(DataBackgroundProperty, value); }
        }

        // Using a DependencyProperty as the backing store for DataBackground.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DataBackgroundProperty =
            DependencyProperty.Register("DataBackground", typeof(ESlideMaster), typeof(CanvasLayoutThumbnail), new PropertyMetadata(null, DatabackgroundCallBack));

        private static void DatabackgroundCallBack(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasLayoutThumbnail _owner = d as CanvasLayoutThumbnail;
            Canvas canvas = new Canvas();
            foreach (var item in _owner.DataBackground.MainLayer.Children)
            {
                if (item.ToString() == (typeof(EStandardElement)).ToString())
                {
                    var _standardElement = ObjectElementsHelper.LoadData(item);
                    if (_standardElement != null) canvas.Children.Add(_standardElement);
                }
            }
            for (int i = 0; i < _owner.Children.Count; i++)
            {
                if (_owner.Children[i] is Canvas)
                {
                    _owner.Children.RemoveAt(i);
                    break;
                }
            }
            Panel.SetZIndex(canvas, -1);
            _owner.Children.Add(canvas);
        }
        #endregion
        #region IsHideBackground
        /// <summary>
        /// Check ẩn hiện ThemeLayout
        /// </summary>
        public bool IsHideBackground
        {
            get { return (bool)GetValue(IsHideBackgroundProperty); }
            set { SetValue(IsHideBackgroundProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsHideBackground.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsHideBackgroundProperty =
            DependencyProperty.Register("IsHideBackground", typeof(bool), typeof(CanvasLayoutThumbnail), new PropertyMetadata(true, IsHideBackgroundCallBack));

        private static void IsHideBackgroundCallBack(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasLayoutThumbnail _owner = d as CanvasLayoutThumbnail;
            for (int i = 0; i < _owner.Children.Count; i++)
            {
                if (_owner.Children[i] is Canvas)
                {
                    if ((bool)e.NewValue)
                        _owner.Children[i].Visibility = Visibility.Collapsed;
                    else
                        _owner.Children[i].Visibility = Visibility.Visible;
                    break;
                }
            }
        }
        #endregion
        #region SelectedThemeColor

        /// <summary>
        /// Màu được chọn
        /// </summary>
        public EColorManagment SelectedColor
        {
            get { return (EColorManagment)GetValue(SelectedColorProperty); }
            set { SetValue(SelectedColorProperty, value); }
        }

        // Using a DependencyProperty as the backing store for SelectedColor.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty SelectedColorProperty =
            DependencyProperty.Register("SelectedColor", typeof(EColorManagment), typeof(CanvasLayoutThumbnail), new PropertyMetadata(null, SelectedColorCallback));

        private static void SelectedColorCallback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasLayoutThumbnail _owner = d as CanvasLayoutThumbnail;
            foreach (var item in _owner.Children)
            {
                if (item is StandardElement standardElement)
                {
                    standardElement.UpdateThemeColor();
                }
                if (item is Canvas canvas)
                {
                    foreach (StandardElement child in canvas.Children)
                    {
                        child.UpdateThemeColor();
                    }
                }
            }
        }


        #endregion
        #region BackgroundStyle


        public BackgroundItem BackgroundStyle
        {
            get { return (BackgroundItem)GetValue(BackgroundStyleProperty); }
            set { SetValue(BackgroundStyleProperty, value); }
        }

        // Using a DependencyProperty as the backing store for BackgroundStyle.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty BackgroundStyleProperty =
            DependencyProperty.Register("BackgroundStyle", typeof(BackgroundItem), typeof(CanvasLayoutThumbnail), new PropertyMetadata(null, BackgroundStyleCallback));

        private static void BackgroundStyleCallback(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            CanvasLayoutThumbnail _owner = d as CanvasLayoutThumbnail;
            BackgroundItem backgroundItem = e.NewValue as BackgroundItem;
            if (backgroundItem != null)
            {
                _owner.Background = backgroundItem.ColorStyle;
            }
        }


        #endregion
    }
}
