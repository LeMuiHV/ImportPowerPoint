using INV.Elearning.Controls;
using INV.Elearning.Core;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.DesignControl.Helper;
using INV.Elearning.DesignControl.UndoRedo;
using INV.Elearning.EOPC;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace INV.Elearning.DesignControl.ViewModel
{
    public class SlideMasterTabControlViewModel : ViewModelBase
    {
        public static EColorManagment BackupColor { get; set; }
        private static bool IsChangeColor = false;
        static ThemesData themesData = new ThemesData();

        public static Gallery ThemeDesign { get; set; }

        ThemesData ThemesData = new ThemesData();

        private EFontfamilyManagment _fonts;

        public EFontfamilyManagment Fonts
        {
            get { return _fonts ?? (_fonts = new EFontfamilyManagment()); }
            set { _fonts = value; OnPropertyChanged("Fonts"); }
        }

        private ObservableCollection<PlaceHolderItem> _placeHolderItem;

        public ObservableCollection<PlaceHolderItem> PlaceHolderItem
        {
            get { return _placeHolderItem ?? (_placeHolderItem = new ObservableCollection<PlaceHolderItem>()); }
            set { _placeHolderItem = value; OnPropertyChanged("PlaceHolderItems"); }
        }

        private ObservableCollection<BackgroundItem> _backgrounds;

        public ObservableCollection<BackgroundItem> Backgrounds
        {
            get { return _backgrounds ?? (_backgrounds = new ObservableCollection<BackgroundItem>()); }
            set { _backgrounds = value; OnPropertyChanged("Backgrounds"); }
        }




        public SlideMasterTabControlViewModel()
        {
            Fonts.ThemeFonts.Add(new EFontfamily() { MajorFont = "Times New Roman", MinorFont = "Arial", Name = "Office" });
            Fonts.ThemeFonts.Add(new EFontfamily() { MajorFont = "Times New Roman", MinorFont = "Arial", Name = "Office" });
            Fonts.ThemeFonts.Add(new EFontfamily() { MajorFont = "Times New Roman", MinorFont = "Arial", Name = "Office" });

            PlaceHolderItem.Add(new ViewModel.PlaceHolderItem() { Name = "Content", UrlIcon = "/INV.Elearning.DesignControl;component/Images/ContentPlaceHolder32.png" });
            PlaceHolderItem.Add(new ViewModel.PlaceHolderItem() { Name = "Text", UrlIcon = "/INV.Elearning.DesignControl;component/Images/TextPlaceHolder32.png" });
            PlaceHolderItem.Add(new ViewModel.PlaceHolderItem() { Name = "Picture", UrlIcon = "/INV.Elearning.DesignControl;component/Images/PicturePlaceHolder32.png" });
            PlaceHolderItem.Add(new ViewModel.PlaceHolderItem() { Name = "Video", UrlIcon = "/INV.Elearning.DesignControl;component/Images/VideoPlaceHolder32.png" });
            PlaceHolderItem.Add(new ViewModel.PlaceHolderItem() { Name = "Chart", UrlIcon = "/INV.Elearning.DesignControl;component/Images/ChartPlaceHolder32.png" });
            GenerateBackgroundStype();
            //(Application.Current as IAppGlobal).ThemesCollection.Clear();
            //foreach (EThemes item in GetThemesCollection())
            //{
            //    (Application.Current as IAppGlobal).ThemesCollection.Add(item);
            //}
        }



        //#region Commands

        public static List<ETheme> GetThemesCollection()
        {
            List<ETheme> _result = new List<ETheme>();
            foreach (string fileName in Directory.GetFiles("Themes/"))
            {

                //Task.Factory.StartNew(() =>
                //{
                //    EThemes eThemes = new EThemes();
                //    eThemes = themesData.Open(fileName);
                //    eThemes.Name = Path.GetFileNameWithoutExtension(fileName);
                //    eThemes.TagName = "Office";
                //    eThemes.ThumbnailIcon = new BitmapImage(new Uri(eThemes.MainTheme.Thumbnail.Path));
                //    foreach (ETheme item in eThemes.Variants)
                //    {
                //        item.ThumbnailBitmap = new BitmapImage(new Uri(item.Thumbnail.Path));
                //    }
                //    _result.Add(eThemes);
                //});                
                ETheme eTheme = new ETheme();
                eTheme = themesData.FirstLoad(fileName);
                eTheme.Name = Path.GetFileNameWithoutExtension(fileName);
                eTheme.ID = ObjectElementsHelper.RandomString(10);
                eTheme.TagName = "Office";
                eTheme.ThemeName = Path.GetFileNameWithoutExtension(fileName);
                eTheme.FilePath = fileName;
                _result.Add(eTheme);
            }
            return _result;
        }


        public static List<ETheme> GetFullThemesCollection()
        {
            List<ETheme> _result = new List<ETheme>();
            foreach (string fileName in Directory.GetFiles("Themes/"))
            {
                if (Path.GetExtension(fileName) == ".athmx")
                {
                    ETheme eTheme = new ETheme();
                    eTheme = themesData.OpenETheme(fileName);
                    eTheme.Name = Path.GetFileNameWithoutExtension(fileName);
                    eTheme.ID = ObjectElementsHelper.RandomString(10);
                    eTheme.TagName = "Office";
                    eTheme.ThemeName = Path.GetFileNameWithoutExtension(fileName);
                    eTheme.FilePath = fileName;
                    eTheme.IsLoaded = false;               
                    _result.Add(eTheme);
                }
            }
            return _result;
        }


        public static void LoadBlankTheme()
        {
            foreach (ETheme item in (Application.Current as IAppGlobal).ThemesCollection)
            {
                if (item.ThemeName == "BlankTheme")
                {
                    Global.BeginInit();
                    ETheme _cloneTheme = new ETheme();
                    _cloneTheme = item.Clone() as ETheme;
                    _cloneTheme.Name = item.Name;
                    _cloneTheme.TagName = "This Presenter";
                    _cloneTheme.ThemeName = item.ThemeName;
                    _cloneTheme.ID = ObjectElementsHelper.RandomString(10);
                    _cloneTheme.Colors = (Application.Current as IAppGlobal).OfficeColors[1];
                    _cloneTheme.SelectedFont = (Application.Current as IAppGlobal).OfficeFonts[1];
                    (Application.Current as IAppGlobal).LocalThemesCollection.Add(_cloneTheme);
                    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = _cloneTheme;
                    (Application.Current as IAppGlobal).DocumentControl.SelectedColors = _cloneTheme.Colors;
                    (Application.Current as IAppGlobal).DocumentControl.SelectedFont = _cloneTheme.SelectedFont;
                    (Application.Current as IAppGlobal).SelectedTheme.SelectedFont = _cloneTheme.SelectedFont;
                    DesignTabControlViewModel.ThemeDesign.SelectedIndex = 0;
                    Global.EndInit();
                }
            }
        }

        #region ColorsCommand
        private static RelayCommand _colorsCommand;
        /// <summary>
        /// Lệnh điều khiển Lựa chọn màu
        /// </summary>
        public static RelayCommand ColorsCommand
        {
            get { return _colorsCommand ?? (_colorsCommand = new RelayCommand(p => ColorsExecute(p))); }
        }

        /// <summary>
        /// Thực thi điều khiển
        /// </summary>
        /// <param name="p"></param>
        private static void ColorsExecute(object p)
        {
            if (p is EColorManagment color)
            {
                (Application.Current as IAppGlobal).SelectedThemeView.Colors = color;
                SlideMaster slide = null;
                if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster slidemaster)
                {
                    slide = slidemaster;
                }
                else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster layoutMaster)
                {
                    slide = layoutMaster.SlideParent;
                }
                if (slide != null) slide.SelectedColor = color;
            }
        }

        #endregion

        //#region PreviewColorCommand
        //private static RelayCommand _previewColorCommand;
        ///// <summary>
        ///// Command khi rê chuột chọn màu
        ///// </summary>
        //public static RelayCommand PreviewColorCommand
        //{
        //    get { return _previewColorCommand ?? (_previewColorCommand = new RelayCommand(p => PreviewExecute(p))); }
        //}

        //private static void PreviewExecute(object p)
        //{
        //    if (p is GalleryItem galleryItem)
        //    {
        //        Global.BeginInit();
        //        EColorManagment eColor = galleryItem.Content as EColorManagment;
        //        if ((Application.Current as IAppGlobal).SelectedTheme != null && (Application.Current as IAppGlobal).DocumentControl != null)
        //        {
        //            (Application.Current as IAppGlobal).SelectedTheme.Colors = eColor;
        //            if ((Application.Current as IAppGlobal).SelectedThemeView != null)
        //                (Application.Current as IAppGlobal).SelectedThemeView.Colors = eColor;

        //            (Application.Current as IAppGlobal).SelectedSlide.UpdateThemeColor();                
        //        }
        //        Global.EndInit();
        //    }
        //}

        //#endregion

        //#region PreviewColorCancelCommand
        //private static RelayCommand _PreviewColorCancelCommand;

        //public static RelayCommand PreviewColorCancelCommand
        //{
        //    get { return _PreviewColorCancelCommand ?? (_PreviewColorCancelCommand = new RelayCommand(p => PreviewColorCancelExecute(p))); }
        //}

        //private static void PreviewColorCancelExecute(object p)
        //{
        //    if (BackupColor != null)
        //    {
        //        Global.BeginInit();
        //        if (!IsChangeColor)
        //        {
        //            if ((Application.Current as IAppGlobal).SelectedTheme != null && (Application.Current as IAppGlobal).DocumentControl != null)
        //            {
        //                (Application.Current as IAppGlobal).SelectedTheme.Colors = BackupColor;
        //                if ((Application.Current as IAppGlobal).SelectedThemeView != null)
        //                    (Application.Current as IAppGlobal).SelectedThemeView.Colors = BackupColor;

        //                foreach (var item in (Application.Current as IAppGlobal).DocumentControl.Slides)
        //                {
        //                    item.UpdateThemeColor();
        //                }
        //            }
        //        }
        //        else
        //        {
        //            IsChangeColor = false;
        //        }
        //        Global.EndInit();
        //    }
        //}

        //#endregion

        //#region ColorDropDownCommand
        //private static RelayCommand _colorDropDownCommand;

        //public static RelayCommand ColorDropDownCommand
        //{
        //    get { return _colorDropDownCommand ?? (_colorDropDownCommand = new RelayCommand(p => ColorDropDownExecute())); }
        //}

        //private static void ColorDropDownExecute()
        //{
        //    BackupColor = (Application.Current as IAppGlobal).SelectedTheme.Colors;
        //}

        //#endregion

        //private RelayCommand _treeviewCommand;

        //public RelayCommand TreeViewCommand
        //{
        //    get { return _treeviewCommand ?? (_treeviewCommand = new RelayCommand(p => TreeViewExecute(p))); }
        //}

        //private void TreeViewExecute(object p)
        //{
        //    SlideMaster slide = null;
        //    if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster slideMaster)
        //    {
        //        slide = slideMaster;
        //    }
        //    else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster layoutMaster)
        //    {
        //        slide = layoutMaster.SlideParent;
        //    }
        //    if (slide != null)
        //    {
        //        foreach (EThemes item in (Application.Current as IAppGlobal).LocalThemesCollection)
        //        {
        //            if (slide.ThemesName == item.Name)
        //            {
        //                (Application.Current as IAppGlobal).SelectedThemes = item;
        //            }
        //        }
        //        ThemeDesign.SelectedValue = (Application.Current as IAppGlobal).SelectedThemes;
        //    }

        //}
        //#endregion

        #region GenerateBackgroundStype
        /// <summary>
        /// Tạo ra các Background Style
        /// </summary>
        private void GenerateBackgroundStype()
        {

            //Style1
            BackgroundItem bgItem = new BackgroundItem();
            ColorGradientBrush gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();


            LinearGradientBrush gradientBrush = new LinearGradientBrush();
            gradientBrush.StartPoint = new Point(0, 0);
            gradientBrush.EndPoint = new Point(1, 1);

            GradientStop gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            gradientStop.Color = Colors.Transparent;

            gradientBrush.GradientStops.Add(gradientStop);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "", Brightness = 1 }, Offset = 0.0 });

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = gradientBrush;
            bgItem.Name = "Style 1";
            Backgrounds.Add(bgItem);

            //Style2
            bgItem = new BackgroundItem();

            gradientBrush = new LinearGradientBrush();
            gradientBrush.StartPoint = new Point(0, 0);
            gradientBrush.EndPoint = new Point(1, 1);

            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            Binding _binding = new Binding("SelectedTheme.Colors.Accent4.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 4", Brightness = 1 }, Offset = 0.0 });
            gradientBrush.GradientStops.Add(gradientStop);

            _binding = new Binding("SelectedTheme.Colors.Accent4.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 4", Brightness = 1 }, Offset = 1.0 });
            gradientBrush.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = gradientBrush;
            bgItem.Name = "Style 2";
            Backgrounds.Add(bgItem);

            //Style3
            bgItem = new BackgroundItem();
            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            gradientBrush = new LinearGradientBrush();
            gradientBrush.StartPoint = new Point(0, 0);
            gradientBrush.EndPoint = new Point(1, 1);

            _binding = new Binding("SelectedTheme.Colors.Accent5.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 5", Brightness = 1 }, Offset = 0.0 });
            gradientBrush.GradientStops.Add(gradientStop);


            _binding = new Binding("SelectedTheme.Colors.Accent5.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 5", Brightness = 1 }, Offset = 1.0 });
            gradientBrush.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = gradientBrush;
            bgItem.Name = "Style 3";
            Backgrounds.Add(bgItem);

            //Style4
            bgItem = new BackgroundItem();
            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            gradientBrush = new LinearGradientBrush();
            gradientBrush.StartPoint = new Point(0, 0);
            gradientBrush.EndPoint = new Point(1, 1);

            _binding = new Binding("SelectedTheme.Colors.Accent6.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 6", Brightness = 1 }, Offset = 0.0 });
            gradientBrush.GradientStops.Add(gradientStop);

            _binding = new Binding("SelectedTheme.Colors.Accent6.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 6", Brightness = 1 }, Offset = 1.0 });
            gradientBrush.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = gradientBrush;
            bgItem.Name = "Style 4";
            Backgrounds.Add(bgItem);

            //Style5
            bgItem = new BackgroundItem();
            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            gradientBrush = new LinearGradientBrush();
            gradientBrush.StartPoint = new Point(0, 0);
            gradientBrush.EndPoint = new Point(1, 0);

            _binding = new Binding("SelectedTheme.Colors.BackgroundLight1.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientBrush.GradientStops.Add(gradientStop);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nền sáng 1", Brightness = 1 }, Offset = 0.0 });

            _binding = new Binding("SelectedTheme.Colors.BackgroundDark1.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nền tối 1", Brightness = 1 }, Offset = 1.0 });
            gradientBrush.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = gradientBrush;
            bgItem.Name = "Style 5";
            Backgrounds.Add(bgItem);


            //Style6
            bgItem = new BackgroundItem();
            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            gradientBrush = new LinearGradientBrush();
            gradientBrush.StartPoint = new Point(0, 0);
            gradientBrush.EndPoint = new Point(1, 0);

            _binding = new Binding("SelectedTheme.Colors.BackgroundLight1.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientBrush.GradientStops.Add(gradientStop);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nền sáng 1", Brightness = 1 }, Offset = 0.0 });

            _binding = new Binding("SelectedTheme.Colors.Accent4.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 4", Brightness = 1 }, Offset = 1.0 });
            gradientBrush.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = gradientBrush;
            bgItem.Name = "Style 6";
            Backgrounds.Add(bgItem);

            //Style7
            bgItem = new BackgroundItem();
            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            gradientBrush = new LinearGradientBrush();
            gradientBrush.StartPoint = new Point(0, 0);
            gradientBrush.EndPoint = new Point(1, 0);

            _binding = new Binding("SelectedTheme.Colors.BackgroundLight1.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientBrush.GradientStops.Add(gradientStop);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nền sáng 1", Brightness = 1 }, Offset = 0.0 });

            _binding = new Binding("SelectedTheme.Colors.Accent5.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 5", Brightness = 1 }, Offset = 1.0 });
            gradientBrush.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = gradientBrush;
            bgItem.Name = "Style 7";
            Backgrounds.Add(bgItem);

            //Style8
            bgItem = new BackgroundItem();
            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            gradientBrush = new LinearGradientBrush();
            gradientBrush.StartPoint = new Point(0, 0);
            gradientBrush.EndPoint = new Point(1, 0);

            _binding = new Binding("SelectedTheme.Colors.BackgroundLight1.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientBrush.GradientStops.Add(gradientStop);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nền sáng 1", Brightness = 1 }, Offset = 0.0 });

            _binding = new Binding("SelectedTheme.Colors.Accent6.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 6", Brightness = 1 }, Offset = 1.0 });
            gradientBrush.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = gradientBrush;
            bgItem.Name = "Style 8";
            Backgrounds.Add(bgItem);

            //Style9
            bgItem = new BackgroundItem();
            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            RadialGradientBrush radialGradient = new RadialGradientBrush();


            _binding = new Binding("SelectedTheme.Colors.BackgroundLight1.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            radialGradient.GradientStops.Add(gradientStop);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nền sáng 1", Brightness = 1 }, Offset = 0.0 });

            _binding = new Binding("SelectedTheme.Colors.BackgroundDark1.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nền tối 1", Brightness = 1 }, Offset = 1.0 });
            radialGradient.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = radialGradient;
            bgItem.Name = "Style 9";
            Backgrounds.Add(bgItem);


            //Style10
            bgItem = new BackgroundItem();
            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            radialGradient = new RadialGradientBrush();


            _binding = new Binding("SelectedTheme.Colors.BackgroundLight1.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            radialGradient.GradientStops.Add(gradientStop);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nền sáng 1", Brightness = 1 }, Offset = 0.0 });

            _binding = new Binding("SelectedTheme.Colors.Accent4.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 4", Brightness = 1 }, Offset = 1.0 });
            radialGradient.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = radialGradient;
            bgItem.Name = "Style 10";
            Backgrounds.Add(bgItem);

            //Style11
            bgItem = new BackgroundItem();
            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            radialGradient = new RadialGradientBrush();


            _binding = new Binding("SelectedTheme.Colors.BackgroundLight1.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            radialGradient.GradientStops.Add(gradientStop);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nền sáng 1", Brightness = 1 }, Offset = 0.0 });

            _binding = new Binding("SelectedTheme.Colors.Accent5.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 5", Brightness = 1 }, Offset = 1.0 });
            radialGradient.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = radialGradient;
            bgItem.Name = "Style 11";
            Backgrounds.Add(bgItem);

            //Style12
            bgItem = new BackgroundItem();
            gradientColor = new ColorGradientBrush();
            gradientColor.GradientStops = new ObservableCollection<CustomGradientStop>();

            radialGradient = new RadialGradientBrush();


            _binding = new Binding("SelectedTheme.Colors.BackgroundLight1.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 0.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            radialGradient.GradientStops.Add(gradientStop);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nền sáng 1", Brightness = 1 }, Offset = 0.0 });

            _binding = new Binding("SelectedTheme.Colors.Accent6.Color");
            _binding.Source = (Application.Current as IAppGlobal);
            _binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
            _binding.Mode = BindingMode.TwoWay;
            gradientStop = new GradientStop();
            gradientStop.Offset = 1.0;
            BindingOperations.SetBinding(gradientStop, GradientStop.ColorProperty, _binding);
            gradientColor.GradientStops.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = gradientStop.Color.ToString(), Name = "Nhấn mạnh 6", Brightness = 1 }, Offset = 1.0 });
            radialGradient.GradientStops.Add(gradientStop);

            bgItem.ColorDataStyle = gradientColor;
            bgItem.ColorStyle = radialGradient;
            bgItem.Name = "Style 12";
            Backgrounds.Add(bgItem);

        }
        #endregion
    }

    public class PlaceHolderItem
    {
        private string _urlIcon;

        public string UrlIcon
        {
            get { return _urlIcon; }
            set { _urlIcon = value; }
        }

        private string _name;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
    }


}
