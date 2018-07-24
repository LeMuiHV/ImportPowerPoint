using INV.Elearning.Controls;
using INV.Elearning.Core;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.Core.ViewModel;
using INV.Elearning.DesignControl.Models;
using INV.Elearning.DesignControl.Views;
using INV.Elearning.EOPC;
using INV.Elearning.Text;
using INV.Elearning.Text.Models;
using INV.Elearning.Text.ViewModels.Text;
using INV.Elearning.Text.Views;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace INV.Elearning.DesignControl.ViewModel
{
    public class DesignTabControlViewModel : RootViewModel
    {
        static bool IsColorChange = false, isClickVariants = false, isClickTheme = false, isFontsSelect = false;
        static ETheme BackupEtheme, BackupEThemeSelectedTheme;
        static EColorManagment BackupColor;
        static ColorBrushBase BackupColorBase;
        static EFontfamily BackupFonts;

        public static InRibbonGallery ThemeDesign { get; set; }

        private ObservableCollection<BackgroundItem> _backgrounds;

        public ObservableCollection<BackgroundItem> Backgrounds
        {
            get { return _backgrounds ?? (_backgrounds = new ObservableCollection<BackgroundItem>()); }
            set { _backgrounds = value; OnPropertyChanged("Backgrounds"); }
        }

        public DesignTabControlViewModel()
        {
            GenerateBackgroundStype();
        }

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

        #region LayoutCommand
        private static RelayCommand _layoutCommand;
        /// <summary>
        /// Layout Command
        /// </summary>
        public static RelayCommand LayoutCommand
        {
            get { return _layoutCommand ?? (_layoutCommand = new RelayCommand(p => LayoutExecute(p))); }
        }

        private static void LayoutExecute(object p)
        {
            if (p is ELayoutMaster eLayoutMaster)
            {
                (Application.Current as IAppGlobal).SelectedSlide.LayoutID = eLayoutMaster.ID;
                //UpdateLayoutMaster(eLayoutMaster, (Application.Current as IAppGlobal).SelectedSlide, (Application.Current as IAppGlobal).SelectedThemes.Clone() as EThemes);
            }
        }
        #endregion

        #region BackgroundStyleCommand
        private static RelayCommand _backgroundStyleCommand;
        /// <summary>
        /// Command backgroundStyle
        /// </summary>
        public static RelayCommand BackgroundStyleCommand
        {
            get { return _backgroundStyleCommand ?? (_backgroundStyleCommand = new RelayCommand(p => BackgroundStyleExecute(p))); }
        }

        private static void BackgroundStyleExecute(object p)
        {
            if (p is BackgroundItem backgroundItem)
            {
                (Application.Current as IAppGlobal).DocumentControl.SelectedBackground = backgroundItem;
            }
        }

        #endregion

        #region HeaderAndFooterCommand
        private static RelayCommand _headerAndFooterCommand;

        public static RelayCommand HeaderAndFooterCommand
        {
            get { return _headerAndFooterCommand ?? (_headerAndFooterCommand = new RelayCommand(p => HeaderAndFooterExecute())); }
        }

        private static void HeaderAndFooterExecute()
        {
            HeaderAndFootersWindow window = new HeaderAndFootersWindow();
            window.Owner = Application.Current?.MainWindow;
            window.ShowDialog();
        }

        /// <summary>
        /// Lấy Footter của Slide
        /// </summary>
        /// <param name="slide"></param>
        /// <returns></returns>
        public static string GetFooterSlide(SlideBase slide)
        {
            string result = null;
            foreach (var item in slide.MainLayout.Elements)
            {
                if (item is TextEditor textEditor && textEditor.RichTextEditor.TextContainer.Document.TypeText == TypeText.Footers)
                {
                    result = Helper.CommandHelper.GetTextDocument(textEditor.RichTextEditor);
                }
            }
            return result;
        }
        #endregion


        // #region UpdateLayout
        /// <summary>
        /// Cap nhat lua chon Layout
        /// </summary>
        /// <param name="slideBase"></param>
        public static void UpdateLayout(SlideBase slideBase)
        {
            Global.BeginInit();
            ESlideMaster eSlideMaster = null;
            bool isDate = slideBase.IsDate;
            bool isFooter = slideBase.IsFooter;
            bool isSlideNumber = slideBase.IsSlideNumber;
            ELayoutMaster eLayoutMaster = FindLayoutMasterByID(slideBase.LayoutID, out eSlideMaster);
            if (eLayoutMaster != null)
            {
                if (eLayoutMaster.MainLayer.Background != null && !slideBase.IsHideBackground && slideBase.MainLayout.Fill == null)
                    slideBase.MainLayout.Fill = ColorHelper.ConverFromColorData(eLayoutMaster.MainLayer.Background);
                slideBase.SelectLayoutType = eLayoutMaster.LayoutType;
                slideBase.MainLayout.ThemeLayoutOwnerID = eLayoutMaster.ID;
                if (slideBase.ThemeLayout == null) slideBase.ThemeLayout = new ThemeLayout();//Cập nhật themeLayout


                if (eLayoutMaster.IsHideBackground)
                {
                    // slideBase.ThemeLayout.UpdateUI(eLayoutMaster.MainLayer);
                }
                else
                {
                    slideBase.ThemeLayout.UpdateWithParent(eLayoutMaster.MainLayer, eSlideMaster.MainLayer);
                }
                slideBase.ThemeLayout.UpdateThemeFont();
                slideBase.ThemeLayout.UpdateThemeColor();
                if (!Global.IsPasting && slideBase is NormalSlide)
                {
                    List<ObjectElementBase> _listObject = eLayoutMaster.MainLayer.Children.ToList();
                    for (int i = 0; i < slideBase.MainLayout.Elements.Count; i++)
                    {
                        ObjectElement item = slideBase.MainLayout.Elements[i];
                        //Kiem tra ton tai cua nhung khung mac dinh nhung khong ton tai trong Layout duoc cau hinh
                        if ((item is TextEditor textEditor && textEditor.RichTextEditor.TextContainer.Document.TypeText != TypeText.Normal))
                        {
                            foreach (ObjectElementBase child in _listObject)
                            {
                                if (child.PlaceHolderType == item.PlaceHolderType)
                                {
                                    item.Width = child.Width;
                                    item.Top = child.Top;
                                    item.Height = child.Height;
                                    item.Left = child.Left;
                                    item.ID = child.ID;
                                    item.IsShow = child.IsShow;
                                    item.Angle = child.Angle;
                                    item.IsScaleX = child.IsScaleX;
                                    item.IsScaleY = child.IsScaleY;
                                    _listObject.Remove(child);
                                    break;
                                }
                            }
                        }
                    }

                    var placeHolders = slideBase.MainLayout.Elements.Where(x => x.PlaceHolderType != PlaceHolderEnum.None).ToList();
                    foreach (var item in placeHolders)
                    {
                        if (eLayoutMaster.MainLayer.Children.FirstOrDefault(x => x.ID == item.ID) == null)
                        {
                            slideBase.MainLayout.Elements.Remove(item);
                        }

                    }

                    //Thêm các đối tượng chưa có vào slide
                    foreach (ObjectElementBase item in _listObject)
                    {
                        if (!item.IsGraphicBackground)
                        {
                            var child = ObjectElementsHelper.LoadData(item);
                            if (child != null)
                                slideBase.MainLayout.Elements.Add(child);
                        }

                    }

                    slideBase.IsDate = isDate;
                    slideBase.IsFooter = isFooter;
                    slideBase.IsSlideNumber = isSlideNumber;

                    //Ẩn hiện các textbox datetime, footer, slidenumber theo thuộc tính của slide
                    foreach (var item in slideBase.MainLayout.Elements)
                    {
                        if (item is TextEditor textEditor)
                        {
                            if (textEditor.RichTextEditor.TextContainer.Document.TypeText == TypeText.DateTime)
                            {
                                if (!slideBase.IsDate) textEditor.IsShow = false;
                                else textEditor.IsShow = true;
                            }
                            else if (textEditor.RichTextEditor.TextContainer.Document.TypeText == TypeText.SlideNumber)
                            {
                                if (!slideBase.IsSlideNumber) textEditor.IsShow = false;
                                else textEditor.IsShow = true;
                            }
                            else if (textEditor.RichTextEditor.TextContainer.Document.TypeText == TypeText.Footers)
                            {
                                if (!slideBase.IsFooter) textEditor.IsShow = false;
                                else textEditor.IsShow = true;
                            }
                        }
                        else if (item is PlaceHolder place)
                        {
                            place.IsMasterSlide = false;
                        }
                    }
                }
            }
            slideBase.MainLayout.UpdateThumbnail();

            foreach (ELayoutMaster item in (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters)
            {
                item.LayountCount = CountLayoutMaster(item);
            }

            Global.EndInit();
        }

        private static ObjectElement CheckPlaceHolderType(SlideBase slideBase, ELayoutMaster eLayoutMaster)
        {
            ObjectElement objectElement = null;

            for (int i = 0; i < slideBase.MainLayout.Elements.Count; i++)
            {
                ObjectElement item = slideBase.MainLayout.Elements[i];
                //Kiem tra ton tai cua nhung khung mac dinh nhung khong ton tai trong Layout duoc cau hinh
                if (item is PlaceHolder || (item is TextEditor textEditor && textEditor.RichTextEditor.TextContainer.Document.TypeText != TypeText.Normal))
                {
                    foreach (ObjectElementBase child in eLayoutMaster.MainLayer.Children)
                    {
                        if (child.PlaceHolderType == item.PlaceHolderType)
                        {
                            item.UpdateUI(child);
                        }
                    }
                }
            }
            return objectElement;
        }


        /// <summary>
        /// Đếm số lượng Object Element có trong Slide
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private static int GetCountElement(Type type, SlideBase slide)
        {
            int count = 0;
            foreach (var item in slide.MainLayout.Elements)
            {
                if (item is ObjectElement objectElement && objectElement.GetType() == type)
                {
                    count = Math.Max(count, objectElement.ElementCount);
                }
            }
            return count + 1;
        }

        /// <summary>
        /// Cập nhật chỉ số trang của Slide
        /// </summary>
        /// <param name="slide"></param>
        public static void UpdateSlideNumber(SlideBase slide)
        {
            if (slide is NormalSlide)
            {
                foreach (var item in slide.MainLayout.Elements)
                {
                    if (item is TextEditor textEditor && textEditor.RichTextEditor.TextContainer.Document.TypeText == TypeText.SlideNumber)
                    {
                        textEditor.RichTextEditor.TextContainer.Document = Helper.CommandHelper.SetTextDocument(textEditor.RichTextEditor.TextContainer.Document, slide.SlideNumber.ToString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                        textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).DocumentControl.SelectedTheme?.SelectedFont.MinorFont;
                        textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                        textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
                        textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
                        textEditor.RichTextEditor.TextContainer.InvalidateVisual();
                    }
                }
            }
        }

        public static void CheckIsShowTextBox(ObjectElement value)
        {
            // Console.WriteLine(value);
            if ((Application.Current as IAppGlobal).SelectedSlide != null)
            {
                if (value is TextEditor textEditor)
                {
                    if (textEditor.RichTextEditor.TextContainer.Document.TypeText == TypeText.DateTime)
                    {
                        (Application.Current as IAppGlobal).SelectedSlide.IsDate = textEditor.IsShow;
                    }
                    else if (textEditor.RichTextEditor.TextContainer.Document.TypeText == TypeText.Footers)
                    {
                        (Application.Current as IAppGlobal).SelectedSlide.IsFooter = textEditor.IsShow;
                    }
                    else if (textEditor.RichTextEditor.TextContainer.Document.TypeText == TypeText.SlideNumber)
                    {
                        (Application.Current as IAppGlobal).SelectedSlide.IsSlideNumber = textEditor.IsShow;
                    }
                }
            }
        }


        /// <summary>
        /// tìm kiếm Layout
        /// </summary>
        /// <param name="id"></param>
        /// <param name="slideMaster"></param>
        /// <returns></returns>
        public static ELayoutMaster FindLayoutMasterByID(string id, out ESlideMaster slideMaster)
        {
            foreach (var slide in (Application.Current as IAppGlobal).SelectedTheme.SlideMasters)
            {
                foreach (var layout in slide.LayoutMasters)
                {
                    if (layout.ID == id)
                    {
                        slideMaster = slide;
                        return layout;
                    }
                }
            }
            slideMaster = null;
            return null;
        }



        public static int CountLayoutMaster(ELayoutMaster eLayout)
        {
            int count = 0;
            foreach (SlideBase item in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {
                if (item.MainLayout.ThemeLayoutOwnerID == eLayout.ID)
                {
                    count++;
                }
            }
            return count;
        }

        //    public static void PreviewUpdateLayoutMaster(ELayoutMaster data, SlideBase slideBase, ESlideMaster eSlideMaster)
        //    {
        //        Global.BeginInit();
        //        if (data != null)
        //        {
        //            if (slideBase.ThemeLayout == null) slideBase.ThemeLayout = new ThemeLayout();
        //            if (data.IsHideBackground)
        //            {
        //                slideBase.ThemeLayout.UpdateUI(data.MainLayer);
        //            }
        //            else
        //            {
        //                slideBase.ThemeLayout.UpdateWithParent(data.MainLayer, eSlideMaster.MainLayer);
        //            }

        //            //slideBase.MainLayout.UpdateLayout(data);

        //            foreach (var item in data.MainLayer.Children)
        //            {
        //                if (item is DataDocument || item is EPlaceHolder)
        //                {
        //                    ObjectElement _element = slideBase.MainLayout.Elements.FirstOrDefault(x => x.ID == item.ID);
        //                    if (_element != null)
        //                    {
        //                        _element.UpdateUI(item);
        //                    }
        //                    else
        //                    {
        //                        _element = ObjectElementsHelper.LoadData(item);
        //                        if (_element != null)
        //                            slideBase.MainLayout.Elements.Add(_element);
        //                        if (_element is IPlaceHolder place) place.IsMasterSlide = false;

        //                    }
        //                }

        //                for (int i = 0; i < slideBase.MainLayout.Elements.Count; i++)
        //                {
        //                    if (data.MainLayer.Children.FirstOrDefault(x => x.ID == slideBase.MainLayout.Elements[i].ID) == null)
        //                    {
        //                        slideBase.MainLayout.Elements.RemoveAt(i);
        //                        i--;
        //                    }
        //                }
        //            }
        //        }
        //        Global.EndInit();
        //    }
        //    #endregion

        #region AddSlideWithLayoutCommand
        private static RelayCommand _addSlideWithLayoutCommand;
        /// <summary>
        /// Thêm Slide với Layout được chọn
        /// </summary>
        public static RelayCommand AddSlideWithLayoutCommand
        {
            get { return _addSlideWithLayoutCommand ?? (_addSlideWithLayoutCommand = new RelayCommand(p => AddSlideWithLayoutExecute(p))); }
        }

        private static void AddSlideWithLayoutExecute(object p)
        {
            if (p is ELayoutMaster eLayoutMaster)
            {
                NormalSlide normalSlide = new NormalSlide();
                normalSlide.IsFooter = (Application.Current as IAppGlobal).SelectedSlide.IsFooter;
                normalSlide.IsDate = (Application.Current as IAppGlobal).SelectedSlide.IsDate;
                normalSlide.IsSlideNumber = (Application.Current as IAppGlobal).SelectedSlide.IsSlideNumber;
                normalSlide.TextFooter = GetFooterSlide((Application.Current as IAppGlobal).SelectedSlide);
                normalSlide.LayoutID = eLayoutMaster.ID;
                (Application.Current as IAppGlobal).DocumentControl.Slides.Add(normalSlide);
            }
        }

        #endregion


        #region AddNewSlide
        private static RelayCommand _addNewSlideCommand;
        /// <summary>
        /// Thêm một slide với với theme đang được chọn
        /// </summary>
        public static RelayCommand AddNewSlideCommand
        {
            get { return _addNewSlideCommand ?? (_addNewSlideCommand = new RelayCommand(p => AddNewSlideExecute())); }
        }

        private static void AddNewSlideExecute()
        {
            NormalSlide normalSlide = new NormalSlide();
            normalSlide.IsFooter = (Application.Current as IAppGlobal).SelectedSlide.IsFooter;
            normalSlide.IsDate = (Application.Current as IAppGlobal).SelectedSlide.IsDate;
            normalSlide.IsSlideNumber = (Application.Current as IAppGlobal).SelectedSlide.IsSlideNumber;
            normalSlide.TextFooter = GetFooterSlide((Application.Current as IAppGlobal).SelectedSlide);
            foreach (ELayoutMaster item in (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].LayoutMasters)
            {
                if (item.LayoutType == LayoutType.Blank)
                {
                    normalSlide.LayoutID = item.ID;
                }
            }
            (Application.Current as IAppGlobal).DocumentControl.Slides.Add(normalSlide);
        }

        #endregion

        public static void UpdateThemeSlide(ETheme eTheme, ETheme oldTheme)
        {
            foreach (ETheme item in (Application.Current as IAppGlobal).LocalThemesCollection)
            {
                item.CountTheme = CountThemeApply(item);
            }
        }

        private static RelayCommand _themeCommand;

        public static RelayCommand ThemeCommand
        {
            get { return _themeCommand ?? (_themeCommand = new RelayCommand(p => ThemeExecute(p))); }
        }

        private static void ThemeExecute(object p)
        {
            //Global.BeginInit();
            if (p is ETheme eTheme)
            {
                //CHeck selected theme
                if (eTheme.TagName == "Office")
                {
                    ETheme cloneTheme = (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Clone() as ETheme;
                    cloneTheme.SlideMasters[0] = eTheme.SlideMasters[0];
                    cloneTheme.Colors = eTheme.Colors;
                    cloneTheme.TagName = "This Presenter";
                    cloneTheme.ThemeName = eTheme.ThemeName;
                    cloneTheme.SelectedFont = eTheme.SelectedFont;
                    cloneTheme.ID = ObjectElementsHelper.RandomString(10);
                    cloneTheme.IsClone = true;
                    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = cloneTheme;

                }
                else
                {
                    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = eTheme;
                }
            }
        }


        public static void DisposeSlideMaster(SlideMaster slide)
        {
            for (int i = 0; i < slide.MainLayout.Elements.Count; i++)
            {
                slide.MainLayout.Elements.RemoveAt(i--);
            }
            for (int j = 0; j < slide.LayoutMasters.Count; j++)
            {
                for (int k = 0; k < slide.LayoutMasters[j].MainLayout.Elements.Count; k++)
                {
                    slide.LayoutMasters[j].MainLayout.Elements.RemoveAt(k--);
                }
                slide.LayoutMasters.RemoveAt(j--);
            }
        }


        public static int CountThemeApply(ETheme eThemes)
        {
            int count = 0;
            foreach (SlideBase slide in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {
                if (slide.ThemeLayout != null && slide.ThemeLayout.ThemeID == eThemes.ID)
                {
                    count++;
                }
            }
            return count;
        }

        /// <summary>
        /// Xuất ra một String ngẫu nhiên
        /// </summary>
        /// <returns></returns>
        public static string RandomString()
        {
            int length = 15;
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            var random = new Random();
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        //    #endregion

        //    #region PreviewThemeCommand
        //    private static RelayCommand _previewThemeCommand;
        //    /// <summary>
        //    /// Preview Theme Command
        //    /// </summary>
        //    public static RelayCommand PreviewThemeCommand
        //    {
        //        get { return _previewThemeCommand ?? (_previewThemeCommand = new RelayCommand(p => PreviewThemeExecute(p))); }
        //    }

        //    private static void PreviewThemeExecute(object p)
        //    {
        //        if (p is GalleryItem galleryItem)
        //        {
        //            Global.BeginInit();
        //            EThemes eThemes = galleryItem.Content as EThemes;
        //            //(Application.Current as IAppGlobal).SelectedTheme.Colors = eThemes.MainTheme.Colors;
        //            if ((Application.Current as IAppGlobal).DocumentControl.SelectedTheme != null)
        //                BackupEThemeSelectedTheme = (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Clone() as ETheme;
        //            SlideBase slideBase = (Application.Current as IAppGlobal).SelectedSlide;
        //            if (slideBase.ThemeLayout == null) slideBase.ThemeLayout = new ThemeLayout();
        //            if (string.IsNullOrEmpty(slideBase.MainLayout.ThemeLayoutOwnerID))
        //            {
        //                slideBase.ThemeLayout.UpdateUI(eThemes.MainTheme.SlideMasters[0].MainLayer);
        //            }
        //            else
        //            {
        //                foreach (ELayoutMaster item in eThemes.MainTheme.SlideMasters[0].LayoutMasters)
        //                {
        //                    if (item.LayoutType == slideBase.SelectLayoutType)
        //                    {
        //                        PreviewUpdateLayoutMaster(item, slideBase, eThemes.MainTheme.SlideMasters[0]);
        //                    }
        //                }

        //            }

        //            BackupColorBase = slideBase.MainLayout.Fill;
        //            slideBase.MainLayout.Fill = ColorHelper.ConverFromColorData(eThemes.MainTheme.SlideMasters[0].MainLayer.Background);
        //            //(Application.Current as IAppGlobal).DocumentControl.SelectedTheme = eThemes.MainTheme;
        //            Global.EndInit();
        //        }
        //    }

        //    #endregion

        //    #region PreviewCancelCommand
        //    private static RelayCommand _previewCancelCommand;
        //    /// <summary>
        //    /// PreviewCancelCommand
        //    /// </summary>
        //    public static RelayCommand PreviewCancelCommand
        //    {
        //        get { return _previewCancelCommand ?? (_previewCancelCommand = new RelayCommand(p => PreviewCancelExecute(p))); }
        //    }

        //    private static void PreviewCancelExecute(object p)
        //    {
        //        if (!isClickTheme)
        //        {
        //            Global.BeginInit();
        //            SlideBase slideBase = (Application.Current as IAppGlobal).SelectedSlide;
        //            slideBase.MainLayout.Fill = BackupColorBase;
        //            if (string.IsNullOrEmpty(slideBase.MainLayout.ThemeLayoutOwnerID))
        //            {
        //                (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = BackupEThemeSelectedTheme;
        //            }
        //            else
        //            {
        //                (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = BackupEThemeSelectedTheme;
        //                foreach (ELayoutMaster item in BackupEThemeSelectedTheme.SlideMasters[0].LayoutMasters)
        //                {
        //                    if (item.ID == slideBase.MainLayout.ThemeLayoutOwnerID)
        //                    {
        //                        PreviewUpdateLayoutMaster(item, slideBase, BackupEThemeSelectedTheme.SlideMasters[0]);
        //                    }
        //                }
        //            }
        //            //EThemes eThemes = (Application.Current as IAppGlobal).SelectedThemes;

        //            Global.EndInit();
        //        }
        //        else isClickTheme = false;
        //    }


        //    #endregion

        //    #region PreviewVariantsCommand
        //    private static RelayCommand _previewVariantsCommand;
        //    /// <summary>
        //    /// Preview Variants Command
        //    /// </summary>
        //    public static RelayCommand PreviewVariantsCommand
        //    {
        //        get { return _previewVariantsCommand ?? (_previewVariantsCommand = new RelayCommand(p => PreviewVariantsExecute(p))); }
        //    }

        //    private static void PreviewVariantsExecute(object p)
        //    {
        //        if (p is GalleryItem galleryItem)
        //        {
        //            //Global.PushUndo(new SelectedThemeStep((Application.Current as IAppGlobal).DocumentControl.SelectedTheme, galleryItem.Content as ETheme));
        //            Global.BeginInit();
        //            if ((Application.Current as IAppGlobal).DocumentControl.SelectedTheme != null)
        //            {
        //                BackupEtheme = (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Clone() as ETheme;
        //                ETheme eThemes = galleryItem.Content as ETheme;
        //                (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = (Application.Current as IAppGlobal).SelectedTheme.Colors = eThemes.Colors;
        //                SlideBase slideBase = (Application.Current as IAppGlobal).SelectedSlide;
        //                slideBase.UpdateThemeColor();
        //                //if (slideBase.ThemeLayout == null) slideBase.ThemeLayout = new ThemeLayout();
        //                //slideBase.ThemeLayout.UpdateUI(eThemes.SlideMasters[0].MainLayer);
        //                //(Application.Current as IAppGlobal).DocumentControl.SelectedTheme = (Application.Current as IAppGlobal).SelectedTheme = eThemes;
        //            }
        //            Global.EndInit();
        //        }
        //    }

        //    #endregion

        //    #region CancelPreviewVariantsCommand
        //    private static RelayCommand _cancelPreviewVariantsCommand;
        //    /// <summary>
        //    /// Cancel Preview Variants Command
        //    /// </summary>
        //    public static RelayCommand CancelPreviewVariantsCommand
        //    {
        //        get { return _cancelPreviewVariantsCommand ?? (_cancelPreviewVariantsCommand = new RelayCommand(p => CancelPreviewVariantsExecute(p))); }
        //    }

        //    private static void CancelPreviewVariantsExecute(object p)
        //    {
        //        if (!isClickVariants)
        //        {
        //            //Global.UndoCommand.Execute(null);
        //            if ((Application.Current as IAppGlobal).SelectedTheme != null && (Application.Current as IAppGlobal).DocumentControl.SelectedTheme != null)
        //            {
        //                Global.BeginInit();
        //                (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = (Application.Current as IAppGlobal).SelectedTheme.Colors = BackupEtheme.Colors;
        //                SlideBase slideBase = (Application.Current as IAppGlobal).SelectedSlide;
        //                slideBase.UpdateThemeColor();
        //                // (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = (Application.Current as IAppGlobal).SelectedTheme = BackupEtheme;
        //                Global.EndInit();
        //            }
        //        }
        //        else
        //        {
        //            isClickVariants = false;
        //        }
        //    }

        //    #endregion

        #region CustomizeColorsCommand
        private RelayCommand _customizeColorsCommand;
        /// <summary>
        /// Lệnh điều khiển mở cửa sổ thay đổi màu theme
        /// </summary>
        public RelayCommand CustomizeColorsCommand
        {
            get { return _customizeColorsCommand ?? (_customizeColorsCommand = new RelayCommand(p => CustomizeColorsExecute())); }
        }

        private void CustomizeColorsExecute()
        {
            ThemeColorWindow _window = new ThemeColorWindow();
            _window.Owner = Application.Current?.MainWindow;
            _window.ShowDialog();
            if (!_window.IsCancelled)
            {
                (Application.Current as IAppGlobal).DocumentControl.SelectedColors = _window.Colors;


            }
        }

        #endregion

        #region CustomizeFontsCommand
        private RelayCommand _customizeFontsCommand;
        /// <summary>
        /// Lệnh điều khiển mở cửa sổ thay đổi màu theme
        /// </summary>
        public RelayCommand CustomizeFontsCommand
        {
            get { return _customizeFontsCommand ?? (_customizeFontsCommand = new RelayCommand(p => CustomizeFontsExecute())); }
        }

        private void CustomizeFontsExecute()
        {
            ThemeFontWindow _window = new ThemeFontWindow();
            _window.Owner = Application.Current?.MainWindow;
            _window.ShowDialog();
            if (!_window.IsCancelled)
            {
                (Application.Current as IAppGlobal).DocumentControl.SelectedFont = _window.DataContext as INV.Elearning.Core.Model.Theme.EFontfamily;
            }
        }

        #endregion

        //    #region VariantsCommand
        //    private static RelayCommand _variantsCommand;
        //    /// <summary>
        //    /// Command khi lựa chọn variant
        //    /// </summary>
        //    public static RelayCommand VariantsCommand
        //    {
        //        get { return _variantsCommand ?? (_variantsCommand = new RelayCommand(p => VariantsExecute(p))); }
        //    }

        //    private static void VariantsExecute(object p)
        //    {
        //        if (p is ETheme eTheme)
        //        {
        //            (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = eTheme;
        //        }
        //        isClickVariants = true;
        //        return;
        //        Global.PushUndo(new SelectedThemeStep((Application.Current as IAppGlobal).SelectedTheme.Colors.Clone() as EColorManagment, (p as ETheme).Colors.Clone() as EColorManagment, (Application.Current as IAppGlobal).SelectedThemes.Clone() as EThemes));
        //        Global.BeginInit();
        //        if (p is ETheme _eTheme)
        //        {
        //            (Application.Current as IAppGlobal).SelectedTheme.Colors = (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = eTheme.Colors;
        //            (Application.Current as IAppGlobal).SelectedThemes.MainTheme.Colors = eTheme.Colors;
        //            foreach (SlideBase slide in (Application.Current as IAppGlobal).DocumentControl.Slides)
        //            {
        //                if (slide.ThemeLayout == null) slide.ThemeLayout = new ThemeLayout();
        //                slide.UpdateThemeColor();
        //            }

        //            foreach (EThemes item in (Application.Current as IAppGlobal).LocalThemesCollection)
        //            {
        //                item.CountTheme = 0;
        //            }


        //            foreach (ELayoutMaster item in (Application.Current as IAppGlobal).SelectedThemes.MainTheme.SlideMasters[0].LayoutMasters)
        //            {
        //                Task.Factory.StartNew(() =>
        //                {
        //                    Application.Current.Dispatcher.Invoke(() =>
        //                    {
        //                        LayoutMaster layoutMaster = new LayoutMaster();
        //                        layoutMaster.LoadData(item);
        //                        if (!layoutMaster.IsHideBackground)
        //                        {
        //                            if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();
        //                            layoutMaster.ThemeLayout.UpdateUIB((Application.Current as IAppGlobal).SelectedThemes.MainTheme.SlideMasters[0].MainLayer);
        //                        }
        //                        layoutMaster.UpdateThemeColor();
        //                        item.ThumbnailBitmap = CaptureLayout(layoutMaster.MainLayout, new Size(341, 192));
        //                        layoutMaster = null;
        //                    });
        //                });
        //            }
        //            SlideMaster slideMaster = new SlideMaster();
        //            slideMaster.MainLayout.Fill = ColorHelper.ConverFromColorData((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer.Background);
        //            slideMaster.ThemeLayout = new ThemeLayout();
        //            slideMaster.ThemeLayout.UpdateUI((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer);
        //            slideMaster.UpdateThemeColor();
        //            (Application.Current as IAppGlobal).DocumentControl.SelectedThemes.ThumbnailIcon = GetThumbnail(slideMaster.MainLayout, eTheme.Colors, new Size(341, 192));
        //        }
        //        isClickVariants = true;
        //        Global.EndInit();
        //    }

        //    #endregion

        #region ColorCommand
        private static RelayCommand _colorsCommand;
        /// <summary>
        /// Colors Command
        /// </summary>
        public static RelayCommand ColorsCommand
        {
            get { return _colorsCommand ?? (_colorsCommand = new RelayCommand(p => ColorExecute(p))); }
        }

        private static void ColorExecute(object p)
        {
            if (p is EColorManagment color)
            {
                //Global.PushUndo(new ColorChangedStep(BackupColor, color));
                //Global.BeginInit();
                IsColorChange = true;
                (Application.Current as IAppGlobal).DocumentControl.SelectedColors = color;

                // Global.EndInit();
            }
        }
        /// <summary>
        /// Cập nhật màu cho slide
        /// </summary>
        /// <param name="color"></param>
        public static void UpdateColorSlide(EColorManagment color)
        {
            (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = color;
            (Application.Current as IAppGlobal).SelectedTheme.Colors = color;
            //SlideMaster slideMaster = new SlideMaster();
            //slideMaster.MainLayout.Fill = ColorHelper.ConverFromColorData((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer.Background);
            //slideMaster.ThemeLayout = new ThemeLayout();
            //slideMaster.ThemeLayout.UpdateUI((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer);
            //slideMaster.UpdateThemeColor();
            SlideBase slide = null;
            //(Application.Current as IAppGlobal).DocumentControl.SelectedTheme.ThumbnailBitmap = GetThumbnail(slideMaster.MainLayout, color, (Application.Current as IAppGlobal).SelectedTheme.SelectedFont, new Size(341, 192));
            if ((Application.Current as IAppGlobal).SelectedSlide != null)
            {
                slide = (Application.Current as IAppGlobal).SelectedSlide;
            }
            ////Cập nhật thumbnail cho Layout Master
            //foreach (ELayoutMaster item in (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].LayoutMasters)
            //{
            //    LayoutMaster layoutMaster = new LayoutMaster();
            //    layoutMaster.LoadData(item);
            //    if (!layoutMaster.IsHideBackground)
            //    {
            //        if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();
            //        layoutMaster.ThemeLayout.UpdateUIB((Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].MainLayer);
            //    }
            //    layoutMaster.MainLayout.Fill = ColorHelper.ConverFromColorData((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer.Background);
            //    layoutMaster.UpdateThemeColor();
            //    item.ThumbnailBitmap = CaptureLayout(layoutMaster.MainLayout, new Size(341, 192));
            //    layoutMaster = null;

            //}
            if (slide != null)
            {
                SlideHelper.UnSlectedAll();
                slide.IsSelected = true;
            }
            foreach (var item in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {
                item.UpdateThemeColor();
            }
        }
        #endregion

        public static void UpdateThumbnailBackground(BackgroundItem item)
        {


        }

        #region PreviewColorCommand
        private static RelayCommand _previewColorCommand;
        /// <summary>
        /// Preview Color Command
        /// </summary>
        public static RelayCommand PreviewColorCommand
        {
            get { return _previewColorCommand ?? (_previewColorCommand = new RelayCommand(p => PreviewColorExecute(p))); }
        }

        private static void PreviewColorExecute(object p)
        {
            if (p is GalleryItem galleryItem)
            {
                Global.BeginInit();
                EColorManagment eColor = galleryItem.Content as EColorManagment;
                SlideBase slide = (Application.Current as IAppGlobal).SelectedSlide;
                (Application.Current as IAppGlobal).SelectedTheme.Colors = eColor;
                slide.MainLayout?.UpdateThemeColor();
                slide.ThemeLayout?.UpdateThemeColor();
                foreach (LayoutBase layout in slide.Layouts)
                {
                    layout.UpdateThemeColor();
                }
                //if ((Application.Current as IAppGlobal).SelectedTheme != null && (Application.Current as IAppGlobal).DocumentControl != null)
                //{
                //    //(Application.Current as IAppGlobal).SelectedTheme.Colors = eColor;
                //    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = eColor;

                //    SlideBase slide = (Application.Current as IAppGlobal).SelectedSlide;
                //    slide.UpdateThemeColor();
                //}
                Global.EndInit();
            }
        }
        #endregion

        #region CancelPreviewColorCommand
        private static RelayCommand _cancelPreviewColorCommand;
        /// <summary>
        /// Cancel Preview Color Command
        /// </summary>
        public static RelayCommand CancelPreviewColorCommand
        {
            get { return _cancelPreviewColorCommand ?? (_cancelPreviewColorCommand = new RelayCommand(p => CancelPreviewColorExecute(p))); }
        }

        private static void CancelPreviewColorExecute(object p)
        {
            if (BackupColor != null)
            {
                Global.BeginInit();
                if (!IsColorChange)
                {
                    SlideBase slide = (Application.Current as IAppGlobal).SelectedSlide;
                    (Application.Current as IAppGlobal).SelectedTheme.Colors = BackupColor;
                    slide.MainLayout?.UpdateThemeColor();
                    slide.ThemeLayout?.UpdateThemeColor();
                    foreach (LayoutBase layout in slide.Layouts)
                    {
                        layout.UpdateThemeColor();
                    }
                }
                else
                {
                    IsColorChange = false;
                }
                Global.EndInit();
            }
        }

        #endregion

        #region DropDownColorCommand
        private static RelayCommand _dropDownColorCommand;
        /// <summary>
        /// Dropdown color command
        /// </summary>
        public static RelayCommand DropDownColorCommand
        {
            get { return _dropDownColorCommand ?? (_dropDownColorCommand = new RelayCommand(p => DropDownColorExecute())); }
        }

        private static void DropDownColorExecute()
        {
            BackupColor = (Application.Current as IAppGlobal).SelectedTheme.Colors.Clone() as EColorManagment;
        }

        #endregion

        #region FontsCommand
        public static void UpdateFontSlide(EFontfamily eFontfamily)
        {
            (Application.Current as IAppGlobal).SelectedTheme.SelectedFont = eFontfamily;
            (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SelectedFont = eFontfamily;
            foreach (SlideBase item in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {
                item.UpdateThemeFont();
                item.MainLayout.UpdateThumbnail();
            }
        }

        private static RelayCommand _fontsCommand;
        /// <summary>
        /// Command lựa chọn fonts
        /// </summary>
        public static RelayCommand FontsCommand
        {
            get { return _fontsCommand ?? (_fontsCommand = new RelayCommand(p => FontsExecute(p))); }
        }

        private static void FontsExecute(object p)
        {
            if (p is EFontfamily eFontfamily)
            {
                isFontsSelect = true;
                (Application.Current as IAppGlobal).DocumentControl.SelectedFont = eFontfamily;
            }
        }

        #endregion

        #region PreviewFontsCommand
        private static RelayCommand _previewFontsCommand;
        /// <summary>
        /// preview fonts
        /// </summary>
        public static RelayCommand PreviewFontsCommand
        {
            get { return _previewFontsCommand ?? (_previewFontsCommand = new RelayCommand(p => PreviewFontsExecute(p))); }
        }

        private static void PreviewFontsExecute(object p)
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SelectedTheme.SelectedFont = p as EFontfamily;
            (Application.Current as IAppGlobal).SelectedSlide.UpdateThemeFont();
            Global.EndInit();
        }

        #endregion

        #region CancelPreviewFontsCommand
        private static RelayCommand _cancelPreviewFontsCommand;
        /// <summary>
        /// Cancel Preview Fonts Command
        /// </summary>
        public static RelayCommand CancelPreviewFontsCommand
        {
            get { return _cancelPreviewFontsCommand ?? (_cancelPreviewFontsCommand = new RelayCommand(p => CancelPreviewFontsExecute())); }
        }

        private static void CancelPreviewFontsExecute()
        {
            Global.BeginInit();
            if (BackupFonts != null)
            {
                if (!isFontsSelect)
                {
                    (Application.Current as IAppGlobal).SelectedTheme.SelectedFont = BackupFonts;
                    (Application.Current as IAppGlobal).SelectedSlide.UpdateThemeFont();
                }
                else
                {
                    isFontsSelect = false;
                }
            }
            Global.EndInit();
        }

        #endregion

        #region DropDownFontsCommand

        private static RelayCommand _dropdownFontsCommand;

        public static RelayCommand DropdownFontsCommand
        {
            get { return _dropdownFontsCommand ?? (_dropdownFontsCommand = new RelayCommand(p => DropDownFontsExecute())); }
        }

        private static void DropDownFontsExecute()
        {
            BackupFonts = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont.Clone() as EFontfamily;
        }


        #endregion
        #region SaveThemesCommand
        private RelayCommand _saveThemeCommand;
        /// <summary>
        /// Lưu theme
        /// </summary>
        public RelayCommand SaveThemeCommand
        {
            get { return _saveThemeCommand ?? (_saveThemeCommand = new RelayCommand(p => SaveThemesExecute())); }
        }

        private void SaveThemesExecute()
        {
            //EThemes eThemes = (Application.Current as IAppGlobal).DocumentControl.SelectedThemes.Clone() as EThemes;
            //eThemes.Variants.Clear();
            ThemesData themesData = new ThemesData();
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "ThemeFile (*.athmx)|*.athmx";
            if (saveFile.ShowDialog() == true)
            {
                themesData.Save((Application.Current as IAppGlobal).DocumentControl.SelectedTheme, saveFile.FileName);
                INV.Elearning.Controls.MessageBox.ShowDialog("Save Themes Success", "Save Success");
            }
        }


        #endregion

        #region LoadThemesCommand
        private RelayCommand _loadThemeCommand;
        /// <summary>
        /// Tải lại themes
        /// </summary>
        public RelayCommand LoadThemeCommand
        {
            get { return _loadThemeCommand ?? (_loadThemeCommand = new RelayCommand(p => LoadThemeExecute())); }
        }

        private void LoadThemeExecute()
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "ThemeFile (*.athmx)|*.athmx";
            if (openFile.ShowDialog() == true)
            {
                ETheme eTheme = new ETheme();
                ThemesData themesData = new ThemesData();
                eTheme = themesData.OpenETheme(openFile.FileName);
                eTheme.TagName = "This Presenter";
                eTheme.ThemeName = Path.GetFileNameWithoutExtension(openFile.FileName);
                eTheme.ID = RandomString();
                eTheme.IsLoaded = true;
                (Application.Current as IAppGlobal).LocalThemesCollection.Add(eTheme);
                (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = eTheme;
                (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = eTheme.Colors;
                (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SelectedFont = eTheme.SelectedFont;
                int index = (Application.Current as IAppGlobal).LocalThemesCollection.IndexOf(eTheme);
                if (index != -1)
                    ThemeDesign.SelectedIndex = index;
                else ThemeDesign.SelectedIndex = 0;
            }
        }

        #endregion

        public static BitmapSource SaveToBitmapImage(LayoutBase element, EColorManagment eColor, EFontfamily eFontfamily, Size imgSize = new Size())
        {
            double _width = element.ActualWidth > 0 ? element.ActualWidth : (Application.Current as IAppGlobal).SlideSize.Width;
            double _height = element.ActualHeight > 0 ? element.ActualHeight : (Application.Current as IAppGlobal).SlideSize.Height;

            if (imgSize.Width == 0 || imgSize.Height == 0)
                imgSize = new Size(_width, _height);

            element.Measure(imgSize);
            element.Arrange(new Rect(imgSize));
            //element.UpdateLayout();

            RenderTargetBitmap renderBitmap =
              new RenderTargetBitmap(
                (int)imgSize.Width,
                (int)imgSize.Height,
                96d * imgSize.Width / _width,
                96d * imgSize.Height / _height,
                PixelFormats.Pbgra32);

            DrawingVisual dv = new DrawingVisual();
            using (DrawingContext ctx = dv.RenderOpen())
            {
                ctx.DrawRectangle(Brushes.White, null, new Rect(new Point(0, 0), new Size(_width, _height)));
                foreach (ObjectElement item in element.Elements)
                {
                    if (!item.IsGraphicBackground)
                    {
                        item.Visibility = Visibility.Hidden;
                    }
                }

                VisualBrush vb = new VisualBrush(element);
                ctx.DrawRectangle(vb, null, new Rect(new Point(0, 0), new Size(_width, _height)));

                FormattedText formatted_text = new FormattedText("A", CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(eFontfamily.MajorFont), 150, Brushes.Black);
                formatted_text.TextAlignment = TextAlignment.Center;
                ctx.DrawText(formatted_text, new Point(180, 250));

                formatted_text = new FormattedText("a", CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(eFontfamily.MinorFont), 150, Brushes.Black);
                formatted_text.TextAlignment = TextAlignment.Center;
                ctx.DrawText(formatted_text, new Point(185 + formatted_text.WidthIncludingTrailingWhitespace, 250));

                ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent1.Color), null, new Rect(130, 420, 120, 50));
                ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent2.Color), null, new Rect(260, 420, 120, 50));
                ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent3.Color), null, new Rect(390, 420, 120, 50));
                ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent4.Color), null, new Rect(520, 420, 120, 50));
                ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent5.Color), null, new Rect(650, 420, 120, 50));
                ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent6.Color), null, new Rect(780, 420, 120, 50));
            }
            renderBitmap.Render(dv);
            return renderBitmap;
        }

        public static BitmapSource SaveToSmallBitmap(LayoutBase element, EColorManagment eColor, EFontfamily eFontfamily, Size size = new Size())
        {
            try
            {
                double _width = element.ActualWidth > 0 ? element.ActualWidth : (Application.Current as IAppGlobal).SlideSize.Width;
                double _height = element.ActualHeight > 0 ? element.ActualHeight : (Application.Current as IAppGlobal).SlideSize.Height;

                double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
                double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;

                if (size.Width == 0 || size.Height == 0)
                    size = new Size(_width, _height);

                element.Measure(size);
                element.Arrange(new Rect(size));

                RenderTargetBitmap renderBitmap =
                  new RenderTargetBitmap(
                    (int)_slideWidth,
                    (int)_slideWidth,
                    96d * size.Width / _width,
                    96d * size.Height / _height,
                    PixelFormats.Pbgra32);
                DrawingVisual dv = new DrawingVisual();
                using (DrawingContext ctx = dv.RenderOpen())
                {
                    ctx.DrawRectangle(Brushes.White, null, new Rect(new Point(10, 10), new Size(_slideWidth - 20, _slideWidth - 20)));

                    foreach (ObjectElement item in element.Elements)
                    {
                        if (!item.IsGraphicBackground)
                        {
                            item.Visibility = Visibility.Hidden;
                        }
                    }

                    ScaleTransform scale = new ScaleTransform();
                    scale.ScaleY = (_slideWidth / _slideHeight);

                    ctx.PushTransform(scale);
                    VisualBrush vb = new VisualBrush(element);
                    ctx.DrawRectangle(vb, null, new Rect(new Point(0, 0), new Size(_width, _height)));
                    ctx.Pop();

                    ctx.DrawLine(new Pen(Brushes.Black, 20), new Point(0, 10), new Point(_slideWidth, 10));
                    ctx.DrawLine(new Pen(Brushes.Black, 20), new Point(0, _slideWidth - 10), new Point(_slideWidth, _slideWidth - 10));
                    ctx.DrawLine(new Pen(Brushes.Black, 20), new Point(10, 0), new Point(10, _slideWidth));
                    ctx.DrawLine(new Pen(Brushes.Black, 20), new Point(_slideWidth - 10, 0), new Point(_slideWidth - 10, _slideWidth));



                    FormattedText formatted_text = new FormattedText("A", CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(eFontfamily.MajorFont), 450, Brushes.Black);
                    formatted_text.TextAlignment = TextAlignment.Center;
                    ctx.DrawText(formatted_text, new Point(300, 250));

                    formatted_text = new FormattedText("a", CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(eFontfamily.MinorFont), 450, Brushes.Black);
                    formatted_text.TextAlignment = TextAlignment.Center;
                    ctx.DrawText(formatted_text, new Point(310 + formatted_text.WidthIncludingTrailingWhitespace, 250));

                    ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent1.Color), null, new Rect(130, 720, 120, 50));
                    ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent2.Color), null, new Rect(260, 720, 120, 50));
                    ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent3.Color), null, new Rect(390, 720, 120, 50));
                    ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent4.Color), null, new Rect(520, 720, 120, 50));
                    ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent5.Color), null, new Rect(650, 720, 120, 50));
                    ctx.DrawRectangle((Brush)new BrushConverter().ConvertFromString(eColor.Accent6.Color), null, new Rect(780, 720, 120, 50));
                }
                renderBitmap.Render(dv);
                foreach (ObjectElement item in element.Elements)
                {
                    if (!item.IsGraphicBackground)
                    {
                        item.Visibility = Visibility.Visible;
                    }
                }
                return renderBitmap;

            }
            catch (Exception e)
            {
                return null;
            }
        }
    }

    public class LayoutItem
    {
        private ELayoutMaster eLayoutMaster;

        public ELayoutMaster ELayoutMaster
        {
            get { return eLayoutMaster; }
            set { eLayoutMaster = value; }
        }

        private string _slideParent;

        public string SlideParent
        {
            get { return _slideParent; }
            set { _slideParent = value; }
        }


        private BitmapSource bitmapImage;

        public BitmapSource Thumbnail
        {
            get { return bitmapImage; }
            set { bitmapImage = value; }
        }

    }
}
