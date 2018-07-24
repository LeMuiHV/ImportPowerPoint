using INV.Elearning.Charts.ViewModel;
using INV.Elearning.Charts.Views;
using INV.Elearning.Controls;
using INV.Elearning.Core;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.DesignControl.Adorners;
using INV.Elearning.DesignControl.Models;
using INV.Elearning.DesignControl.ThemeBackground;
using INV.Elearning.DesignControl.UndoRedo;
using INV.Elearning.DesignControl.ViewModel;
using INV.Elearning.DesignControl.Views;
using INV.Elearning.EOPC;
using INV.Elearning.ImageProcess.Views;
using INV.Elearning.Text;
using INV.Elearning.Text.Models;
using INV.Elearning.Text.ViewModels.Text;
using INV.Elearning.Text.Views;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace INV.Elearning.DesignControl.Helper
{
    public static class CommandHelper
    {
        static DrawingPlaceHolderAdorner _drawingAdorner;
        static ColorBrushBase BackupBackground;
        static bool IsBackgroundSelect = false, isFontsSelect = false;
        static EFontfamily BackupFonts;

        #region PlaceHolderAddVideoCommand

        private static RelayCommand _placeHolderAddVideoCommand;
        /// <summary>
        /// Thêm Video trên PlaceHolder
        /// </summary>
        public static RelayCommand PlaceHolderAddVideoCommand
        {
            get { return _placeHolderAddVideoCommand ?? (_placeHolderAddVideoCommand = new RelayCommand(p => AddVideoExecute(p))); }
        }

        private static void AddVideoExecute(object p)
        {
            var values = (object[])p;
            var width = (double)values[0];
            var height = (double)values[1];
            var top = (double)values[2];
            var left = (double)values[3];
            PlaceHolder placeHolder = new PlaceHolder();
            if ((Application.Current as IAppGlobal).SelectedItem is PlaceHolder holder)
            {
                placeHolder = holder;
            }

            if (VideoElementControl.Helper.CommandHelper.AddVideoWithValue(width, height, top, left))
            {
                Global.BeginInit();
                (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Elements.Remove(placeHolder);
                Global.EndInit();
                Global.PushUndo(new PlaceHolderStep(placeHolder, (Application.Current as IAppGlobal).SelectedItem, (Application.Current as IAppGlobal).SelectedSlide));
            }
        }

        #endregion
        #region PlaceHolderAddTextCommand
        private static RelayCommand _placeHolderAddTextCommand;

        public static RelayCommand PlaceHolderAddTextCommand
        {
            get { return _placeHolderAddTextCommand ?? (_placeHolderAddTextCommand = new RelayCommand(p => PlaceHolderTextExecute(p))); }
        }

        private static void PlaceHolderTextExecute(object p)
        {
            var values = (object[])p;
            var width = (double)values[0];
            var height = (double)values[1];
            var top = (double)values[2];
            var left = (double)values[3];
            PlaceHolder placeHolder = new PlaceHolder();
            if ((Application.Current as IAppGlobal).SelectedItem is PlaceHolder holder)
            {
                placeHolder = holder;
            }
            TextEditor textEditor = new TextEditor() { IsSelected = true };

            if (textEditor.RichTextEditor?.TextContainer != null)
            {
                textEditor.RichTextEditor.TextContainer.Document = CreateData("", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Top, textEditor.RichTextEditor.TextContainer);
                textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                textEditor.Focusable = true;
                Keyboard.Focus(textEditor);
                textEditor.RichTextEditor.Focus();

            }
            if (textEditor.RichTextEditor?.TextContainer?.Document != null)
            {
                textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = TypeTextContainer.None;
                textEditor.Width = width;
                textEditor.Height = height;
                textEditor.Top = top;
                textEditor.Left = left;


            }
            Global.BeginInit();
            (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Elements.Add(textEditor);
            (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Elements.Remove(placeHolder);
            Global.EndInit();
            Global.PushUndo(new PlaceHolderStep(placeHolder, (Application.Current as IAppGlobal).SelectedItem, (Application.Current as IAppGlobal).SelectedSlide));
        }

        #endregion
        #region PlaceHolderAddImageCommand
        private static RelayCommand _placeHolderAddImageCommand;
        /// <summary>
        /// Thêm hình ảnh place holder
        /// </summary>
        public static RelayCommand PlaceHolderAddImageCommand
        {
            get { return _placeHolderAddImageCommand ?? (_placeHolderAddImageCommand = new RelayCommand(p => PlaceHolderAddImageExecute(p))); }
        }

        private static void PlaceHolderAddImageExecute(object p)
        {
            var values = (object[])p;
            var width = (double)values[0];
            var height = (double)values[1];
            var top = (double)values[2];
            var left = (double)values[3];

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Open Image";
            openFileDialog.Filter = "ImageFiles (*.png;*.jpg;*.jpeg;*.bmp;*.ico)|*.png;*.jpg;*.jpeg;*.bmp;*.ico|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                PlaceHolder placeHolder = new PlaceHolder();
                if ((Application.Current as IAppGlobal).SelectedItem is PlaceHolder holder)
                {
                    placeHolder = holder;
                }

                ImageContent imageContent = new ImageContent();
                imageContent.OpenImage(openFileDialog.FileName);
                imageContent.Width = width;
                imageContent.Height = height;
                imageContent.Top = top;
                imageContent.Left = left;
                Global.BeginInit();
                (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Elements.Add(imageContent);
                (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Elements.Remove(placeHolder);
                ObjectElementsHelper.UnSelectedAll();
                imageContent.IsSelected = true;
                Global.EndInit();
                Global.PushUndo(new PlaceHolderStep(placeHolder, (Application.Current as IAppGlobal).SelectedItem, (Application.Current as IAppGlobal).SelectedSlide));

            }
        }

        #endregion
        #region PlaceHolderAddChartCommand
        private static RelayCommand _placeHolderChartCommand;
        /// <summary>
        /// Add Chart Place Holder
        /// </summary>
        public static RelayCommand PlaceHolderChartCommand
        {
            get { return _placeHolderChartCommand ?? (_placeHolderChartCommand = new RelayCommand(p => PlaceHolderChartExecute(p))); }
        }

        private static void PlaceHolderChartExecute(object p)
        {
            var values = (object[])p;
            var width = (double)values[0];
            var height = (double)values[1];
            var top = (double)values[2];
            var left = (double)values[3];

            PlaceHolder placeHolder = new PlaceHolder();
            if ((Application.Current as IAppGlobal).SelectedItem is PlaceHolder holder)
            {
                placeHolder = holder;
            }

            //TChart chart = new TChart();
            //chart.Width = width;
            //chart.Height = height;
            //chart.Top = top;
            //chart.Left = left;
            var chart = TMainViewModel.AddChartClick(width, height, top, left);
            if (chart != null)
            {
                Global.BeginInit();
                (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Elements.Add(chart);
                (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Elements.Remove(placeHolder);
                ObjectElementsHelper.UnSelectedAll();
                chart.IsSelected = true;
                Global.EndInit();
                Global.PushUndo(new PlaceHolderStep(placeHolder, (Application.Current as IAppGlobal).SelectedItem, (Application.Current as IAppGlobal).SelectedSlide));
            }

        }

        #endregion
        #region SaveThemeCommand
        private static RelayCommand _saveThemeCommand;
        /// <summary>
        /// Lưu theme hiện tại
        /// </summary>
        public static RelayCommand SaveThemeCommand
        {
            get { return _saveThemeCommand ?? (_saveThemeCommand = new RelayCommand(p => SaveThemeExecute())); }
        }

        private static void SaveThemeExecute()
        {
            ThemesData themesData = new ThemesData();
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "ThemeFile (*.athmx)|*.athmx";
            if (saveFile.ShowDialog() == true)
            {
                themesData.Save((Application.Current as IAppGlobal).SelectedTheme, saveFile.FileName);
                INV.Elearning.Controls.MessageBox.ShowDialog("Save Themes Success", "Save Success");
            }
        }

        #endregion
        #region LoadThemeCommand
        private static RelayCommand _loadThemeCommand;
        /// <summary>
        /// Tải lại theme
        /// </summary>
        public static RelayCommand LoadThemeCommand
        {
            get { return _loadThemeCommand ?? (_loadThemeCommand = new RelayCommand(p => LoadThemeExecute())); }
        }

        private static void LoadThemeExecute()
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "ThemeFile (*.athmx)|*.athmx";
            if (openFile.ShowDialog() == true)
            {
                ETheme eTheme = new ETheme();
                ThemesData themesData = new ThemesData();
                eTheme = themesData.OpenETheme(openFile.FileName);
                eTheme.Name = Path.GetFileNameWithoutExtension(openFile.FileName);
                eTheme.TagName = "This Presenter";
                eTheme.ThemeName = Path.GetFileNameWithoutExtension(openFile.FileName);
                eTheme.ID = ObjectElementsHelper.RandomString(10);
                eTheme.IsLoaded = true;
                eTheme.SlideMasters[0].SlideName = eTheme.Name;
                eTheme.SlideMasters[0].ThemesName = eTheme.Name;
                SlideMaster slideMaster = new SlideMaster();
                slideMaster.LoadData(eTheme.SlideMasters[0]);
                slideMaster.UpdateThemeColor();
                //int count = 0;
                //foreach (ELayoutMaster eLayout in eTheme.SlideMasters[0].LayoutMasters)
                //{
                //    slideMaster.LayoutMasters[count].UpdateLayoutTheme();
                //    slideMaster.LayoutMasters[count].UpdateThemeColor();
                //    slideMaster.LayoutMasters[count].MainLayout.Fill = ColorHelper.ConverFromColorData(eTheme.SlideMasters[0].MainLayer.Background);
                //    eLayout.ThumbnailBitmap = DesignTabControlViewModel.CaptureLayout(slideMaster.LayoutMasters[count].MainLayout, new Size(341, 192));
                //    count++;
                //    count = Math.Min(slideMaster.LayoutMasters.Count - 1, count);
                //}
                (Application.Current as IAppGlobal).LocalThemesCollection.Add(eTheme);
                (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = eTheme;
                foreach (EColorManagment item in (Application.Current as IAppGlobal).OfficeColors)
                {
                    if (item.Name == eTheme.Colors.Name)
                    {
                        (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = item;
                    }
                }
                foreach (EFontfamily font in (Application.Current as IAppGlobal).OfficeFonts)
                {
                    if (font.Name == eTheme.SelectedFont.Name)
                    {
                        (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SelectedFont = font;
                    }
                }
                (Application.Current as IAppGlobal).SelectedThemeView.Colors = eTheme.Colors;
                (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont = eTheme.SelectedFont;

                (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters.Add(slideMaster);
                int index = (Application.Current as IAppGlobal).LocalThemesCollection.IndexOf(eTheme);
                if (index != -1)
                {
                    SlideMasterTabControlViewModel.ThemeDesign.SelectedIndex = index;
                }
                else
                {
                    SlideMasterTabControlViewModel.ThemeDesign.SelectedIndex = 0;
                }

                //(Application.Current as IAppGlobal).SelectedSlide.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
                //{
                //    SlideMaster slideMaster = new SlideMaster();
                //    slideMaster.LoadData((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].Clone() as ESlideMaster);
                //    slideMaster.MainLayout.Fill = ColorHelper.ConverFromColorData((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer.Background);
                //    slideMaster.ThemeLayout = new ThemeLayout();
                //    slideMaster.ThemeLayout.UpdateUI((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer);
                //    slideMaster.UpdateThemeColor();
                //    int count = 0;
                //    foreach (ELayoutMaster eLayout in (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].LayoutMasters)
                //    {
                //        slideMaster.LayoutMasters[count].UpdateLayoutTheme();
                //        slideMaster.LayoutMasters[count].UpdateThemeColor();
                //        slideMaster.LayoutMasters[count].MainLayout.Fill = ColorHelper.ConverFromColorData((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer.Background);
                //        eLayout.ThumbnailBitmap = DesignTabControlViewModel.CaptureLayout(slideMaster.LayoutMasters[count].MainLayout, new Size(341, 192));
                //        count++;
                //        count = Math.Min(slideMaster.LayoutMasters.Count - 1, count);
                //    }
                //    slideMaster = null;
                //    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.ThumbnailBitmap = new BitmapImage(new Uri(eTheme.Thumbnail.Path));

                //}));
            }
        }

        #endregion
        #region ColorCommand
        private static RelayCommand _colorCommand;
        /// <summary>
        /// Thay đổi Color
        /// </summary>
        public static RelayCommand ColorCommand
        {
            get { return _colorCommand ?? (_colorCommand = new RelayCommand(p => ColorExecute(p))); }
        }

        private static void ColorExecute(object p)
        {
            //if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster)
            //{
            //    ((Application.Current as IAppGlobal).SelectedSlide as SlideMaster).Color = p as EColorManagment;
            //}
            //else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster)
            //{
            //    ((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).SlideParent.Color = p as EColorManagment;
            //}
        }
        #endregion

        #region InsertMasterSlideCommand

        private static RelayCommand _InsertMasterSLide;
        /// <summary>
        /// Thêm Slide Master
        /// </summary>
        public static RelayCommand InsertMasterSlideCommand
        {
            get { return _InsertMasterSLide ?? (_InsertMasterSLide = new RelayCommand(p => InsertMasterSlideExecute())); }
        }

        private static void InsertMasterSlideExecute()
        {
            AddNewSlideMaster();
        }

        /// <summary>
        /// Thêm một SlideMaster vào Theme đang được chọn
        /// </summary>
        private static void AddNewSlideMaster()
        {
            Global.BeginInit();
            //Add Slide Master
            (Application.Current as IAppGlobal).SelectedThemeView.TempTheme = (Application.Current as IAppGlobal).SelectedTheme;
            SlideMaster slide = InsertSlideMaster();
            if (CheckDuplicateSlideMaster())
            {
                slide.SlideName = CountCustomSlideMaster().ToString() + "_Custom Design";
            }
            else
            {
                slide.SlideName = "Custom Design";
            }
            slide.IsCustomSlide = true;

            //Add Layout Master
            LayoutMaster layout = InsertTitleLayoutMaster();
            layout.IsCustomLayout = true;
            layout.SlideParent = slide;
            layout.LayoutType = LayoutType.Title;
            slide.LayoutMasters.Add(layout);

            layout = InsertTitleAndContentLayoutMaster(slide);
            layout.IsCustomLayout = true;
            layout.LayoutType = LayoutType.TitleContent;
            slide.LayoutMasters.Add(layout);

            layout = InsertSectionHeaderLayoutMaster(slide);
            layout.IsCustomLayout = true;
            layout.LayoutType = LayoutType.SectionHeader;
            slide.LayoutMasters.Add(layout);

            layout = InsertTwoContentLayoutMaster(slide);
            layout.IsCustomLayout = true;
            layout.LayoutType = LayoutType.TwoContent;
            slide.LayoutMasters.Add(layout);

            layout = InsertComparisionLayoutMaster(slide);
            layout.IsCustomLayout = true;
            layout.LayoutType = LayoutType.Comparision;
            slide.LayoutMasters.Add(layout);

            layout = InsertTitleOnlyLayoutMaster(slide);
            layout.IsCustomLayout = true;
            layout.LayoutType = LayoutType.TitleOnly;
            slide.LayoutMasters.Add(layout);

            layout = InsertBlankLayoutMaster(slide);
            layout.IsCustomLayout = true;
            layout.IsTitle = false;
            layout.LayoutType = LayoutType.Blank;
            slide.LayoutMasters.Add(layout);

            layout = InsertContentWithCaptionLayoutMaster(slide);
            layout.IsCustomLayout = true;
            layout.LayoutType = LayoutType.ContentWithCaption;
            slide.LayoutMasters.Add(layout);

            layout = InsertPictureWithCaptionLayoutMaster(slide);
            layout.IsCustomLayout = true;
            layout.LayoutType = LayoutType.PictureWithCaption;
            slide.LayoutMasters.Add(layout);


            slide.RefreshData();
            ETheme eTheme = new ETheme();
            eTheme.SlideMasters.Add(slide.Data as ESlideMaster);
            if (CheckDuplicateThemes())
            {
                eTheme.Name = CountCustomThemes().ToString() + "_Custom Design";
            }
            else
            {
                eTheme.Name = "Custom Design";
            }
            eTheme.TagName = "This Presenter";
            slide.ThemesName = eTheme.Name;
            eTheme.ThemeName = eTheme.Name;
            eTheme.Colors = (Application.Current as IAppGlobal).SelectedThemeView.Colors;
            eTheme.SelectedFont = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont;
            eTheme.ID = DesignTabControlViewModel.RandomString();
            slide.ID = eTheme.ID;
            (Application.Current as IAppGlobal).LocalThemesCollection.Add(eTheme);
            (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = eTheme;
            Global.EndInit();
            (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters.Add(slide);
            Global.BeginInit();
            slide.IsSelected = true;
            Global.EndInit();
            Global.PushUndo(new AddSlideMasterStep((Application.Current as IAppGlobal).SelectedSlide, slide, eTheme, (Application.Current as IAppGlobal).SelectedTheme));
        }



        private static bool CheckDuplicateSlideMaster()
        {
            foreach (SlideMaster item in (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters)
            {
                if (item.SlideName.Contains("Custom Design"))
                {
                    return true;
                }
            }
            return false;
        }

        private static int CountCustomSlideMaster()
        {
            int count = 0;
            foreach (SlideMaster item in (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters)
            {
                if (item.SlideName.Contains("Custom Design"))
                {
                    count++;
                }
            }
            return count;
        }

        private static bool CheckDuplicateThemes()
        {
            foreach (ETheme item in (Application.Current as IAppGlobal).LocalThemesCollection)
            {
                if (item.Name.Contains("Custom Design"))
                {
                    return true;
                }
            }
            return false;
        }

        private static int CountCustomThemes()
        {
            int count = 0;
            foreach (ETheme item in (Application.Current as IAppGlobal).LocalThemesCollection)
            {
                if (item.Name.Contains("Custom Design"))
                {
                    count++;
                }
            }
            return count;
        }

        /// <summary>
        /// Chèn một SlideMaster mới
        /// </summary>
        /// <returns></returns>
        private static SlideMaster InsertSlideMaster()
        {
            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;
            double _marginLeft = 80, _marginTop = 30, _verSpace = 20, _titleHeight = 120, _footerHeight = 36, _dateWidth = 200,
                _textHeight = _slideHeight - _titleHeight - _footerHeight - _marginTop - 3 * _verSpace,
                _textTop = _marginTop + _titleHeight + _verSpace,
                _footerTop = _slideHeight - _footerHeight - _verSpace,
                _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
                _footerLeft = _marginLeft + _dateWidth + _marginLeft,
                _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            SlideMaster slide = new SlideMaster() { SlideName = "Custom SlideMaster" };
            slide.MainLayout = new LayoutBase();
            var textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title" };
            textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
            textEditor.IsGraphicBackground = false;

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;


            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 36, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
            textEditor.PlaceHolderType = PlaceHolderEnum.Title;
            textEditor.RichTextEditor.TextContainer.InvalidateVisual();
            slide.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _textHeight, Top = _textTop, Left = _marginLeft, TargetName = "Sub Title" };
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master text styles", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Top, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.PlaceHolderType = PlaceHolderEnum.Subtitle;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Text;

            slide.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };
            textEditor.IsGraphicBackground = false;
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
            textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
            slide.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
            textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
            slide.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };
            textEditor.IsGraphicBackground = false;
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
            textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
            slide.MainLayout.Elements.Add(textEditor);

            return slide;
        }

        private static LayoutMaster InsertTitleLayoutMaster()
        {
            LayoutMaster layout = new LayoutMaster() { LayoutName = "Title Slide Layout", LayoutCount = 0 };
            layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);
            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;

            double _marginLeft = 80, _verSpace = 20, _footerHeight = 36, _dateWidth = 200,
               _footerTop = _slideHeight - _footerHeight - _verSpace,
               _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
               _footerLeft = _marginLeft + _dateWidth + _marginLeft,
               _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            layout.MainLayout = new LayoutBase();
            var textEditor = new TextEditor() { Width = (_slideWidth - 120 * 2), Height = 200, Top = 100, Left = 120, TargetName = "Title" };
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 36, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Bottom, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
            textEditor.PlaceHolderType = PlaceHolderEnum.Title;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = (_slideWidth - 120 * 2), Height = _slideHeight - 410, Top = 310, Left = 120, TargetName = "Sub Title" };
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master text styles", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Top, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Text;
            textEditor.PlaceHolderType = PlaceHolderEnum.Subtitle;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
            textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
            textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
            textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
            layout.MainLayout.Elements.Add(textEditor);

            return layout;
        }

        private static LayoutMaster InsertTitleAndContentLayoutMaster(SlideMaster slide)
        {
            LayoutMaster layout = new LayoutMaster() { SlideParent = slide, LayoutName = "Title Content Layout", LayoutCount = 0 };
            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;
            double _marginLeft = 80, _marginTop = 30, _verSpace = 20, _titleHeight = 120, _footerHeight = 36, _dateWidth = 200,
                _textHeight = _slideHeight - _titleHeight - _footerHeight - _marginTop - 3 * _verSpace,
                _textTop = _marginTop + _titleHeight + _verSpace,
                _footerTop = _slideHeight - _footerHeight - _verSpace,
                _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
                _footerLeft = _marginLeft + _dateWidth + _marginLeft,
                _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            layout.MainLayout = new Core.View.LayoutBase();
            layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);

            TextEditor textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 36, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
            textEditor.PlaceHolderType = PlaceHolderEnum.Title;
            layout.MainLayout.Elements.Add(textEditor);

            PlaceHolder placeHolder = new PlaceHolder() { Width = (_slideWidth - _marginLeft * 2), Height = _textHeight, Top = _textTop, Left = _marginLeft, TargetName = "Content", PlaceHolderType = PlaceHolderEnum.Normal };
            placeHolder.IsGraphicBackground = false;
            layout.MainLayout.Elements.Add(placeHolder);


            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
            textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
            textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
            textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
            layout.MainLayout.Elements.Add(textEditor);

            return layout;
        }

        private static LayoutMaster InsertSectionHeaderLayoutMaster(SlideMaster slide)
        {
            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;
            double _marginLeft = 80, _marginTop = 150, _verSpace = 20, _titleHeight = 200, _footerHeight = 36, _dateWidth = 200,
                _textHeight = _slideHeight - _titleHeight - _footerHeight - _marginTop - 3 * _verSpace,
                _textTop = _marginTop + _titleHeight + _verSpace,
                _footerTop = _slideHeight - _footerHeight - _verSpace,
                _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
                _footerLeft = _marginLeft + _dateWidth + _marginLeft,
                _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            LayoutMaster layout = new LayoutMaster() { SlideName = "Section Header Layout", SlideParent = slide, LayoutName = "Section Header Layout", LayoutCount = 0 };
            layout.MainLayout = new Core.View.LayoutBase();
            layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);

            TextEditor textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title" };
            Binding binding = new Binding()
            {
                Source = layout,
                Path = new PropertyPath("IsTitle"),
                Converter = new INV.Elearning.Core.Helper.BooleanToVisibility()
            };
            textEditor.SetBinding(TextEditor.VisibilityProperty, binding);
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 36, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Bottom, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
            textEditor.PlaceHolderType = PlaceHolderEnum.Title;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _textHeight, Top = _textTop, Left = _marginLeft, TargetName = "Section" };
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master text styles", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Top, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Text;
            textEditor.PlaceHolderType = PlaceHolderEnum.Subtitle;

            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
            textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
            textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
            textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
            layout.MainLayout.Elements.Add(textEditor);
            return layout;
        }

        private static LayoutMaster InsertTwoContentLayoutMaster(SlideMaster slide)
        {
            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;
            double _marginLeft = 80, _marginTop = 30, _verSpace = 20, _titleHeight = 120, _footerHeight = 36, _dateWidth = 200,
                _textHeight = _slideHeight - _titleHeight - _footerHeight - _marginTop - 3 * _verSpace,
                _textTop = _marginTop + _titleHeight + _verSpace,
                _footerTop = _slideHeight - _footerHeight - _verSpace,
                _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
                _footerLeft = _marginLeft + _dateWidth + _marginLeft,
                _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            LayoutMaster layout = new LayoutMaster() { SlideParent = slide, LayoutName = "Two Content Layout", LayoutCount = 0 };
            layout.MainLayout = new Core.View.LayoutBase();
            layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);

            TextEditor textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 36, (Application.Current as IAppGlobal).OfficeFonts[2].MajorFont, new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
            textEditor.PlaceHolderType = PlaceHolderEnum.Title;
            layout.MainLayout.Elements.Add(textEditor);

            PlaceHolder placeHolder = new PlaceHolder() { Width = (_slideWidth - _marginLeft * 2) / 2, Height = _textHeight, Top = _textTop, Left = _marginLeft, TargetName = "Content 1", Type = PlaceHolderType.Content, PlaceHolderType = PlaceHolderEnum.Normal };
            placeHolder.IsGraphicBackground = false;
            layout.MainLayout.Elements.Add(placeHolder);

            placeHolder = new PlaceHolder() { Width = (_slideWidth - _marginLeft * 2) / 2, Height = _textHeight, Top = _textTop, Left = _marginLeft + (_slideWidth - _marginLeft * 2) / 2, TargetName = "Content 2", Type = PlaceHolderType.Content, PlaceHolderType = PlaceHolderEnum.Normal };
            placeHolder.IsGraphicBackground = false;
            layout.MainLayout.Elements.Add(placeHolder);


            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
            textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
            textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
            textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
            layout.MainLayout.Elements.Add(textEditor);
            return layout;
        }

        private static LayoutMaster InsertComparisionLayoutMaster(SlideMaster slide)
        {
            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;
            double _marginLeft = 80, _marginTop = 30, _verSpace = 20, _titleHeight = 120, _footerHeight = 36, _dateWidth = 200, _contentHeight = 250,
                _textHeight = _slideHeight - _titleHeight - _footerHeight - _marginTop - 3 * _verSpace - _contentHeight,
                _textTop = _marginTop + _titleHeight + _verSpace,
                _contentTop = _textTop + _textHeight,
                _footerTop = _slideHeight - _footerHeight - _verSpace,
                _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
                _footerLeft = _marginLeft + _dateWidth + _marginLeft,
                _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            LayoutMaster layout = new LayoutMaster() { SlideParent = slide, LayoutName = "Comparison Layout", LayoutCount = 0 };
            layout.MainLayout = new Core.View.LayoutBase();
            layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);

            TextEditor textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 36, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
            textEditor.PlaceHolderType = PlaceHolderEnum.Title;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2) / 2, Height = _textHeight, Top = _textTop, Left = _marginLeft, TargetName = "Header 1" };
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master text styles", 24, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Bottom, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Text;
            textEditor.PlaceHolderType = PlaceHolderEnum.Subtitle;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2) / 2, Height = _textHeight, Top = _textTop, Left = _marginLeft + (_slideWidth - _marginLeft * 2) / 2, TargetName = "Header 2" };
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master text styles", 24, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Bottom, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Text;
            textEditor.PlaceHolderType = PlaceHolderEnum.Subtitle;
            layout.MainLayout.Elements.Add(textEditor);

            PlaceHolder placeHolder = new PlaceHolder() { Width = (_slideWidth - _marginLeft * 2) / 2, Height = _contentHeight, Top = _contentTop, Left = _marginLeft, TargetName = "Content 1", Type = PlaceHolderType.Content, PlaceHolderType = PlaceHolderEnum.Normal };
            placeHolder.IsGraphicBackground = false;
            layout.MainLayout.Elements.Add(placeHolder);

            placeHolder = new PlaceHolder() { Width = (_slideWidth - _marginLeft * 2) / 2, Height = _contentHeight, Top = _contentTop, Left = _marginLeft + (_slideWidth - _marginLeft * 2) / 2, TargetName = "Content 2", Type = PlaceHolderType.Content, PlaceHolderType = PlaceHolderEnum.Normal };
            placeHolder.IsGraphicBackground = false;
            layout.MainLayout.Elements.Add(placeHolder);


            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
            textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
            textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
            textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
            layout.MainLayout.Elements.Add(textEditor);
            return layout;
        }

        private static LayoutMaster InsertTitleOnlyLayoutMaster(SlideMaster slide)
        {
            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;
            double _marginLeft = 80, _marginTop = 30, _verSpace = 20, _titleHeight = 120, _footerHeight = 36, _dateWidth = 200, _contentHeight = 250,
                _textHeight = _slideHeight - _titleHeight - _footerHeight - _marginTop - 3 * _verSpace - _contentHeight,
                _textTop = _marginTop + _titleHeight + _verSpace,
                _contentTop = _textTop + _textHeight,
                _footerTop = _slideHeight - _footerHeight - _verSpace,
                _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
                _footerLeft = _marginLeft + _dateWidth + _marginLeft,
                _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            LayoutMaster layout = new LayoutMaster() { SlideParent = slide, LayoutName = "Title Only Layout", LayoutCount = 0 };
            layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);
            layout.MainLayout = new Core.View.LayoutBase();

            TextEditor textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 36, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
            textEditor.PlaceHolderType = PlaceHolderEnum.Title;
            layout.MainLayout.Elements.Add(textEditor);


            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
            textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
            textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
            textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
            layout.MainLayout.Elements.Add(textEditor);
            return layout;
        }

        private static LayoutMaster InsertBlankLayoutMaster(SlideMaster slide)
        {
            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;
            double _marginLeft = 80, _marginTop = 30, _verSpace = 20, _titleHeight = 120, _footerHeight = 36, _dateWidth = 200, _contentHeight = 250,
                _textHeight = _slideHeight - _titleHeight - _footerHeight - _marginTop - 3 * _verSpace - _contentHeight,
                _textTop = _marginTop + _titleHeight + _verSpace,
                _contentTop = _textTop + _textHeight,
                _footerTop = _slideHeight - _footerHeight - _verSpace,
                _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
                _footerLeft = _marginLeft + _dateWidth + _marginLeft,
                _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            LayoutMaster layout = new LayoutMaster() { SlideParent = slide, LayoutName = "Blank Layout", LayoutCount = 0 };
            layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);

            layout.IsTitle = false;
            layout.MainLayout = new Core.View.LayoutBase();

            TextEditor textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title", Visibility = Visibility.Hidden, IsShow = false };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 36, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
            textEditor.PlaceHolderType = PlaceHolderEnum.Title;
            layout.MainLayout.Elements.Add(textEditor);


            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
            textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

            textEditor.IsGraphicBackground = false;
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
            textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
            textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
            layout.MainLayout.Elements.Add(textEditor);
            return layout;
        }

        private static LayoutMaster InsertContentWithCaptionLayoutMaster(SlideMaster slide)
        {
            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;
            double _marginLeft = 80, _titleWidth = 320, _marginTop = 30, _verSpace = 20, _titleHeight = 150, _footerHeight = 36, _dateWidth = 200, _marginContent = 60,
                _textHeight = _slideHeight - _titleHeight - _footerHeight - _marginTop - 3 * _verSpace,
                _textTop = _marginTop + _titleHeight,
                _contentWidth = _slideWidth - 2 * _marginLeft - _titleWidth - _verSpace,
                _contentLeft = _titleWidth + _verSpace + _marginLeft,
                _contentHeight = _slideHeight - _marginContent - _footerHeight - 3 * _verSpace,
                _contentTop = _marginContent,
                _footerTop = _slideHeight - _footerHeight - _verSpace,
                _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
                _footerLeft = _marginLeft + _dateWidth + _marginLeft,
                _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            LayoutMaster layout = new LayoutMaster() { SlideParent = slide, LayoutName = "Content With Caption Layout", LayoutCount = 0 };
            layout.MainLayout = new Core.View.LayoutBase();
            layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);

            TextEditor textEditor = new TextEditor() { Width = _titleWidth, Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;
            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Bottom, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
            textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
            textEditor.PlaceHolderType = PlaceHolderEnum.Title;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _titleWidth, Height = _textHeight, Top = _textTop, Left = _marginLeft, TargetName = "Sub Title" };
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master text styles", 16, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Top, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Text;
            textEditor.PlaceHolderType = PlaceHolderEnum.Subtitle;
            layout.MainLayout.Elements.Add(textEditor);

            PlaceHolder placeHolder = new PlaceHolder() { Width = _contentWidth, Height = _contentHeight, Left = _contentLeft, Top = _contentTop, Type = PlaceHolderType.Content, TargetName = "Content", PlaceHolderType = PlaceHolderEnum.Normal };
            placeHolder.IsGraphicBackground = false;
            layout.MainLayout.Elements.Add(placeHolder);


            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
            textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
            textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
            layout.MainLayout.Elements.Add(textEditor);


            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
            textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
            layout.MainLayout.Elements.Add(textEditor);
            return layout;
        }

        private static LayoutMaster InsertPictureWithCaptionLayoutMaster(SlideMaster slide)
        {
            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;
            double _marginLeft = 80, _titleWidth = 320, _marginTop = 30, _verSpace = 20, _titleHeight = 150, _footerHeight = 36, _dateWidth = 200, _marginContent = 60,
                _textHeight = _slideHeight - _titleHeight - _footerHeight - _marginTop - 3 * _verSpace,
                _textTop = _marginTop + _titleHeight,
                _contentWidth = _slideWidth - 2 * _marginLeft - _titleWidth - _verSpace,
                _contentLeft = _titleWidth + _verSpace + _marginLeft,
                _contentHeight = _slideHeight - _marginContent - _footerHeight - 3 * _verSpace,
                _contentTop = _marginContent,
                _footerTop = _slideHeight - _footerHeight - _verSpace,
                _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
                _footerLeft = _marginLeft + _dateWidth + _marginLeft,
                _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            LayoutMaster layout = new LayoutMaster() { SlideParent = slide, LayoutName = "Picture With Caption Layout", LayoutCount = 0 };
            layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);

            layout.MainLayout = new Core.View.LayoutBase();

            TextEditor textEditor = new TextEditor() { Width = _titleWidth, Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Bottom, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
            textEditor.PlaceHolderType = PlaceHolderEnum.Title;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _titleWidth, Height = _textHeight, Top = _textTop, Left = _marginLeft, TargetName = "Sub Title" };
            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master text styles", 16, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Top, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Text;
            textEditor.PlaceHolderType = PlaceHolderEnum.Subtitle;
            layout.MainLayout.Elements.Add(textEditor);

            PlaceHolder placeHolder = new PlaceHolder() { Width = _contentWidth, Height = _contentHeight, Left = _contentLeft, Top = _contentTop, Type = PlaceHolderType.Picture, TargetName = "Picture", PlaceHolderType = PlaceHolderEnum.Normal };
            placeHolder.IsGraphicBackground = false;
            layout.MainLayout.Elements.Add(placeHolder);


            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
            textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
            layout.MainLayout.Elements.Add(textEditor);

            textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
            textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
            layout.MainLayout.Elements.Add(textEditor);


            textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

            textEditor.Thickness = 1;
            textEditor.DashType = Controls.Enums.DashType.DashDot;

            textEditor.IsGraphicBackground = false;
            textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
            textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
            textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
            textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
            textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
            textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
            layout.MainLayout.Elements.Add(textEditor);
            return layout;
        }
        public static Document CreateData(string Text, double fontSize, TextContainer textContainer, HorizontalAlign align, VerticalAlign verticalAlign = VerticalAlign.Top)
        {
            Document _result = new Document();
            _result.IsQuestion = false;
            _result.Padding = new EThickness() { Top = 5, Bottom = 5, Left = 5, Right = 5 };
            _result.VerticalAlign = verticalAlign;
            Paragraph _para = new Paragraph();
            Run _run = new Run();
            _run.Fontfamily = "Arial";
            _run.FontSize = fontSize;
            _run.FontStyle = INV.Elearning.Text.FontStyle.Normal;
            _run.Foreground = new ColorBrushBase() { Brush = Brushes.Black };
            _run.Text = Text + "\r";
            _run.StrikeThrough = new StrikeThrough();
            _run.StrikeThrough = StrikeThrough.None;
            _run.Underline = new Underline();
            _run.Underline.Style = UnderlineStyle.None;
            _run.Underline.Color = null;
            _run.Parent = _para;
            _para.Inlines.Add(_run);
            _para.TextIndent = 0;
            _para.Parent = _result;
            _para.LineHeight = 1;
            _para.TextAlign = align;
            _para.Margin = new EThickness() { Top = 0, Left = 0 };
            _result.Blocks.Add(_para);
            textContainer.Selection.Start = textContainer.Selection.End = new TextPointer(_run, 0);

            return _result;

        }

        /// <summary>
        /// Tạo nội dung mẫu cho khung
        /// </summary>
        /// <param name="content"></param>
        /// <param name="fontSize"></param>
        /// <param name="fontFamily"></param>
        /// <param name="foreground"></param>
        /// <param name="align"></param>
        /// <param name="verticalAlign"></param>
        /// <param name="textContainer"></param>
        /// <returns></returns>
        public static Document CreateData(string content, double fontSize, string fontFamily, ColorBrushBase foreground, HorizontalAlign align, VerticalAlign verticalAlign, TextContainer textContainer)
        {
            Document _result = new Document();
            _result.IsQuestion = false;
            _result.Padding = new EThickness() { Top = 5, Bottom = 5, Left = 5, Right = 5 };
            _result.VerticalAlign = verticalAlign;
            Paragraph _para = new Paragraph();
            Run _run = new Run();
            _run.Fontfamily = fontFamily;
            _run.FontSize = fontSize;
            _run.FontStyle = Text.FontStyle.Normal;
            _run.Foreground = foreground;
            _run.Text = content + "\r";
            _run.StrikeThrough = StrikeThrough.None;
            _run.Underline = new Underline();
            _run.Underline.Style = UnderlineStyle.None;
            _run.Parent = _para;
            _para.Inlines.Add(_run);
            _para.TextIndent = 0;
            _para.Parent = _result;
            _para.LineHeight = 1;
            _para.TextAlign = align;
            _para.Margin = new EThickness();
            _result.Blocks.Add(_para);
            textContainer.Selection.Start = textContainer.Selection.End = new TextPointer(_run, 0);

            return _result;

        }


        public static Document SetTextDocument(Document document, string content, double fontSize, string fontFamily, ColorBrushBase foreground, HorizontalAlign align, VerticalAlign verticalAlign, TextContainer textContainer)
        {
            if (document?.Blocks?.Count > 0)
                for (int i = 0; i < document.Blocks.Count; i++)
                {
                    document.Blocks.RemoveAt(i);
                    i--;
                }

            document.IsQuestion = false;
            document.Padding = new EThickness() { Top = 5, Bottom = 5, Left = 5, Right = 5 };
            document.VerticalAlign = verticalAlign;
            Paragraph _para = new Paragraph();
            Run _run = new Run();
            _run.Fontfamily = fontFamily;
            _run.FontSize = fontSize;
            _run.FontStyle = Text.FontStyle.Normal;
            _run.Foreground = foreground;
            _run.Text = content + "\r";
            _run.StrikeThrough = StrikeThrough.None;
            _run.Underline = new Underline();
            _run.Underline.Style = UnderlineStyle.None;
            _run.Parent = _para;
            _para.Inlines.Add(_run);
            _para.TextIndent = 0;
            _para.Parent = document;
            _para.LineHeight = 1;
            _para.TextAlign = align;
            _para.Margin = new EThickness();
            document.Blocks.Add(_para);
            textContainer.Selection.Start = textContainer.Selection.End = new TextPointer(_run, 0);

            return document;

        }
        public static string GetTextDocument(RichTextEditor richTextEditor)
        {
            string result = "";
            if (richTextEditor?.TextContainer?.Words?.Count > 0)
            {
                int _countWord = richTextEditor.TextContainer.Words.Count;
                for (int i = 0; i < _countWord; i++)
                {
                    if (!(richTextEditor.TextContainer.Words[i] is WordBullet))
                    {
                        if (richTextEditor.TextContainer.Words[i] is WordRun) result += (richTextEditor.TextContainer.Words[i] as WordRun).Text;
                        else
                            result += " ";
                    }
                }
            }
            return result;
        }
        #endregion

        #region InsertMasterLayoutCommand
        private static RelayCommand _insertLayoutCommand;
        /// <summary>
        /// Thêm layout trong Slide Master
        /// </summary>
        public static RelayCommand InsertLayoutCommand
        {
            get { return _insertLayoutCommand ?? (_insertLayoutCommand = new RelayCommand(p => InsertLayoutExecute())); }
        }

        private static void InsertLayoutExecute()
        {

            double _slideWidth = (Application.Current as IAppGlobal).SlideSize.Width;
            double _slideHeight = (Application.Current as IAppGlobal).SlideSize.Height;
            double _marginLeft = 80, _marginTop = 30, _verSpace = 20, _titleHeight = 120, _footerHeight = 36, _dateWidth = 200, _contentHeight = 250,
                _textHeight = _slideHeight - _titleHeight - _footerHeight - _marginTop - 3 * _verSpace - _contentHeight,
                _textTop = _marginTop + _titleHeight + _verSpace,
                _contentTop = _textTop + _textHeight,
                _footerTop = _slideHeight - _footerHeight - _verSpace,
                _footerWidth = _slideWidth - 4 * _marginLeft - 2 * _dateWidth,
                _footerLeft = _marginLeft + _dateWidth + _marginLeft,
                _pageLeft = _slideWidth - _marginLeft - _dateWidth;

            if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster slideMaster)
            {
                LayoutMaster layout = new LayoutMaster() { SlideParent = slideMaster, LayoutName = "Custom Layout" };
                layout.LayoutType = LayoutType.Custom;
                layout.MainLayout = new LayoutBase();

                if (CheckDuplicateName(slideMaster))
                {
                    layout.LayoutName = CountCustomLayout(slideMaster).ToString() + "_Custom Layout";
                }
                layout.LayoutCount = 0;
                layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);
                TextEditor textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title" };

                textEditor.Thickness = 1;
                textEditor.DashType = Controls.Enums.DashType.DashDot;

                textEditor.IsGraphicBackground = false;
                textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 36, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
                textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
                textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
                textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
                textEditor.PlaceHolderType = PlaceHolderEnum.Title;
                layout.MainLayout.Elements.Add(textEditor);


                textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

                textEditor.Thickness = 1;
                textEditor.DashType = Controls.Enums.DashType.DashDot;

                textEditor.IsGraphicBackground = false;
                textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
                textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
                textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
                textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
                layout.MainLayout.Elements.Add(textEditor);

                textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

                textEditor.Thickness = 1;
                textEditor.DashType = Controls.Enums.DashType.DashDot;

                textEditor.IsGraphicBackground = false;
                textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
                textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
                textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
                textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
                layout.MainLayout.Elements.Add(textEditor);

                textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

                textEditor.Thickness = 1;
                textEditor.DashType = Controls.Enums.DashType.DashDot;

                textEditor.IsGraphicBackground = false;
                textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
                textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
                textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
                textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
                layout.MainLayout.Elements.Add(textEditor);
                layout.IsSelected = true;
                Global.PushUndo(new AddLayoutMasterStep(layout, slideMaster));
                slideMaster.LayoutMasters.Add(layout);
                (Application.Current as IAppGlobal).SelectedThemeView.RefreshData();
                (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0] = (Application.Current as IAppGlobal).SelectedThemeView.Data.SlideMasters[0];
            }
            else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster layoutMaster)
            {
                LayoutMaster layout = new LayoutMaster() { SlideParent = layoutMaster.SlideParent, LayoutName = "Custom Layout" };
                layout.LayoutType = LayoutType.Custom;
                if (CheckDuplicateName(layoutMaster.SlideParent))
                {
                    layout.LayoutName = CountCustomLayout(layoutMaster.SlideParent).ToString() + "_Custom Layout";
                }
                layout.LayoutCount = 0;
                layout.SlideName = string.Format("{0}: uses by {1} slides", layout.LayoutName, layout.LayoutCount);

                layout.MainLayout = new LayoutBase();
                TextEditor textEditor = new TextEditor() { Width = (_slideWidth - _marginLeft * 2), Height = _titleHeight, Top = _marginTop, Left = _marginLeft, TargetName = "Title" };

                textEditor.Thickness = 1;
                textEditor.DashType = Controls.Enums.DashType.DashDot;

                textEditor.IsGraphicBackground = false;
                textEditor.RichTextEditor.TextContainer.Document = CreateData("Click to edit Master title styles", 36, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MajorFont;
                textEditor.RichTextEditor.TextContainer.Document.IsTitle = true;
                textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.Title;
                textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Title;
                textEditor.PlaceHolderType = PlaceHolderEnum.Title;
                layout.MainLayout.Elements.Add(textEditor);


                textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _marginLeft, TargetName = "DateTime" };

                textEditor.Thickness = 1;
                textEditor.DashType = Controls.Enums.DashType.DashDot;

                textEditor.IsGraphicBackground = false;
                textEditor.RichTextEditor.TextContainer.Document = CreateData(DateTime.Now.ToShortDateString(), 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
                textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
                textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.DateTime;
                textEditor.PlaceHolderType = PlaceHolderEnum.DateTime;
                layout.MainLayout.Elements.Add(textEditor);

                textEditor = new TextEditor() { Width = _footerWidth, Height = _footerHeight, Top = _footerTop, Left = _footerLeft, TargetName = "Footer" };

                textEditor.Thickness = 1;
                textEditor.DashType = Controls.Enums.DashType.DashDot;

                textEditor.IsGraphicBackground = false;
                textEditor.RichTextEditor.TextContainer.Document = CreateData("Footer", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
                textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
                textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
                textEditor.PlaceHolderType = PlaceHolderEnum.Footer;
                layout.MainLayout.Elements.Add(textEditor);

                textEditor = new TextEditor() { Width = _dateWidth, Height = _footerHeight, Top = _footerTop, Left = _pageLeft, TargetName = "Slide Number" };

                textEditor.Thickness = 1;
                textEditor.DashType = Controls.Enums.DashType.DashDot;
                textEditor.IsGraphicBackground = false;
                textEditor.RichTextEditor.TextContainer.Document = CreateData("<#>", 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Right, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont.MinorFont;
                textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
                textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.SlideNumber;
                textEditor.PlaceHolderType = PlaceHolderEnum.SlideNumber;
                layout.MainLayout.Elements.Add(textEditor);
                layout.IsSelected = true;
                int index = layoutMaster.SlideParent.LayoutMasters.IndexOf(layoutMaster);
                Global.PushUndo(new AddLayoutMasterIndexStep(layout, layoutMaster.SlideParent, index));

                layoutMaster.SlideParent.LayoutMasters.Insert(index + 1, layout);
                (Application.Current as IAppGlobal).SelectedThemeView.RefreshData();
                (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0] = (Application.Current as IAppGlobal).SelectedThemeView.Data.SlideMasters[0];
            }
        }

        private static bool CheckDuplicateName(SlideMaster slideMaster)
        {
            if (slideMaster != null)
            {
                foreach (LayoutMaster item in slideMaster.LayoutMasters)
                {
                    if (item.SlideName.Contains("Custom Layout"))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static int CountCustomLayout(SlideMaster slideMaster)
        {
            int count = 0;
            foreach (LayoutMaster item in slideMaster.LayoutMasters)
            {
                if (item.SlideName?.Contains("Custom Layout") == true)
                {
                    count++;
                }
            }
            return count;
        }

        #endregion

        #region DeleteLayoutCommand
        private static RelayCommand _deleteLayoutCommand;

        /// <summary>
        /// Xóa layout trong slide Master
        /// </summary>
        public static RelayCommand DeleteLayoutCommand
        {
            get { return _deleteLayoutCommand ?? (_deleteLayoutCommand = new RelayCommand(p => DeleteLayoutExecute(), o => DeleteCanExecute())); }
        }

        private static bool DeleteCanExecute()
        {
            if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster)
            {
                return true;
            }
            else
            {
                if ((Application.Current as IAppGlobal).LocalThemesCollection.Count > 1)
                {
                    return true;
                }
                else return false;
            }
        }

        private static void DeleteLayoutExecute()
        {
            if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster layoutMaster)
            {
                string idLayoutMaster = layoutMaster.ID;
                int index = (layoutMaster.SlideParent.LayoutMasters.IndexOf(layoutMaster));
                Global.PushUndo(new DeleteLayoutMasterStep(layoutMaster, index));
                Global.BeginInit();
                layoutMaster.SlideParent.LayoutMasters.Remove(layoutMaster);
                (Application.Current as IAppGlobal).DocumentControl.Slides.Remove(layoutMaster);
                for (int i = 0; i < (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters.Count; i++)
                {
                    if (idLayoutMaster == (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters[i].ID)
                    {
                        (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters.Remove((Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters[i]);
                        i--;
                    }
                }
                Global.EndInit();
            }
            else if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster slideMaster)
            {
                int index = (Application.Current as IAppGlobal).LocalThemesCollection.IndexOf((Application.Current as IAppGlobal).SelectedTheme);
                ESlideMaster cloneSlide = new ESlideMaster();
                slideMaster.RefreshData();
                cloneSlide = slideMaster.Data.Clone() as ESlideMaster;
                (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters.Remove(slideMaster);
                //Global.PushUndo(new DeleteSlideMasterStep(slideMaster, cloneSlide, (Application.Current as IAppGlobal).SelectedTheme, index));
                Global.BeginInit();
                DesignTabControlViewModel.DisposeSlideMaster(slideMaster);
                (Application.Current as IAppGlobal).LocalThemesCollection.Remove((Application.Current as IAppGlobal).SelectedTheme);
                foreach (LayoutMaster item in slideMaster.LayoutMasters)
                {
                    (Application.Current as IAppGlobal).DocumentControl.Slides.Remove(item);
                }
                 (Application.Current as IAppGlobal).DocumentControl.Slides.Remove(slideMaster);
                slideMaster = null;
                Global.EndInit();
            }

        }

        #endregion

        #region RenameCommand
        private static RelayCommand _renameCommand;
        /// <summary>
        /// Thay đổi tên của Layout hoặc Slide trong Slide master
        /// </summary>
        public static RelayCommand RenameCommand
        {
            get { return _renameCommand ?? (_renameCommand = new RelayCommand(p => RenameExecute())); }
        }

        private static void RenameExecute()
        {
            if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster slideMaster)
            {
                RenameWindow rename = new RenameWindow(slideMaster.SlideName);
                rename.Owner = (Application.Current.MainWindow);
                rename.ShowDialog();
            }
            else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster layoutMaster)
            {
                RenameWindow rename = new RenameWindow(layoutMaster.SlideName);
                rename.Owner = (Application.Current.MainWindow);
                rename.ShowDialog();
            }
        }

        #endregion

        #region PreserveCommand
        private static RelayCommand _preserveCommand;
        /// <summary>
        /// Ghim Slide master
        /// </summary>
        public static RelayCommand PreserveCommand
        {
            get { return _preserveCommand ?? (_preserveCommand = new RelayCommand(p => PreserveExecute(), o => PreserveCanExecute())); }
        }

        private static bool PreserveCanExecute()
        {
            return (Application.Current as IAppGlobal).SelectedSlide is SlideMaster;
        }

        private static void PreserveExecute()
        {
           // InsertThemeBackground(GenerateThemeShape.Theme10((Application.Current as IAppGlobal).OfficeColors[2]), (Application.Current as IAppGlobal).SelectedSlide as SlideMaster, (Application.Current as IAppGlobal).OfficeColors[2]);
        }

        #endregion

        #region MasterLayoutCommand

        private static RelayCommand _masterLayoutCommand;
        /// <summary>
        /// Chỉnh sửa thông số của Slide Master
        /// </summary>
        public static RelayCommand MasterLayoutCommand
        {
            get { return _masterLayoutCommand ?? (_masterLayoutCommand = new RelayCommand(p => MasterLayoutExecute(), o => MasterLayoutCanExecute())); }
        }

        private static bool MasterLayoutCanExecute()
        {
            return (Application.Current as IAppGlobal).SelectedSlide is SlideMaster;
        }

        private static void MasterLayoutExecute()
        {
            MasterLayoutConfig masterLayoutConfig = new MasterLayoutConfig((Application.Current as IAppGlobal).SelectedSlide as SlideMaster);
            masterLayoutConfig.Owner = (Application.Current.MainWindow);
            masterLayoutConfig.ShowDialog();
        }


        #endregion

        #region IsInsertPlaceHolder
        private static RelayCommand _isInsertPlaceHolder;

        public static RelayCommand IsInsertPlaceHolder
        {
            get { return _isInsertPlaceHolder ?? (_isInsertPlaceHolder = new RelayCommand(p => IsInsertPlaceHolderExecute(), o => IsInsertPlaceHolderCanExecute())); }
        }

        private static bool IsInsertPlaceHolderCanExecute()
        {
            return ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster);
        }

        private static void IsInsertPlaceHolderExecute()
        {

        }

        #endregion

        #region InsertPlaceHolder
        private static RelayCommand _insertPlaceHolderComand;
        /// <summary>
        /// Thêm placeHolder
        /// </summary>
        public static RelayCommand InsertPlaceHolderCommand
        {
            get { return _insertPlaceHolderComand ?? (_insertPlaceHolderComand = new RelayCommand(p => InsertPlaceHolderExecute(p), o => InsertPlaceHolderCanExecute(o))); }
        }

        private static bool InsertPlaceHolderCanExecute(object o)
        {
            return ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster);
        }

        private static void InsertPlaceHolderExecute(object p)
        {
            PlaceHolderType placeHolderType = PlaceHolderType.Content;
            switch ((p as PlaceHolderItem).Name)
            {
                case "Content":
                    placeHolderType = PlaceHolderType.Content;
                    break;
                case "Video":
                    placeHolderType = PlaceHolderType.Video;
                    break;
                case "Text":
                    placeHolderType = PlaceHolderType.Text;
                    break;
                case "Picture":
                    placeHolderType = PlaceHolderType.Picture;
                    break;
                case "Chart":
                    placeHolderType = PlaceHolderType.Chart;
                    break;
                default:
                    break;
            }
            LayoutMaster layoutMaster = (Application.Current as IAppGlobal).SelectedSlide as LayoutMaster;
            var adornerLayer = System.Windows.Documents.AdornerLayer.GetAdornerLayer(layoutMaster.MainLayout);
            if (_drawingAdorner == null)
            {
                _drawingAdorner = new DrawingPlaceHolderAdorner(layoutMaster.MainLayout);
                _drawingAdorner.PlaceHolderType = placeHolderType;
                _drawingAdorner.ActionCompleted += _drawingAdorner_ActionCompleted;
                adornerLayer.Add(_drawingAdorner);
            }
        }

        private static void _drawingAdorner_ActionCompleted(object sender, PlaceHolderEventArg e)
        {
            if (e.Rectangle != null)
            {
                Point _position = new Point(e.Position.Left, e.Position.Top);
                _position = (sender as System.Windows.Documents.Adorner).TranslatePoint(_position, ((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).MainLayout);
                PlaceHolder placeHolder = new PlaceHolder() { Type = e.PlaceHolderType, Width = e.Position.Width, Height = e.Position.Height, Top = _position.Y, Left = _position.X, IsGraphicBackground = false, PlaceHolderType = PlaceHolderEnum.Normal };
                ((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).MainLayout.Elements.Add(placeHolder);
                var _adornerLayer = System.Windows.Documents.AdornerLayer.GetAdornerLayer(((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).MainLayout);
                if (_adornerLayer != null)
                {
                    _adornerLayer.Remove(sender as System.Windows.Documents.Adorner);
                }
                placeHolder.IsSelected = true;
                placeHolder.RefreshData();
            }
            else
            {
                var _adornerLayer = System.Windows.Documents.AdornerLayer.GetAdornerLayer(((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).MainLayout);
                if (_adornerLayer != null)
                {
                    _adornerLayer.Remove(sender as System.Windows.Documents.Adorner);
                }
            }
            _drawingAdorner = null;
        }


        #endregion

        #region TitleCommand
        private static RelayCommand _titleCommand;
        /// <summary>
        /// Check xem có hiện khung title
        /// </summary>
        public static RelayCommand TitleCommand
        {
            get { return _titleCommand ?? (_titleCommand = new RelayCommand(p => TitleExecute(p), o => TitleCanExecute(o))); }
        }

        private static bool TitleCanExecute(object o)
        {
            return (Application.Current as IAppGlobal).SelectedSlide is LayoutMaster;
        }

        private static void TitleExecute(object p)
        {
            ((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).IsTitle = (bool)p;
        }

        #endregion

        #region FootersCommand
        private static RelayCommand _footersCommand;
        /// <summary>
        /// Check xem có hiện khung footer hay không
        /// </summary>
        public static RelayCommand FootersCommand
        {
            get { return _footersCommand ?? (_footersCommand = new RelayCommand(p => FootersExecute(p), o => FootersCanExecute(o))); }
        }

        private static bool FootersCanExecute(object o)
        {
            return (Application.Current as IAppGlobal).SelectedSlide is LayoutMaster;
        }

        private static void FootersExecute(object p)
        {
            ((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).IsFooter = (bool)p;
        }

        #endregion

        #region FontsCommand
        private static RelayCommand _fontsCommand;
        /// <summary>
        /// Thay đổi font
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
                (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont = eFontfamily;
                SlideMaster slide = null;
                if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster slidemaster)
                {
                    slide = slidemaster;
                }
                else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster layoutMaster)
                {
                    slide = layoutMaster.SlideParent;
                }
                if (slide != null) slide.SelectedFont = eFontfamily;
            }
        }

        #endregion

        #region PreviewFontsCommand
        private static RelayCommand _previewFontsCommand;

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

        #region DropDownFontCommand
        private static RelayCommand _dropDownFontsCommand;

        public static RelayCommand DropdownFontsCommand
        {
            get { return _dropDownFontsCommand ?? (_dropDownFontsCommand = new RelayCommand(p => DropdownFontsExecute())); }
        }

        private static void DropdownFontsExecute()
        {
            BackupFonts = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont.Clone() as EFontfamily;
        }

        #endregion
        #region BackgroundStylesCommand
        private static RelayCommand _backgroundStyleCommand;
        /// <summary>
        /// Thay đổi background style
        /// </summary>
        public static RelayCommand BackgroundStyleCommand
        {
            get { return _backgroundStyleCommand ?? (_backgroundStyleCommand = new RelayCommand(p => BackgroundStylesExecute(p))); }
        }

        private static void BackgroundStylesExecute(object p)
        {
            IsBackgroundSelect = true;
            BackgroundItem backgroundItem = p as BackgroundItem;
            (Application.Current as IAppGlobal).SelectedThemeView.Background = backgroundItem;

        }

        #endregion

        #region OpenFormatBackgroundWindow
        private static RelayCommand _openFormatBackgroundCommand;

        public static RelayCommand OpenFormatBackgroundCommand
        {
            get { return _openFormatBackgroundCommand ?? (_openFormatBackgroundCommand = new RelayCommand(p => OpenFormatBackgroundExecute())); }
        }

        private static void OpenFormatBackgroundExecute()
        {
            (Application.Current as IAppGlobal).SlideBackgroundFormatCommand.Execute(null);
        }

        #endregion

        #region PreviewBackgroundCommand
        private static RelayCommand _previewBackgroundCommand;
        /// <summary>
        /// Preview Background Command
        /// </summary>
        public static RelayCommand PreviewBackgroundCommand
        {
            get { return _previewBackgroundCommand ?? (_previewBackgroundCommand = new RelayCommand(p => PreviewBackgroundExecute(p))); }
        }

        private static void PreviewBackgroundExecute(object p)
        {
            Global.BeginInit("bg");
            if (p is GalleryItem galleryItem)
            {
                if (galleryItem.Content is BackgroundItem backgroundItem)
                {
                    ColorGradientBrush gradientColor = new ColorGradientBrush();
                    ObservableCollection<CustomGradientStop> listCustom = new ObservableCollection<CustomGradientStop>();
                    if (backgroundItem.ColorStyle is LinearGradientBrush)
                    {
                        foreach (CustomGradientStop item in backgroundItem.ColorDataStyle.GradientStops)
                        {
                            foreach (GradientStop item1 in (backgroundItem.ColorStyle as LinearGradientBrush).GradientStops)
                            {
                                if (backgroundItem.ColorDataStyle.GradientStops.IndexOf(item) == (backgroundItem.ColorStyle as LinearGradientBrush).GradientStops.IndexOf(item1))
                                {
                                    listCustom.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = item1.Color.ToString(), Name = item.Color.Name, Brightness = 1 }, Offset = item.Offset });
                                }
                            }
                        }
                    }
                    else if (backgroundItem.ColorStyle is RadialGradientBrush)
                    {
                        gradientColor.GradientType = GradientColorType.Radial;
                        foreach (CustomGradientStop item in backgroundItem.ColorDataStyle.GradientStops)
                        {
                            foreach (GradientStop item1 in (backgroundItem.ColorStyle as RadialGradientBrush).GradientStops)
                            {
                                if (backgroundItem.ColorDataStyle.GradientStops.IndexOf(item) == (backgroundItem.ColorStyle as RadialGradientBrush).GradientStops.IndexOf(item1))
                                {
                                    listCustom.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = item1.Color.ToString(), Name = item.Color.Name, Brightness = 1 }, Offset = item.Offset });
                                }
                            }
                        }
                    }



                    gradientColor.GradientStops = listCustom;
                    backgroundItem.ColorDataStyle = gradientColor;
                    (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Fill = backgroundItem.ColorDataStyle;
                }
            }
            Global.EndInit("bg");
        }

        #endregion

        #region CancelPreviewBackgroundCommand
        private static RelayCommand _cancelPreviewBackgroundCommand;
        /// <summary>
        /// Cancel preview Background command
        /// </summary>
        public static RelayCommand CancelPreviewBackgroundCommand
        {
            get { return _cancelPreviewBackgroundCommand ?? (_cancelPreviewBackgroundCommand = new RelayCommand(p => CancelPreviewBackgroundExecute(p))); }
        }

        private static void CancelPreviewBackgroundExecute(object p)
        {
            Global.BeginInit("ca");
            if (BackupBackground != null)
            {
                if (!IsBackgroundSelect)
                {
                    (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Fill = BackupBackground;
                }
                else
                {
                    IsBackgroundSelect = false;
                }
            }
            Global.EndInit("ca");
        }

        #endregion

        #region DropDownBackgroundCommand
        private static RelayCommand _dropDownBackgroundCommand;
        /// <summary>
        /// Dropdown Background Execute
        /// </summary>
        public static RelayCommand DropDownBackgroundCommand
        {
            get { return _dropDownBackgroundCommand ?? (_dropDownBackgroundCommand = new RelayCommand(p => DropDownBackgroundExecute())); }
        }

        private static void DropDownBackgroundExecute()
        {
            BackupBackground = (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Fill;
        }

        #endregion

        #region HideBackgroundCommand
        private static RelayCommand _hideBackgroundCommand;
        /// <summary>
        /// Kiểm tra xem có ẩn hiện background hay ko
        /// </summary>
        public static RelayCommand HideBackgroundCommand
        {
            get { return _hideBackgroundCommand ?? (_hideBackgroundCommand = new RelayCommand(p => HideBackgroundExecute(p), o => HideBackgroundCanExecute(o))); }
        }

        private static bool HideBackgroundCanExecute(object o)
        {
            return ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster);
        }

        private static void HideBackgroundExecute(object p)
        {
            (Application.Current as IAppGlobal).SelectedSlide.IsHideBackground = (bool)p;

        }

        #endregion

        #region CloseMasterViewCommand
        private static RelayCommand _closeMasterViewCommand;
        /// <summary>
        /// Đóng tab masterview
        /// </summary>
        public static RelayCommand CloseMasterViewCommand
        {
            get { return _closeMasterViewCommand ?? (_closeMasterViewCommand = new RelayCommand(p => CloseMasterViewExecute())); }
        }

        private static void CloseMasterViewExecute()
        {
            //Close SlideMasterView
            //(Application.Current as IAppGlobal).SelectedThemeView.SelectedFont = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont;
            //(Application.Current as IAppGlobal).SelectedThemeView.RefreshData();

            // Xu lys
            //(Application.Current as IAppGlobal).DocumentControl.SelectedTheme = (Application.Current as IAppGlobal).SelectedTheme;
            (Application.Current as IAppGlobal).SlideViewMode = SlideViewMode.Normal;

            //(Application.Current as IAppGlobal).SelectedThemeView.RefreshData();
            //var _data = (Application.Current as IAppGlobal).SelectedThemeView.Data;
            ////(Application.Current as IAppGlobal).AllTheme = _data;
            //(Application.Current as IAppGlobal).SelectedTheme.Colors = _data.Colors;
            //_data = null;

            //Console.WriteLine((Application.Current as IAppGlobal).SelectedTheme);
            //foreach (var item in (Application.Current as IAppGlobal).DocumentControl.Slides)
            //{
            //    item.UpdateThemeColor();
            //}

            //ColorGradientBrush gradientColor = new ColorGradientBrush();
            //ObservableCollection<CustomGradientStop> listCustom = new ObservableCollection<CustomGradientStop>();
            //listCustom.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = "#FF549938", Name = "Red", Brightness = -0.11 }, Offset = 0.0 });
            //listCustom.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = "Yellow", Name = "Yellow", Brightness = -0.11 }, Offset = 1.0 });
            //gradientColor.GradientStops = listCustom;
            //gradientColor.Angle = 90;
            //(Application.Current as IAppGlobal).SelectedSlide.MainLayout.Fill = new ColorSolidBrush() { Color = Colors.Transparent, ColorSpecialName = "Transparent" };


            ////Generate Theme
            //ThemesData themesData = new ThemesData();

            ////ETheme eTheme = new ETheme();
            ////eTheme = themesData.OpenETheme("E:\\BlankTheme.zip");
            //(Application.Current as IAppGlobal).SelectedThemeView.RefreshData();
            //ETheme eThemes = new ETheme();
            //eThemes = (Application.Current as IAppGlobal).SelectedThemeView.Data;
            //eThemes.Colors = (Application.Current as IAppGlobal).SelectedTheme.Colors;
            //eThemes.SelectedFont = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont;



            //themesData.Save(eThemes, "E:\\Theme09.athmx");
            //// SlideMasterTabControlViewModel.GetThemesCollection();

        }

        private static ObservableCollection<CustomGradientStop> GetStopsPreset5(INV.Elearning.Core.Model.SolidColor color)
        {
            ObservableCollection<CustomGradientStop> _result = new ObservableCollection<CustomGradientStop>();
            Color _argColor = (Color)ColorConverter.ConvertFromString(color.Color);
            _result.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = "Red", Name = color.Name, Brightness = -0.11 }, Offset = 0.0 });
            _result.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = "Blue", Name = color.Name, Brightness = -0.11 }, Offset = 0.23 });
            _result.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = "Black", Name = color.Name, Brightness = -0.25 }, Offset = 0.69 });
            _result.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = "Yellow", Name = color.Name, Brightness = -0.30 }, Offset = 0.97 });

            return _result;
        }


        #endregion

        #region ThemeCommand
        private static RelayCommand _selectThemeCommand;
        /// <summary>
        /// Command lựa chọn theme
        /// </summary>
        public static RelayCommand SelectThemeCommand
        {
            get { return _selectThemeCommand ?? (_selectThemeCommand = new RelayCommand(p => SelectThemeExecute(p))); }
        }

        private static void SelectThemeExecute(object p)
        {
            ETheme eTheme = p as ETheme;
            //CHeck selected theme
            if (eTheme.TagName == "Office")
            {            
                SlideMaster slide = null;
                if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster slidemaster)
                {
                    slide = slidemaster;
                }
                else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster layoutMaster)
                {
                    slide = layoutMaster.SlideParent;
                }
                if (slide != null)
                {
                    slide.SelectedThemeId = eTheme.ID;                 
                }
            }
            else
            {
                //Global.PushUndo(new SlideMasterThisPresenterStep(_beforeThemes, selectedThemes));
                foreach (SlideMaster item in (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters)
                {
                    if (item.ThemesName == eTheme.Name)
                    {
                        item.IsSelected = true;
                        break;
                    }
                }
            }
        }
      

        /// <summary>
        /// Thay đổi background khi lựa chọn theme
        /// </summary>
        /// <param name="eTheme"></param>
        public static void ChangeThemeBackground(ETheme eTheme)
        {
            Global.BeginInit();
            SlideMaster slide = null;
            if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster)
            {
                slide = (Application.Current as IAppGlobal).SelectedSlide as SlideMaster;
            }
            else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster)
            {
                slide = ((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).SlideParent;
            }

            if (slide.ThemeLayout == null) slide.ThemeLayout = new ThemeLayout();

            RemoveBackgroundSlideMasterShapes(slide);

            ObservableCollection<StandardElement> StandardElements = new ObservableCollection<StandardElement>();
            foreach (var item in eTheme.SlideMasters[0].MainLayer.Children)
            {
                if (item is EStandardElement eThemeShape && eThemeShape.IsGraphicBackground)
                {
                    StandardElement StandardElement = new StandardElement();
                    StandardElement.UpdateUI(eThemeShape);
                    StandardElements.Add(StandardElement);
                }
            }

            slide.MainLayout.Fill = new ColorSolidBrush() { Color = Colors.Transparent, ColorSpecialName = "Transparent" };



            int maxZindex = GetMaxZIndexShapeBackground(StandardElements);
            foreach (ObjectElement objectelement in slide.MainLayout.Elements)
            {
                objectelement.ZIndex += (maxZindex + 1);
            }
            foreach (StandardElement item in StandardElements)
            {

                slide.MainLayout.Elements.Add(item as StandardElement);
            }

            slide.MainLayout.UpdateThumbnail();

            foreach (LayoutMaster layout in slide.LayoutMasters)
            {
                RemoveBackgroundLayoutMasterShapes(layout);
                if (layout.ThemeLayout == null) layout.ThemeLayout = new ThemeLayout();
                slide.RefreshData();
                layout.ThemeLayout.UpdateUI(slide.Data.MainLayer);
                foreach (ELayoutMaster item in eTheme.SlideMasters[0].LayoutMasters)
                {
                    if (item.SlideName == layout.SlideName)
                    {
                        layout.IsHideBackground = item.IsHideBackground;

                        foreach (EStandardElement eStandardElement in item.MainLayer.Children)
                        {
                            if (eStandardElement is DataDocument data)
                            {
                                TextEditor textEditor = new TextEditor();
                                textEditor.UpdateUI(data);
                                layout.MainLayout.Elements.Add(textEditor);
                            }
                            else if (eStandardElement is EPlaceHolder ePlaceHolder)
                            {
                                PlaceHolder placeHolder = new PlaceHolder();
                                placeHolder.UpdateUI(ePlaceHolder);
                                layout.MainLayout.Elements.Add(placeHolder);
                            }
                            else
                            {
                                StandardElement layoutElement = new StandardElement();
                                layoutElement.UpdateUI(eStandardElement);
                                layout.MainLayout.Elements.Add(layoutElement);
                            }
                        }
                    }
                }

                layout.MainLayout.Fill = new ColorSolidBrush() { Color = Colors.Transparent, ColorSpecialName = "Transparent" };

                layout.UpdateLayoutTheme();
                layout.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
                {
                    layout.MainLayout.UpdateThumbnail();
                }));
            }
            Global.EndInit();
        }

        /// <summary>
        /// Thay đổi màu của Slide đang được chọn
        /// </summary>
        /// <param name="eColor"></param>
        /// <param name="slide"></param>
        public static void UpdateColor(EColorManagment eColor, SlideMaster slide)
        {
            (Application.Current as IAppGlobal).SelectedTheme.Colors = eColor;
            int index = (Application.Current as IAppGlobal).OfficeColors.IndexOf(eColor);
            slide.UpdateThemeColor();
            foreach (LayoutMaster item in slide.LayoutMasters)
            {
                if (item.ThemeLayout == null) item.ThemeLayout = new ThemeLayout();
                item.UpdateThemeColor();
            }
        }

        /// <summary>
        /// Thay đổi font của slide đang được chọn
        /// </summary>
        /// <param name="eFont"></param>
        /// <param name="slide"></param>
        public static void UpdateFont(EFontfamily eFont, SlideMaster slide)
        {
            (Application.Current as IAppGlobal).SelectedTheme.SelectedFont = eFont;
            slide.UpdateThemeFont();
            foreach (LayoutMaster item in slide.LayoutMasters)
            {
                item.UpdateThemeFont();
            }
        }

        /// <summary>
        /// Cập nhật màu cho background
        /// </summary>
        /// <param name="backgroundItem"></param>
        /// <param name="slide"></param>
        public static void UpdateBackground(BackgroundItem backgroundItem, SlideBase slide)
        {
            ColorGradientBrush gradientColor = new ColorGradientBrush();
            ObservableCollection<CustomGradientStop> listCustom = new ObservableCollection<CustomGradientStop>();
            if (backgroundItem.ColorStyle is LinearGradientBrush)
            {
                foreach (CustomGradientStop item in backgroundItem.ColorDataStyle.GradientStops)
                {
                    foreach (GradientStop item1 in (backgroundItem.ColorStyle as LinearGradientBrush).GradientStops)
                    {
                        if (backgroundItem.ColorDataStyle.GradientStops.IndexOf(item) == (backgroundItem.ColorStyle as LinearGradientBrush).GradientStops.IndexOf(item1))
                        {
                            listCustom.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = item1.Color.ToString(), Name = item.Color.Name, Brightness = 1 }, Offset = item.Offset });
                        }
                    }
                }
            }
            else if (backgroundItem.ColorStyle is RadialGradientBrush)
            {
                gradientColor.GradientType = GradientColorType.Radial;
                foreach (CustomGradientStop item in backgroundItem.ColorDataStyle.GradientStops)
                {
                    foreach (GradientStop item1 in (backgroundItem.ColorStyle as RadialGradientBrush).GradientStops)
                    {
                        if (backgroundItem.ColorDataStyle.GradientStops.IndexOf(item) == (backgroundItem.ColorStyle as RadialGradientBrush).GradientStops.IndexOf(item1))
                        {
                            listCustom.Add(new CustomGradientStop() { Color = new INV.Elearning.Controls.SolidColor() { Color = item1.Color.ToString(), Name = item.Color.Name, Brightness = 1 }, Offset = item.Offset });
                        }
                    }
                }
            }



            gradientColor.GradientStops = listCustom;
            backgroundItem.ColorDataStyle = gradientColor;
            if (slide is LayoutMaster layoutMaster)
            {
                // Global.BeginInit();
                layoutMaster.MainLayout.Fill = backgroundItem.ColorDataStyle;
                layoutMaster.MainLayout.UpdateThumbnail();
                // Global.EndInit();
            }
            else if (slide is SlideMaster slideMaster)
            {
                // Global.BeginInit();
                slideMaster.MainLayout.Fill = backgroundItem.ColorDataStyle;
                slideMaster.MainLayout.UpdateThumbnail();
                foreach (LayoutMaster item in slideMaster.LayoutMasters)
                {
                    item.MainLayout.Fill = backgroundItem.ColorDataStyle;
                    item.MainLayout.UpdateThumbnail();
                }
                (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].MainLayer.Background = ColorHelper.ConverterFromColor(backgroundItem.ColorDataStyle);

                //Global.EndInit();
            }
            else if (slide is NormalSlide normalSlide)
            {
                normalSlide.MainLayout.Fill = backgroundItem.ColorDataStyle;
                normalSlide.MainLayout.UpdateThumbnail();
            }

        }

        /// <summary>
        /// Cập nhật Layout Master
        /// </summary>
        /// <param name="slideMaster"></param>
        /// <param name="eTheme"></param>
        public static void UpdateLayoutsMaster(SlideMaster slideMaster, ETheme eTheme)
        {
            slideMaster.RefreshData();
            foreach (LayoutMaster layout in slideMaster.LayoutMasters)
            {
                RemoveBackgroundLayoutMasterShapes(layout);
                if (layout.ThemeLayout == null) layout.ThemeLayout = new ThemeLayout();
                layout.ThemeLayout.UpdateUI(slideMaster.Data.MainLayer);
                if (layout.LayoutType != LayoutType.Custom)
                {
                    foreach (ELayoutMaster item in eTheme.SlideMasters[0].LayoutMasters)
                    {
                        if (item.LayoutType == layout.LayoutType)
                        {
                            layout.IsHideBackground = item.IsHideBackground;

                            foreach (EStandardElement eStandardElement in item.MainLayer.Children)
                            {
                                if (eStandardElement is DataDocument data)
                                {
                                    TextEditor textEditor = new TextEditor();
                                    textEditor.UpdateUI(data);
                                    layout.MainLayout.Elements.Add(textEditor);
                                }
                                else if (eStandardElement is EPlaceHolder ePlaceHolder)
                                {
                                    PlaceHolder placeHolder = new PlaceHolder();
                                    placeHolder.UpdateUI(ePlaceHolder);
                                    layout.MainLayout.Elements.Add(placeHolder);
                                }
                                else
                                {
                                    StandardElement layoutElement = new StandardElement();
                                    layoutElement.UpdateUI(eStandardElement);
                                    layout.MainLayout.Elements.Add(layoutElement);
                                }
                            }
                        }
                    }
                }
                else
                {
                    foreach (ELayoutMaster item in eTheme.SlideMasters[0].LayoutMasters)
                    {
                        if (item.LayoutType == LayoutType.TitleOnly)
                        {
                            layout.IsHideBackground = item.IsHideBackground;

                            foreach (EStandardElement eStandardElement in item.MainLayer.Children)
                            {
                                if (eStandardElement is DataDocument data)
                                {
                                    TextEditor textEditor = new TextEditor();
                                    textEditor.UpdateUI(data);
                                    layout.MainLayout.Elements.Add(textEditor);
                                }
                                else if (eStandardElement is EPlaceHolder ePlaceHolder)
                                {
                                    PlaceHolder placeHolder = new PlaceHolder();
                                    placeHolder.UpdateUI(ePlaceHolder);
                                    layout.MainLayout.Elements.Add(placeHolder);
                                }
                                else
                                {
                                    StandardElement layoutElement = new StandardElement();
                                    layoutElement.UpdateUI(eStandardElement);
                                    layout.MainLayout.Elements.Add(layoutElement);
                                }
                            }
                        }
                    }
                }

            }
        }

        /// <summary>
        /// Thay đổi thuộc tính IsTitle của Slide
        /// </summary>
        /// <param name="value"></param>
        /// <param name="layout"></param>
        public static void IsTitleSlideChanged(bool value, SlideBase slide)
        {
            if (slide is LayoutMaster layout)
            {
                foreach (var item in layout.MainLayout.Elements)
                {
                    if (item is TextEditor editor && editor.RichTextEditor.TextContainer.Document.TypeText == TypeText.Title)
                    {
                        editor.IsShow = value;
                    }
                }
                layout.MainLayout.UpdateThumbnail();
            }
            else if (slide is SlideMaster slideMaster)
            {
                foreach (var item in slideMaster.MainLayout.Elements)
                {
                    if (item is TextEditor editor && editor.RichTextEditor.TextContainer.Document.TypeText == TypeText.Title)
                    {
                        editor.IsShow = value;
                    }
                }
                slideMaster.MainLayout.UpdateThumbnail();
            }

        }

        /// <summary>
        /// Thay đổi thuộc tính IsFooters của Slide
        /// </summary>
        /// <param name="value"></param>
        /// <param name="layout"></param>
        public static void IsFootersSlideChanged(bool value, SlideBase slide)
        {
            if (slide is LayoutMaster layout)
            {
                foreach (var item in layout.MainLayout.Elements)
                {
                    if (item is TextEditor editor && (editor.RichTextEditor.TextContainer.Document.TypeText == TypeText.DateTime || editor.RichTextEditor.TextContainer.Document.TypeText == TypeText.SlideNumber || editor.RichTextEditor.TextContainer.Document.TypeText == TypeText.Footers))
                    {
                        editor.IsShow = value;
                    }
                }
            }
            else if (slide is SlideMaster slideMaster)
            {
                foreach (var item in slideMaster.MainLayout.Elements)
                {
                    if (item is TextEditor editor1 && (editor1.RichTextEditor.TextContainer.Document.TypeText == TypeText.Footers))
                    {
                        editor1.IsShow = value;
                    }
                }
            }
            else if (slide is NormalSlide normalSlide)
            {
                foreach (var item in normalSlide.MainLayout.Elements)
                {
                    if (item is TextEditor editor1 && (editor1.RichTextEditor.TextContainer.Document.TypeText == TypeText.Footers))
                    {
                        editor1.IsShow = value;
                    }
                }
            }
            slide.MainLayout.UpdateThumbnail();

        }

        /// <summary>
        /// Thay đổi thuộc tính date của SlideMaster
        /// </summary>
        /// <param name="value"></param>
        /// <param name="slideMaster"></param>
        public static void IsDateSlideChanged(bool value, SlideBase slide)
        {
            if (slide is SlideMaster slideMaster)
            {
                foreach (var item in slideMaster.MainLayout.Elements)
                {
                    if (item is TextEditor editor1 && (editor1.RichTextEditor.TextContainer.Document.TypeText == TypeText.DateTime))
                    {
                        editor1.IsShow = value;
                    }
                }
            }
            else if (slide is NormalSlide normalSlide)
            {
                foreach (var item in normalSlide.MainLayout.Elements)
                {
                    if (item is TextEditor editor1 && (editor1.RichTextEditor.TextContainer.Document.TypeText == TypeText.DateTime))
                    {
                        editor1.IsShow = value;
                    }
                }
            }
            slide.MainLayout.UpdateThumbnail();
        }

        /// <summary>
        /// Thay đổi thuộc tính slideNumber của SlideMaster
        /// </summary>
        /// <param name="value"></param>
        /// <param name="slideMaster"></param>
        public static void IsSlideNumberSlideChanged(bool value, SlideBase slide)
        {
            if (slide is SlideMaster slideMaster)
            {
                foreach (var item in slideMaster.MainLayout.Elements)
                {
                    if (item is TextEditor editor1 && (editor1.RichTextEditor.TextContainer.Document.TypeText == TypeText.SlideNumber))
                    {
                        editor1.IsShow = value;
                    }
                }
            }
            else if (slide is NormalSlide normalSlide)
            {
                foreach (var item in normalSlide.MainLayout.Elements)
                {
                    if (item is TextEditor editor1 && (editor1.RichTextEditor.TextContainer.Document.TypeText == TypeText.SlideNumber))
                    {
                        editor1.IsShow = value;
                    }
                }
            }
            slide.MainLayout.UpdateThumbnail();
        }

        /// <summary>
        /// Thay đổi thuộc tính IsText của SlideMaster
        /// </summary>
        /// <param name="value"></param>
        /// <param name="slideMaster"></param>
        public static void IsTextNumberSlideChanged(bool value, SlideBase slide)
        {
            if (slide is SlideMaster slideMaster)
            {
                foreach (var item in slideMaster.MainLayout.Elements)
                {
                    if (item is TextEditor editor1 && (editor1.RichTextEditor.TextContainer.Document.TypeText == TypeText.Text))
                    {
                        editor1.IsShow = value;
                    }
                }
            }

            slide.MainLayout.UpdateThumbnail();
        }

        private static void InsertThemeBackground(ObservableCollection<StandardElement> shapeBackgrounds, SlideMaster slide, EColorManagment color)
        {

            RemoveBackgroundSlideMasterShapes(slide);

            int maxZindex = GetMaxZIndexShapeBackground(shapeBackgrounds);
            foreach (ObjectElement objectelement in slide.MainLayout.Elements)
            {
                objectelement.ZIndex += (maxZindex + 1);
            }

            foreach (StandardElement item in shapeBackgrounds)
            {
                if (slide.ThemeLayout == null) slide.ThemeLayout = new ThemeLayout();
                slide.MainLayout.Elements.Add(item as StandardElement);
            }

            GenerateThemeShape.GenerateLayoutTheme10(slide, maxZindex, color);
        }

        /// <summary>
        /// Get max Zindex của hình nền
        /// </summary>
        /// <param name="shapeBackgrounds"></param>
        /// <returns></returns>
        private static int GetMaxZIndexShapeBackground(ObservableCollection<StandardElement> shapeBackgrounds)
        {
            int max = 0;
            foreach (StandardElement item in shapeBackgrounds)
            {
                if (item.ZIndex > max)
                {
                    max = item.ZIndex;
                }
            }
            return max;
        }

        /// <summary>
        /// Xóa toàn bộ hình nền của Slide Master
        /// </summary>
        /// <param name="slide"></param>
        private static void RemoveBackgroundSlideMasterShapes(SlideMaster slide)
        {
            Global.BeginInit();
            for (int i = 0; i < slide.MainLayout.Elements.Count; i++)
            {
                if (slide.MainLayout.Elements[i] is StandardElement && ((slide.MainLayout.Elements[i] as StandardElement).IsGraphicBackground))
                {
                    slide.MainLayout.Elements.Remove(slide.MainLayout.Elements[i] as StandardElement);
                    i--;
                }
            }
            Global.EndInit();
        }

        /// <summary>
        /// Xóa toàn bộ hình nền của Layout Master
        /// </summary>
        /// <param name="slide"></param>
        private static void RemoveBackgroundLayoutMasterShapes(LayoutMaster slide)
        {
            Global.BeginInit();         
            for (int i = 0; i < slide.MainLayout.Children.Count; i++)
            {
                var element = slide.MainLayout.Children[i];
                if (element is ThemeLayout) continue;
                slide.MainLayout.Children.Remove(element);
                element = null;
                i--;
            }
            slide.MainLayout.Elements.Clear();
            Global.EndInit();
        }


        #endregion


        public static void ChangeNormalMode()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = null;
            (Application.Current as IAppGlobal).SelectedTheme.Colors = (Application.Current as IAppGlobal).SelectedThemeView.Colors;
            (Application.Current as IAppGlobal).SelectedTheme.SelectedFont = (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont;

            foreach (ETheme item in (Application.Current as IAppGlobal).LocalThemesCollection)
            {
                if (item.ID == (Application.Current as IAppGlobal).SelectedTheme.ID)
                {
                    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = item;
                }
            }


            int index = (Application.Current as IAppGlobal).LocalThemesCollection.IndexOf((Application.Current as IAppGlobal).DocumentControl.SelectedTheme);
            if (index != -1)
            {
                DesignTabControlViewModel.ThemeDesign.SelectedIndex = index;
            }
            else DesignTabControlViewModel.ThemeDesign.SelectedIndex = 0;


            foreach (SlideBase slideBase in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {

                slideBase.MainLayout.Fill = ColorHelper.ConverFromColorData((Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].MainLayer.Background);
            }
            foreach (ETheme item in (Application.Current as IAppGlobal).LocalThemesCollection)
            {
                item.CountTheme = DesignTabControlViewModel.CountThemeApply(item);
            }
            Global.EndInit();

            Global.PushUndo(new NormalViewModeStep());
        }

        public static void ChangeSlideMasterMode()
        {
            if ((Application.Current as IAppGlobal).SelectedThemeView == null) (Application.Current as IAppGlobal).SelectedThemeView = new Theme();
            Global.BeginInit();
            var _themeData = new ETheme();
            foreach (ETheme item in (Application.Current as IAppGlobal).LocalThemesCollection)
            {
                item.SlideMasters[0].SlideName = item.Name;
                item.SlideMasters[0].ThemesName = item.Name;
                item.SlideMasters[0].ID = item.ID;
                _themeData.SlideMasters.Add(item.SlideMasters[0]);
            }


            Theme _theme = new Theme();
            _theme.UpdateUI(_themeData);
            _theme.Colors = (Application.Current as IAppGlobal).SelectedTheme.Colors;
            _theme.SelectedFont = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont;
            (Application.Current as IAppGlobal).SelectedThemeView = _theme;
            _theme.SlideMasters[0].MainLayout.UpdateThemeColor();
            _theme.SlideMasters[0].UpdateThemeFont();
            foreach (LayoutMaster item in _theme.SlideMasters[0].LayoutMasters)
            {
                item.UpdateLayoutTheme();
                item.MainLayout.UpdateThemeColor();
                item.UpdateThemeFont();
            }

            SlideHelper.UnSlectedAll();
            (Application.Current as IAppGlobal).DocumentControl.Slides[0].IsSelected = true;

            foreach (SlideBase item in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {
                item.MainLayout.UpdateThumbnail();
                if (item is SlideMaster slideMaster)
                {
                    slideMaster.SelectedThemeId = slideMaster.ID;
                }
            }
            (Application.Current as IAppGlobal).SelectedThemeView.SelectedTheme = (Application.Current as IAppGlobal).SelectedTheme;
            (Application.Current as IAppGlobal).SelectedThemeView.Colors = (Application.Current as IAppGlobal).SelectedTheme.Colors;
            (Application.Current as IAppGlobal).SelectedThemeView.SelectedFont = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont;

            INV.Elearning.Timeline.Timeline.TimelineViewModel.UpdateTimeline();
            Global.EndInit();
            Global.PushUndo(new SlideMasterViewModeStep(_theme));
        }

    }
}
