using DocumentFormat.OpenXml.Packaging;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.DesignControl.Models;
using INV.Elearning.Text.Models;
using Microsoft.Office.Core;
using System.Windows;
using pp = Microsoft.Office.Interop.PowerPoint;

namespace INV.Elearning.ImportPowerPoint.Helper
{
    /// <summary>
    /// Enum phân biệt Part
    /// </summary>
    public enum PartType
    {
        NormalSlide,
        SlideMaster,
        LayoutMaster
    }
    public partial class HelperClass
    {
        #region GetColorTheme
        /// <summary>
        /// Lấy màu của theme
        /// </summary>
        /// <param name="colorScheme"></param>
        /// <returns></returns>
        public EColorManagment GetEColorManagment(ThemeColorScheme colorScheme)
        {
            EColorManagment eColor = new EColorManagment();
            eColor.Accent1 = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeAccent1);
            eColor.Accent2 = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeAccent2);
            eColor.Accent3 = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeAccent3);
            eColor.Accent4 = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeAccent4);
            eColor.Accent5 = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeAccent5);
            eColor.Accent6 = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeAccent6);
            eColor.BackgroundDark1 = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeDark1);
            eColor.BackgroundDark2 = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeDark2);
            eColor.BackgroundLight1 = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeLight1);
            eColor.BackgroundLight2 = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeLight2);
            eColor.Hyperlink = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeHyperlink);
            eColor.FollowedHyperlink = GetSolidColor(colorScheme, MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink);
            ////Lấy color accent 1
            //Core.Model.SolidColor color = new Core.Model.SolidColor();
            //color.Brightness = 1;
            //color.Color = color

            return eColor;
        }

        /// <summary>
        /// Lấy màu Solid Color
        /// </summary>
        /// <param name="themeColor"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public Core.Model.SolidColor GetSolidColor(ThemeColorScheme themeColor, MsoThemeColorSchemeIndex index)
        {
            Core.Model.SolidColor solid = new Core.Model.SolidColor();
            solid.Brightness = 1;
            solid.Color = ConvertColor(themeColor.Colors(index).RGB);
            switch (index)
            {
                case MsoThemeColorSchemeIndex.msoThemeDark1:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounDark1").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounDark1").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeLight1:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounDark2").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounDark2").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeDark2:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounLight1").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounLight1").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeLight2:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounLight2").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounLight2").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeAccent1:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent1").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent1").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeAccent2:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent2").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent2").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeAccent3:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent3").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent3").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeAccent4:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent4").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent4").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeAccent5:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent5").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent5").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeAccent6:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent6").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_Accent6").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeHyperlink:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_Hyperlink").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_Hyperlink").ToString();
                    break;
                case MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink:
                    solid.Name = System.Windows.Application.Current.TryFindResource("ColorGallery_FollowedHyperlink").ToString();
                    solid.SpecialName = System.Windows.Application.Current.TryFindResource("ColorGallery_FollowedHyperlink").ToString();
                    break;
                default:
                    break;
            }
            return solid;
        }


        #endregion

        #region GetFontTheme
        public EFontfamily GetFontTheme(ThemeFontScheme themeFontScheme)
        {
            EFontfamily fontfamily = new EFontfamily();
            fontfamily.MajorFont = themeFontScheme.MajorFont.Item(MsoFontLanguageIndex.msoThemeLatin).Name;
            fontfamily.MinorFont = themeFontScheme.MinorFont.Item(MsoFontLanguageIndex.msoThemeLatin).Name;
            fontfamily.TagName = fontfamily.MajorFont + "-" + fontfamily.MinorFont;
            return fontfamily;
        }
        #endregion

        #region GetTheme
        /// <summary>
        /// Lấy chủ đề
        /// </summary>
        /// <param name="design"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        public ETheme GetETheme(pp.Design design, int count)
        {
            ETheme etheme = new ETheme();
            etheme.Name = design.Name;
            etheme.ID = ObjectElementsHelper.RandomString(10);
            etheme.ThemeName = design.Name;
            EColorManagment eColorManagment = new EColorManagment();
            eColorManagment = GetEColorManagment(design.SlideMaster.Theme.ThemeColorScheme);
            etheme.Colors = eColorManagment;
            etheme.SelectedFont = GetFontTheme(design.SlideMaster.Theme.ThemeFontScheme);
            ESlideMaster eSlideMaster = new ESlideMaster();
            eSlideMaster.ThemesName = etheme.Name;
            eSlideMaster.Name = etheme.Name;
            eSlideMaster.MainLayer = new PageLayer();
            SlideMasterPart slideMasterPart = GetSlideMasterPart(count);
            eSlideMaster.MainLayer.Background = GetFillColorSlideMaster(design.SlideMaster.Background.Fill, slideMasterPart.SlideMaster?.CommonSlideData, slideMasterPart);
            eSlideMaster.ID = ObjectElementsHelper.RandomString(12);
            foreach (pp.Shape shape in design.SlideMaster.Shapes)
            {
                eSlideMaster.MainLayer.Children.Add(GetShape(shape, shape.Type, design.SlideMaster.TimeLine, slideMasterPart));
            }
            SetIsBackgroundShape(eSlideMaster.MainLayer);

            int countLayoutMaster = 0;
            foreach (pp.CustomLayout layout in design.SlideMaster.CustomLayouts)
            {
                SlideLayoutPart slideLayoutPart = GetSlideLayoutPart(countLayoutMaster++, slideMasterPart);
                ELayoutMaster eLayoutMaster = new ELayoutMaster();
                eLayoutMaster.ID = ObjectElementsHelper.RandomString(13);
                eLayoutMaster.MainLayer = new PageLayer();
                eLayoutMaster.MainLayer.Background = GetFillColorSlideMaster(layout.Background.Fill, slideLayoutPart.SlideLayout?.CommonSlideData, slideLayoutPart);
                if (eLayoutMaster.MainLayer.Background == null)
                {
                    eLayoutMaster.MainLayer.Background = eSlideMaster.MainLayer.Background;
                }
                foreach (pp.Shape shape in layout.Shapes)
                {
                    eLayoutMaster.MainLayer.Children.Add(GetShape(shape, shape.Type, design.SlideMaster.TimeLine, slideLayoutPart));
                }
                SetIsBackgroundShape(eLayoutMaster.MainLayer);
                eLayoutMaster.SlideName = layout.Name;
                eLayoutMaster.LayoutName = layout.Name;
                eLayoutMaster.IsHideBackground = layout.DisplayMasterShapes != MsoTriState.msoTrue;

                eSlideMaster.LayoutMasters.Add(eLayoutMaster);
            }

            etheme.SlideMasters.Add(eSlideMaster);
            return etheme;
        }

        /// <summary>
        /// Set thuộc tính có phải là background hay không cho các object trong theme
        /// </summary>
        /// <param name="pageLayer"></param>
        private void SetIsBackgroundShape(PageLayer pageLayer)
        {
            foreach (ObjectElementBase item in pageLayer.Children)
            {
                if (!(item is DataDocument || item is EPlaceHolder))
                {
                    item.IsGraphicBackground = true;
                }
            }
        }


        #endregion
    }
}
