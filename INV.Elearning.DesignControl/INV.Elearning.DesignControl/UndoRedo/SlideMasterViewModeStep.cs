using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.Core.ViewModel;
using INV.Elearning.DesignControl.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace INV.Elearning.DesignControl.UndoRedo
{

    public class SlideMasterViewModeStep : StepBase
    {
        public Theme Theme { get; set; }
        public SlideMasterViewModeStep(Theme theme)
        {
            Theme = theme;
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SlideViewMode = SlideViewMode.Normal;
            (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = (Application.Current as IAppGlobal).SelectedTheme;
            //(Application.Current as IAppGlobal).SelectedTheme = (Application.Current as IAppGlobal).SelectedThemes.MainTheme;
            //SlideMaster slideMaster = new SlideMaster();
            //slideMaster.LoadData((Application.Current as IAppGlobal).DocumentControl.SelectedThemes.MainTheme.SlideMasters[0]);
            //int count = 0;
            //(Application.Current as IAppGlobal).DocumentControl.SelectedThemes.ThumbnailIcon = DesignTabControlViewModel.GetThumbnail(slideMaster.MainLayout, (Application.Current as IAppGlobal).DocumentControl.SelectedThemes.MainTheme.Colors, new Size(341, 192));

            //foreach (ELayoutMaster item in (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters)
            //{
            //    slideMaster.LayoutMasters[count].UpdateLayoutTheme();
            //    item.ThumbnailBitmap = DesignTabControlViewModel.CaptureLayout(slideMaster.LayoutMasters[count].MainLayout, new Size(341, 192));
            //    count++;
            //    count = Math.Min(slideMaster.LayoutMasters.Count - 1, count);
            //}
            //foreach (SlideBase slideBase in (Application.Current as IAppGlobal).DocumentControl.Slides)
            //{
            //    slideBase.ThemeLayout.ThemeID = (Application.Current as IAppGlobal).SelectedThemes.ID;
            //    if (slideBase.ThemeLayout == null) slideBase.ThemeLayout = new ThemeLayout();
            //    slideBase.ThemeLayout.UpdateUI((Application.Current as IAppGlobal).SelectedThemes.MainTheme.SlideMasters[0].MainLayer);
            //    slideBase.MainLayout.Fill = ColorHelper.ConverFromColorData((Application.Current as IAppGlobal).SelectedThemes.MainTheme.SlideMasters[0].MainLayer.Background);
            //    slideBase.MainLayout.UpdateThumbnail();
            //    slideBase.MainLayout.UpdateThemeColor();
            //    slideBase.UpdateThemeFont();
            //}
            //foreach (EThemes item in (Application.Current as IAppGlobal).LocalThemesCollection)
            //{
            //    item.CountTheme = DesignTabControlViewModel.CountThemeApply(item);
            //}
            //SlideHelper.UnSlectedAll();
            //(Application.Current as IAppGlobal).DocumentControl.Slides[0].IsSelected = true;
            //slideMaster = null;
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SlideViewMode = SlideViewMode.SlideMaster;
            SlideHelper.UnSlectedAll();
            (Application.Current as IAppGlobal).DocumentControl.Slides[0].IsSelected = true;
            Global.EndInit();
        }
    }
}
