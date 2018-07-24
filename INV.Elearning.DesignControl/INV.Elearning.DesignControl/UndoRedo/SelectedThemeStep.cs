using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
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
using System.Windows.Media;

namespace INV.Elearning.DesignControl.UndoRedo
{
    public class SelectedThemeStep : StepBase
    {
        //public EColorManagment OldColor { get; set; }
        //public EColorManagment NewColor { get; set; }
        //public EThemes EThemes { get; set; }
        //public SelectedThemeStep(EColorManagment oldColor, EColorManagment newColor, EThemes ethemes)
        //{
        //    OldColor = oldColor;
        //    NewColor = newColor;
        //    EThemes = ethemes;
        //}

        //public override void UndoExcute()
        //{
        //    Global.BeginInit();
        //    (Application.Current as IAppGlobal).SelectedTheme.Colors = (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = OldColor;
        //    (Application.Current as IAppGlobal).SelectedThemes.MainTheme.Colors = OldColor;
        //    foreach (SlideBase slide in (Application.Current as IAppGlobal).DocumentControl.Slides)
        //    {
        //        slide.UpdateThemeColor();
        //    }

        //    foreach (EThemes item in (Application.Current as IAppGlobal).LocalThemesCollection)
        //    {
        //        item.CountTheme = 0;
        //    }

        //    //foreach (ELayoutMaster item in EThemes.MainTheme.SlideMasters[0].LayoutMasters)
        //    //{
        //    //    Task.Factory.StartNew(() =>
        //    //    {
        //    //        Application.Current.Dispatcher.Invoke(() =>
        //    //        {
        //    //            LayoutMaster layoutMaster = new LayoutMaster();
        //    //            layoutMaster.LoadData(item);
        //    //            if (!layoutMaster.IsHideBackground)
        //    //            {
        //    //                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();
        //    //                layoutMaster.ThemeLayout.UpdateUIB(EThemes.MainTheme.SlideMasters[0].MainLayer);
        //    //            }
        //    //            layoutMaster.UpdateThemeColor();
        //    //            item.ThumbnailBitmap = DesignTabControlViewModel.CaptureLayout(layoutMaster.MainLayout, new Size(341, 192));
        //    //        });
        //    //    });
        //    //}
        //    //SlideMaster slideMaster = new SlideMaster();
        //    //slideMaster.MainLayout.Fill = ColorHelper.ConverFromColorData(EThemes.MainTheme.SlideMasters[0].MainLayer.Background);
        //    //slideMaster.ThemeLayout = new ThemeLayout();
        //    //slideMaster.ThemeLayout.UpdateUI((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer);
        //    //slideMaster.UpdateThemeColor();
        //    //(Application.Current as IAppGlobal).DocumentControl.SelectedThemes.ThumbnailIcon = DesignTabControlViewModel.GetThumbnail(slideMaster.MainLayout, OldColor, new Size(341, 192));
        //    Global.EndInit();
        //}

        //public override void RedoExcute()
        //{
        //    Global.BeginInit();
        //    (Application.Current as IAppGlobal).SelectedTheme.Colors = (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = NewColor;
        //    (Application.Current as IAppGlobal).SelectedThemes.MainTheme.Colors = NewColor;
        //    foreach (SlideBase slide in (Application.Current as IAppGlobal).DocumentControl.Slides)
        //    {
        //        slide.UpdateThemeColor();
        //    }

        //    foreach (EThemes item in (Application.Current as IAppGlobal).LocalThemesCollection)
        //    {
        //        item.CountTheme = 0;
        //    }
        //    //foreach (ELayoutMaster item in EThemes.MainTheme.SlideMasters[0].LayoutMasters)
        //    //{
        //    //    Task.Factory.StartNew(() =>
        //    //    {
        //    //        Application.Current.Dispatcher.Invoke(() =>
        //    //        {
        //    //            LayoutMaster layoutMaster = new LayoutMaster();
        //    //            layoutMaster.LoadData(item);
        //    //            if (!layoutMaster.IsHideBackground)
        //    //            {
        //    //                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();
        //    //                layoutMaster.ThemeLayout.UpdateUIB(EThemes.MainTheme.SlideMasters[0].MainLayer);
        //    //            }
        //    //            layoutMaster.UpdateThemeColor();
        //    //            item.ThumbnailBitmap = DesignTabControlViewModel.CaptureLayout(layoutMaster.MainLayout, new Size(341, 192));
        //    //        });
        //    //    });
        //    //}
        //    //SlideMaster slideMaster = new SlideMaster();
        //    //slideMaster.MainLayout.Fill = ColorHelper.ConverFromColorData(EThemes.MainTheme.SlideMasters[0].MainLayer.Background);
        //    //slideMaster.ThemeLayout = new ThemeLayout();
        //    //slideMaster.ThemeLayout.UpdateUI((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer);
        //    //slideMaster.UpdateThemeColor();
        //    //(Application.Current as IAppGlobal).DocumentControl.SelectedThemes.ThumbnailIcon = DesignTabControlViewModel.GetThumbnail(slideMaster.MainLayout, NewColor, new Size(341, 192));
        //    Global.EndInit();
        //}
    }
}
