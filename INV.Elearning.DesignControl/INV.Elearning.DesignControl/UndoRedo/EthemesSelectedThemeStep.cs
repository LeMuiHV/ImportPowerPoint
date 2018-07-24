using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.Core.ViewModel;
using INV.Elearning.DesignControl.ViewModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace INV.Elearning.DesignControl.UndoRedo
{
    public class EthemesSelectedThemeStep : StepBase
    {
        //public EThemes OldEThemes { get; set; }
        //public EThemes NewEThemes { get; set; }
        //public ObservableCollection<EThemes> OldThemes { get; set; }
        //public ObservableCollection<EThemes> NewThemes { get; set; }
        //public EthemesSelectedThemeStep(EThemes oldEThemes, EThemes newEThemes, ObservableCollection<EThemes> oldThemes, ObservableCollection<EThemes> newThemes)
        //{
        //    OldEThemes = oldEThemes;
        //    NewEThemes = newEThemes;
        //    OldThemes = oldThemes;
        //    NewThemes = newThemes;
        //}

        //public override void UndoExcute()
        //{
        //    Global.BeginInit();
        //    (Application.Current as IAppGlobal).LocalThemesCollection = OldThemes;
        //    (Application.Current as IAppGlobal).SelectedThemes = OldEThemes;
        //    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = OldEThemes.MainTheme;
        //    (Application.Current as IAppGlobal).SelectedTheme = OldEThemes.MainTheme;
        //    (Application.Current as IAppGlobal).DocumentControl.SelectedThemes = OldEThemes;

        //    //(Application.Current as IAppGlobal).DocumentControl.SelectedTheme = OldEThemes.MainTheme;
        //    //Add themelayout của theme vào các slide
        //    foreach (SlideBase slide in (Application.Current as IAppGlobal).DocumentControl.Slides)
        //    {
        //        if (slide.ThemeLayout == null) slide.ThemeLayout = new ThemeLayout();
        //        if (slide.SelectLayoutType == LayoutType.None)
        //        {
        //            if (OldEThemes.MainTheme.SlideMasters.Count > 0)
        //            {
        //                slide.ThemeLayout.UpdateUI(OldEThemes.MainTheme.SlideMasters[0].MainLayer);
        //                slide.MainLayout.Fill = ColorHelper.ConverFromColorData(OldEThemes.MainTheme.SlideMasters[0].MainLayer.Background);
        //            }
        //        }
        //        else
        //        {
        //            foreach (ELayoutMaster item in OldEThemes.MainTheme.SlideMasters[0].LayoutMasters)
        //            {
        //                if (item.LayoutType == slide.SelectLayoutType)
        //                {
        //                    DesignTabControlViewModel.UpdateLayoutMaster(item, slide, OldEThemes.Clone() as EThemes);
        //                }
        //            }
        //        }
        //        slide.UpdateThemeFont();
        //        slide.MainLayout.UpdateThumbnail();
        //    }
        //    foreach (EThemes item in (Application.Current as IAppGlobal).LocalThemesCollection)
        //    {
        //        item.CountTheme = DesignTabControlViewModel.CountThemeApply(item);
        //    }
        //    //(Application.Current as IAppGlobal).DocumentControl.SelectedThemes = OldEThemes;
        //}

        //public override void RedoExcute()
        //{
        //    Global.BeginInit();
        //    (Application.Current as IAppGlobal).LocalThemesCollection = NewThemes;
        //    (Application.Current as IAppGlobal).SelectedThemes = NewEThemes;
        //    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme = NewEThemes.MainTheme;
        //    (Application.Current as IAppGlobal).SelectedTheme = NewEThemes.MainTheme;
        //    (Application.Current as IAppGlobal).DocumentControl.SelectedThemes = NewEThemes;

        //    foreach (SlideBase slide in (Application.Current as IAppGlobal).DocumentControl.Slides)
        //    {
        //        if (slide.ThemeLayout == null) slide.ThemeLayout = new ThemeLayout();
        //        if (slide.SelectLayoutType == LayoutType.None)
        //        {
        //            if (NewEThemes.MainTheme.SlideMasters.Count > 0)
        //            {
        //                slide.ThemeLayout.UpdateUI(NewEThemes.MainTheme.SlideMasters[0].MainLayer);
        //                slide.MainLayout.Fill = ColorHelper.ConverFromColorData(NewEThemes.MainTheme.SlideMasters[0].MainLayer.Background);
        //            }
        //        }
        //        else
        //        {
        //            foreach (ELayoutMaster item in NewEThemes.MainTheme.SlideMasters[0].LayoutMasters)
        //            {
        //                if (item.LayoutType == slide.SelectLayoutType)
        //                {
        //                    DesignTabControlViewModel.UpdateLayoutMaster(item, slide, NewEThemes.Clone() as EThemes);
        //                }
        //            }
        //        }
        //        slide.UpdateThemeFont();
        //        slide.MainLayout.UpdateThumbnail();
        //    }
        //    foreach (EThemes item in (Application.Current as IAppGlobal).LocalThemesCollection)
        //    {
        //        item.CountTheme = DesignTabControlViewModel.CountThemeApply(item);
        //    }
        //    //(Application.Current as IAppGlobal).DocumentControl.SelectedTheme = NewEThemes.MainTheme;
        //    //(Application.Current as IAppGlobal).DocumentControl.SelectedThemes = NewEThemes;
        //    Global.EndInit();
        //}
    }
}
