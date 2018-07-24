using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
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
    public class ColorChangedStep : StepBase
    {
        //public EColorManagment OldColor { get; set; }
        //public EColorManagment NewColor { get; set; }
        //public ColorChangedStep(EColorManagment oldColor, EColorManagment newColor)
        //{
        //    OldColor = oldColor;
        //    NewColor = newColor;
        //}

        //public override void UndoExcute()
        //{
        //    Global.BeginInit();
        //    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = (Application.Current as IAppGlobal).SelectedTheme.Colors = OldColor;
        //    (Application.Current as IAppGlobal).SelectedTheme.Colors = OldColor;
        //    SlideMaster slideMaster = new SlideMaster();
        //    slideMaster.ThemeLayout = new ThemeLayout();
        //    slideMaster.ThemeLayout.UpdateUI((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer);
        //    slideMaster.UpdateThemeColor();

        //    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.ThumbnailBitmap = DesignTabControlViewModel.GetThumbnail(slideMaster.MainLayout, OldColor, new Size(341, 192));

        //    foreach (ELayoutMaster item in (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters)
        //    {
        //        Task.Factory.StartNew(() =>
        //        {
        //            Application.Current.Dispatcher.Invoke(() =>
        //            {
        //                LayoutMaster layoutMaster = new LayoutMaster();
        //                layoutMaster.LoadData(item);
        //                if (!layoutMaster.IsHideBackground)
        //                {
        //                    if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();
        //                    layoutMaster.ThemeLayout.UpdateUIB((Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].MainLayer);
        //                }
        //                layoutMaster.UpdateThemeColor();
        //                item.ThumbnailBitmap = DesignTabControlViewModel.CaptureLayout(layoutMaster.MainLayout, new Size(341, 192));
        //            });
        //        });
        //    }
        //    foreach (var item in (Application.Current as IAppGlobal).DocumentControl.Slides)
        //    {
        //        item.UpdateThemeColor();
        //    }
        //    foreach (var item in (Application.Current as IAppGlobal).DocumentControl.Slides)
        //    {
        //        item.UpdateThemeColor();
        //    }
        //    Global.EndInit();
        //}

        //public override void RedoExcute()
        //{
        //    Global.BeginInit();
        //    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.Colors = (Application.Current as IAppGlobal).SelectedTheme.Colors = NewColor;
        //    (Application.Current as IAppGlobal).SelectedTheme.Colors = NewColor;
        //    SlideMaster slideMaster = new SlideMaster();
        //    slideMaster.ThemeLayout = new ThemeLayout();
        //    slideMaster.ThemeLayout.UpdateUI((Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SlideMasters[0].MainLayer);
        //    slideMaster.UpdateThemeColor();

        //    (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.ThumbnailBitmap = DesignTabControlViewModel.GetThumbnail(slideMaster.MainLayout, NewColor, new Size(341, 192));

        //    foreach (ELayoutMaster item in (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters)
        //    {
        //        Task.Factory.StartNew(() =>
        //        {
        //            Application.Current.Dispatcher.Invoke(() =>
        //            {
        //                LayoutMaster layoutMaster = new LayoutMaster();
        //                layoutMaster.LoadData(item);
        //                if (!layoutMaster.IsHideBackground)
        //                {
        //                    if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();
        //                    layoutMaster.ThemeLayout.UpdateUIB((Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].MainLayer);
        //                }
        //                layoutMaster.UpdateThemeColor();
        //                item.ThumbnailBitmap = DesignTabControlViewModel.CaptureLayout(layoutMaster.MainLayout, new Size(341, 192));
        //            });
        //        });
        //    }

        //    foreach (var item in (Application.Current as IAppGlobal).DocumentControl.Slides)
        //    {
        //        item.UpdateThemeColor();
        //    }
        //    Global.EndInit();
        //}
    }
}
