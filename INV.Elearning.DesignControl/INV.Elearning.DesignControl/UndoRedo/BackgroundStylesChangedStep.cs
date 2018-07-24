using INV.Elearning.Core;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.Core.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace INV.Elearning.DesignControl.UndoRedo
{
    //public class BackgroundStylesChangedStep : StepBase
    //{
    //    public ColorBrushBase OldBackground { get; set; }
    //    public ColorBrushBase NewBackground { get; set; }
    //    public SlideBase Slide { get; set; }
    //    public BackgroundStylesChangedStep(ColorBrushBase oldBackground, ColorBrushBase newBackground, SlideBase source)
    //    {
    //        OldBackground = oldBackground;
    //        NewBackground = newBackground;
    //        Slide = source;
    //    }

    //    public override void UndoExcute()
    //    {
    //        Global.BeginInit();
    //        if (Slide is LayoutMaster layoutMaster)
    //        {
    //            layoutMaster.MainLayout.Fill = OldBackground;
    //            layoutMaster.MainLayout.UpdateThumbnail();
    //        }
    //        else if (Slide is SlideMaster slideMaster)
    //        {
    //            slideMaster.MainLayout.Fill = OldBackground;
    //            slideMaster.MainLayout.UpdateThumbnail();

    //            slideMaster.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
    //            {
    //                (Application.Current as IAppGlobal).SelectedTheme.ThumbnailBitmap = FileHelper.SaveToThumbnailIcon(slideMaster.MainLayout, (Application.Current as IAppGlobal).SelectedTheme.Colors, new Size(341, 192));
    //                (Application.Current as IAppGlobal).SelectedTheme.ThumbnailBitmapSmall = FileHelper.SaveToThumbnailIconSmall(slideMaster.MainLayout, (Application.Current as IAppGlobal).SelectedTheme.Colors);
    //            }));

    //            foreach (LayoutMaster item in slideMaster.LayoutMasters)
    //            {
    //                item.MainLayout.Fill = OldBackground;
    //                item.MainLayout.UpdateThumbnail();
    //            }
    //        }
    //        Global.EndInit();
    //    }

    //    public override void RedoExcute()
    //    {
    //        Global.BeginInit();
    //        if (Slide is LayoutMaster layoutMaster)
    //        {
    //            layoutMaster.MainLayout.Fill = NewBackground;
    //            layoutMaster.MainLayout.UpdateThumbnail();
    //        }
    //        else if (Slide is SlideMaster slideMaster)
    //        {
    //            slideMaster.MainLayout.Fill = NewBackground;
    //            slideMaster.MainLayout.UpdateThumbnail();

    //            slideMaster.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
    //            {
    //                (Application.Current as IAppGlobal).SelectedTheme.ThumbnailBitmap = FileHelper.SaveToThumbnailIcon(slideMaster.MainLayout, (Application.Current as IAppGlobal).SelectedTheme.Colors, new Size(341, 192));
    //                (Application.Current as IAppGlobal).SelectedTheme.ThumbnailBitmapSmall = FileHelper.SaveToThumbnailIconSmall(slideMaster.MainLayout, (Application.Current as IAppGlobal).SelectedTheme.Colors);
    //            }));

    //            foreach (LayoutMaster item in slideMaster.LayoutMasters)
    //            {
    //                item.MainLayout.Fill = NewBackground;
    //                item.MainLayout.UpdateThumbnail();
    //            }
    //        }
    //        Global.EndInit();
    //    }
    //}
}
