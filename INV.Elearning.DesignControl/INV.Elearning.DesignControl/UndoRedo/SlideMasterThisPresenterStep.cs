using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model.Theme;
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
    public class SlideMasterThisPresenterStep : StepBase
    {
        //public EThemes OldEThemes { get; set; }
        //public EThemes NewEThemes { get; set; }
        //public SlideMasterThisPresenterStep(EThemes oldEThemes, EThemes newEThemes)
        //{
        //    OldEThemes = oldEThemes;
        //    NewEThemes = newEThemes;
        //}

        //public override void UndoExcute()
        //{
        //    Global.BeginInit();
        //    foreach (SlideMaster item in (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters)
        //    {
        //        if (item.ThemesName == OldEThemes.Name)
        //        {
        //            item.IsSelected = true;
        //            break;
        //        }
        //    }
        //    (Application.Current as IAppGlobal).SelectedThemes = OldEThemes;
        //    Global.EndInit();
        //}

        //public override void RedoExcute()
        //{
        //    Global.BeginInit();
        //    foreach (SlideMaster item in (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters)
        //    {
        //        if (item.ThemesName == NewEThemes.Name)
        //        {
        //            item.IsSelected = true;
        //            break;
        //        }
        //    }
        //    (Application.Current as IAppGlobal).SelectedThemes = NewEThemes;
        //    Global.EndInit();
        //}
    }
}
