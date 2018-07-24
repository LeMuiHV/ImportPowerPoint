using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model.Theme;
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
    public class AddSlideMasterStep : StepBase
    {
        public SlideBase OldSlideMaster { get; set; }
        public SlideMaster NewSlideMaster { get; set; }
        public ETheme OldTheme { get; set; }
        public ETheme NewTheme { get; set; }
        public AddSlideMasterStep(SlideBase oldSlideMaster, SlideMaster newSlideMaster, ETheme newTheme, ETheme oldTheme)
        {
            OldSlideMaster = oldSlideMaster;
            NewSlideMaster = newSlideMaster;
            NewTheme = newTheme;
            OldTheme = oldTheme;
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters.Remove(NewSlideMaster);
            (Application.Current as IAppGlobal).LocalThemesCollection.Remove(NewTheme);
            SlideHelper.UnSlectedAll();
            (Application.Current as IAppGlobal).SelectedSlide = OldSlideMaster;
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters.Add(NewSlideMaster);
            (Application.Current as IAppGlobal).LocalThemesCollection.Add(OldTheme);
            //(Application.Current as IAppGlobal).SelectedTheme = NewTheme;
            SlideHelper.UnSlectedAll();
            NewSlideMaster.IsSelected = true;
            Global.EndInit();
        }
    }
}
