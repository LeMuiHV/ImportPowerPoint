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
    public class DeleteSlideMasterStep : StepBase
    {
        public SlideMaster SlideMaster { get; set; }
        public ESlideMaster ESlideMaster { get; set; }
        public ETheme ETheme { get; set; }
        public int Index { get; set; }
        public DeleteSlideMasterStep(SlideMaster slideMaster, ESlideMaster eSlide, ETheme etheme, int index)
        {
            SlideMaster = slideMaster;
            ETheme = etheme;
            Index = index;
            ESlideMaster = eSlide;
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).LocalThemesCollection.Insert(Index, ETheme);
            SlideMaster slide = new SlideMaster();
            slide.IsSelected = true;
            slide.LoadData(ESlideMaster);
            (Application.Current as IAppGlobal).DocumentControl.Slides.Add(slide);
            foreach (LayoutMaster item in slide.LayoutMasters)
            {
                (Application.Current as IAppGlobal).DocumentControl.Slides.Add(item);
            }
             (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters.Insert(Index, slide);
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster slide)
            {
                (Application.Current as IAppGlobal).LocalThemesCollection.Remove(ETheme);
                foreach (LayoutMaster item in slide.LayoutMasters)
                {
                    (Application.Current as IAppGlobal).DocumentControl.Slides.Remove(item);
                }
                (Application.Current as IAppGlobal).DocumentControl.Slides.Remove(slide);
                SlideMaster.LayoutMasters.Clear();
                (Application.Current as IAppGlobal).SelectedThemeView.SlideMasters.Remove(slide);
            }

            Global.EndInit();
        }
    }
}
