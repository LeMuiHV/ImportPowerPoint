using INV.Elearning.Core.Helper;
using INV.Elearning.Core.View;
using INV.Elearning.Core.ViewModel;
using System.Windows;

namespace INV.Elearning.DesignControl.UndoRedo
{
    public class NormalViewModeStep : StepBase
    {
        public NormalViewModeStep()
        { 
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SlideViewMode = SlideViewMode.SlideMaster;
            SlideHelper.UnSlectedAll();
            (Application.Current as IAppGlobal).DocumentControl.Slides[0].IsSelected = true;               
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SlideViewMode = SlideViewMode.Normal;
            SlideHelper.UnSlectedAll();
            (Application.Current as IAppGlobal).DocumentControl.Slides[0].IsSelected = true;
            Global.EndInit();
        }
    }
}
