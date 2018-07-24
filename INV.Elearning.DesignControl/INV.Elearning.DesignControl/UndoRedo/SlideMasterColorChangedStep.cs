using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace INV.Elearning.DesignControl.UndoRedo
{
    public class SlideMasterColorChangedStep : StepBase
    {
        public EColorManagment OldColor { get; set; }
        public EColorManagment NewColor { get; set; }
        public SlideMasterColorChangedStep(EColorManagment oldColor, EColorManagment newColor)
        {
            OldColor = oldColor;
            NewColor = newColor;
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SelectedTheme.Colors = OldColor;
            if ((Application.Current as IAppGlobal).SelectedThemeView != null)
                (Application.Current as IAppGlobal).SelectedThemeView.Colors = OldColor;
            foreach (var item in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {
                item.UpdateThemeColor();
            }
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SelectedTheme.Colors = NewColor;
            if ((Application.Current as IAppGlobal).SelectedThemeView != null)
                (Application.Current as IAppGlobal).SelectedThemeView.Colors = NewColor;
            foreach (var item in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {
                item.UpdateThemeColor();
            }
            Global.EndInit();
        }
    }
}
