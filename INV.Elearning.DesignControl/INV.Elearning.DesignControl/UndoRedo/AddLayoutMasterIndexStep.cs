using INV.Elearning.Core.Helper;
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
    public class AddLayoutMasterIndexStep : StepBase
    {
        public LayoutMaster LayoutMaster { get; set; }
        public SlideMaster SlideMaster { get; set; }
        public int Index { get; set; }
        public AddLayoutMasterIndexStep(LayoutMaster layoutMaster, SlideMaster slideMaster, int index)
        {
            LayoutMaster = layoutMaster;
            SlideMaster = slideMaster;
            Index = index;
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            SlideMaster.LayoutMasters.Remove(LayoutMaster);
            (Application.Current as IAppGlobal).SelectedThemeView.RefreshData();
            (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0] = (Application.Current as IAppGlobal).SelectedThemeView.Data.SlideMasters[0];
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            SlideMaster.LayoutMasters.Insert(Index + 1, LayoutMaster);
            SlideHelper.UnSlectedAll();
            LayoutMaster.IsSelected = true;
            (Application.Current as IAppGlobal).SelectedThemeView.RefreshData();
            (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0] = (Application.Current as IAppGlobal).SelectedThemeView.Data.SlideMasters[0];
            Global.EndInit();
        }
    }
}
