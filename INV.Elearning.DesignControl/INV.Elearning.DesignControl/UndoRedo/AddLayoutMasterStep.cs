using INV.Elearning.Core.Helper;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.Core.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace INV.Elearning.DesignControl.UndoRedo
{
    public class AddLayoutMasterStep : StepBase
    {
        public LayoutMaster LayoutMaster { get; set; }
        public SlideMaster SlideMaster { get; set; }
        public AddLayoutMasterStep(LayoutMaster layoutMaster, SlideMaster slideMaster)
        {
            LayoutMaster = layoutMaster;
            SlideMaster = slideMaster;
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            SlideMaster.LayoutMasters.Remove(LayoutMaster);
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            SlideMaster.LayoutMasters.Add(LayoutMaster);
            LayoutMaster.IsSelected = true;
            Global.EndInit();
        }
    }
}
