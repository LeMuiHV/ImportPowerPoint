using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace INV.Elearning.DesignControl.UndoRedo
{
    public class ThisPresenterSlideMasterStep : StepBase
    {
        public ESlideMaster OldESlideMaster { get; set; }
        public ESlideMaster NewESlideMaster { get; set; }
        public ETheme ETheme { get; set; }
        public ThisPresenterSlideMasterStep(ESlideMaster oldESlideMaster, ESlideMaster newESlideMaster, ETheme eTheme)
        {
            OldESlideMaster = oldESlideMaster;
            NewESlideMaster = newESlideMaster;
            ETheme = eTheme;
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            ETheme.SlideMasters[0] = OldESlideMaster;
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            ETheme.SlideMasters[0] = NewESlideMaster;
            Global.EndInit();
        }
    }
}
