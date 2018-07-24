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
    public class DeleteLayoutMasterStep : StepBase
    {
        public LayoutMaster LayoutMaster { get; set; }
        public int Index { get; set; }
        public DeleteLayoutMasterStep(LayoutMaster newLayoutMaster, int index)
        {
            LayoutMaster = newLayoutMaster;
            Index = index;
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            LayoutMaster.SlideParent.LayoutMasters.Insert(Index, LayoutMaster);
            (Application.Current as IAppGlobal).DocumentControl.Slides.Insert(Index, LayoutMaster);
            LayoutMaster.RefreshData();
            (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters.Insert(Index, LayoutMaster.Data as ELayoutMaster);
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            string idLayoutMaster = LayoutMaster.ID;
            LayoutMaster.SlideParent.LayoutMasters.Remove(LayoutMaster);
            (Application.Current as IAppGlobal).DocumentControl.Slides.Remove(LayoutMaster);
            for (int i = 0; i < (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters.Count; i++)
            {
                if (idLayoutMaster == (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters[i].ID)
                {
                    (Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters.Remove((Application.Current as IAppGlobal).SelectedTheme.SlideMasters[0].LayoutMasters[i]);
                    i--;
                }
            }
            Global.EndInit();
        }
    }
}
