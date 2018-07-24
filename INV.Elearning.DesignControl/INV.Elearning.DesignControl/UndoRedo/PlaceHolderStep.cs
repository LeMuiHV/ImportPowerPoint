using INV.Elearning.Core.Helper;
using INV.Elearning.Core.View;
using INV.Elearning.Core.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace INV.Elearning.DesignControl.UndoRedo
{
    public class PlaceHolderStep : StepBase
    {
        public ObjectElement OldItem { get; set; }
        public ObjectElement NewItem { get; set; }
        public SlideBase Slide { get; set; }
        public PlaceHolderStep(ObjectElement oldItem, ObjectElement newItem, SlideBase slide)
        {
            OldItem = oldItem;
            NewItem = newItem;
            Slide = slide;
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            Slide.IsSelected = true;
            Slide.MainLayout.Elements.Remove(NewItem);
            Slide.MainLayout.Elements.Add(OldItem);
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            Slide.IsSelected = true;
            Slide.MainLayout.Elements.Remove(OldItem);
            Slide.MainLayout.Elements.Add(NewItem);
            Global.EndInit();
        }
    }
}
