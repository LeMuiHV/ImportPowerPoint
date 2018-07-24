using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model.Theme;
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
    public class FontsStep : StepBase
    {
        public EFontfamily OldFont { get; set; }
        public EFontfamily NewFont { get; set; }
        public FontsStep(EFontfamily oldFont, EFontfamily newFont)
        {
            OldFont = oldFont;
            NewFont = newFont;
        }

        public override void UndoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SelectedTheme.SelectedFont = OldFont;
            foreach (SlideBase item in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {
                item.UpdateThemeFont();
                item.MainLayout.UpdateThumbnail();
            }
            Global.EndInit();
        }

        public override void RedoExcute()
        {
            Global.BeginInit();
            (Application.Current as IAppGlobal).SelectedTheme.SelectedFont = NewFont;
            foreach (SlideBase item in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {
                item.UpdateThemeFont();
                item.MainLayout.UpdateThumbnail();
            }
            Global.EndInit();
        }
    }
}
