using INV.Elearning.Controls;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.DesignControl.ViewModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace INV.Elearning.DesignControl
{
    /// <summary>
    /// Interaction logic for SlideMasterTabControl.xaml
    /// </summary>
    public partial class SlideMasterTabControl : RibbonTabItem
    {
        public SlideMasterTabControl()
        {
            InitializeComponent();
            DataContext = new SlideMasterTabControlViewModel();
            SlideMasterTabControlViewModel.ThemeDesign = themeDesign;
            IsVisibleChanged += SlideMasterTabControl_IsVisibleChanged;
        }

        private void SlideMasterTabControl_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
           //if((bool)e.NewValue && Global.UndoProcessing)
           // {
           //     Application.Current.Dispatcher.Invoke(() =>
           //     {
           //         (Application.Current as IAppGlobal).DocumentControl.Slides[0].IsSelected = true;
           //         foreach (SlideBase item in (Application.Current as IAppGlobal).DocumentControl.Slides)
           //         {
           //             item.MainLayout.UpdateThumbnail();
           //         }
           //     });
               
           // }
        }

     

        private void themeDesign_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((Application.Current as IAppGlobal).SelectedTheme != null)
            {
                themeDesign.SelectedValue = (Application.Current as IAppGlobal).SelectedTheme;
            }
        }

    }
}
