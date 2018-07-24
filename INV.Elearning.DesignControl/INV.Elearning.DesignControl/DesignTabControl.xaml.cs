using INV.Elearning.Controls;
using INV.Elearning.Core.Helper;
using INV.Elearning.DesignControl.ViewModel;
using System.Windows;
using System.Windows.Controls;

namespace INV.Elearning.DesignControl
{
    /// <summary>
    /// Interaction logic for DesignTabControl.xaml
    /// </summary>
    public partial class DesignTabControl : RibbonTabItem
    {
        public DesignTabControl()
        {
            InitializeComponent();
            DataContext = new DesignTabControlViewModel();
            DesignTabControlViewModel.ThemeDesign = themes;

        }

        private void InRibbonGallery_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((Application.Current as IAppGlobal).DocumentControl.SelectedTheme != null)
            {
                themes.SelectedValue = (Application.Current as IAppGlobal).DocumentControl.SelectedTheme;
            }         
        }

     
    }
}
