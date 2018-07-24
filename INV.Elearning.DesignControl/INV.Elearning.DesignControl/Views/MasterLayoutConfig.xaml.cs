using INV.Elearning.Controls;
using INV.Elearning.Core.View.Theme;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace INV.Elearning.DesignControl.Views
{
    /// <summary>
    /// Interaction logic for MasterLayoutConfig.xaml
    /// </summary>
    public partial class MasterLayoutConfig : ElearningWindow
    {
        SlideMaster slideMaster;
        public MasterLayoutConfig(SlideMaster slideMaster)
        {
            InitializeComponent();
            this.slideMaster = slideMaster;
            cbTitle.IsChecked = slideMaster.IsTitle;
            cbText.IsChecked = slideMaster.IsText;
            cbDate.IsChecked = slideMaster.IsDate;
            cbSlideNumber.IsChecked = slideMaster.IsSlideNumber;
            cbFooters.IsChecked = slideMaster.IsFooter;
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            slideMaster.IsText = cbText.IsChecked.HasValue ? cbText.IsChecked.Value : false;
            slideMaster.IsTitle = cbTitle.IsChecked.HasValue ? cbTitle.IsChecked.Value : false;
            slideMaster.IsDate = cbDate.IsChecked.HasValue ? cbDate.IsChecked.Value : false;
            slideMaster.IsSlideNumber = cbSlideNumber.IsChecked.HasValue ? cbSlideNumber.IsChecked.Value : false;
            slideMaster.IsFooter = cbFooters.IsChecked.HasValue ? cbFooters.IsChecked.Value : false;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
