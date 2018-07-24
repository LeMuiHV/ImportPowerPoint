using INV.Elearning.Controls;
using INV.Elearning.Core.Helper;
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
    /// Interaction logic for RenameWindow.xaml
    /// </summary>
    public partial class RenameWindow : ElearningWindow
    {
        public RenameWindow(string layoutName)
        {
            InitializeComponent();
            txtName.Text = layoutName;
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            (Application.Current as IAppGlobal).SelectedSlide.SlideName = txtName.Text;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
