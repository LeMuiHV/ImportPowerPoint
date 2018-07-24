using INV.Elearning.Controls;
using INV.Elearning.ImportPowerPoint.ViewModel;
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

namespace INV.Elearning.ImportPowerPoint.View
{
    /// <summary>
    /// Interaction logic for NotifyOfficeVersion.xaml
    /// </summary>
    public partial class NotifyOfficeVersion : ElearningWindow
    {
        public NotifyOfficeVersion()
        {
            InitializeComponent();
        }
        private void OK_Click(object sender, RoutedEventArgs e)
        {
            base.Close();
        }
    }
}
