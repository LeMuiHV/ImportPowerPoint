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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace INV.Elearning.DesignControl.Views
{
    /// <summary>
    /// Interaction logic for GridTreeView.xaml
    /// </summary>
    public partial class GridTreeView : Grid
    {
        public ScrollViewer ScrollViewer { get; set; }
        public GridTreeView()
        {
            InitializeComponent();
            ScrollViewer = scroll;
        }

        private void Thumb_DragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            Width = ActualWidth;
            double dragDelta = Math.Min(-e.HorizontalChange, Width - 50);
            Width -= dragDelta;

        }
    }
}
