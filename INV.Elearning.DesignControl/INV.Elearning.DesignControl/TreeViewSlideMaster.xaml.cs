using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.DesignControl.ViewModel;
using INV.Elearning.DesignControl.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    /// Interaction logic for TreeViewSlideMaster.xaml
    /// </summary>
    public partial class TreeViewSlideMaster : UserControl
    {
        public TreeViewSlideMaster()
        {
            InitializeComponent();
            DataContext = new SlideMasterTabControlViewModel();
        }

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is SlideMaster slideMaster)
            {
                AllIsSelected();
                slideMaster.IsSelected = true;
            }
            else if (e.NewValue is LayoutMaster layoutMaster)
            {
                AllIsSelected();
                layoutMaster.IsSelected = true;
            }
        }

        private void AllIsSelected()
        {
            foreach (SlideMaster item in ((Application.Current as IAppGlobal).SelectedThemeView.SlideMasters))
            {
                item.IsSelected = false;
                foreach (LayoutMaster item1 in item.LayoutMasters)
                {
                    item.IsSelected = false;
                }
            }
        }

        private void TreeView_SelectedItemChanged_1(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            SlideMaster slide = null;
            if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster slideMaster)
            {
                slide = slideMaster;
            }
            else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster layoutMaster)
            {
                slide = layoutMaster.SlideParent;
            }
            if (slide != null)
            {
                foreach (ETheme item in (Application.Current as IAppGlobal).LocalThemesCollection)
                {
                    if (slide.ID == item.ID)
                    {
                        (Application.Current as IAppGlobal).SelectedTheme = item;
                    }
                }

                ICollectionView view = CollectionViewSource.GetDefaultView((Application.Current as IAppGlobal).LocalThemesCollection);
                view.Refresh();
            }
            INV.Elearning.Timeline.Timeline.TimelineViewModel.UpdateTimeline();


        }


    }
}
