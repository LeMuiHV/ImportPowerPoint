using INV.Elearning.Controls;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using System.Windows;

namespace INV.Elearning.DesignControl.Views
{
    /// <summary>
    /// Interaction logic for ThemeColorWindow.xaml
    /// </summary>
    public partial class ThemeColorWindow : ElearningWindow
    {
        private bool _isCancelled = true;
        private RelayCommand _saveCommand = null;
        private RelayCommand _cancelCommand = null;
        private RelayCommand _resetCommand = null;

        #region Colors

        /// <summary>
        /// Các quản lý màu
        /// </summary>
        public EColorManagment Colors
        {
            get { return (EColorManagment)GetValue(ColorsProperty); }
            set { SetValue(ColorsProperty, value); }
        }

        /// <summary>
        /// Hàm hủy
        /// </summary>
        public bool IsCancelled { get => _isCancelled; }

        // Using a DependencyProperty as the backing store for Colors.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty ColorsProperty =
            DependencyProperty.Register("Colors", typeof(EColorManagment), typeof(ThemeColorWindow), new PropertyMetadata(null));


        /// <summary>
        /// Lệnh điều khiển lúc bấm lưu
        /// </summary>
        public RelayCommand SaveCommand { get => _saveCommand ?? (_saveCommand = new RelayCommand(o => ExitExcute(false), p => !string.IsNullOrWhiteSpace(this.Colors?.Name))); }

        /// <summary>
        /// Thoát
        /// </summary>
        /// <param name="isCancelled"></param>
        private void ExitExcute(bool isCancelled = true)
        {
            this._isCancelled = isCancelled;
            this.Close();
        }

        /// <summary>
        /// Lệnh điều khiển lúc bấm hủy
        /// </summary>
        public RelayCommand CancelCommand { get => _cancelCommand ?? (_cancelCommand = new RelayCommand(o => ExitExcute())); }

        /// <summary>
        /// So sánh dữ liệu
        /// </summary>
        /// <returns></returns>
        private bool EqualData()
        {
            if (this.DataContext is EColorManagment && this.Colors != null)
            {
                var _rootColors = this.DataContext as EColorManagment;
                if (_rootColors.Accent1.Color != Colors.Accent1.Color) //So sánh màu
                    return false;
                if (_rootColors.Accent2.Color != Colors.Accent2.Color) //So sánh màu
                    return false;
                if (_rootColors.Accent3.Color != Colors.Accent3.Color) //So sánh màu
                    return false;
                if (_rootColors.Accent4.Color != Colors.Accent4.Color) //So sánh màu
                    return false;
                if (_rootColors.Accent5.Color != Colors.Accent5.Color) //So sánh màu
                    return false;
                if (_rootColors.Accent6.Color != Colors.Accent6.Color) //So sánh màu
                    return false;
                if (_rootColors.BackgroundDark1.Color != Colors.BackgroundDark1.Color) //So sánh màu
                    return false;
                if (_rootColors.BackgroundDark2.Color != Colors.BackgroundDark2.Color) //So sánh màu
                    return false;
                if (_rootColors.BackgroundLight1.Color != Colors.BackgroundLight1.Color) //So sánh màu
                    return false;
                if (_rootColors.BackgroundLight2.Color != Colors.BackgroundLight2.Color) //So sánh màu
                    return false;
                if (_rootColors.Hyperlink.Color != Colors.Hyperlink.Color) //So sánh màu
                    return false;
                if (_rootColors.FollowedHyperlink.Color != Colors.FollowedHyperlink.Color) //So sánh màu
                    return false;

            }

            return true;
        }

        /// <summary>
        /// Lệnh điều khiển làm lại
        /// </summary>
        public RelayCommand ResetCommand { get => _resetCommand ?? (_resetCommand = new RelayCommand(o => ResetExcute(), p => !EqualData())); }

        /// <summary>
        /// Thực thi lệnh reset
        /// </summary>
        private void ResetExcute()
        {
            if (this.DataContext is EColorManagment)
            {
                this.Colors = (this.DataContext as EColorManagment).Clone() as EColorManagment;
                this.Colors.Name = "Customs Colors";
            }
        }

        #endregion

        /// <summary>
        /// Hàm khởi tạo mặc định
        /// </summary>
        public ThemeColorWindow()
        {
            InitializeComponent();
            ResetExcute();
        }
    }

}
