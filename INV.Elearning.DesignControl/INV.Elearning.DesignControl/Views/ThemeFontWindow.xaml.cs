

namespace INV.Elearning.DesignControl.Views
{
    using INV.Elearning.Controls;
    using INV.Elearning.Core.Helper;
    using INV.Elearning.Core.Model.Theme;

    /// <summary>
    /// Interaction logic for ThemeFontWindow.xaml
    /// </summary>
    public partial class ThemeFontWindow : ElearningWindow
    {
        private RelayCommand _saveCommand = null;
        private RelayCommand _cancelCommand = null;
        private bool _isCancelled = true;
        private EFontfamily _themeFontFamily = null;

        /// <summary>
        /// Lệnh điều khiển lúc bấm lưu
        /// </summary>
        public RelayCommand SaveCommand { get => _saveCommand ?? (_saveCommand = new RelayCommand(o => ExitExcute(false), p => !string.IsNullOrWhiteSpace((this.DataContext as EFontfamily)?.Name))); }
        /// <summary>
        /// Lệnh điều khiển lúc bấm hủy
        /// </summary>
        public RelayCommand CancelCommand { get => _cancelCommand ?? (_cancelCommand = new RelayCommand(o => ExitExcute())); }
        /// <summary>
        /// Lấy thuộc tính đã hủy dữ liệu hay chưa
        /// </summary>
        public bool IsCancelled { get => _isCancelled; }
        /// <summary>
        /// Lấy thuộc tính phông chữ cần tạo
        /// </summary>
        public EFontfamily ThemeFontFamily { get => _themeFontFamily; }

        private void ExitExcute(bool isCancelled = true)
        {
            this._isCancelled = isCancelled;
            this.Close();
        }

        /// <summary>
        /// Hàm khởi tạo
        /// </summary>
        public ThemeFontWindow()
        {
            InitializeComponent();
            this.DataContext = this._themeFontFamily = new EFontfamily() { TagName = "Customize" };
        }
    }
}
