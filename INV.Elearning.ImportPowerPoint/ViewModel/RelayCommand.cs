namespace INV.Elearning.ImportPowerPoint.ViewModel
{
    using System;
    using System.Windows;
    using System.Windows.Input;

    /// <summary>
    /// Lớp Command tùy chỉnh, sử dụng trong các đối tượng cần Binding Command
    /// </summary>
    public class RelayCommand : ICommand
    {
        private Action<object> execute;
        private Action<object, IInputElement> executeWithTarget;
        private Func<object, bool> canExecute;

        /// <summary>
        /// Sự kiện thay đổi việc có thể thực thi command hay không
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        /// <summary>
        /// Khởi tạo một command
        /// </summary>
        /// <param name="execute">Hàm thực thi</param>
        /// <param name="canExecute">Hàm kiểm tra việc có thể thực thi hàm</param>
        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            this.execute = execute;
            this.canExecute = canExecute;
        }

        /// <summary>
        /// Khởi tạo hàm với tham số thứ 2 có thể không khai báo
        /// </summary>
        /// <param name="execute">Hàm thực thi</param>
        /// <param name="canExecute">Hàm kiểm tra việc có thể thực thi hàm</param>
        public RelayCommand(Action<object, IInputElement> execute, Func<object, bool> canExecute = null)
        {
            this.executeWithTarget = execute;
            this.canExecute = canExecute;
        }

        /// <summary>
        /// Phương thức kiểm tra có thể thực thi command đối với một đối tượng
        /// </summary>
        /// <param name="parameter">Đối tượng cần kiểm tra</param>
        /// <returns></returns>
        public bool CanExecute(object parameter)
        {
            return this.canExecute == null || this.canExecute(parameter);
        }

        /// <summary>
        /// Thực thi command với một đối tượng
        /// </summary>
        /// <param name="parameter">Đối tượng cần thực thi</param>
        public void Execute(object parameter)
        {
            this.execute(parameter);
        }

        /// <summary>
        /// Thực thi command với một đối tượng
        /// </summary>
        /// <param name="parameter"></param>
        /// <param name="target"></param>
        public void Execute(object parameter, IInputElement target)
        {
            this.executeWithTarget(parameter, target);
        }
    }
}
