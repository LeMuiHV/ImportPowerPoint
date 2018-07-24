using INV.Elearning.ImportPowerPoint.Model;
using System;
using System.Windows;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace INV.Elearning.ImportPowerPoint.ViewModel
{
    public class ImportPowerPointViewModel : PropertyChangedBase
    {
        public Action CloseForm { get; set; }
        private RelayCommand _cancelCommand = null;
        public RelayCommand CancelCommand
        {
            get { return _cancelCommand ?? new RelayCommand(o => CancelExecute()); }
        }
        private void CancelExecute()
        {
            CloseForm();
        }
        /// <summary>
        /// Command Select All Slide
        /// </summary>
        private RelayCommand _selectAllCommand = null;
        public RelayCommand SelectAllCommand
        {
            get { return _selectAllCommand ?? new RelayCommand(o => SelectAllExecute(true)); }
        }
        /// <summary>
        /// Command Deselect All Slide
        /// </summary>
        private RelayCommand _deSelectAllCommand = null;
        public RelayCommand DeSelectAllCommand
        {
            get { return _deSelectAllCommand ?? new RelayCommand(o => SelectAllExecute(false)); }
        }
        private void SelectAllExecute(bool obj)
        {
            foreach (var sld in Slides)
            {
                sld.IsSelect = obj;
            }
            OnPropertyChanged("NumberSelected");
        }
        private RelayCommand _itemUpdateCommand;

        public RelayCommand ItemUpdateCommand
        {
            get { return _itemUpdateCommand ?? new RelayCommand(o => ItemUpdateExcute()); }
        }
        /// <summary>
        /// Xử lý cập nhật lúc chọn vào đối tượng Slide
        /// </summary>
        private void ItemUpdateExcute()
        {
            OnPropertyChanged("NumberSelected");
        }
        private ObservableCollection<Slide> _slides = new ObservableCollection<Slide>();
        /// <summary>
        /// Danh sách Slide truyền vào
        /// </summary>
        public ObservableCollection<Slide> Slides
        {
            get { return _slides; }
        }

        private string _numberSelected;
        /// <summary>
        /// Nội dung label
        /// </summary>
        public string NumberSelected
        {
            get
            {
                if(_slides!=null)
                _lstSelected = _slides.Where(o=>o.IsSelect).ToList();
                int count = _lstSelected.Count;
                var content = GetContent("ImportPP_PageSelected");
                _numberSelected = string.Format(content, count, _slides.Count);
                return _numberSelected;
            }
        }
        private List<Slide> _lstSelected = new List<Slide>();
        /// <summary>
        /// Danh sách slide đang được chọn
        /// </summary>
        public List<Slide> LstSelected
        {
            get
            {
                return _lstSelected;
            }
        }
        /// <summary>
        /// Kiểm tra và thêm Resource
        /// </summary>
        private void CheckAndAddLanguageResource()
        {
            if (Application.Current == null) return;
            var _content = (Application.Current.TryFindResource("ImportPP_Title")) as string;
            if (_content == null)
            {
                Application.Current.Resources.MergedDictionaries.Add(new ResourceDictionary() { Source = new Uri(@"pack://application:,,,/INV.Elearning.ImportPowerPoint;Component/Resources/Language/en-us/Language.xaml", UriKind.RelativeOrAbsolute) });
            }
        }
        /// <summary>
        /// Lấy nội dung trong file Language.xaml
        /// </summary>
        /// <param name="_key"></param>
        /// <returns></returns>
        private string GetContent(string _key)
        {
            string content = (Application.Current.TryFindResource(_key)) as string;
            return content;
        }
        public ImportPowerPointViewModel()
        {
            CheckAndAddLanguageResource();
        }
    }
}
