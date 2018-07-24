
namespace INV.Elearning.ImportPowerPoint.Model
{
    //---------------------------------------------------------------------------
    // Copyright (C)Huong Viet Group.  All rights reserved.
    // File: SlideInfo.cs
    // Description: Thông tin Slide
    // Develope by : Le Mui
    // History:
    // 10/05/2018 : Create
    //---------------------------------------------------------------------------
    public class Slide : PropertyChangedBase
    {
        private int _slideIndex;
        /// <summary>
        /// Số thư tự slide
        /// </summary>
        public int SlideIndex
        {
            get { return _slideIndex; }
            set { _slideIndex = value; }
        }

        private bool _isSelect = false;
        /// <summary>
        /// Slide đang được chọn hay không?
        /// </summary>
        public bool IsSelect
        {
            get { return _isSelect; }
            set
            {
                _isSelect = value;
                OnPropertyChanged("IsSelect");
            }
        }
        private string _thumbnail;
        /// <summary>
        /// Đường dẫn tới ảnh chụp slide
        /// </summary>
        public string Thumbnail
        {
            get { return _thumbnail; }
            set { _thumbnail = value; }
        }
    }
}
