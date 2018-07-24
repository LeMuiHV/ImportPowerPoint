using System;
using System.Windows;
using System.IO;
using INV.Elearning.Controls;
using INV.Elearning.ImportPowerPoint.Helper;
using INV.Elearning.ImportPowerPoint.ViewModel;
using pp = Microsoft.Office.Interop.PowerPoint;
using office = Microsoft.Office.Core;
using INV.Elearning.Core.Model;
using INV.Elearning.ImportPowerPoint.Process;
using INV.Elearning.Core.Helper;
using System.Diagnostics;
using Microsoft.Win32;
using DocumentFormat.OpenXml.Packaging;
using INV.Elearning.Core.Model.Theme;
using Microsoft.Office.Core;
using System.Collections.ObjectModel;
using System.Linq;

namespace INV.Elearning.ImportPowerPoint.View
{
    /// <summary>
    /// Interaction logic for ImportPowerPointView.xaml
    /// </summary>
    public partial class ImportPowerPointView : ElearningWindow
    {
        pp.Application app = null;
        pp.Presentation presentation = null;
        ElearningDocument _documentMain = null;
        public ObservableCollection<ETheme> listETheme = new ObservableCollection<ETheme>();
        public static bool loaded = false;
        public static ImportPowerPointViewModel viewmodel;
        private string presentationFile = string.Empty;

        public ImportPowerPointView(string _fileName)
        {
            HelperClass helperClass = new HelperClass();
            InitializeComponent();
            _documentMain = new ElearningDocument();
            app = new pp.Application();
            presentationFile = _fileName;
            presentation = app.Presentations.Open(_fileName, office.MsoTriState.msoTrue, office.MsoTriState.msoTrue, office.MsoTriState.msoFalse);
            if (System.IO.Path.GetExtension(_fileName) == ".ppt")
            {
                presentationFile = System.IO.Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, DateTime.Now.Ticks.ToString() + ".pptx");
                presentation.SaveCopyAs(presentationFile, pp.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, office.MsoTriState.msoCTrue);
            }
            Utils.tlw = 1024 / presentation.PageSetup.SlideWidth;
            Utils.tlh = 576 / presentation.PageSetup.SlideHeight;
            Utils.LstSlide = presentation.Slides;


            Opening oForm = new Opening();
            oForm.ShowDialog();
            this.DataContext = viewmodel;
            if (DataContext != null)
            {
                if ((DataContext as ImportPowerPointViewModel).CloseForm == null)
                {
                    (DataContext as ImportPowerPointViewModel).CloseForm = new Action(Close);
                }
            }
        }


        /// <summary>
        /// Chèn vào tài liệu
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OK_Click(object sender, RoutedEventArgs e)
        {
            //---Lấy danh sách slide được chọn
            Utils.LstSlideSelected.Clear();
            foreach (Model.Slide slideSelected in ((ImportPowerPointViewModel)(this.DataContext)).LstSelected)
            {
                foreach (pp.Slide slide in Utils.LstSlide)
                {
                    if (slideSelected.SlideIndex == slide.SlideIndex)
                        Utils.LstSlideSelected.Add(slide);
                }
            }
            Importing importing = new Importing(_documentMain, presentation, presentationFile, (bool)rdBackground.IsChecked);
            importing.ShowDialog();
            this.Close();
        }
        /// <summary>
        /// Trả về Elearningdocument và hiển thị lên phần mềm
        /// </summary>
        /// <returns></returns>
        public ElearningDocument ShowAndReturnData()
        {
            (this as Window).ShowDialog();
            return _documentMain;
        }
        /// <summary>
        /// Giải phóng tài nguyên
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ElearningWindow_Closed(object sender, EventArgs e)
        {
            if (presentation != null) presentation.Close();
            if (app != null) app.Quit();
            GC.Collect();
            GC.SuppressFinalize(this);
        }



        #region Check Office Version
        private const string RegKey = @"Software\Microsoft\Windows\CurrentVersion\App Paths";
        /// <summary>
        /// gets the component's path from the registry. if it can't find it - retuns an empty string
        /// </summary>
        /// <param name="_component"></param>
        /// <returns></returns>
        private string GetComponentPath()
        {
            string toReturn = string.Empty;
            string _key = "powerpnt.exe";
            //looks inside CURRENT_USER:
            RegistryKey _mainKey = Registry.CurrentUser;
            try
            {
                _mainKey = _mainKey.OpenSubKey(RegKey + "\\" + _key, false);
                if (_mainKey != null)
                {
                    toReturn = _mainKey.GetValue(string.Empty).ToString();
                }
            }
            catch
            { }
            //if not found, looks inside LOCAL_MACHINE:
            _mainKey = Registry.LocalMachine;
            if (string.IsNullOrEmpty(
                _toReturn))
            {
                try
                {
                    _mainKey = _mainKey.OpenSubKey(RegKey + "\\" + _key, false);
                    if (_mainKey != null)
                    {
                        toReturn = _mainKey.GetValue(string.Empty).ToString();
                    }
                }
                catch
                {
                }
            }
            //closing the handle:
            if (_mainKey != null)
                _mainKey.Close();
            return toReturn;
        }
        /// <summary>
        /// Gets the major version of the path. if file not found (or any other exception occures
        /// - returns 0
        /// </summary>
        private int GetMajorVersion(string _path)
        {
            int toReturn = 0;
            if (File.Exists(_path))
            {
                try
                {
                    FileVersionInfo _fileVersion = FileVersionInfo.GetVersionInfo(_path);
                    toReturn = _fileVersion.FileMajorPart;
                }
                catch
                { }
            }

            return toReturn;
        }
        public int OfficeVersion()
        {
            int version = 11;
            version = GetMajorVersion(GetComponentPath());
            return version;
        }
        public string _toReturn { get; set; }
        #endregion
    }
}
