using DocumentFormat.OpenXml.Packaging;
using INV.Elearning.Controls;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.ImportPowerPoint.Helper;
using INV.Elearning.ImportPowerPoint.Process;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace INV.Elearning.ImportPowerPoint.View
{
    /// <summary>
    /// Interaction logic for Converting.xaml
    /// </summary>
    public partial class Importing : ElearningWindow
    {
        BackgroundWorker backgroundWorker;
        ElearningDocument documentMain = null;
        string presentationFile = string.Empty;
        bool isBackground;
        Presentation presentation;
        public Importing()
        {
            InitializeComponent();
        }
        public Importing(ElearningDocument _documentMain, Presentation _presentation, string _presentationFile, bool _isBackground)
        {
            InitializeComponent();
            this.presentation = _presentation;
            this.documentMain = _documentMain;
            this.presentationFile = _presentationFile;
            this.isBackground = _isBackground;
        }
        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (!backgroundWorker.CancellationPending)
            {
                progressbarConvert.Value = e.ProgressPercentage;
            }
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            base.Close();
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            int slideCount = 1;
            ConvertPresentation cvPresentation = new ConvertPresentation();
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                HelperClass helperClass = new HelperClass();
                Utils.GetPresentationPart(presentationDocument);
                #region GetTheme
                var ss = presentation.SlideMaster.Theme;
                int count = 0;

                foreach (Design item in presentation.Designs)
                {
                    ETheme eTheme = new HelperClass().GetETheme(item, count++);
                    documentMain.ListETheme.Add(eTheme);
                }
                documentMain.SelectedTheme = documentMain.ListETheme.FirstOrDefault();
                #endregion

                foreach (Slide slide in Helper.Utils.LstSlideSelected)
                {
                    NormalPage page = null;
                    //--Chuyển slide PP thành ảnh nền trang
                    if (isBackground)
                    {
                        page = new NormalPage();
                        page.MainLayer = new PageLayer();
                        page.MainLayer.ID = ObjectElementsHelper.RandomString(13);
                        page.MainLayer.ThemeLayoutOwnerID = ObjectElementsHelper.RandomString(13);
                        page.CanShowInMenu = true;
                        page.PageConfig = new PageConfig();
                        page.PageConfig.PreviousButtonEnable = true;
                        cvPresentation.SlideAsBackground(page, slide);
                        slideCount++;
                    }
                    else
                    {
                        page = helperClass.GetNormalPage(slide, helperClass.GetSlidePart(slide.SlideNumber - 1));
                        page.IDLayout = documentMain.SelectedTheme.SlideMasters[0].LayoutMasters.FirstOrDefault(x => x.SlideName == slide.CustomLayout.MatchingName)?.ID;
                        page.IsHideBackground = slide.DisplayMasterShapes != MsoTriState.msoTrue;
                        slideCount++;
                    }
                (sender as BackgroundWorker).ReportProgress(slideCount * 100 / Helper.Utils.LstSlideSelected.Count());
                    documentMain.Pages.Add(page);
                }
            }
        }

        private void ElearningWindow_ContentRendered(object sender, EventArgs e)
        {
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;//Cho phép báo cáo tiến trình
            backgroundWorker.WorkerSupportsCancellation = true;//Cho phép dừng tiến trình
            backgroundWorker.DoWork += BackgroundWorker_DoWork;
            backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
            backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
            backgroundWorker.RunWorkerAsync();
        }
    }
}
