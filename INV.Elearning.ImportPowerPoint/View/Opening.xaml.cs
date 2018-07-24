    using INV.Elearning.ImportPowerPoint.Helper;
using INV.Elearning.ImportPowerPoint.Model;
using INV.Elearning.ImportPowerPoint.ViewModel;
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
using System.Windows.Threading;
using System.IO;
using INV.Elearning.Controls;
using INV.Elearning.Core.Helper;

namespace INV.Elearning.ImportPowerPoint.View
{
    /// <summary>
    /// Interaction logic for Opening.xaml
    /// </summary>
    public partial class Opening : ElearningWindow
    {
        BackgroundWorker backgroundWorker;
        public Opening()
        {
            InitializeComponent();
        }

        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (!backgroundWorker.CancellationPending)
            {
                progressbarOpen.Value = e.ProgressPercentage;
            }
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            base.Close();
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            int count = Utils.LstSlide.Count;
            ImportPowerPointView.viewmodel = new ImportPowerPointViewModel();
            int i = 1;
            foreach (Microsoft.Office.Interop.PowerPoint.Slide sld in Utils.LstSlide)
            {
                i++;
                string imgSlide = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, "Slide_" + sld.SlideIndex.ToString() + ".png");
                //---Export slide to image
                sld.Export(imgSlide, "png", 120, 80);
                (sender as BackgroundWorker).ReportProgress(i * 100 / count);
                ImportPowerPointView.viewmodel.Slides.Add(new Slide()
                {
                    IsSelect = true,
                    SlideIndex = sld.SlideIndex,
                    Thumbnail = imgSlide
                });
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
