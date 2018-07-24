using INV.Elearning.Controls;
using INV.Elearning.Core;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.View;
using INV.Elearning.DesignControl.ViewModel;
using INV.Elearning.Text;
using INV.Elearning.Text.Views;
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
using System.Windows.Shapes;

namespace INV.Elearning.DesignControl.Views
{
    /// <summary>
    /// Interaction logic for HeaderAndFootersWindow.xaml
    /// </summary>
    public partial class HeaderAndFootersWindow : ElearningWindow
    {
        public HeaderAndFootersWindow()
        {
            InitializeComponent();
            SlideBase slide = (Application.Current as IAppGlobal).SelectedSlide;
            txtFooter.Text = DesignTabControlViewModel.GetFooterSlide(slide);
            cbDateTime.IsChecked = slide.IsDate;
            cbFooter.IsChecked = slide.IsFooter;
            cbSlideNumber.IsChecked = slide.IsSlideNumber;
        }

        private void applyBtn_Click(object sender, RoutedEventArgs e)
        {
            SlideBase slide = (Application.Current as IAppGlobal).SelectedSlide;
            slide.IsDate = cbDateTime.IsChecked.HasValue ? cbDateTime.IsChecked.Value : false;
            slide.IsSlideNumber = cbSlideNumber.IsChecked.HasValue ? cbSlideNumber.IsChecked.Value : false;
            slide.IsFooter = cbFooter.IsChecked.HasValue ? cbFooter.IsChecked.Value : false;
            slide.TextFooter = txtFooter.Text;
            foreach (var item in slide.MainLayout.Elements)
            {
                if (item is TextEditor textEditor && textEditor.RichTextEditor.TextContainer.Document.TypeText == Text.TypeText.Footers)
                {
                    textEditor.RichTextEditor.TextContainer.Document = Helper.CommandHelper.SetTextDocument(textEditor.RichTextEditor.TextContainer.Document, txtFooter.Text, 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                    textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SelectedFont.MinorFont;
                    textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                    textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
                    textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
                    textEditor.RichTextEditor.TextContainer.InvalidateVisual();
                }
            }
            slide.MainLayout.UpdateThumbnail();
            Close();
        }

        private void applyAllbtn_Click(object sender, RoutedEventArgs e)
        {
            foreach (SlideBase slide in (Application.Current as IAppGlobal).DocumentControl.Slides)
            {
                slide.IsDate = cbDateTime.IsChecked.HasValue ? cbDateTime.IsChecked.Value : false;
                slide.IsSlideNumber = cbSlideNumber.IsChecked.HasValue ? cbSlideNumber.IsChecked.Value : false;
                slide.IsFooter = cbFooter.IsChecked.HasValue ? cbFooter.IsChecked.Value : false;
                slide.TextFooter = txtFooter.Text;
                foreach (var item in slide.MainLayout.Elements)
                {
                    if (item is TextEditor textEditor && textEditor.RichTextEditor.TextContainer.Document.TypeText == Text.TypeText.Footers)
                    {
                        textEditor.RichTextEditor.TextContainer.Document = Helper.CommandHelper.SetTextDocument(textEditor.RichTextEditor.TextContainer.Document, txtFooter.Text, 12, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Center, VerticalAlign.Middle, textEditor.RichTextEditor.TextContainer);
                        textEditor.RichTextEditor.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).DocumentControl.SelectedTheme.SelectedFont.MinorFont;
                        textEditor.RichTextEditor.TextContainer.Document.Container = textEditor.RichTextEditor.TextContainer;
                        textEditor.RichTextEditor.TextContainer.Document.TypeTextContainer = Text.TypeTextContainer.None;
                        textEditor.RichTextEditor.TextContainer.Document.TypeText = Text.TypeText.Footers;
                        textEditor.RichTextEditor.TextContainer.InvalidateVisual();
                    }
                }
                slide.MainLayout.UpdateThumbnail();
            }
            Close();
        }

        private void cancelBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
