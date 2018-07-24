using INV.Elearning.Core;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.DesignControl.Converters;
using INV.Elearning.DesignControl.Helper;
using INV.Elearning.DesignControl.Models;
using INV.Elearning.Text;
using INV.Elearning.Text.Models;
using INV.Elearning.Text.ViewModels.Text;
using INV.Elearning.Text.Views;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace INV.Elearning.DesignControl.Views
{
    public class PlaceHolder : StandardElement, IPlaceHolder
    {
        VisualCollection visualChildren;
        public RichTextEditor text;
        PlaceHolderButton btnVideo, btnText, btnChart, btnImage;
        Grid _buttonGrid;
        /// <summary>
        /// Kiểu Place Holder
        /// </summary>
        public PlaceHolderType Type
        {
            get { return (PlaceHolderType)GetValue(TypeProperty); }
            set { SetValue(TypeProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Type.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty TypeProperty =
            DependencyProperty.Register("Type", typeof(PlaceHolderType), typeof(PlaceHolder), new FrameworkPropertyMetadata(INV.Elearning.Core.Model.Theme.PlaceHolderType.Content, FrameworkPropertyMetadataOptions.AffectsRender));




        /// <summary>
        /// Document của PlaceHolder
        /// </summary>
        public Document Document
        {
            get { return (Document)GetValue(DocumentProperty); }
            set { SetValue(DocumentProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Document.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DocumentProperty =
            DependencyProperty.Register("Document", typeof(Document), typeof(PlaceHolder), new PropertyMetadata(null));




        /// <summary>
        /// Kiểm tra xem có phải là place holder nằm trong MasterSlide hay không
        /// </summary>
        public bool IsMasterSlide
        {
            get { return (bool)GetValue(IsMasterSlideProperty); }
            set { SetValue(IsMasterSlideProperty, value); }
        }

        // Using a DependencyProperty as the backing store for IsMasterSlide.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty IsMasterSlideProperty =
            DependencyProperty.Register("IsMasterSlide", typeof(bool), typeof(PlaceHolder), new PropertyMetadata(true));




        public PlaceHolder()
        {
            this.Data = new EPlaceHolder();
            visualChildren = new VisualCollection(this);
            Background = Brushes.Transparent;
            Loaded += PlaceHolder_Loaded;
            SelectionChanged += PlaceHolder_SelectionChanged;
            btnVideo = new PlaceHolderButton();
            btnImage = new PlaceHolderButton();
            btnText = new PlaceHolderButton();
            btnChart = new PlaceHolderButton();
            text = new RichTextEditor();
            this.Thickness = 0;
            DashType = Controls.Enums.DashType.DashDot;
        }

        static PlaceHolder()
        {
            RuleConfigurationProperty.OverrideMetadata(typeof(PlaceHolder), new PropertyMetadata(new ObjectElementRule(true, false)));
            IconProperty.OverrideMetadata(typeof(PlaceHolder), new PropertyMetadata(@"pack://application:,,,/INV.Elearning.DesignControl;Component/Images/InsertPlaceHolder16.png"));
            ElementTypeProperty.OverrideMetadata(typeof(PlaceHolder), new PropertyMetadata("Place Holder"));
        }

        private void PlaceHolder_SelectionChanged(object sender, ObjectElementEventArg e)
        {
            if (IsMasterSlide)
            {
                if ((bool)e.NewValue)
                {
                    btnVideo.Visibility = Visibility.Collapsed;
                    btnChart.Visibility = Visibility.Collapsed;
                    btnText.Visibility = Visibility.Collapsed;
                    btnImage.Visibility = Visibility.Collapsed;
                }
                else
                {
                    btnVideo.Visibility = Visibility.Visible;
                    btnChart.Visibility = Visibility.Visible;
                    btnText.Visibility = Visibility.Visible;
                    btnImage.Visibility = Visibility.Visible;
                }
            }
            if ((bool)e.NewValue)
            {
                foreach (ObjectElement item in (Application.Current as IAppGlobal).SelectedSlide.MainLayout.Elements)
                {
                    if (item != this)
                    {
                        item.IsSelected = false;
                    }
                }
            }
        }

        private void PlaceHolder_Loaded(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.Image image = new System.Windows.Controls.Image() { Source = new BitmapImage(new Uri(@"/INV.Elearning.DesignControl;Component/Images/Video32.png", UriKind.Relative)) };
            image.Width = 32;
            image.Height = 32;
            btnVideo.Command = CommandHelper.PlaceHolderAddVideoCommand;
            MultiBinding multiBinding = new MultiBinding();
            multiBinding.Converter = new Converters.MuitiCommandParameterConverter();
            multiBinding.Bindings.Add(new Binding("ActualWidth") { Source = this });
            multiBinding.Bindings.Add(new Binding("ActualHeight") { Source = this });
            multiBinding.Bindings.Add(new Binding("Top") { Source = this });
            multiBinding.Bindings.Add(new Binding("Left") { Source = this });
            multiBinding.NotifyOnSourceUpdated = true;
            btnVideo.SetBinding(PlaceHolderButton.CommandParameterProperty, multiBinding);
            btnVideo.ToolTip = "Video";
            btnVideo.Content = image;

            image = new System.Windows.Controls.Image() { Source = new BitmapImage(new Uri(@"/INV.Elearning.DesignControl;Component/Images/Picture32.png", UriKind.Relative)) };
            image.Width = 32;
            image.Height = 32;
            btnImage.Command = CommandHelper.PlaceHolderAddImageCommand;
            multiBinding = new MultiBinding();
            multiBinding.Converter = new Converters.MuitiCommandParameterConverter();
            multiBinding.Bindings.Add(new Binding("ActualWidth") { Source = this });
            multiBinding.Bindings.Add(new Binding("ActualHeight") { Source = this });
            multiBinding.Bindings.Add(new Binding("Top") { Source = this });
            multiBinding.Bindings.Add(new Binding("Left") { Source = this });
            multiBinding.NotifyOnSourceUpdated = true;
            btnImage.SetBinding(PlaceHolderButton.CommandParameterProperty, multiBinding);
            btnImage.ToolTip = "Image";
            btnImage.Content = image;

            image = new System.Windows.Controls.Image() { Source = new BitmapImage(new Uri(@"/INV.Elearning.DesignControl;Component/Images/Text32.png", UriKind.Relative)) };
            image.Width = 32;
            image.Height = 32;
            btnText.Command = CommandHelper.PlaceHolderAddTextCommand;
            multiBinding = new MultiBinding();
            multiBinding.Converter = new Converters.MuitiCommandParameterConverter();
            multiBinding.Bindings.Add(new Binding("ActualWidth") { Source = this });
            multiBinding.Bindings.Add(new Binding("ActualHeight") { Source = this });
            multiBinding.Bindings.Add(new Binding("Top") { Source = this });
            multiBinding.Bindings.Add(new Binding("Left") { Source = this });
            multiBinding.NotifyOnSourceUpdated = true;
            btnText.SetBinding(PlaceHolderButton.CommandParameterProperty, multiBinding);
            btnText.ToolTip = "Text";
            btnText.Content = image;

            image = new System.Windows.Controls.Image() { Source = new BitmapImage(new Uri(@"/INV.Elearning.DesignControl;Component/Images/Chart32.png", UriKind.Relative)) };
            image.Width = 32;
            image.Height = 32;
            btnChart.Command = CommandHelper.PlaceHolderChartCommand;
            multiBinding = new MultiBinding();
            multiBinding.Converter = new Converters.MuitiCommandParameterConverter();
            multiBinding.Bindings.Add(new Binding("ActualWidth") { Source = this });
            multiBinding.Bindings.Add(new Binding("ActualHeight") { Source = this });
            multiBinding.Bindings.Add(new Binding("Top") { Source = this });
            multiBinding.Bindings.Add(new Binding("Left") { Source = this });
            multiBinding.NotifyOnSourceUpdated = true;
            btnChart.SetBinding(PlaceHolderButton.CommandParameterProperty, multiBinding);
            btnChart.ToolTip = "Chart";
            btnChart.Content = image;

            if (IsMasterSlide)
            {
                text.IsHitTestVisible = true;
                Binding binding = new Binding()
                {
                    Source = this,
                    Path = new PropertyPath("ActualWidth")
                };
                text.SetBinding(RichTextEditor.WidthProperty, binding);

                Binding binding2 = new Binding()
                {
                    Source = this,
                    Path = new PropertyPath("ActualHeight")
                };
                text.SetBinding(RichTextEditor.HeightProperty, binding2);
                text.HorizontalAlignment = HorizontalAlignment.Left;
                text.VerticalAlignment = VerticalAlignment.Top;
                btnVideo.IsHitTestVisible = false;
                btnImage.IsHitTestVisible = false;
                btnText.IsHitTestVisible = false;
                btnChart.IsHitTestVisible = false;
            }
            else text.IsHitTestVisible = false;
            InvalidateVisual();
        }

        public static Document CreateData(string Text, double fontSize, TextContainer textContainer, HorizontalAlign align)
        {
            Document _result = new Document();
            _result.IsQuestion = false;
            _result.Padding = new EThickness() { Top = 0, Bottom = 0, Left = 0, Right = 0 };
            _result.VerticalAlign = VerticalAlign.Top;
            Paragraph _para = new Paragraph();
            Run _run = new Run();
            _run.Fontfamily = "Arial";
            _run.FontSize = fontSize;
            _run.FontStyle = INV.Elearning.Text.FontStyle.Normal;
            _run.Foreground = new ColorBrushBase() { Brush = Brushes.Black };
            _run.Text = Text + "\r";
            _run.StrikeThrough = new StrikeThrough();
            _run.StrikeThrough = StrikeThrough.None;
            _run.Underline = new Underline();
            _run.Underline.Style = UnderlineStyle.None;
            _run.Underline.Color = null;
            _run.Parent = _para;
            _para.Inlines.Add(_run);
            _para.TextIndent = 0;
            _para.Parent = _result;
            _para.LineHeight = 1;
            _para.TextAlign = align;
            _para.Margin = new EThickness() { Top = 0, Left = 0 };
            _result.Blocks.Add(_para);
            textContainer.Selection.Start = textContainer.Selection.End = new TextPointer(_run, 0);

            return _result;

        }

        /// <summary>
        /// Tạo nội dung mẫu cho khung
        /// </summary>
        /// <param name="content"></param>
        /// <param name="fontSize"></param>
        /// <param name="fontFamily"></param>
        /// <param name="foreground"></param>
        /// <param name="align"></param>
        /// <param name="verticalAlign"></param>
        /// <param name="textContainer"></param>
        /// <returns></returns>
        public static Document CreateData(string content, double fontSize, string fontFamily, ColorBrushBase foreground, HorizontalAlign align, VerticalAlign verticalAlign, TextContainer textContainer)
        {
            Document _result = new Document();
            _result.IsQuestion = false;
            _result.Padding = new EThickness() { Top = 5, Bottom = 5, Left = 5, Right = 5 };
            _result.VerticalAlign = verticalAlign;
            Paragraph _para = new Paragraph();
            Run _run = new Run();
            _run.Fontfamily = fontFamily;
            _run.FontSize = fontSize;
            _run.FontStyle = Text.FontStyle.Normal;
            _run.Foreground = foreground;
            _run.Text = content + "\r";
            _run.StrikeThrough = StrikeThrough.None;
            _run.Underline = new Underline();
            _run.Underline.Style = UnderlineStyle.None;
            _run.Parent = _para;
            _para.Inlines.Add(_run);
            _para.TextIndent = 0;
            _para.Parent = _result;
            _para.LineHeight = 1;
            _para.TextAlign = align;
            _para.Margin = new EThickness();
            _result.Blocks.Add(_para);
            textContainer.Selection.Start = textContainer.Selection.End = new TextPointer(_run, 0);

            return _result;

        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            _buttonGrid = new Grid();
            _buttonGrid.Width = 80;
            _buttonGrid.Height = 80;
            ColumnDefinition c1 = new ColumnDefinition();
            c1.Width = new GridLength(1, GridUnitType.Star);
            ColumnDefinition c2 = new ColumnDefinition();
            c2.Width = new GridLength(1, GridUnitType.Star);

            _buttonGrid.ColumnDefinitions.Add(c1);
            _buttonGrid.ColumnDefinitions.Add(c2);

            RowDefinition r1 = new RowDefinition();
            r1.Height = new GridLength(1, GridUnitType.Star);
            RowDefinition r2 = new RowDefinition();
            r2.Height = new GridLength(1, GridUnitType.Star);

            _buttonGrid.RowDefinitions.Add(r1);
            _buttonGrid.RowDefinitions.Add(r2);


            if (Container != null)
            {
                if (text.Parent == null)
                {
                    Container.Children.Add(text);
                }
                switch (Type)
                {
                    case INV.Elearning.Core.Model.Theme.PlaceHolderType.Content:
                        text.TextContainer.Document = CreateData("Click to select content", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Top, text.TextContainer);
                        text.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont.MinorFont;
                        text.TextContainer.Document.TypeTextContainer = TypeTextContainer.None;
                        text.TextContainer.Document.Container = text.TextContainer;
                        if (btnVideo.Parent == null)
                        {
                            //Container.Children.Add(btnVideo);
                            _buttonGrid.Children.Add(btnVideo);

                        }
                        if (btnImage.Parent == null)
                        {
                            Grid.SetRow(btnImage, 1);
                            _buttonGrid.Children.Add(btnImage);
                            //Container.Children.Add(btnImage);
                        }
                        if (btnText.Parent == null)
                        {
                            Grid.SetColumn(btnText, 1);
                            _buttonGrid.Children.Add(btnText);
                            //Container.Children.Add(btnText);
                        }
                        if (btnChart.Parent == null)
                        {
                            Grid.SetRow(btnChart, 1);
                            Grid.SetColumn(btnChart, 1);
                            _buttonGrid.Children.Add(btnChart);
                            //Container.Children.Add(btnChart);
                        }
                        break;
                    case INV.Elearning.Core.Model.Theme.PlaceHolderType.Text:
                        text.TextContainer.Document = CreateData("Click to edit Master text styles", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Top, text.TextContainer);
                        text.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont.MinorFont;

                        text.TextContainer.Document.Container = text.TextContainer;
                        text.TextContainer.Document.TypeTextContainer = TypeTextContainer.None;
                        if (btnText.Parent == null)
                        {
                            Container.Children.Add(btnText);
                        }
                        break;
                    case INV.Elearning.Core.Model.Theme.PlaceHolderType.Picture:
                        text.TextContainer.Document = CreateData("Picture", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Top, text.TextContainer);
                        text.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont.MinorFont;
                        text.TextContainer.Document.TypeTextContainer = TypeTextContainer.None;
                        text.TextContainer.Document.Container = text.TextContainer;
                        if (btnImage.Parent == null)
                        {
                            Container.Children.Add(btnImage);
                        }
                        break;
                    case INV.Elearning.Core.Model.Theme.PlaceHolderType.Video:
                        text.TextContainer.Document = CreateData("Video", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Top, text.TextContainer);
                        text.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont.MinorFont;
                        text.TextContainer.Document.TypeTextContainer = TypeTextContainer.None;
                        text.TextContainer.Document.Container = text.TextContainer;
                        if (btnText.Parent == null)
                        {
                            Container.Children.Add(btnVideo);
                        }
                        break;
                    case INV.Elearning.Core.Model.Theme.PlaceHolderType.Chart:
                        text.TextContainer.Document = CreateData("Chart", 28, "Arial", new ColorBrushBase() { Brush = Brushes.Black }, HorizontalAlign.Left, VerticalAlign.Top, text.TextContainer);
                        text.TextContainer.Document.Fontfamily = (Application.Current as IAppGlobal).SelectedTheme.SelectedFont.MinorFont;
                        text.TextContainer.Document.TypeTextContainer = TypeTextContainer.None;
                        text.TextContainer.Document.Container = text.TextContainer;
                        if (btnChart.Parent == null)
                        {
                            Container.Children.Add(btnChart);
                        }
                        break;
                    default:
                        break;
                }
            }
            Container.Children.Add(_buttonGrid);

        }

        public void SetTextBlank()
        {
            text.TextContainer.Document = CreateData("", 28, text.TextContainer, HorizontalAlign.Left);
            text.TextContainer.Document.Container = text.TextContainer;
        }

        protected override void OnPreviewMouseDown(MouseButtonEventArgs e)
        {
            base.OnMouseDown(e);
            IsSelected = true;
        }



        protected override void OnRender(DrawingContext drawingContext)
        {
            base.OnRender(drawingContext);

            if (text != null)
            {
                Document = text.TextContainer.Document;
            }
            drawingContext.DrawRectangle(Brushes.Transparent, new Pen(Brushes.Transparent, 1), new Rect(0, 0, ActualWidth, ActualHeight));

        }

        public override void UpdateUI(ObjectElementBase data)
        {
            base.UpdateUI(data);
            if (data is EPlaceHolder ePlaceHolder)
            {
                Type = ePlaceHolder.Type;
                IsMasterSlide = ePlaceHolder.IsMasterSlide;
                Document document = new Document();
                document = Text.Helper.ConverterData.ConverXmlToData(ePlaceHolder.Document);
                Document = document;
            }

        }

        public override void RefreshData()
        {
            base.RefreshData();
            if (this.Data is EPlaceHolder ePlaceHolder)
            {
                ePlaceHolder.Type = Type;
                ePlaceHolder.IsMasterSlide = IsMasterSlide;
                DataDocument dataDocument = new INV.Elearning.Text.Models.DataDocument() { Foreground = new DataColor() };
                Text.Helper.DataText.CopyFormatDocument(this.Document, dataDocument);
                ePlaceHolder.Document = dataDocument;
            }
        }

    }
}
