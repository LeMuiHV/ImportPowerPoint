using INV.Elearning.Core;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.View;
using INV.Elearning.ImportPowerPoint.Helper;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using pp = Microsoft.Office.Interop.PowerPoint;
using office = Microsoft.Office.Core;
using INV.Elearning.Controls.Audio;
using D = DocumentFormat.OpenXml.Drawing;
using oxml = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using INV.Elearning.Controls.Flash.Controls;
using INV.Elearning.VideoElementControl.Model;
using INV.Elearning.Text.Views;
using System.Windows.Media;
using INV.Elearning.Text.Models;
using INV.Elearning.Text;
using INV.Elearning.Animations.UI;
using INV.Elearning.Animations.Enums;
using INV.Elearning.ImageProcess.Models;
using INV.Elearning.Animations;
using INV.Elearing.Controls.Shapes;
using DocumentFormat.OpenXml;

namespace INV.Elearning.ImportPowerPoint.Process
{
    public class ConvertPresentation
    {

        /// <summary>
        /// Convert toàn bộ các trang tài liệu dạng ảnh nền
        /// </summary>
        public void SlideAsBackground(NormalPage _page, pp.Slide _slide)
        {
            HelperClass helperClass = new HelperClass();
            ImageColor imgColor = new ImageColor();
            imgColor.Source = new Image();
            imgColor.ScaleX = 1;
            imgColor.ScaleY = 1;
            imgColor.Source.OpenFile(Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, "Slide_" + _slide.SlideIndex.ToString() + ".png"), false);
            _page.MainLayer = new PageLayer();
            _page.MainLayer.Background = imgColor;
            _page.ID = ObjectElementsHelper.RandomString(13);
            //----Hiệu ứng trang
            helperClass.GetTransition(_slide, _page);


        }
        /// <summary>
        /// Nhập đối tượng tài liệu PowerPoint
        /// </summary>
        /// <param name="_page"></param>
        /// <param name="_slide"></param>
        public void ObjectsInSlides(NormalPage _page, pp.Slide _slide, int count)
        {
            HelperClass helperClass = new HelperClass();
         
            //----Hiệu ứng trang
            helperClass.GetTransition(_slide, _page);
            //---Chọn kiểu hiệu ứng trang và hướng chiếu
           
            //-- Thiết lập hình nền màu nền => quy ra ảnh nền object để lấy cả chủ đề, dải màu//loop: Tránh vòng lặp vô hạn
            int loop = 0;
            while (_slide.Shapes.Count > 0)
            {
                if (loop > 200)
                {
                    loop = 0;
                    goto BreakLoop;
                }
                else
                {
                    _slide.Shapes[1].Delete();
                    loop++;
                }
            }
            BreakLoop:
            try
            {
                string PathBackground = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, "Slide_" + DateTime.Now.Ticks.ToString() + ".png");
                _slide.Export(PathBackground, "png");
                if (File.Exists(PathBackground))
                {
                    ImageColor imgColor = new ImageColor();
                    imgColor.Source = new Image();
                    imgColor.Source.OpenFile(PathBackground, false);
                    _page.MainLayer.Background = imgColor;
                }
            }
            catch (Exception ex)
            {
            }
        }




        private void ShapeIsPicture(PageLayer _page, pp.Shape _shape, EAnimation _animation, string _idShape)
        {
            EImageContent eImage = new EImageContent();
            eImage.Name = "Image";
            eImage.ID = _idShape;
            eImage.Width = _shape.Width * Utils.tlw;
            eImage.Height = _shape.Height * Utils.tlh;
            eImage.Top = _shape.Top * Utils.tlh;
            eImage.Left = _shape.Left * Utils.tlw;
            eImage.Angle = _shape.Rotation;
            eImage.ZOder = (int)_shape.ZOrderPosition;
            eImage.Animations = _animation;
            eImage.RectCrop = new ESuperPoint();
            var lstPic = Utils.slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
            P.Picture picTag = GetPictureTag(lstPic, _shape.Id.ToString());
            string rID = picTag.BlipFill.Blip.Embed.Value;
            ImagePart part = (ImagePart)Utils.slidePart.GetPartById(rID);
            string pathImage = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + Path.GetExtension(part.Uri.ToString()));
            Utils.CopyStream(part.GetStream(), pathImage);
            eImage.Image = new Image();
            eImage.Image.Path = pathImage;
            _page.Children.Add(eImage);
        }
        /// <summary>
        /// Nhập đối tượng là audio
        /// </summary>
        /// <param name="_page"></param>
        /// <param name="_shape"></param>
        private void ShapeIsAudio(PageLayer _page, pp.Shape _shape, EAnimation _animation, string _idShape)
        {
            EAudioElement eAudio = new EAudioElement();
            eAudio.Name = "Audio";
            eAudio.ID = _idShape;
            eAudio.Width = _shape.Width * Utils.tlw;
            eAudio.Height = _shape.Height * Utils.tlh;
            eAudio.Top = _shape.Top * Utils.tlh;
            eAudio.Left = _shape.Left * Utils.tlw;
            eAudio.Angle = _shape.Rotation;
            eAudio.ZOder = (int)_shape.ZOrderPosition;
            eAudio.Animations = _animation;
            //Lấy file audio trong gói sử dụng openxml
            var lstPic = Utils.slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
            P.Picture audioTag = GetPictureTag(lstPic, _shape.Id.ToString());
            var audioFromFile = audioTag.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties.Descendants<D.AudioFromFile>().Select(p => p);
            ReferenceRelationship relationship = Utils.slidePart.GetReferenceRelationship(audioFromFile.First().Link.Value);
            DataPartReferenceRelationship data = (DataPartReferenceRelationship)relationship;
            string pathAudio = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + Path.GetExtension(data.Uri.ToString()));
            Utils.CopyStream(data.DataPart.GetStream(), pathAudio);
            eAudio.AudioFileName = new Core.Model.Media.Audio();
            eAudio.AudioFileName.Path = pathAudio;
            _page.Children.Add(eAudio);
        }
        /// <summary>
        /// Nhập đối tượng là video
        /// </summary>
        /// <param name="_page"></param>
        /// <param name="_shape"></param>
        private void ShapeIsVideo(PageLayer _page, pp.Shape _shape, EAnimation _animation, string _idShape)
        {
            var lstPic = Utils.slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
            P.Picture videoTag = GetPictureTag(lstPic, _shape.Id.ToString());
            var videoFromFile = videoTag.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties.Descendants<D.VideoFromFile>().Select(p => p);
            ReferenceRelationship relationship = Utils.slidePart.GetReferenceRelationship(videoFromFile.First().Link.Value);
            if (videoTag.NonVisualPictureProperties.NonVisualDrawingProperties.HyperlinkOnClick == null)
                VideoAsFlash(_page, _shape, relationship, _animation, _idShape);
            else
            {
                EVideoPlayer eVideoPlayer = new EVideoPlayer();
                eVideoPlayer.Name = "Video";
                eVideoPlayer.ID = _idShape;
                eVideoPlayer.Angle = _shape.Rotation;
                eVideoPlayer.Width = _shape.Width * Utils.tlw;
                eVideoPlayer.Height = _shape.Height * Utils.tlh;
                eVideoPlayer.Top = _shape.Top * Utils.tlh;
                eVideoPlayer.Left = _shape.Left * Utils.tlw;
                eVideoPlayer.ZOder = _shape.ZOrderPosition;
                eVideoPlayer.Animations = _animation;
                //Lấy file video trong gói sử dụng openxml
                DataPartReferenceRelationship data = (DataPartReferenceRelationship)relationship;
                string pathVideo = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + Path.GetExtension(data.Uri.ToString()));
                Utils.CopyStream(data.DataPart.GetStream(), pathVideo);
                eVideoPlayer.Source = new Core.Model.Media.Video();
                eVideoPlayer.Source.Path = pathVideo;
                //Lấy avata video
                eVideoPlayer.ImageFirst = new Image();
                string rId = videoTag.BlipFill.Blip.Embed.Value;
                OpenXmlPart avtPart = Utils.slidePart.GetPartById(rId);
                string pathAvata = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + Path.GetExtension(avtPart.Uri.ToString()));
                Utils.CopyStream(avtPart.GetStream(), pathAvata);
                eVideoPlayer.ImageFirst.Path = pathAvata;
                _page.Children.Add(eVideoPlayer);
            }
        }
        /// <summary>
        /// Chèn Flash như video
        /// </summary>
        /// <param name=""></param>
        /// <param name="_shape"></param>
        /// <param name="_flashTag"></param> 
        private void VideoAsFlash(PageLayer _page, pp.Shape _shape, ReferenceRelationship _rls, EAnimation _animation, string _idShape)
        {
            EFlashControl eFlash = new EFlashControl();
            eFlash.Name = "Flash";
            eFlash.ID = _idShape;
            eFlash.Angle = _shape.Rotation;
            eFlash.Height = _shape.Height * Utils.tlh;
            eFlash.Left = _shape.Left * Utils.tlw;
            eFlash.Width = _shape.Width * Utils.tlw;
            eFlash.Top = _shape.Top * Utils.tlh;
            eFlash.ZOder = _shape.ZOrderPosition;
            eFlash.Animations = _animation;
            //Lấy file flash trong gói sử dụng openxml
            eFlash.Source = new Core.Model.Media.Flash();
            eFlash.Source.Path = _rls.Uri.AbsolutePath;
            _page.Children.Add(eFlash);
        }
        private void ShapeIsText(PageLayer _page, pp.Shape _shape, EAnimation _animation, string _idShape)
        {
            //Thiếu textindent và strike words 
            DataDocument document = new DataDocument();
            document.Name = "TextEditor";
            List<DataParaPositionBullet> ListParaPosition = new List<DataParaPositionBullet>();
            DataParaPositionBullet paraPositionBullet = null;
            document.Blocks.Clear();
            document.Container = new DataTextContainer();

            if (_shape.Type == office.MsoShapeType.msoTextBox)
                document.TypeTextContainer = TypeTextContainer.AutoChangeHeight;
            else
                document.TypeTextContainer = TypeTextContainer.ChangeSize;


            //Lấy nội dung 
            pp.BulletFormat hasBullet = _shape.TextFrame.TextRange.ParagraphFormat.Bullet;
            if (hasBullet.Visible == office.MsoTriState.msoCTrue || hasBullet.Visible == office.MsoTriState.msoTrue)
            {
                foreach (pp.TextRange paragraph in _shape.TextFrame.TextRange.Paragraphs())
                {
                    DataDocumentList docList = new DataDocumentList();
                    int level = paragraph.IndentLevel;
                    DataDocumentListItem docListItem = new DataDocumentListItem();
                    pp.BulletFormat prgBullet = paragraph.ParagraphFormat.Bullet;
                    docListItem.TypeWordBullet = new Text.ViewModels.Text.TypeWordBullet();

                    docListItem.TypeWordBullet.StarIndex = prgBullet.StartValue;
                    if (prgBullet.Visible == office.MsoTriState.msoCTrue || prgBullet.Visible == office.MsoTriState.msoTrue)
                    {
                        docListItem.TypeWordBullet = new Text.ViewModels.Text.TypeWordBullet();
                        docListItem.TypeWordBullet.SizeOffset = 100;
                        if (prgBullet.Type == PpBulletType.ppBulletNumbered)
                        {
                            docListItem.TypeWordBullet.Fontfamily = "Arial";
                            docListItem.TypeWordBullet.Text = "1.";
                            docListItem.TypeWordBullet.ListType = Text.ListType.Decimal;
                            switch (prgBullet.Style)
                            {
                                case PpNumberedBulletStyle.ppBulletStyleMixed:
                                    break;
                                case PpNumberedBulletStyle.ppBulletAlphaLCPeriod:
                                    docListItem.TypeWordBullet.Fontfamily = "Arial";
                                    docListItem.TypeWordBullet.Text = "a.";
                                    docListItem.TypeWordBullet.ListType = Text.ListType.LowerLatin;
                                    break;
                                case PpNumberedBulletStyle.ppBulletAlphaUCPeriod:
                                    docListItem.TypeWordBullet.Fontfamily = "Arial";
                                    docListItem.TypeWordBullet.Text = "A.";
                                    docListItem.TypeWordBullet.ListType = Text.ListType.UpperLatin;
                                    break;
                                case PpNumberedBulletStyle.ppBulletArabicParenRight:
                                    docListItem.TypeWordBullet.Fontfamily = "Arial";
                                    docListItem.TypeWordBullet.Text = "1)";
                                    docListItem.TypeWordBullet.ListType = Text.ListType.Decimal2;
                                    break;
                                case PpNumberedBulletStyle.ppBulletArabicPeriod:
                                    docListItem.TypeWordBullet.Fontfamily = "Arial";
                                    docListItem.TypeWordBullet.Text = "1.";
                                    docListItem.TypeWordBullet.ListType = Text.ListType.Decimal;
                                    break;
                                case PpNumberedBulletStyle.ppBulletRomanLCParenBoth:
                                    break;
                                case PpNumberedBulletStyle.ppBulletRomanLCParenRight:
                                    break;
                                case PpNumberedBulletStyle.ppBulletRomanLCPeriod:
                                    docListItem.TypeWordBullet.Fontfamily = "Arial";
                                    docListItem.TypeWordBullet.Text = "i.";
                                    docListItem.TypeWordBullet.ListType = Text.ListType.LowerRoman;
                                    break;
                                case PpNumberedBulletStyle.ppBulletRomanUCPeriod:
                                    docListItem.TypeWordBullet.Fontfamily = "Arial";
                                    docListItem.TypeWordBullet.Text = "I.";
                                    docListItem.TypeWordBullet.ListType = Text.ListType.UpperRoman;
                                    break;
                                case PpNumberedBulletStyle.ppBulletAlphaLCParenBoth:
                                    break;
                                case PpNumberedBulletStyle.ppBulletAlphaLCParenRight:
                                    docListItem.TypeWordBullet.Fontfamily = "Arial";
                                    docListItem.TypeWordBullet.Text = "a)";
                                    docListItem.TypeWordBullet.ListType = Text.ListType.LowerLatin2;
                                    break;
                                case PpNumberedBulletStyle.ppBulletAlphaUCParenBoth:
                                    break;
                                case PpNumberedBulletStyle.ppBulletAlphaUCParenRight:
                                    break;
                                case PpNumberedBulletStyle.ppBulletArabicParenBoth:
                                    break;
                                case PpNumberedBulletStyle.ppBulletArabicPlain:
                                    break;
                                case PpNumberedBulletStyle.ppBulletRomanUCParenBoth:
                                    break;
                                case PpNumberedBulletStyle.ppBulletRomanUCParenRight:
                                    break;
                                case PpNumberedBulletStyle.ppBulletSimpChinPlain:
                                    break;
                                case PpNumberedBulletStyle.ppBulletSimpChinPeriod:
                                    break;
                                case PpNumberedBulletStyle.ppBulletCircleNumDBPlain:
                                    break;
                                case PpNumberedBulletStyle.ppBulletCircleNumWDWhitePlain:
                                    break;
                                case PpNumberedBulletStyle.ppBulletCircleNumWDBlackPlain:
                                    break;
                                case PpNumberedBulletStyle.ppBulletTradChinPlain:
                                    break;
                                case PpNumberedBulletStyle.ppBulletTradChinPeriod:
                                    break;
                                case PpNumberedBulletStyle.ppBulletArabicAlphaDash:
                                    break;
                                case PpNumberedBulletStyle.ppBulletArabicAbjadDash:
                                    break;
                                case PpNumberedBulletStyle.ppBulletHebrewAlphaDash:
                                    break;
                                case PpNumberedBulletStyle.ppBulletKanjiKoreanPlain:
                                    break;
                                case PpNumberedBulletStyle.ppBulletKanjiKoreanPeriod:
                                    break;
                                case PpNumberedBulletStyle.ppBulletArabicDBPlain:
                                    break;
                                case PpNumberedBulletStyle.ppBulletArabicDBPeriod:
                                    break;
                                case PpNumberedBulletStyle.ppBulletThaiAlphaPeriod:
                                    break;
                                case PpNumberedBulletStyle.ppBulletThaiAlphaParenRight:
                                    break;
                                case PpNumberedBulletStyle.ppBulletThaiAlphaParenBoth:
                                    break;
                                case PpNumberedBulletStyle.ppBulletThaiNumPeriod:
                                    break;
                                case PpNumberedBulletStyle.ppBulletThaiNumParenRight:
                                    break;
                                case PpNumberedBulletStyle.ppBulletThaiNumParenBoth:
                                    break;
                                case PpNumberedBulletStyle.ppBulletHindiAlphaPeriod:
                                    break;
                                case PpNumberedBulletStyle.ppBulletHindiNumPeriod:
                                    break;
                                case PpNumberedBulletStyle.ppBulletKanjiSimpChinDBPeriod:
                                    break;
                                case PpNumberedBulletStyle.ppBulletHindiNumParenRight:
                                    break;
                                case PpNumberedBulletStyle.ppBulletHindiAlpha1Period:
                                    break;
                                default:
                                    break;
                            }
                        }
                        else if (prgBullet.Type == PpBulletType.ppBulletUnnumbered)
                        {
                            if (prgBullet.Character == 8226)
                            {
                                docListItem.TypeWordBullet.Fontfamily = "Wingdings";
                                docListItem.TypeWordBullet.Text = "l";
                                docListItem.TypeWordBullet.ListType = Text.ListType.Bullet;
                            }
                            else if (prgBullet.Character == 111)
                            {
                                docListItem.TypeWordBullet.Fontfamily = "Raleway";
                                docListItem.TypeWordBullet.Text = "o";
                                docListItem.TypeWordBullet.ListType = Text.ListType.Bullet;
                            }
                            else if (prgBullet.Character == 167)
                            {
                                docListItem.TypeWordBullet.Fontfamily = "Wingdings";
                                docListItem.TypeWordBullet.Text = ((char)167).ToString();
                                docListItem.TypeWordBullet.ListType = Text.ListType.Bullet;
                            }
                            else if (prgBullet.Character == 113)
                            {
                                docListItem.TypeWordBullet.Fontfamily = "Wingdings";
                                docListItem.TypeWordBullet.Text = "q";
                                docListItem.TypeWordBullet.ListType = Text.ListType.Bullet;
                            }
                            else if (prgBullet.Character == 118)
                            {
                                docListItem.TypeWordBullet.Fontfamily = "Wingdings";
                                docListItem.TypeWordBullet.Text = "v";
                                docListItem.TypeWordBullet.ListType = Text.ListType.Bullet;
                            }
                            else if (prgBullet.Character == 216)
                            {
                                docListItem.TypeWordBullet.Fontfamily = "Wingdings";
                                docListItem.TypeWordBullet.Text = ((char)216).ToString();
                                docListItem.TypeWordBullet.ListType = Text.ListType.Bullet;
                            }
                            else if (prgBullet.Character == 252)
                            {
                                docListItem.TypeWordBullet.Fontfamily = "Wingdings";
                                docListItem.TypeWordBullet.Text = ((char)252).ToString();
                                docListItem.TypeWordBullet.ListType = Text.ListType.Bullet;
                            }
                            else
                            {
                                docListItem.TypeWordBullet.Fontfamily = "Wingdings";
                                docListItem.TypeWordBullet.Text = "l";
                                docListItem.TypeWordBullet.ListType = Text.ListType.Bullet;
                            }
                        }
                    }
                    paraPositionBullet = new DataParaPositionBullet();
                    paraPositionBullet.Para = GetEParagraph(paragraph);
                    paraPositionBullet.ListType = docListItem.TypeWordBullet;
                    paraPositionBullet.Position = level;
                    ListParaPosition.Add(paraPositionBullet);
                }
            }
            else
            {
                foreach (pp.TextRange paragraph in _shape.TextFrame.TextRange.Paragraphs())
                {
                    /*
                    var _para = GetEParagraph(paragraph);
                    if (_para.Inlines?.Count > 0)
                    {
                        if ((_para.Inlines.Last() as DataRun)?.Text != null)
                            (_para.Inlines.Last() as DataRun).Text = (_para.Inlines.Last() as DataRun).Text + "\r";
                    }
                    document.Blocks.Add(_para);*/
                    paraPositionBullet = new DataParaPositionBullet();
                    paraPositionBullet.Para = GetEParagraph(paragraph);
                    if (paraPositionBullet.Para?.Inlines?.Count > 0)
                    {
                        if ((paraPositionBullet.Para.Inlines.Last() as DataRun)?.Text != null)
                            (paraPositionBullet.Para.Inlines.Last() as DataRun).Text = (paraPositionBullet.Para.Inlines.Last() as DataRun).Text + "\r";
                    }
                    paraPositionBullet.ListType = null;
                    paraPositionBullet.Position = 0;
                    ListParaPosition.Add(paraPositionBullet);
                }
                //goto addToDoc;
            }
            if (ListParaPosition?.Count == 0)
            {
                paraPositionBullet = new DataParaPositionBullet();
                paraPositionBullet.Para = new DataParagraph();
                DataRun _run = new DataRun();
                _run.Text = "\r";
                _run.FontSize = 14;
                _run.Fontfamily = "Times New Roman";
                _run.Parent = paraPositionBullet.Para;
                paraPositionBullet.Para.Inlines.Add(_run);
                ListParaPosition.Add(paraPositionBullet);
            }
            ConverterListParaToDataDocument(ListParaPosition, document);

            _page.Children.Add(document);
        }
        /// <summary>
        /// Hàm lấy đoạn văn bản của PowerPoint
        /// </summary>
        /// <param name="_paragraph"></param>
        /// <returns></returns>
        private DataParagraph GetEParagraph(pp.TextRange _paragraph)
        {
            DataParagraph eParagraph = new DataParagraph();
            //Căn vị trí đoạn văn bản - mặc định sẽ căn trái
            eParagraph.LineHeight = _paragraph.ParagraphFormat.SpaceWithin;
            eParagraph.Margin.Top = _paragraph.ParagraphFormat.SpaceBefore;
            eParagraph.TextAlign = Text.HorizontalAlign.Left;
            switch (_paragraph.ParagraphFormat.Alignment)
            {
                case PpParagraphAlignment.ppAlignmentMixed:
                    break;
                case PpParagraphAlignment.ppAlignLeft:
                    eParagraph.TextAlign = Text.HorizontalAlign.Left;
                    break;
                case PpParagraphAlignment.ppAlignCenter:
                    eParagraph.TextAlign = Text.HorizontalAlign.Center;
                    break;
                case PpParagraphAlignment.ppAlignRight:
                    eParagraph.TextAlign = Text.HorizontalAlign.Right;
                    break;
                case PpParagraphAlignment.ppAlignJustify:
                    eParagraph.TextAlign = Text.HorizontalAlign.Justify;
                    break;
                case PpParagraphAlignment.ppAlignDistribute:
                    break;
                case PpParagraphAlignment.ppAlignThaiDistribute:
                    break;
                case PpParagraphAlignment.ppAlignJustifyLow:
                    break;
                default:
                    break;
            }
            foreach (pp.TextRange textRun in _paragraph.Runs())
            {
                DataRun eRun = new DataRun();
                if (textRun.Font.NameAscii == "Cambria Math" && textRun.Font.NameOther == "Cambria Math")
                {
                    //xuất ảnh chứa công thức toán học
                    return null;
                }
                else
                {
                    //Tạm rào để sử dụng app demo
                    bool result = ((System.Windows.Application.Current as IAppGlobal).SelectedTheme.FontFamilies.Fonts).Any(o => o.Name == textRun.Font.NameAscii);
                    if (result)
                        eRun.Fontfamily = textRun.Font.NameAscii;
                    else eRun.Fontfamily = "Arial";
                    eRun.FontSize = textRun.Font.Size;
                    if (textRun.Font.Bold == Microsoft.Office.Core.MsoTriState.msoTrue || textRun.Font.Bold == Microsoft.Office.Core.MsoTriState.msoCTrue)
                        eRun.FontWeight = Text.FontWeight.Bold;
                    else eRun.FontWeight = Text.FontWeight.Normal;
                    if (textRun.Font.Italic == office.MsoTriState.msoTrue || textRun.Font.Italic == office.MsoTriState.msoCTrue)
                        eRun.FontStyle = Text.FontStyle.Italic;
                    else eRun.FontStyle = Text.FontStyle.Normal;
                    if (textRun.Font.Underline == office.MsoTriState.msoTrue || textRun.Font.Underline == office.MsoTriState.msoCTrue)
                    {
                        DataUnderline underline = new DataUnderline();
                        underline.Color = new SolidColor() { Color = ConvertColor(textRun.Font.Color.RGB) };
                        underline.Style = Text.UnderlineStyle.Single;
                        eRun.Underline = underline;
                    }
                    eRun.Foreground = new DataColor() { Name = ConvertColor(textRun.Font.Color.RGB) };
                    eRun.Text = textRun.Text;
                    eRun.ScriptOffset = (int)(textRun.Font.BaselineOffset * 100);
                    eParagraph.Inlines.Add(eRun);
                }
            }
            return eParagraph;
        }
        /// <summary>
        /// Lấy vị trí para trước nó có cấp bằng hoặc nhỏ hơn gần nhất so với para hiện tại
        /// </summary>
        /// <param name="list"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        public static int GetIndexList(List<DataParaPositionBullet> list, int index)
        {
            for (int i = index - 1; i >= 0; i--)
            {
                if (list[i].Position <= list[index].Position)
                    return i;
            }
            return -1;
        }
        public static void ConverterListParaToDataDocument(List<DataParaPositionBullet> list, DataDocument dataDocument)
        {
            DataParagraph _paraAdd = new DataParagraph();
            DataParagraph _para = new DataParagraph();
            //TextRange _textRange = new TextRange();
            int _countListPara = list.Count;
            int positionBlock = 0;
            for (int t = 0; t < _countListPara; t++)
            {
                //Trường hợp là Para
                if (list[t].Position == 0)
                {
                    list[t].Para.Parent = dataDocument;
                    dataDocument.Blocks.Add(list[t].Para);
                    positionBlock++;
                }
                else//Trường hợp là List
                {
                    //Nếu para trước nó là bullet=null hoặc nó là list đầu tiên thì tạo mới List
                    if (t == 0 || list[t - 1].Position == 0)
                    {
                        DataDocumentList _list = new DataDocumentList();
                        DataDocumentListItem _listItem = new DataDocumentListItem();
                        list[t].Para.Parent = _listItem;
                        _listItem.Blocks.Add(list[t].Para);
                        _listItem.TypeWordBullet = list[t].ListType;
                        _listItem.Parent = _list;
                        _list.Items.Add(_listItem);
                        dataDocument.Blocks.Add(_list);
                        positionBlock++;
                    }
                    else
                    {
                        int index = GetIndexList(list, t);
                        if (index != -1)
                        {
                            //Trường hợp không cùng cấp
                            if (list[t].Position - list[index].Position > 0)
                            {
                                list[t].Position = list[index].Position + 1;
                                if (list[index].Para.Parent is DataDocumentListItem)
                                {
                                    DataDocumentListItem _listItem = new DataDocumentListItem();
                                    DataDocumentList _listData = new DataDocumentList();
                                    DataDocumentListItem _listItem2 = new DataDocumentListItem();
                                    list[t].Para.Parent = _listItem2;
                                    _listItem2.Blocks.Add(list[t].Para);
                                    _listItem2.TypeWordBullet = list[t].ListType;
                                    _listItem2.Parent = _listData;
                                    _listData.Items.Add(_listItem2);
                                    _listData.Parent = _listItem;
                                    (list[index].Para.Parent as DataDocumentListItem).Blocks.Add(_listData);
                                }
                            }
                            //Trường hợp cùng cấp
                            else
                            {
                                if (list[index].Para.Parent is DataDocumentListItem && list[index].Para.Parent.Parent is DataDocumentList)
                                {
                                    DataDocumentListItem _listItem = new DataDocumentListItem();
                                    list[t].Para.Parent = _listItem;
                                    _listItem.Blocks.Add(list[t].Para);
                                    _listItem.TypeWordBullet = list[t].ListType;
                                    _listItem.Parent = list[index].Para.Parent.Parent;
                                    (list[index].Para.Parent.Parent as DataDocumentList).Items.Add(_listItem);
                                }
                            }
                        }
                    }
                }
            }
        }
        private DataDocumentListItem GetListItem(DataTextElementBase _parent, DataBlockBase _child)
        {
            DataDocumentListItem docListItem = new DataDocumentListItem();
            _child.Parent = docListItem;
            docListItem.Blocks.Add(_child);
            docListItem.Parent = _parent;
            return docListItem;
        }
        private void ShapeIsFlash(PageLayer _page, pp.Shape _shape, EAnimation _animation, string _idShape)
        {
            EFlashControl eFlash = new EFlashControl();
            eFlash.Name = "Flash";
            eFlash.ID = _idShape;
            eFlash.Angle = _shape.Rotation;
            eFlash.Height = _shape.Height * Utils.tlh;
            eFlash.Left = _shape.Left * Utils.tlw;
            eFlash.Width = _shape.Width * Utils.tlw;
            eFlash.Top = _shape.Top * Utils.tlh;
            eFlash.ZOder = _shape.ZOrderPosition;
            eFlash.Animations = _animation;
            //Lấy file flash trong gói sử dụng openxml
            string pathFlash = _shape.OLEFormat.Object.Movie;
            eFlash.Source = new Core.Model.Media.Flash();
            eFlash.Source.Path = pathFlash;
            _page.Children.Add(eFlash);
        }
        private void ShapeToImage(PageLayer _page, pp.Shape _shape, EAnimation _animation, string _idShape)
        {
            EImageContent eImage = new EImageContent();
            eImage.ID = _idShape;
            eImage.Width = _shape.Width * Utils.tlw;
            eImage.Height = _shape.Height * Utils.tlh;
            eImage.Top = _shape.Top * Utils.tlh;
            eImage.Left = _shape.Left * Utils.tlw;
            eImage.Angle = _shape.Rotation;
            eImage.ZOder = (int)_shape.ZOrderPosition;
            eImage.Animations = _animation;
            eImage.RectCrop = new ESuperPoint();
            //Export Shape ra ảnh
            string pathImage = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + ".png");
            _shape.Export(pathImage, pp.PpShapeFormat.ppShapeFormatPNG, 0, 0, pp.PpExportMode.ppScaleXY);
            eImage.Image = new Image();
            eImage.Image.Path = pathImage;
            _page.Children.Add(eImage);
        }

        //private void ShapeToStandardElement(PageLayer _page, pp.Shape _shape, EAnimation _animation, string _idShape)
        //{
        //    HelperClass helperClass = new HelperClass();
        //    EStandardElement eStandardElement = new EStandardElement();
        //    eStandardElement.ID = _idShape;
        //    eStandardElement.Width = _shape.Width * Utils.tlw;
        //    eStandardElement.Height = _shape.Height * Utils.tlh;
        //    eStandardElement.Top = _shape.Top * Utils.tlh;
        //    eStandardElement.Left = _shape.Left * Utils.tlw;
        //    eStandardElement.Angle = _shape.Rotation;
        //    eStandardElement.ZOder = (int)_shape.ZOrderPosition;
        //    eStandardElement.Animations = _animation;
        //    eStandardElement.ShapePresent = helperClass.GetShapePresent(_shape);
        //    eStandardElement.Border = new EBorder();
        //    eStandardElement.Border.Fill = helperClass.GetFillColor(_shape);
        //    eStandardElement.IsLikeSlideFill = eStandardElement.Border.Fill == null;

        //    //đối tượng này chứa path-- a dùng documentformat.openxml để đọc gói zip của PP

        //    _page.Children.Add(eStandardElement);
        //}

        /// <summary>
        /// RGB to hexa
        /// </summary>
        /// <param name="_oleColor"></param>
        /// <returns></returns>
        public string ConvertColor(int _oleColor)
        {
            System.Drawing.Color stdColor = System.Drawing.ColorTranslator.FromOle(_oleColor);
            string hexa = "#" + stdColor.R.ToString("X2") + stdColor.G.ToString("X2") + stdColor.B.ToString("X2");
            return hexa;
        }
        private P.Picture GetPictureTag(IEnumerable<P.Picture> _lstPic, string _idShape)
        {
            foreach (P.Picture pic in _lstPic)
            {
                if (pic.NonVisualPictureProperties.NonVisualDrawingProperties.Id == _idShape)
                    return pic;
            }
            return null;
        }




       
        #region Danh sách hàm tạo hiệu ứng đối tượng


        private int GetIndexAns(pp.Master _slide, pp.Shape _shape)
        {
            for (int i = 1; i <= _slide.TimeLine.MainSequence.Count; i++)
            {
                if (_slide.TimeLine.MainSequence[i].Shape == _shape)
                    return i;
            }
            return 0;
        }

        private int GetIndexAns(pp.CustomLayout _slide, pp.Shape _shape)
        {
            for (int i = 1; i <= _slide.TimeLine.MainSequence.Count; i++)
            {
                if (_slide.TimeLine.MainSequence[i].Shape == _shape)
                    return i;
            }
            return 0;
        }

        #endregion
    }
}
