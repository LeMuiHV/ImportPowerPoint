using INV.Elearing.Controls.Shapes;
using INV.Elearning.Animations;
using INV.Elearning.Animations.Enums;
using INV.Elearning.Animations.UI;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using office = Microsoft.Office.Core;
using pp = Microsoft.Office.Interop.PowerPoint;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using INV.Elearning.Core.Model.ObjectElement;
using DocumentFormat.OpenXml.Packaging;
using INV.Elearning.Text.Models;
using INV.Elearning.Text;
using INV.Elearning.Controls.Flash.Controls;
using INV.Elearning.ImageProcess.Models;
using INV.Elearning.Core;
using System.IO;
using DocumentFormat.OpenXml.Office2010.Drawing;
using INV.Elearning.DesignControl.Models;
using INV.Elearning.Text.ViewModels.Text;
using INV.Elearning.VideoElementControl.Model;
using INV.Elearning.Controls.Audio;

namespace INV.Elearning.ImportPowerPoint.Helper
{
    public partial class HelperClass
    {
        #region GetShape
        /// <summary>
        /// Lấy tất cả các Shape
        /// </summary>
        /// <param name="pageLayer"></param>
        /// <param name="shapes"></param>
        /// <param name="timeLine"></param>
        public ObjectElementBase GetShape(pp.Shape shape, MsoShapeType shapeType, pp.TimeLine timeLine, OpenXmlPart openXmlPart, bool returnImageIfEmpty = true)
        {
            string idShape = ObjectElementsHelper.RandomString(13);
            P.Shape shapeTag = GetShapeTag(shape.Id.ToString(), openXmlPart);
            switch (shapeType)
            {
                case office.MsoShapeType.msoFreeform:
                case office.MsoShapeType.msoAutoShape:
                    //1307 Anh Mùi
                    if (shape.HasTextFrame == MsoTriState.msoTrue && !string.IsNullOrEmpty(shape.TextFrame.TextRange.Text))
                    {
                        if (MathOrWordArt(shape, shapeTag))
                            return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                        else
                            return GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                    }
                    return GetEStandardElement(shape, timeLine, shapeTag, openXmlPart);
                case office.MsoShapeType.msoCallout:
                    return GetEStandardElement(shape, timeLine, shapeTag, openXmlPart);
                case office.MsoShapeType.msoGroup:
                    return GetGroupElement(shape, timeLine, openXmlPart);
                case office.MsoShapeType.msoLine:
                    return GetLineElement(shape, timeLine, shapeTag, openXmlPart);
                case office.MsoShapeType.msoOLEControlObject:
                    return GetFlashControl(shape, timeLine);
                case office.MsoShapeType.msoPicture:
                    return GetPicture(shape, timeLine, openXmlPart);
                case office.MsoShapeType.msoPlaceholder:
                    return GetPlaceHolder(shape, timeLine, shapeTag, openXmlPart);
                case office.MsoShapeType.msoMedia:
                    switch (shape.MediaType)
                    {
                        case pp.PpMediaType.ppMediaTypeMixed:
                            break;
                        case pp.PpMediaType.ppMediaTypeOther:
                            break;
                        case pp.PpMediaType.ppMediaTypeSound:
                            var elementAudio = GetAudioElement(shape, timeLine, openXmlPart);
                            if (elementAudio != null) return elementAudio;
                            break;
                        case pp.PpMediaType.ppMediaTypeMovie:
                            var element = GetVideoElement(shape, timeLine, openXmlPart);
                            if (element != null) return element;
                            break;
                    }
                    break;
                case office.MsoShapeType.msoTextBox:
                    //1307 Anh Mùi
                    if (shape.HasTextFrame == office.MsoTriState.msoTrue || shape.HasTextFrame == office.MsoTriState.msoCTrue)
                    {
                        if (MathOrWordArt(shape, shapeTag))
                            return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                        else
                            return GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                    }
                    break;
                case MsoShapeType.msoWebVideo:
                case MsoShapeType.msoSlicer:
                case MsoShapeType.msoInkComment:
                case MsoShapeType.msoInk:
                case MsoShapeType.msoDiagram:
                case MsoShapeType.msoCanvas:
                case MsoShapeType.msoScriptAnchor:
                case office.MsoShapeType.msoTextEffect:
                case MsoShapeType.msoLinkedPicture:
                case MsoShapeType.msoLinkedOLEObject:
                case MsoShapeType.msoFormControl:
                case MsoShapeType.msoEmbeddedOLEObject:
                case MsoShapeType.msoComment:
                case MsoShapeType.msoShapeTypeMixed:
                case office.MsoShapeType.msoChart:
                case office.MsoShapeType.msoTable:
                case office.MsoShapeType.msoSmartArt:
                    return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
            }
            return returnImageIfEmpty ? ShapeToImage(shape, timeLine, shapeTag, openXmlPart) : null;
        }
        #endregion

        /// <summary>
        /// Lấy Object Element Line
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="timeLine"></param>
        /// <param name="openXmlPart"></param>
        /// <returns></returns>
        private EStandardElement GetLineElement(pp.Shape shape, pp.TimeLine timeLine, P.Shape shapeTag, OpenXmlPart openXmlPart)
        {
            EStandardElement eStandardElement = new EStandardElement();
            SetBasicProperties(shape, eStandardElement, timeLine);
            LineInfo lineInfo = new LineInfo();
            lineInfo.HeadEnd = ArrowTypesConverter(shape.Line.BeginArrowheadStyle);
            lineInfo.TailEnd = ArrowTypesConverter(shape.Line.EndArrowheadStyle);
            lineInfo.StartPoint = new System.Windows.Point(0, 0);
            lineInfo.EndPoint = new System.Windows.Point(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
            eStandardElement.ShapePresent = lineInfo;
            eStandardElement.Border = GetBorderShape(shape, shapeTag, openXmlPart);
            eStandardElement.IsLikeSlideFill = eStandardElement.Border.Fill == null;
            return eStandardElement;
        }
        /// <summary>
        /// Lấy đối tượng văn bản
        /// </summary>
        /// <param name="_shape"></param>
        /// <param name="timeLine"></param>
        /// <param name="openXmlPart"></param>
        /// <returns></returns>
        private DataDocument GetDataDocument(pp.Shape _shape, pp.TimeLine timeLine, P.Shape shapeTag, OpenXmlPart openXmlPart)
        {
            //Thiếu textindent và strike words 
            DataDocument document = new DataDocument();
            SetBasicProperties(_shape, document, timeLine);
            SetStandardElementProperties(_shape, shapeTag, document, openXmlPart);
            document.Name = "TextEditor";
            document.TypeText = TypeText.Text;
            List<DataParaPositionBullet> ListParaPosition = new List<DataParaPositionBullet>();
            DataParaPositionBullet paraPositionBullet = null;
            document.Blocks.Clear();
            document.Container = new DataTextContainer();
            if (_shape.Type == MsoShapeType.msoTextBox)
                document.TypeTextContainer = TypeTextContainer.AutoChangeHeight;
            else
            document.TypeTextContainer = TypeTextContainer.ChangeSize;
            //Lấy nội dung 
            try
            {
                document.VerticalAlign = ConvertVerticalAlignText(_shape.TextFrame.VerticalAnchor);
                pp.BulletFormat hasBullet = _shape.TextFrame.TextRange.ParagraphFormat.Bullet;
                if (hasBullet.Visible == office.MsoTriState.msoCTrue || hasBullet.Visible == office.MsoTriState.msoTrue)
                {
                    foreach (pp.TextRange paragraph in _shape.TextFrame.TextRange.Paragraphs())
                    {
                        //Check đoạn văn bản trắng thì không thêm ký tự đầu dòng - trường hợp có thêm ký tự đầu dong
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
                        if (paraPositionBullet.Para?.Inlines?.Count > 0)
                        {
                            if ((paraPositionBullet.Para.Inlines.Last() as DataRun)?.Text.Length > 0 && (paraPositionBullet.Para.Inlines.Last() as DataRun).Text[(paraPositionBullet.Para.Inlines.Last() as DataRun).Text.Length-1]!='\r')
                                (paraPositionBullet.Para.Inlines.Last() as DataRun).Text = (paraPositionBullet.Para.Inlines.Last() as DataRun).Text + "\r";
                        }
                        paraPositionBullet.ListType = docListItem.TypeWordBullet;
                        paraPositionBullet.Position = level;
                        ListParaPosition.Add(paraPositionBullet);
                    }
                }
                else
                {
                    foreach (pp.TextRange paragraph in _shape.TextFrame.TextRange.Paragraphs())
                    {
                        paraPositionBullet = new DataParaPositionBullet();
                        paraPositionBullet.Para = GetEParagraph(paragraph);
                        if (paraPositionBullet.Para?.Inlines?.Count > 0)
                        {
                            if ((paraPositionBullet.Para.Inlines.Last() as DataRun)?.Text.Length>0&&(paraPositionBullet.Para.Inlines.Last() as DataRun).Text[(paraPositionBullet.Para.Inlines.Last() as DataRun).Text.Length - 1] != '\r')
                                (paraPositionBullet.Para.Inlines.Last() as DataRun).Text = (paraPositionBullet.Para.Inlines.Last() as DataRun).Text + "\r";
                        }
                        paraPositionBullet.ListType = null;
                        paraPositionBullet.Position = 0;
                        ListParaPosition.Add(paraPositionBullet);
                    }
                }
            }
            catch
            {
                //Exception prop of TextFrame
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
            return document;

        }

        /// <summary>
        /// hàm chuyển đổi vertical Align của text
        /// </summary>
        /// <param name="msoVerticalAnchor"></param>
        /// <returns></returns>
        private VerticalAlign ConvertVerticalAlignText(MsoVerticalAnchor msoVerticalAnchor)
        {
            switch (msoVerticalAnchor)
            {
                case MsoVerticalAnchor.msoAnchorTopBaseline:
                case MsoVerticalAnchor.msoAnchorTop:
                    return VerticalAlign.Top;
                case MsoVerticalAnchor.msoAnchorMiddle:
                    return VerticalAlign.Middle;
                case MsoVerticalAnchor.msoAnchorBottom:
                case MsoVerticalAnchor.msoAnchorBottomBaseLine:
                    return VerticalAlign.Bottom;
                default:
                    break;
            }
            return VerticalAlign.Middle;
        }

        /// <summary>
        /// Xuất hình ảnh đối tượng
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="timeLine"></param>
        /// <param name="openXmlPart"></param>
        /// <returns></returns>
        private EImageContent ShapeToImage(pp.Shape shape, pp.TimeLine timeLine, P.Shape shapeTag, OpenXmlPart openXmlPart)
        {
            EImageContent eImage = new EImageContent();
            SetBasicProperties(shape, eImage, timeLine);
            SetStandardElementProperties(shape, shapeTag, eImage, openXmlPart);
            eImage.RectCrop = new ESuperPoint();
            //1307 Anh Mùi - không set border cho shape phải export ra ảnh
            eImage.Border = new EBorder();
            //Export Shape ra ảnh
            string pathImage = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + ".png");
            try
            {
                shape.Export(pathImage, pp.PpShapeFormat.ppShapeFormatPNG, 0, 0, pp.PpExportMode.ppScaleXY);
            }
            catch
            {
                //Export shape to image
            }
            eImage.Image = new Image();
            if (File.Exists(pathImage))
                eImage.Image.Path = pathImage;
            eImage.Image.OriginPath = pathImage;
            return eImage;
        }

        private ObjectElementBase GetPlaceHolder(pp.Shape shape, pp.TimeLine timeLine, P.Shape shapeTag, OpenXmlPart openXmlPart)
        {
            switch (shape.PlaceholderFormat.Type)
            {
                case PpPlaceholderType.ppPlaceholderTitle:
                case PpPlaceholderType.ppPlaceholderBody:
                case PpPlaceholderType.ppPlaceholderCenterTitle:
                case PpPlaceholderType.ppPlaceholderVerticalTitle:
                    //1307 Anh Mùi
                    if (shape.HasTextFrame == MsoTriState.msoTrue && !string.IsNullOrEmpty(shape.TextFrame.TextRange.Text) && MathOrWordArt(shape, shapeTag))
                        return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                    DataDocument document = GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                    document.PlaceHolderType = Core.View.PlaceHolderEnum.Normal;
                    // SetBasicProperties(shape, document, timeLine);
                    // SetStandardElementProperties(shape, openXmlPart, document);
                    //document.TypeText = TypeText.Title;
                    //document.TypeTextContainer = TypeTextContainer.Title;
                    return document;
                case PpPlaceholderType.ppPlaceholderSubtitle:
                    if (shape.HasTextFrame == MsoTriState.msoTrue && !string.IsNullOrEmpty(shape.TextFrame.TextRange.Text) && MathOrWordArt(shape, shapeTag))
                        return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                    DataDocument documentSubTitle = GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                    documentSubTitle.PlaceHolderType = Core.View.PlaceHolderEnum.Normal;
                    SetBasicProperties(shape, documentSubTitle, timeLine);
                    SetStandardElementProperties(shape, shapeTag, documentSubTitle, openXmlPart);
                    return documentSubTitle;
                case PpPlaceholderType.ppPlaceholderVerticalBody:
                case PpPlaceholderType.ppPlaceholderHeader:
                case PpPlaceholderType.ppPlaceholderMixed:
                    //1307 Anh Mùi
                    if (shape.HasTextFrame == MsoTriState.msoTrue && !string.IsNullOrEmpty(shape.TextFrame.TextRange.Text) && MathOrWordArt(shape, shapeTag))
                        return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                    DataDocument documentText = GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                    SetBasicProperties(shape, documentText, timeLine);
                    documentText.PlaceHolderType = Core.View.PlaceHolderEnum.Normal;
                    SetStandardElementProperties(shape, shapeTag, documentText, openXmlPart);
                    //documentText.TypeText = TypeText.Text;
                    //documentText.TypeTextContainer = TypeTextContainer.None;
                    return documentText;
                case PpPlaceholderType.ppPlaceholderTable:
                case PpPlaceholderType.ppPlaceholderVerticalObject:
                case PpPlaceholderType.ppPlaceholderObject:
                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        //1307 Anh Mùi
                        if (!string.IsNullOrEmpty(shape.TextFrame.TextRange.Text) && MathOrWordArt(shape, shapeTag))
                            return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                        DataDocument documentObject = GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                        SetBasicProperties(shape, documentObject, timeLine);
                        SetStandardElementProperties(shape, shapeTag, documentObject, openXmlPart);
                        documentObject.PlaceHolderType = Core.View.PlaceHolderEnum.Normal;
                        //documentObject.TypeText = TypeText.Text;
                        //documentObject.TypeTextContainer = TypeTextContainer.None;
                        return documentObject;
                    }
                    else
                    {
                        if (shape.PlaceholderFormat.ContainedType != MsoShapeType.msoAutoShape)
                        {
                            var element = GetShape(shape, shape.PlaceholderFormat.ContainedType, timeLine, openXmlPart, false);
                            if (element != null) return element;
                        }
                        else
                        {
                            EPlaceHolder ePlaceHolder = new EPlaceHolder();
                            SetBasicProperties(shape, ePlaceHolder, timeLine);
                            SetStandardElementProperties(shape, shapeTag, ePlaceHolder, openXmlPart);
                            ePlaceHolder.Type = Core.Model.Theme.PlaceHolderType.Content;
                            ePlaceHolder.PlaceHolderType = Core.View.PlaceHolderEnum.Normal;
                            return ePlaceHolder;
                        }
                    }
                    break;
                case PpPlaceholderType.ppPlaceholderOrgChart:
                case PpPlaceholderType.ppPlaceholderChart:
                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        //1307 Anh Mùi
                        if (!string.IsNullOrEmpty(shape.TextFrame.TextRange.Text) && MathOrWordArt(shape, shapeTag))
                            return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                        DataDocument documentChart = GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                        SetBasicProperties(shape, documentChart, timeLine);
                        SetStandardElementProperties(shape, shapeTag, documentChart, openXmlPart);
                        //documentChart.TypeText = TypeText.Text;
                        //documentChart.TypeTextContainer = TypeTextContainer.None;
                        return documentChart;
                    }
                    else if (shape.PlaceholderFormat.ContainedType != MsoShapeType.msoAutoShape)
                    {
                        var element2 = GetShape(shape, shape.PlaceholderFormat.ContainedType, timeLine, openXmlPart, false);
                        if (element2 != null) return element2;
                    }
                    else
                    {
                        EPlaceHolder ePlaceHolderChart = new EPlaceHolder();
                        SetBasicProperties(shape, ePlaceHolderChart, timeLine);
                        SetStandardElementProperties(shape, shapeTag, ePlaceHolderChart, openXmlPart);
                        ePlaceHolderChart.Type = Core.Model.Theme.PlaceHolderType.Chart;
                        ePlaceHolderChart.PlaceHolderType = Core.View.PlaceHolderEnum.Normal;
                        return ePlaceHolderChart;
                    }
                    break;
                case PpPlaceholderType.ppPlaceholderBitmap:
                case PpPlaceholderType.ppPlaceholderPicture:
                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        //1307 Anh Mùi
                        if (!string.IsNullOrEmpty(shape.TextFrame.TextRange.Text) && MathOrWordArt(shape, shapeTag))
                            return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                        DataDocument documentPicture = GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                        SetBasicProperties(shape, documentPicture, timeLine);
                        SetStandardElementProperties(shape, shapeTag, documentPicture, openXmlPart);
                        documentPicture.PlaceHolderType = Core.View.PlaceHolderEnum.Normal;
                        //documentPicture.TypeText = TypeText.Text;
                        //documentPicture.TypeTextContainer = TypeTextContainer.None;
                        return documentPicture;
                    }
                    else if (shape.PlaceholderFormat.ContainedType != MsoShapeType.msoAutoShape)
                    {
                        var element3 = GetShape(shape, shape.PlaceholderFormat.ContainedType, timeLine, openXmlPart, false);
                        if (element3 != null) return element3;
                    }
                    else
                    {
                        EPlaceHolder ePlaceHolderImage = new EPlaceHolder();
                        SetBasicProperties(shape, ePlaceHolderImage, timeLine);
                        SetStandardElementProperties(shape, shapeTag, ePlaceHolderImage, openXmlPart);
                        ePlaceHolderImage.Type = Core.Model.Theme.PlaceHolderType.Picture;
                        ePlaceHolderImage.PlaceHolderType = Core.View.PlaceHolderEnum.Normal;
                        return ePlaceHolderImage;
                    }
                    break;
                case PpPlaceholderType.ppPlaceholderMediaClip:
                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        //1307 Anh Mùi
                        if (!string.IsNullOrEmpty(shape.TextFrame.TextRange.Text) && MathOrWordArt(shape, shapeTag))
                            return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                        DataDocument documentMedia = GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                        SetBasicProperties(shape, documentMedia, timeLine);
                        SetStandardElementProperties(shape, shapeTag, documentMedia, openXmlPart);
                        documentMedia.PlaceHolderType = Core.View.PlaceHolderEnum.Normal;
                        //documentMedia.TypeText = TypeText.Text;
                        //documentMedia.TypeTextContainer = TypeTextContainer.None;
                        return documentMedia;
                    }
                    else if (shape.PlaceholderFormat.ContainedType != MsoShapeType.msoAutoShape)
                    {
                        var element4 = GetShape(shape, shape.PlaceholderFormat.ContainedType, timeLine, openXmlPart, false);
                        if (element4 != null) return element4;
                    }
                    else
                    {
                        EPlaceHolder ePlaceHolderVideo = new EPlaceHolder();
                        SetBasicProperties(shape, ePlaceHolderVideo, timeLine);
                        SetStandardElementProperties(shape, shapeTag, ePlaceHolderVideo, openXmlPart);
                        ePlaceHolderVideo.Type = Core.Model.Theme.PlaceHolderType.Video;
                        ePlaceHolderVideo.PlaceHolderType = Core.View.PlaceHolderEnum.Normal;
                        return ePlaceHolderVideo;
                    }
                    break;
                case PpPlaceholderType.ppPlaceholderSlideNumber:
                    //1307 Anh Mùi
                    if (shape.HasTextFrame == MsoTriState.msoTrue && !string.IsNullOrEmpty(shape.TextFrame.TextRange.Text) && MathOrWordArt(shape, shapeTag))
                        return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                    DataDocument documentSlideNumber = GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                    documentSlideNumber.TypeText = TypeText.SlideNumber;
                    documentSlideNumber.PlaceHolderType = Core.View.PlaceHolderEnum.SlideNumber;
                    //documentSlideNumber.TypeTextContainer = TypeTextContainer.None;
                    return documentSlideNumber;
                case PpPlaceholderType.ppPlaceholderFooter:
                    //1307 Anh Mùi
                    if (shape.HasTextFrame == MsoTriState.msoTrue && !string.IsNullOrEmpty(shape.TextFrame.TextRange.Text) && MathOrWordArt(shape, shapeTag))
                        return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                    DataDocument documentFooter = GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                    documentFooter.TypeText = TypeText.Footers;
                    documentFooter.PlaceHolderType = Core.View.PlaceHolderEnum.Footer;
                    //documentFooter.TypeTextContainer = TypeTextContainer.None;
                    return documentFooter;
                case PpPlaceholderType.ppPlaceholderDate:
                    //1307 Anh Mùi
                    if (shape.HasTextFrame == MsoTriState.msoTrue && !string.IsNullOrEmpty(shape.TextFrame.TextRange.Text) && MathOrWordArt(shape, shapeTag))
                        return ShapeToImage(shape, timeLine, shapeTag, openXmlPart);
                    DataDocument documentDate = GetDataDocument(shape, timeLine, shapeTag, openXmlPart);
                    documentDate.TypeText = TypeText.DateTime;
                    documentDate.PlaceHolderType = Core.View.PlaceHolderEnum.DateTime;
                    //documentDate.TypeTextContainer = TypeTextContainer.None;
                    return documentDate;
                default:
                    break;
            }
            return null;
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
            eParagraph.Margin.Top = _paragraph.ParagraphFormat.SpaceWithin;
            eParagraph.Margin.Bottom = _paragraph.ParagraphFormat.SpaceBefore;
            eParagraph.TextAlign = Text.HorizontalAlign.Left;
            switch (_paragraph.ParagraphFormat.Alignment)
            {
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
            }
            foreach (pp.TextRange textRun in _paragraph.Runs())
            {
                DataRun eRun = new DataRun();
                if (textRun.Font.NameAscii == "Cambria Math" && textRun.Font.NameOther == "Cambria Math")
                {
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
        /// Lấy đối tượng Flash
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="timeLine"></param>
        /// <param name="openXmlPart"></param>
        /// <returns></returns>
        public EFlashControl GetFlashControl(pp.Shape shape, pp.TimeLine timeLine)
        {
            EFlashControl eFlash = new EFlashControl();
            eFlash.Name = "Flash";
            SetBasicProperties(shape, eFlash, timeLine);
            //Lấy file flash trong gói sử dụng openxml
            string pathFlash = shape.OLEFormat.Object.Movie;
            eFlash.Source = new Core.Model.Media.Flash();
            eFlash.Source.Path = pathFlash;
            return eFlash;
        }

        /// <summary>
        /// Lấy đối tượng Picture
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="timeLine"></param>
        /// <param name="openXmlPart"></param>
        /// <returns></returns>
        public EImageContent GetPicture(pp.Shape shape, pp.TimeLine timeLine, OpenXmlPart openXmlPart)
        {
            EImageContent eImage = new EImageContent();
            eImage.Name = "Image";
            eImage.RectCrop = new ESuperPoint();
            P.Picture picTag = null;
            IEnumerable<DocumentFormat.OpenXml.Presentation.Picture> lstPic = null;
            ImagePart part = null;
            if (openXmlPart is SlidePart slidePart)
            {
                lstPic = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    picTag = GetPictureTag(lstPic, shape.Id.ToString());
                    string rID = picTag.BlipFill.Blip.Embed.Value;
                    part = (ImagePart)slidePart.GetPartById(rID);
                }
            }
            else if (openXmlPart is SlideMasterPart slideMasterPart)
            {
                lstPic = slideMasterPart.SlideMaster.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    picTag = GetPictureTag(lstPic, shape.Id.ToString());
                    string rID = picTag.BlipFill.Blip.Embed.Value;
                    part = (ImagePart)slideMasterPart.GetPartById(rID);
                }
            }
            else if (openXmlPart is SlideLayoutPart slideLayoutPart)
            {
                lstPic = slideLayoutPart.SlideLayout.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    picTag = GetPictureTag(lstPic, shape.Id.ToString());
                    string rID = picTag.BlipFill.Blip.Embed.Value;
                    part = (ImagePart)slideLayoutPart.GetPartById(rID);
                }
            }
            if (lstPic != null && part != null && picTag != null)
            {
                SetBasicProperties(shape, eImage, timeLine);
                SetStandardElementProperties(shape, picTag, eImage, openXmlPart);
                string pathImage = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, Path.GetFileName(part.Uri.ToString()));
                if (!File.Exists(pathImage))
                    Utils.CopyStream(part.GetStream(), pathImage);
                eImage.Image = new Image(pathImage);
            }
            return eImage;
        }

        /// <summary>
        /// Lấy đối tượng Video
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="timeLine"></param>
        /// <param name="openXmlPart"></param>
        /// <returns></returns>
        public EVideoPlayer GetVideoElement(pp.Shape shape, pp.TimeLine timeLine, OpenXmlPart openXmlPart)
        {
            IEnumerable<DocumentFormat.OpenXml.Presentation.Picture> lstPic = null;
            P.Picture videoTag = null;
            if (openXmlPart is SlidePart slidePart)
            {
                lstPic = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                    videoTag = GetPictureTag(lstPic, shape.Id.ToString());
            }
            else if (openXmlPart is SlideMasterPart slideMasterPart)
            {
                lstPic = slideMasterPart.SlideMaster.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                    videoTag = GetPictureTag(lstPic, shape.Id.ToString());
            }
            else if (openXmlPart is SlideLayoutPart slideLayoutPart)
            {
                lstPic = slideLayoutPart.SlideLayout.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                    videoTag = GetPictureTag(lstPic, shape.Id.ToString());
            }
            if (lstPic != null && videoTag != null)
            {
                var videoFromFile = videoTag.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties.Descendants<D.VideoFromFile>().Select(p => p);
                ReferenceRelationship relationship = openXmlPart.GetReferenceRelationship(videoFromFile.First().Link.Value); ;
                
                EVideoPlayer eVideoPlayer = new EVideoPlayer();
                SetBasicProperties(shape, eVideoPlayer, timeLine);
                //Lấy file video trong gói sử dụng openxml
                if (relationship is VideoReferenceRelationship)
                {
                    string pathVideo = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + Path.GetExtension(relationship.Uri.ToString()));
                    Utils.CopyStream(((VideoReferenceRelationship)relationship).DataPart.GetStream(), pathVideo);
                    eVideoPlayer.Source = new Core.Model.Media.Video();
                    eVideoPlayer.Source.Path = pathVideo;
                }
                else return null;
                //Lấy avata video - trường hợp video trong file ppt convert sang pptx sẽ chỉ lấy được avata
                eVideoPlayer.ImageFirst = new Image();
                string rId = videoTag.BlipFill.Blip.Embed.Value;
                OpenXmlPart avtPart = openXmlPart.GetPartById(rId);
                string pathAvata = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + Path.GetExtension(avtPart.Uri.ToString()));
                Utils.CopyStream(avtPart.GetStream(), pathAvata);
                eVideoPlayer.ImageFirst.Path = pathAvata;
                return eVideoPlayer;
            }
            return null;
        }

        public EAudioElement GetAudioElement(pp.Shape shape, pp.TimeLine timeLine, OpenXmlPart openXmlPart)
        {
            EAudioElement eAudio = new EAudioElement();
            SetBasicProperties(shape, eAudio, timeLine);
            //Lấy file audio trong gói sử dụng openxml
            IEnumerable<DocumentFormat.OpenXml.Presentation.Picture> lstPic = null;
            P.Picture audioTag = null;
            if (openXmlPart is SlidePart slidePart)
            {
                lstPic = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                    audioTag = GetPictureTag(lstPic, shape.Id.ToString());
            }
            else if (openXmlPart is SlideMasterPart slideMasterPart)
            {
                lstPic = slideMasterPart.SlideMaster.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                    audioTag = GetPictureTag(lstPic, shape.Id.ToString());
            }
            else if (openXmlPart is SlideLayoutPart slideLayoutPart)
            {
                lstPic = slideLayoutPart.SlideLayout.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                    audioTag = GetPictureTag(lstPic, shape.Id.ToString());
            }

            var audioFromFile = audioTag.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties.Descendants<D.AudioFromFile>().Select(p => p);
            ReferenceRelationship relationship = openXmlPart.GetReferenceRelationship(audioFromFile.First().Link.Value);
            if (relationship is AudioReferenceRelationship)
            {
                string pathAudio = Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + Path.GetExtension(relationship.Uri.ToString()));
                Utils.CopyStream(((AudioReferenceRelationship)relationship).DataPart.GetStream(), pathAudio);
                eAudio.AudioFileName = new Core.Model.Media.Audio();
                eAudio.AudioFileName.Path = pathAudio;
            }
            else return null;
            return eAudio;
        }
        private void SetEffectImage(P.Picture picTag, EImageContent eImage)
        {
            foreach (var blip in picTag.BlipFill.Blip)
            {
                if (blip is D.BlipExtensionList extensionList)
                {
                    foreach (var ext in extensionList)
                    {
                        if (ext is D.BlipExtension extension)
                        {
                            foreach (var prop in extension)
                            {
                                if (prop is ImageProperties imageProperties)
                                {
                                    foreach (ImageEffect effect in imageProperties.ImageLayer)
                                    {
                                        SetImageEffect(effect, eImage);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void SetImageEffect(ImageEffect imageEffect, EImageContent eImage)
        {
            if (imageEffect.BrightnessContrast != null)
            {
                eImage.ContrastAdjustEffect = new EContrastAdjustEffect();
                if (imageEffect.BrightnessContrast.Bright != null)
                {
                    eImage.ContrastAdjustEffect.Brightness = imageEffect.BrightnessContrast.Bright;
                }
                if (imageEffect.BrightnessContrast.Contrast != null)
                {
                    eImage.ContrastAdjustEffect.Contrast = imageEffect.BrightnessContrast.Contrast;
                }
            }
            if (imageEffect.SharpenSoften != null)
            {
                eImage.SharpenEffect = new ESharpenEffect();
                eImage.SharpenEffect.Amount = imageEffect.SharpenSoften.Amount;
            }
            if (imageEffect.Saturation != null)
            {
                eImage.ColorSaturationEffect = new EColorSaturationEffect();
                eImage.ColorSaturationEffect.Saturation = imageEffect.Saturation.SaturationAmount;
            }
            if (imageEffect.ColorTemperature != null)
            {
                eImage.ToonShaderEffect = new EToonShaderEffect();
                eImage.ToonShaderEffect.Levels = imageEffect.ColorTemperature.ColorTemperatureValue;
            }
        }

        /// <summary>
        /// hàm lấy picture tag
        /// </summary>
        /// <param name="_lstPic"></param>
        /// <param name="_idShape"></param>
        /// <returns></returns>
        private P.Picture GetPictureTag(IEnumerable<P.Picture> _lstPic, string _idShape)
        {
            foreach (P.Picture pic in _lstPic)
            {
                if (pic.NonVisualPictureProperties.NonVisualDrawingProperties.Id == _idShape)
                    return pic;
            }
            return null;
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

        /// <summary>
        /// Hàm chuyển đổi từ ListPara sang DataDocument
        /// </summary>
        /// <param name="list"></param>
        /// <param name="dataDocument"></param>
        public void ConverterListParaToDataDocument(List<DataParaPositionBullet> list, DataDocument dataDocument)
        {
            DataParagraph _paraAdd = new DataParagraph();
            DataParagraph _para = new DataParagraph();
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



        public DataDocument CreateDataDocument(string content, double fontSize, string fontFamily, DataColor foreground, HorizontalAlign align, VerticalAlign verticalAlign)
        {
            DataDocument _result = new DataDocument();
            _result.IsQuestion = false;
            _result.Padding = new EThickness() { Top = 5, Bottom = 5, Left = 5, Right = 5 };
            _result.VerticalAlign = verticalAlign;
            DataParagraph _para = new DataParagraph();
            DataRun _run = new DataRun();
            _run.Fontfamily = fontFamily;
            _run.FontSize = fontSize;
            _run.FontStyle = Text.FontStyle.Normal;
            _run.Foreground = foreground;
            _run.Text = content + "\r";
            _run.StrikeThrough = StrikeThrough.None;
            _run.Underline = new DataUnderline();
            _run.Underline.Style = UnderlineStyle.None;
            _run.Parent = _para;
            _para.Inlines.Add(_run);
            _para.TextIndent = 0;
            _para.Parent = _result;
            _para.LineHeight = 1;
            _para.TextAlign = align;
            _para.Margin = new EThickness();
            _result.Blocks.Add(_para);


            return _result;

        }


    }
}
