using DocumentFormat.OpenXml.Packaging;
using INV.Elearning.Core.Model;
using Microsoft.Office.Core;
using pp = Microsoft.Office.Interop.PowerPoint;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using INV.Elearning.Core.Helper;
using DocumentFormat.OpenXml;

namespace INV.Elearning.ImportPowerPoint.Helper
{
    /// <summary>
    /// Lớp hỗ trợ lấy màu
    /// </summary>
    public partial class HelperClass
    {
        /// <summary>
        /// Lấy màu 
        /// </summary>
        /// <returns></returns>
        public ColorBase GetFillColor(pp.Shape shape, OpenXmlCompositeElement shapeTag, OpenXmlPart openXmlPart)
        {
            if (shapeTag == null) return null;
            if (shapeTag is P.Shape _shape)
            {
                foreach (var item in _shape.ShapeProperties)
                {
                    if (item is NoFill noFill)
                    {
                        return new SolidColor() { Name = Brushes.Transparent.ToString() };
                    }
                    if (item is SolidFill solidFill)
                    {
                        return GetSolidColor(shape.Fill.ForeColor, shape.Fill.Transparency);
                    }
                    if (item is PatternFill patternFill)
                    {
                        return GetPatternColor(shape.Fill.Pattern, shape.Fill.BackColor, shape.Fill.ForeColor);
                    }
                    if (item is GradientFill gradientFill)
                    {
                        if (shape.Fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalDown || shape.Fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalUp || shape.Fill.GradientStyle == MsoGradientStyle.msoGradientHorizontal || shape.Fill.GradientStyle == MsoGradientStyle.msoGradientVertical)
                        {
                            return GetLinearColor(shape.Fill.GradientStops, shape.Fill.GradientStyle, shape.Fill.GradientAngle);
                        }
                        else
                        {
                            return GetRadialColor(shape.Fill.GradientStops, shape.Fill.GradientStyle);
                        }
                    }
                    if (item is BlipFill blipFill)
                    {
                        return GetImageColor(shape, openXmlPart, blipFill);
                    }
                }
            }
            else if (shapeTag is P.Picture picture)
            {
                foreach (var item in picture.ShapeProperties)
                {

                    if (item is NoFill noFill)
                    {
                        return new SolidColor() { Name = Brushes.Transparent.ToString() };
                    }
                    if (item is SolidFill solidFill)
                    {
                        return GetSolidColor(shape.Fill.ForeColor);
                    }
                    if (item is PatternFill patternFill)
                    {
                        return GetPatternColor(shape.Fill.Pattern, shape.Fill.BackColor, shape.Fill.ForeColor);
                    }
                    if (item is GradientFill gradientFill)
                    {
                        if (shape.Fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalDown || shape.Fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalUp || shape.Fill.GradientStyle == MsoGradientStyle.msoGradientHorizontal || shape.Fill.GradientStyle == MsoGradientStyle.msoGradientVertical)
                        {
                            return GetLinearColor(shape.Fill.GradientStops, shape.Fill.GradientStyle, shape.Fill.GradientAngle);
                        }
                        else
                        {
                            return GetRadialColor(shape.Fill.GradientStops, shape.Fill.GradientStyle);
                        }
                    }
                    if (item is BlipFill blipFill)
                    {
                        return GetImageColor(shape, openXmlPart, blipFill);
                    }
                }
            }
            switch (shape.Fill.Type)
            {
                case MsoFillType.msoFillSolid:
                    if (shape.Fill.Visible == MsoTriState.msoTrue)
                        return GetSolidColor(shape.Fill.ForeColor, shape.Fill.Transparency);
                    break;
                case MsoFillType.msoFillPatterned:
                    return GetPatternColor(shape.Fill.Pattern, shape.Fill.BackColor, shape.Fill.ForeColor);
                case MsoFillType.msoFillGradient:
                    if (shape.Fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalDown || shape.Fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalUp || shape.Fill.GradientStyle == MsoGradientStyle.msoGradientHorizontal || shape.Fill.GradientStyle == MsoGradientStyle.msoGradientVertical)
                    {
                        return GetLinearColor(shape.Fill.GradientStops, shape.Fill.GradientStyle, shape.Fill.GradientAngle);
                    }
                    else
                    {
                        return GetRadialColor(shape.Fill.GradientStops, shape.Fill.GradientStyle);
                    }
            }
            return new SolidColor() { Name = Brushes.Transparent.ToString() };

        }
        /// <summary>
        /// lấy màu background Slide
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="slideTag"></param>
        /// <param name="openXmlPart"></param>
        /// <returns></returns>
        public ColorBase GetFillColor(pp.FillFormat fill, P.Background background, OpenXmlPart openXmlPart)
        {
            if (background == null || background.BackgroundProperties == null)
            {
                switch (fill.Type)
                {
                    case MsoFillType.msoFillMixed:
                        return new SolidColor() { Name = Brushes.Transparent.ToString() };
                    case MsoFillType.msoFillSolid:
                        return GetSolidColor(fill.ForeColor);
                    case MsoFillType.msoFillPatterned:
                        return GetPatternColor(fill.Pattern, fill.BackColor, fill.ForeColor);
                    case MsoFillType.msoFillGradient:
                        if (fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalDown || fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalUp || fill.GradientStyle == MsoGradientStyle.msoGradientHorizontal || fill.GradientStyle == MsoGradientStyle.msoGradientVertical)
                        {
                            return GetLinearColor(fill.GradientStops, fill.GradientStyle, fill.GradientAngle);
                        }
                        else
                        {
                            return GetRadialColor(fill.GradientStops, fill.GradientStyle);
                        }
                    case MsoFillType.msoFillPicture:

                        break;
                }
            }
            else
            {
                foreach (var item in background.BackgroundProperties)
                {
                    if (item is NoFill noFill)
                    {
                        return new SolidColor() { Name = Brushes.Transparent.ToString() };
                    }
                    if (item is SolidFill solidFill)
                    {
                        return GetSolidColor(fill.ForeColor);
                    }
                    if (item is PatternFill patternFill)
                    {
                        return GetPatternColor(fill.Pattern, fill.BackColor, fill.ForeColor);
                    }
                    if (item is GradientFill gradientFill)
                    {
                        if (fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalDown || fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalUp || fill.GradientStyle == MsoGradientStyle.msoGradientHorizontal || fill.GradientStyle == MsoGradientStyle.msoGradientVertical)
                        {
                            return GetLinearColor(fill.GradientStops, fill.GradientStyle, fill.GradientAngle);
                        }
                        else
                        {
                            return GetRadialColor(fill.GradientStops, fill.GradientStyle);
                        }
                    }
                    if (item is BlipFill blipFill)
                    {
                        return GetImageColor(fill, openXmlPart, blipFill);
                    }
                }

            }
            return null;
        }

        public ColorBase GetFillColorSlideMaster(pp.FillFormat fill, P.CommonSlideData commonSlide, OpenXmlPart openXmlPart)
        {
            if (commonSlide.Background != null && commonSlide.Background.BackgroundProperties != null)
            {
                switch (fill.Type)
                {
                    case MsoFillType.msoFillMixed:
                        return new SolidColor() { Name = Brushes.Transparent.ToString() };
                    case MsoFillType.msoFillSolid:
                        return GetSolidColor(fill.ForeColor);
                    case MsoFillType.msoFillPatterned:
                        return GetPatternColor(fill.Pattern, fill.BackColor, fill.ForeColor);
                    case MsoFillType.msoFillGradient:
                        if (fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalDown || fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalUp || fill.GradientStyle == MsoGradientStyle.msoGradientVertical)
                        {
                            return GetLinearColor(fill.GradientStops, fill.GradientStyle, fill.GradientAngle);
                        }
                        else
                        {
                            return GetRadialColor(fill.GradientStops, fill.GradientStyle);
                        }
                    case MsoFillType.msoFillTextured:
                    case MsoFillType.msoFillPicture:
                        foreach (var item in commonSlide.Background.BackgroundProperties)
                        {
                            if (item is DocumentFormat.OpenXml.Drawing.BlipFill blipFill)
                                return GetImageColor(fill, openXmlPart, blipFill);

                        }
                        break;
                }
            }
            else
            {
                switch (fill.Type)
                {
                    case MsoFillType.msoFillMixed:
                        return new SolidColor() { Name = Brushes.Transparent.ToString() };
                    case MsoFillType.msoFillSolid:
                        return GetSolidColor(fill.ForeColor);
                    case MsoFillType.msoFillPatterned:
                        return GetPatternColor(fill.Pattern, fill.BackColor, fill.ForeColor);
                    case MsoFillType.msoFillGradient:
                        if (fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalDown || fill.GradientStyle == MsoGradientStyle.msoGradientDiagonalUp || fill.GradientStyle == MsoGradientStyle.msoGradientVertical)
                        {
                            return GetLinearColor(fill.GradientStops, fill.GradientStyle, fill.GradientAngle);
                        }
                        else
                        {
                            return GetRadialColor(fill.GradientStops, fill.GradientStyle);
                        }
                    case MsoFillType.msoFillPicture:

                        break;
                }
            }

            return null;
        }

        /// <summary>
        /// Lấy màu border
        /// </summary>
        /// <returns></returns>
        public ColorBase GetBorderBrushColor(pp.Shape shape, OpenXmlCompositeElement shapeTag, OpenXmlPart openXmlPart)
        {
            if (shapeTag == null) return null;
            if (shapeTag is P.Shape _shape)
            {
                if (_shape.ShapeProperties != null)
                {
                    foreach (var item in _shape.ShapeProperties)
                    {
                        if (item is Outline outline)
                        {
                            return GetSolidColor(shape.Line.ForeColor);

                        }
                    }
                    if (_shape.ShapeStyle?.FillReference?.SchemeColor != null)
                    {
                        return GetSolidColor(shape.Line.ForeColor);
                    }

                }
            }
            return new SolidColor() { Name = "Transparent" };
        }

        /// <summary>
        /// Lấy màu shape
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="openXmlPart"></param>
        /// <param name="blipFill"></param>
        /// <returns></returns>
        private ImageColor GetImageColor(pp.Shape shape, OpenXmlPart openXmlPart, BlipFill blipFill)
        {
            ImagePart part = null;
            IEnumerable<DocumentFormat.OpenXml.Presentation.Picture> lstPic = null;

            if (openXmlPart is SlidePart slidePart)
            {
                lstPic = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    string rID = blipFill.Blip.Embed.Value;
                    part = (ImagePart)slidePart.GetPartById(rID);
                }
            }
            else if (openXmlPart is SlideMasterPart slideMasterPart)
            {
                lstPic = slideMasterPart.SlideMaster.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    string rID = blipFill.Blip.Embed.Value;
                    part = (ImagePart)slideMasterPart.GetPartById(rID);
                }
            }
            else if (openXmlPart is SlideLayoutPart slideLayoutPart)
            {
                lstPic = slideLayoutPart.SlideLayout.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    string rID = blipFill.Blip.Embed.Value;
                    part = (ImagePart)slideLayoutPart.GetPartById(rID);
                }
            }
            if (lstPic != null && part != null)
            {
                string pathImage = System.IO.Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + System.IO.Path.GetExtension(part.Uri.ToString()));
                Utils.CopyStream(part.GetStream(), pathImage);
                ImageColor imageColor = new ImageColor();
                imageColor.Source = new Core.Model.Image(pathImage);
                imageColor.Opacity = 1 - shape.Fill.Transparency;
                imageColor.OffsetX = shape.Fill.TextureOffsetX;
                imageColor.OffsetY = shape.Fill.TextureOffsetY;
                imageColor.IsTiled = shape.Fill.TextureTile == MsoTriState.msoTrue;
                imageColor.Alignment = ConvertImageAligment(shape.Fill.TextureAlignment);
                imageColor.ScaleX = shape.Fill.TextureHorizontalScale;
                imageColor.ScaleY = shape.Fill.TextureVerticalScale;
                imageColor.RotateWidthShape = shape.Fill.RotateWithObject == MsoTriState.msoTrue;
                foreach (var item in blipFill)
                {
                    if (item is Stretch stretch)
                    {
                        if (stretch.FillRectangle.Top != null)
                            imageColor.OffsetTop = stretch.FillRectangle.Top / 100000;
                        if (stretch.FillRectangle.Left != null)
                            imageColor.OffsetLeft = stretch.FillRectangle.Left / 100000;
                        if (stretch.FillRectangle.Right != null)
                            imageColor.OffsetRight = stretch.FillRectangle.Right / 100000;
                        if (stretch.FillRectangle.Bottom != null)
                            imageColor.OffsetBottom = stretch.FillRectangle.Bottom / 100000;
                    }
                }
                return imageColor;
            }
            return null;
        }

        /// <summary>
        /// Lấy màu image background Slidemaster
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="openXmlPart"></param>
        /// <param name="blipFill"></param>
        /// <returns></returns>
        private ImageColor GetImageColor(pp.FillFormat fill, OpenXmlPart openXmlPart, DocumentFormat.OpenXml.Presentation.BlipFill blipFill)
        {
            ImagePart part = null;
            IEnumerable<DocumentFormat.OpenXml.Presentation.Picture> lstPic = null;

            if (openXmlPart is SlidePart slidePart)
            {
                lstPic = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    string rID = blipFill.Blip.Embed.Value;
                    part = (ImagePart)slidePart.GetPartById(rID);
                }
            }
            else if (openXmlPart is SlideMasterPart slideMasterPart)
            {
                lstPic = slideMasterPart.SlideMaster.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    string rID = blipFill.Blip.Embed.Value;
                    part = (ImagePart)slideMasterPart.GetPartById(rID);
                }
            }
            else if (openXmlPart is SlideLayoutPart slideLayoutPart)
            {
                lstPic = slideLayoutPart.SlideLayout.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    string rID = blipFill.Blip.Embed.Value;
                    part = (ImagePart)slideLayoutPart.GetPartById(rID);
                }
            }
            if (lstPic != null && part != null)
            {
                string pathImage = System.IO.Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + System.IO.Path.GetExtension(part.Uri.ToString()));
                Utils.CopyStream(part.GetStream(), pathImage);
                ImageColor imageColor = new ImageColor();
                imageColor.Source = new Core.Model.Image(pathImage);
                imageColor.Opacity = 1 - fill.Transparency;
                imageColor.OffsetX = fill.TextureOffsetX;
                imageColor.OffsetY = fill.TextureOffsetY;
                imageColor.IsTiled = fill.TextureTile == MsoTriState.msoTrue;
                imageColor.Alignment = ConvertImageAligment(fill.TextureAlignment);
                imageColor.ScaleX = fill.TextureHorizontalScale;
                imageColor.ScaleY = fill.TextureVerticalScale;
                imageColor.RotateWidthShape = fill.RotateWithObject == MsoTriState.msoTrue;
                foreach (var item in blipFill)
                {
                    if (item is Stretch stretch)
                    {
                        if (stretch.FillRectangle.Top != null)
                            imageColor.OffsetTop = stretch.FillRectangle.Top / 100000;
                        if (stretch.FillRectangle.Left != null)
                            imageColor.OffsetLeft = stretch.FillRectangle.Left / 100000;
                        if (stretch.FillRectangle.Right != null)
                            imageColor.OffsetRight = stretch.FillRectangle.Right / 100000;
                        if (stretch.FillRectangle.Bottom != null)
                            imageColor.OffsetBottom = stretch.FillRectangle.Bottom / 100000;
                    }
                }
                return imageColor;
            }
            return null;
        }

        /// <summary>
        /// Lấy màu Image Slide
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="openXmlPart"></param>
        /// <param name="blipFill"></param>
        /// <returns></returns>
        private ImageColor GetImageColor(pp.FillFormat fill, OpenXmlPart openXmlPart, BlipFill blipFill)
        {
            ImagePart part = null;
            IEnumerable<DocumentFormat.OpenXml.Presentation.Picture> lstPic = null;

            if (openXmlPart is SlidePart slidePart)
            {
                lstPic = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    string rID = blipFill.Blip.Embed.Value;
                    part = (ImagePart)slidePart.GetPartById(rID);
                }
            }
            else if (openXmlPart is SlideMasterPart slideMasterPart)
            {
                lstPic = slideMasterPart.SlideMaster.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    string rID = blipFill.Blip.Embed.Value;
                    part = (ImagePart)slideMasterPart.GetPartById(rID);
                }
            }
            else if (openXmlPart is SlideLayoutPart slideLayoutPart)
            {
                lstPic = slideLayoutPart.SlideLayout.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Select(p => p);
                if (lstPic != null)
                {
                    string rID = blipFill.Blip.Embed.Value;
                    part = (ImagePart)slideLayoutPart.GetPartById(rID);
                }
            }
            if (lstPic != null && part != null)
            {
                string pathImage = System.IO.Path.Combine((System.Windows.Application.Current as IAppGlobal).DocumentTempFolder, ObjectElementsHelper.RandomString(13) + System.IO.Path.GetExtension(part.Uri.ToString()));
                Utils.CopyStream(part.GetStream(), pathImage);
                ImageColor imageColor = new ImageColor();
                imageColor.Source = new Core.Model.Image(pathImage);
                imageColor.Opacity = 1 - fill.Transparency;
                imageColor.OffsetX = fill.TextureOffsetX;
                imageColor.OffsetY = fill.TextureOffsetY;
                imageColor.IsTiled = fill.TextureTile == MsoTriState.msoTrue;
                imageColor.Alignment = ConvertImageAligment(fill.TextureAlignment);
                imageColor.ScaleX = fill.TextureHorizontalScale;
                imageColor.ScaleY = fill.TextureVerticalScale;
                imageColor.RotateWidthShape = fill.RotateWithObject == MsoTriState.msoTrue;
                foreach (var item in blipFill)
                {
                    if (item is Stretch stretch)
                    {
                        if (stretch.FillRectangle.Top != null)
                            imageColor.OffsetTop = stretch.FillRectangle.Top / 100000;
                        if (stretch.FillRectangle.Left != null)
                            imageColor.OffsetLeft = stretch.FillRectangle.Left / 100000;
                        if (stretch.FillRectangle.Right != null)
                            imageColor.OffsetRight = stretch.FillRectangle.Right / 100000;
                        if (stretch.FillRectangle.Bottom != null)
                            imageColor.OffsetBottom = stretch.FillRectangle.Bottom / 100000;
                    }
                }
                return imageColor;
            }
            return null;
        }

        /// <summary>
        /// hàm  chuyển đổi ImageAlignment
        /// </summary>
        /// <param name="msoTextureAlignment"></param>
        /// <returns></returns>
        private ImageAlignement ConvertImageAligment(MsoTextureAlignment msoTextureAlignment)
        {
            switch (msoTextureAlignment)
            {
                case MsoTextureAlignment.msoTextureAlignmentMixed:
                    break;
                case MsoTextureAlignment.msoTextureTopLeft:
                    return ImageAlignement.TopLeft;
                case MsoTextureAlignment.msoTextureTop:
                    return ImageAlignement.Top;
                case MsoTextureAlignment.msoTextureTopRight:
                    return ImageAlignement.TopRight;
                case MsoTextureAlignment.msoTextureLeft:
                    return ImageAlignement.Left;
                case MsoTextureAlignment.msoTextureCenter:
                    return ImageAlignement.Center;
                case MsoTextureAlignment.msoTextureRight:
                    return ImageAlignement.Right;
                case MsoTextureAlignment.msoTextureBottomLeft:
                    return ImageAlignement.BottomLeft;
                case MsoTextureAlignment.msoTextureBottom:
                    return ImageAlignement.Bottom;
                case MsoTextureAlignment.msoTextureBottomRight:
                    return ImageAlignement.BottomRight;
                default:
                    break;
            }
            return ImageAlignement.Left;
        }

        /// <summary>
        /// Lấy Solid Color
        /// </summary>
        /// <param name="colorFormat"></param>
        /// <returns></returns>
        private SolidColor GetSolidColor(pp.ColorFormat colorFormat, double transparent = 0.0)
        {
            SolidColor solidColor = new SolidColor() { Color = ConvertColor(colorFormat.RGB, transparent), Name = GetSpecialName(colorFormat.ObjectThemeColor), SpecialName = GetSpecialName(colorFormat.ObjectThemeColor) };
            return solidColor;
        }

        /// <summary>
        /// Lấy màu Linear
        /// </summary>
        /// <param name="gradientStops"></param>
        /// <param name="msoGradientStyle"></param>
        /// <param name="gradientAngle"></param>
        /// <returns></returns>
        private GradientColor GetLinearColor(GradientStops gradientStops, MsoGradientStyle msoGradientStyle, float gradientAngle)
        {
            GradientColor gradientColor = new GradientColor();
            gradientColor.Angle = (int)gradientAngle;
            gradientColor.Type = GradientColorType.Linear;
            
            for (int i = 1; i <= gradientStops.Count; i++)
            {
                gradientColor.LinearElements.Add(new LinearElement() { Color = ConvertColor(gradientStops[i].Color.RGB), SpecialName = GetSpecialName(gradientStops[i].Color.ObjectThemeColor), Brightness = gradientStops[i].Color.Brightness, Offset = gradientStops[i].Position });
            }
            gradientColor.LinearElements = gradientColor.LinearElements.OrderBy(s => s.Offset).ToList();
            return gradientColor;
        }

        /// <summary>
        /// Lấy màu radiant
        /// </summary>
        /// <param name="gradientStops"></param>
        /// <param name="msoGradientStyle"></param>
        /// <returns></returns>
        private GradientColor GetRadialColor(GradientStops gradientStops, MsoGradientStyle msoGradientStyle)
        {
            GradientColor gradientColor = new GradientColor();
            gradientColor.Type = GradientColorType.Radial;
            for (int i = 1; i <= gradientStops.Count; i++)
            {
                gradientColor.LinearElements.Add(new LinearElement() { Color = ConvertColor(gradientStops[i].Color.RGB), SpecialName = GetSpecialName(gradientStops[i].Color.ObjectThemeColor), Brightness = gradientStops[i].Color.Brightness, Offset = gradientStops[i].Position });
            }
            gradientColor.LinearElements = gradientColor.LinearElements.OrderBy(s => s.Offset).ToList();
            return gradientColor;
        }

        /// <summary>
        /// Lấy màu pattern
        /// </summary>
        /// <param name="msoPatternType"></param>
        /// <param name="background"></param>
        /// <param name="foreground"></param>
        /// <returns></returns>
        private PatternColor GetPatternColor(MsoPatternType msoPatternType, pp.ColorFormat background, pp.ColorFormat foreground)
        {
            PatternColor patternColor = new PatternColor();
            patternColor.Background = GetSolidColor(background);
            patternColor.Foreground = GetSolidColor(foreground);
            switch (msoPatternType)
            {
                case MsoPatternType.msoPattern5Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT05;
                    break;
                case MsoPatternType.msoPattern10Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT10;
                    break;
                case MsoPatternType.msoPattern20Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT20;
                    break;
                case MsoPatternType.msoPattern25Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT25;
                    break;
                case MsoPatternType.msoPattern30Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT30;
                    break;
                case MsoPatternType.msoPattern40Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT40;
                    break;
                case MsoPatternType.msoPattern50Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT50;
                    break;
                case MsoPatternType.msoPattern60Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT60;
                    break;
                case MsoPatternType.msoPattern70Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT70;
                    break;
                case MsoPatternType.msoPattern75Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT75;
                    break;
                case MsoPatternType.msoPattern80Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT80;
                    break;
                case MsoPatternType.msoPattern90Percent:
                    patternColor.HatchStyle = HatchStyle.HS_PERCENT90;
                    break;
                case MsoPatternType.msoPatternDarkHorizontal:
                    patternColor.HatchStyle = HatchStyle.HS_DARKHORIZONTAL;
                    break;
                case MsoPatternType.msoPatternDarkVertical:
                    patternColor.HatchStyle = HatchStyle.HS_DARKVERTICAL;
                    break;
                case MsoPatternType.msoPatternDarkDownwardDiagonal:
                    patternColor.HatchStyle = HatchStyle.HS_DARKDOWNWARDDIAGONAL;
                    break;
                case MsoPatternType.msoPatternDarkUpwardDiagonal:
                    patternColor.HatchStyle = HatchStyle.HS_DARKUPWARDDIAGONAL;
                    break;
                case MsoPatternType.msoPatternSmallCheckerBoard:
                    patternColor.HatchStyle = HatchStyle.HS_SMALLCHECKERBOARD;
                    break;
                case MsoPatternType.msoPatternTrellis:
                    patternColor.HatchStyle = HatchStyle.HS_TRELLIS;
                    break;
                case MsoPatternType.msoPatternLightHorizontal:
                    patternColor.HatchStyle = HatchStyle.HS_LIGHTHORIZONTAL;
                    break;
                case MsoPatternType.msoPatternLightVertical:
                    patternColor.HatchStyle = HatchStyle.HS_LIGHTVERTICAL;
                    break;
                case MsoPatternType.msoPatternLightDownwardDiagonal:
                    patternColor.HatchStyle = HatchStyle.HS_LIGHTDOWNWARDDIAGONAL;
                    break;
                case MsoPatternType.msoPatternLightUpwardDiagonal:
                    patternColor.HatchStyle = HatchStyle.HS_LIGHTUPWARDDIAGONAL;
                    break;
                case MsoPatternType.msoPatternSmallGrid:
                    patternColor.HatchStyle = HatchStyle.HS_SMALLGRID;
                    break;
                case MsoPatternType.msoPatternDottedDiamond:
                    patternColor.HatchStyle = HatchStyle.HS_DOTTEDDIAMOND;
                    break;
                case MsoPatternType.msoPatternWideDownwardDiagonal:
                    patternColor.HatchStyle = HatchStyle.HS_WIDEDOWNWARDDIAGONAL;
                    break;
                case MsoPatternType.msoPatternWideUpwardDiagonal:
                    patternColor.HatchStyle = HatchStyle.HS_WIDEUPWARDDIAGONAL;
                    break;
                case MsoPatternType.msoPatternDashedUpwardDiagonal:
                    patternColor.HatchStyle = HatchStyle.HS_DASHEDUPWARDDIAGONAL;
                    break;
                case MsoPatternType.msoPatternDashedDownwardDiagonal:
                    patternColor.HatchStyle = HatchStyle.HS_DASHEDDOWNWARDDIAGONAL;
                    break;
                case MsoPatternType.msoPatternNarrowVertical:
                    patternColor.HatchStyle = HatchStyle.HS_NARROWVERTICAL;
                    break;
                case MsoPatternType.msoPatternNarrowHorizontal:
                    patternColor.HatchStyle = HatchStyle.HS_NARROWHORIZONTAL;
                    break;
                case MsoPatternType.msoPatternDashedVertical:
                    patternColor.HatchStyle = HatchStyle.HS_DASHEDVERTICAL;
                    break;
                case MsoPatternType.msoPatternDashedHorizontal:
                    patternColor.HatchStyle = HatchStyle.HS_DASHEDHORIZONTAL;
                    break;
                case MsoPatternType.msoPatternLargeConfetti:
                    patternColor.HatchStyle = HatchStyle.HS_LARGECONFETTI;
                    break;
                case MsoPatternType.msoPatternLargeGrid:
                    patternColor.HatchStyle = HatchStyle.HS_LARGEGRID;
                    break;
                case MsoPatternType.msoPatternHorizontalBrick:
                    patternColor.HatchStyle = HatchStyle.HS_HORIZONTALBRICK;
                    break;
                case MsoPatternType.msoPatternLargeCheckerBoard:
                    patternColor.HatchStyle = HatchStyle.HS_LARGECHECKERBOARD;
                    break;
                case MsoPatternType.msoPatternSmallConfetti:
                    patternColor.HatchStyle = HatchStyle.HS_SMALLCONFETTI;
                    break;
                case MsoPatternType.msoPatternZigZag:
                    patternColor.HatchStyle = HatchStyle.HS_ZIGZAG;
                    break;
                case MsoPatternType.msoPatternSolidDiamond:
                    patternColor.HatchStyle = HatchStyle.HS_SOLIDDIAMOND;
                    break;
                case MsoPatternType.msoPatternDiagonalBrick:
                    patternColor.HatchStyle = HatchStyle.HS_DIAGONALBRICK;
                    break;
                case MsoPatternType.msoPatternOutlinedDiamond:
                    patternColor.HatchStyle = HatchStyle.HS_OUTLINEDDIAMOND;
                    break;
                case MsoPatternType.msoPatternPlaid:
                    patternColor.HatchStyle = HatchStyle.HS_PLAID;
                    break;
                case MsoPatternType.msoPatternSphere:
                    patternColor.HatchStyle = HatchStyle.HS_SPHERE;
                    break;
                case MsoPatternType.msoPatternWeave:
                    patternColor.HatchStyle = HatchStyle.HS_WEAVE;
                    break;
                case MsoPatternType.msoPatternDottedGrid:
                    patternColor.HatchStyle = HatchStyle.HS_DOTTEDGRID;
                    break;
                case MsoPatternType.msoPatternDivot:
                    patternColor.HatchStyle = HatchStyle.HS_DIVOT;
                    break;
                case MsoPatternType.msoPatternShingle:
                    patternColor.HatchStyle = HatchStyle.HS_SHINGLE;
                    break;
                case MsoPatternType.msoPatternWave:
                    patternColor.HatchStyle = HatchStyle.HS_WAVE;
                    break;
                case MsoPatternType.msoPatternHorizontal:
                    patternColor.HatchStyle = HatchStyle.HS_HORIZONTAL;
                    break;
                case MsoPatternType.msoPatternVertical:
                    patternColor.HatchStyle = HatchStyle.HS_VERTICAL;
                    break;
                case MsoPatternType.msoPatternCross:
                    patternColor.HatchStyle = HatchStyle.HS_DIAGONALCROSS;
                    break;
                case MsoPatternType.msoPatternDownwardDiagonal:
                    patternColor.HatchStyle = HatchStyle.HS_DASHEDDOWNWARDDIAGONAL;
                    break;
                case MsoPatternType.msoPatternUpwardDiagonal:
                    patternColor.HatchStyle = HatchStyle.HS_DASHEDUPWARDDIAGONAL;
                    break;
                case MsoPatternType.msoPatternDiagonalCross:
                    patternColor.HatchStyle = HatchStyle.HS_DIAGONALCROSS;
                    break;
                default:
                    break;
            }

            return patternColor;
        }


        /// <summary>
        /// Lấy Special Name của color
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        private string GetSpecialName(MsoThemeColorIndex index)
        {
            switch (index)
            {
                case MsoThemeColorIndex.msoThemeColorDark1:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounDark1").ToString();
                case MsoThemeColorIndex.msoThemeColorLight1:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounDark2").ToString();
                case MsoThemeColorIndex.msoThemeColorDark2:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounLight1").ToString();
                case MsoThemeColorIndex.msoThemeColorLight2:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_BackgrounLight2").ToString();
                case MsoThemeColorIndex.msoThemeColorAccent1:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_Accent1").ToString();
                case MsoThemeColorIndex.msoThemeColorAccent2:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_Accent2").ToString();
                case MsoThemeColorIndex.msoThemeColorAccent3:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_Accent3").ToString();
                case MsoThemeColorIndex.msoThemeColorAccent4:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_Accent4").ToString();
                case MsoThemeColorIndex.msoThemeColorAccent5:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_Accent5").ToString();
                case MsoThemeColorIndex.msoThemeColorAccent6:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_Accent6").ToString();
                case MsoThemeColorIndex.msoThemeColorHyperlink:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_Hyperlink").ToString();

                case MsoThemeColorIndex.msoThemeColorFollowedHyperlink:
                    return System.Windows.Application.Current.TryFindResource("ColorGallery_FollowedHyperlink").ToString();
                default:
                    break;
            }
            return string.Empty;
        }

        /// <summary>
        /// RGB to hexa
        /// </summary>
        /// <param name="_oleColor"></param>
        /// <returns></returns>
        private string ConvertColor(int _oleColor, double transparent = 0.0)
        {
            System.Drawing.Color stdColor = System.Drawing.ColorTranslator.FromOle(_oleColor);
            string hexa = "#" + ((int)(255 * (1 - transparent))).ToString("X2") + stdColor.R.ToString("X2") + stdColor.G.ToString("X2") + stdColor.B.ToString("X2");
            return hexa;
        }



    }
}
