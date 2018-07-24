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
using INV.Elearing.Controls.Shapes.Lines.Enums;
using DocumentFormat.OpenXml.Drawing;

namespace INV.Elearning.ImportPowerPoint.Helper
{
    public partial class HelperClass
    {
        /// <summary>
        /// Lấy model ShapeInfo
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        public ShapeInfo GetShapePresent(pp.Shape shape, OpenXmlCompositeElement shapeTag)
        {
            if (shape.Type == MsoShapeType.msoFreeform)
            {
                StaticShapeInfo staticShape = new StaticShapeInfo();
                staticShape.PathData = GetPathDataCustomGeometry(shape, shapeTag);
                return staticShape;
            }
            switch (shape.AutoShapeType)
            {
                case office.MsoAutoShapeType.msoShapeMixed:
                    if (shape.Connector == MsoTriState.msoTrue)
                    {
                        LineInfo lineInfo = new LineInfo();
                        lineInfo.HeadEnd = ArrowTypesConverter(shape.Line.BeginArrowheadStyle);
                        lineInfo.TailEnd = ArrowTypesConverter(shape.Line.EndArrowheadStyle);
                        lineInfo.StartPoint = new System.Windows.Point(0, 0);
                        lineInfo.EndPoint = new System.Windows.Point(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                        return lineInfo;
                    }
                    break;
                case office.MsoAutoShapeType.msoShapeRectangle:
                    RectangleInfo rectangleInfo = new RectangleInfo();
                    return rectangleInfo;
                case office.MsoAutoShapeType.msoShapeParallelogram:
                    ParallelogramInfo parallelogramInfo = new ParallelogramInfo();
                    parallelogramInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    parallelogramInfo.FactorBodySize = parallelogramInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return parallelogramInfo;
                case office.MsoAutoShapeType.msoShapeTrapezoid:
                    TrapezoidInfo trapezoidInfo = new TrapezoidInfo();
                    trapezoidInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    trapezoidInfo.FactorBodySize = trapezoidInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return trapezoidInfo;
                case office.MsoAutoShapeType.msoShapeDiamond:
                    DiamondInfo diamondInfo = new DiamondInfo();
                    return diamondInfo;
                case office.MsoAutoShapeType.msoShapeRoundedRectangle:
                    RoundedRectangleInfo rounded = new RoundedRectangleInfo();
                    rounded.CornerSize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    rounded.Factor = rounded.CornerSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return rounded;
                case office.MsoAutoShapeType.msoShapeOctagon:
                    OctagonInfo octagonInfo = new OctagonInfo();
                    octagonInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    octagonInfo.FactorBodySize = octagonInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return octagonInfo;
                case office.MsoAutoShapeType.msoShapeIsoscelesTriangle:
                    TriangleInfo triangleInfo = new TriangleInfo();
                    triangleInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    triangleInfo.FactorBodySize = triangleInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return triangleInfo;
                case office.MsoAutoShapeType.msoShapeRightTriangle:
                    RightTriangleInfo rightTriangle = new RightTriangleInfo();
                    return rightTriangle;
                case office.MsoAutoShapeType.msoShapeOval:
                    EllipseInfo ellipseInfo = new EllipseInfo();
                    return ellipseInfo;
                case office.MsoAutoShapeType.msoShapeHexagon:
                    HexagonInfo hexagonInfo = new HexagonInfo();
                    hexagonInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    hexagonInfo.FactorBodySize = hexagonInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return hexagonInfo;
                case office.MsoAutoShapeType.msoShapeCross:
                    CrossInfo crossInfo = new CrossInfo();
                    crossInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    crossInfo.FactorBodySize = crossInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return crossInfo;
                case office.MsoAutoShapeType.msoShapeRegularPentagon:
                    RegularPentagonInfo regularPentagonInfo = new RegularPentagonInfo(5);
                    return regularPentagonInfo;
                case office.MsoAutoShapeType.msoShapeCan:
                    CanInfo canInfo = new CanInfo();
                    canInfo.CoverHeight = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    canInfo.FactorCoverHeight = canInfo.CoverHeight / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return canInfo;
                case office.MsoAutoShapeType.msoShapeCube:
                    CubeInfo cubeInfo = new CubeInfo();
                    cubeInfo.FaceHeight = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    cubeInfo.FactorFaceHeight = cubeInfo.FaceHeight / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return cubeInfo;
                case office.MsoAutoShapeType.msoShapeBevel:
                    BevelInfo bevelInfo = new BevelInfo();
                    bevelInfo.FaceHeight = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    bevelInfo.FactorFaceHeight = bevelInfo.FaceHeight / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return bevelInfo;
                case office.MsoAutoShapeType.msoShapeFoldedCorner:
                    FoldedCornerInfo foldedCornerInfo = new FoldedCornerInfo();
                    foldedCornerInfo.CornerSize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    foldedCornerInfo.Factor = foldedCornerInfo.CornerSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return foldedCornerInfo;
                case office.MsoAutoShapeType.msoShapeSmileyFace:
                    SmileFaceInfo smileFaceInfo = new SmileFaceInfo();
                    smileFaceInfo.MouthHeight = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    smileFaceInfo.Emotion = smileFaceInfo.MouthHeight > 0 ? Elearing.Controls.Shapes.Basics.Emotions.Happy : Elearing.Controls.Shapes.Basics.Emotions.UnHappy;
                    smileFaceInfo.MouthHeight = Math.Abs(smileFaceInfo.MouthHeight);
                    smileFaceInfo.Factor = smileFaceInfo.MouthHeight / (shape.Height / Utils.tlh) * 2;
                    return smileFaceInfo;
                case office.MsoAutoShapeType.msoShapeDonut:
                    DonutInfo donutInfo = new DonutInfo();
                    donutInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    donutInfo.FactorBodySize = donutInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return donutInfo;
                case office.MsoAutoShapeType.msoShapeNoSymbol:
                    NoSymbolInfo noSymbolInfo = new NoSymbolInfo();
                    noSymbolInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    noSymbolInfo.FactorBodySize = noSymbolInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return noSymbolInfo;
                case office.MsoAutoShapeType.msoShapeBlockArc:
                    BlockArcInfo blockArcInfo = new BlockArcInfo();
                    blockArcInfo.EndAngle = shape.Adjustments[1] > 0 ? 360 - shape.Adjustments[1] : -shape.Adjustments[1];
                    blockArcInfo.StartAngle = shape.Adjustments[2] > 0 ? 360 - shape.Adjustments[2] : -shape.Adjustments[2];
                    blockArcInfo.BodySize = shape.Adjustments[3] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    blockArcInfo.FactorBodySize = blockArcInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return blockArcInfo;
                case office.MsoAutoShapeType.msoShapeHeart:
                    //for (int i = 1; i <= shape.Nodes.Count; i++)
                    //{
                    //    pp.ShapeNode shapeNode = shape.Nodes[i];

                    //}
                    break;
                case office.MsoAutoShapeType.msoShapeLightningBolt:
                    LightnightBoltInfo lightnightBoltInfo = new LightnightBoltInfo();
                    return lightnightBoltInfo;
                case office.MsoAutoShapeType.msoShapeSun:

                    break;
                case office.MsoAutoShapeType.msoShapeMoon:
                    MoonInfo moonInfo = new MoonInfo();
                    moonInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    moonInfo.FactorBodySize = moonInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return moonInfo;
                case office.MsoAutoShapeType.msoShapeArc:
                    ArcInfo arcInfo = new ArcInfo();
                    arcInfo.EndAngle = shape.Adjustments[1] > 0 ? 360 - shape.Adjustments[1] : -shape.Adjustments[1];
                    arcInfo.StartAngle = shape.Adjustments[2] > 0 ? 360 - shape.Adjustments[2] : -shape.Adjustments[2];
                    return arcInfo;
                case office.MsoAutoShapeType.msoShapeDoubleBracket:

                    break;
                case office.MsoAutoShapeType.msoShapeDoubleBrace:

                    break;
                case office.MsoAutoShapeType.msoShapePlaque:
                    PlaqueInfo plaqueInfo = new PlaqueInfo();
                    plaqueInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    plaqueInfo.FactorBodySize = plaqueInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return plaqueInfo;
                case office.MsoAutoShapeType.msoShapeLeftBracket:

                    break;
                case office.MsoAutoShapeType.msoShapeRightBracket:
                    break;
                case office.MsoAutoShapeType.msoShapeLeftBrace:
                    break;
                case office.MsoAutoShapeType.msoShapeRightBrace:
                    break;
                case office.MsoAutoShapeType.msoShapeRightArrow:
                    RightArrowInfo rightArrowInfo = new RightArrowInfo();
                    rightArrowInfo.ArrowSize = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    rightArrowInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    rightArrowInfo.FactorArrowSize = rightArrowInfo.ArrowSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    rightArrowInfo.FactorBodySize = rightArrowInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return rightArrowInfo;
                case office.MsoAutoShapeType.msoShapeLeftArrow:
                    LeftArrowInfo leftArrowInfo = new LeftArrowInfo();
                    leftArrowInfo.ArrowSize = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    leftArrowInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    leftArrowInfo.FactorArrowSize = leftArrowInfo.ArrowSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    leftArrowInfo.FactorBodySize = leftArrowInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return leftArrowInfo;
                case office.MsoAutoShapeType.msoShapeUpArrow:
                    UpArrowInfo upArrowInfo = new UpArrowInfo();
                    upArrowInfo.ArrowSize = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    upArrowInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    upArrowInfo.FactorArrowSize = upArrowInfo.ArrowSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    upArrowInfo.FactorBodySize = upArrowInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return upArrowInfo;
                case office.MsoAutoShapeType.msoShapeDownArrow:
                    DownArrowInfo downArrowInfo = new DownArrowInfo();
                    downArrowInfo.ArrowSize = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    downArrowInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    downArrowInfo.FactorArrowSize = downArrowInfo.ArrowSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    downArrowInfo.FactorBodySize = downArrowInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return downArrowInfo;
                case office.MsoAutoShapeType.msoShapeLeftRightArrow:
                    LeftRightArrowInfo leftRightArrowInfo = new LeftRightArrowInfo();
                    leftRightArrowInfo.ArrowSize = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    leftRightArrowInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    leftRightArrowInfo.FactorArrowSize = leftRightArrowInfo.ArrowSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    leftRightArrowInfo.FactorBodySize = leftRightArrowInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return leftRightArrowInfo;
                case office.MsoAutoShapeType.msoShapeUpDownArrow:
                    UpDownArrowInfo upDownArrowInfo = new UpDownArrowInfo();
                    upDownArrowInfo.ArrowSize = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    upDownArrowInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    upDownArrowInfo.FactorArrowSize = upDownArrowInfo.ArrowSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    upDownArrowInfo.FactorBodySize = upDownArrowInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return upDownArrowInfo;
                case office.MsoAutoShapeType.msoShapeQuadArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeLeftRightUpArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeBentArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeUTurnArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeLeftUpArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeBentUpArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeCurvedRightArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeCurvedLeftArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeCurvedUpArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeCurvedDownArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeStripedRightArrow:
                    StripedRightArrowInfo stripedRightArrowInfo = new StripedRightArrowInfo();
                    stripedRightArrowInfo.ArrowSize = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    stripedRightArrowInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    stripedRightArrowInfo.FactorArrowSize = stripedRightArrowInfo.ArrowSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    stripedRightArrowInfo.FactorBodySize = stripedRightArrowInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return stripedRightArrowInfo;
                case office.MsoAutoShapeType.msoShapeNotchedRightArrow:
                    NotChedRightArrowInfo notChedRightArrowInfo = new NotChedRightArrowInfo();
                    notChedRightArrowInfo.ArrowSize = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    notChedRightArrowInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    notChedRightArrowInfo.FactorArrowSize = notChedRightArrowInfo.ArrowSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    notChedRightArrowInfo.FactorBodySize = notChedRightArrowInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return notChedRightArrowInfo;
                case office.MsoAutoShapeType.msoShapePentagon:
                    PentagonInfo pentagonInfo = new PentagonInfo();
                    pentagonInfo.ArrowSize = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    pentagonInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    pentagonInfo.FactorArrowSize = pentagonInfo.ArrowSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    pentagonInfo.FactorBodySize = pentagonInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return pentagonInfo;
                case office.MsoAutoShapeType.msoShapeChevron:
                    ChevronInfo chevronInfo = new ChevronInfo();
                    chevronInfo.ArrowSize = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    chevronInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    chevronInfo.FactorArrowSize = chevronInfo.ArrowSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    chevronInfo.FactorBodySize = chevronInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return chevronInfo;
                case office.MsoAutoShapeType.msoShapeRightArrowCallout:
                    break;
                case office.MsoAutoShapeType.msoShapeLeftArrowCallout:
                    break;
                case office.MsoAutoShapeType.msoShapeUpArrowCallout:
                    break;
                case office.MsoAutoShapeType.msoShapeDownArrowCallout:
                    break;
                case office.MsoAutoShapeType.msoShapeLeftRightArrowCallout:
                    break;
                case office.MsoAutoShapeType.msoShapeUpDownArrowCallout:
                    break;
                case office.MsoAutoShapeType.msoShapeQuadArrowCallout:
                    break;
                case office.MsoAutoShapeType.msoShapeCircularArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartProcess:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartAlternateProcess:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartDecision:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartData:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartPredefinedProcess:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartInternalStorage:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartDocument:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartMultidocument:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartTerminator:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartPreparation:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartManualInput:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartManualOperation:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartConnector:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartOffpageConnector:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartCard:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartPunchedTape:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartSummingJunction:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartOr:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartCollate:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartSort:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartExtract:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartMerge:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartStoredData:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartDelay:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartSequentialAccessStorage:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartMagneticDisk:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartDirectAccessStorage:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartDisplay:
                    break;
                case office.MsoAutoShapeType.msoShapeExplosion1:
                    ExplosionInfo1 explosionInfo1 = new ExplosionInfo1();
                    return explosionInfo1;
                case office.MsoAutoShapeType.msoShapeExplosion2:
                    ExplosionInfo2 explosionInfo2 = new ExplosionInfo2();
                    return explosionInfo2;
                case office.MsoAutoShapeType.msoShape4pointStar:
                    Star4Info star4Info = new Star4Info();
                    star4Info.CountPoint = 4;
                    star4Info.Radius = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    star4Info.Factor = star4Info.Radius / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return star4Info;
                case office.MsoAutoShapeType.msoShape5pointStar:
                    Star5Info star5Info = new Star5Info();
                    star5Info.CountPoint = 5;
                    star5Info.Radius = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    star5Info.Factor = star5Info.Radius / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return star5Info;
                case office.MsoAutoShapeType.msoShape8pointStar:
                    Star8Info star8Info = new Star8Info();
                    star8Info.CountPoint = 8;
                    star8Info.Radius = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    star8Info.Factor = star8Info.Radius / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return star8Info;
                case office.MsoAutoShapeType.msoShape16pointStar:
                    Star16Info star16Info = new Star16Info();
                    star16Info.CountPoint = 16;
                    star16Info.Radius = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    star16Info.Factor = star16Info.Radius / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return star16Info;
                case office.MsoAutoShapeType.msoShape24pointStar:
                    Star24Info star24Info = new Star24Info();
                    star24Info.CountPoint = 24;
                    star24Info.Radius = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    star24Info.Factor = star24Info.Radius / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return star24Info;
                case office.MsoAutoShapeType.msoShape32pointStar:
                    Star32Info star32Info = new Star32Info();
                    star32Info.CountPoint = 32;
                    star32Info.Radius = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    star32Info.Factor = star32Info.Radius / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return star32Info;
                case office.MsoAutoShapeType.msoShapeUpRibbon:
                    break;
                case office.MsoAutoShapeType.msoShapeDownRibbon:
                    break;
                case office.MsoAutoShapeType.msoShapeCurvedUpRibbon:
                    break;
                case office.MsoAutoShapeType.msoShapeCurvedDownRibbon:
                    break;
                case office.MsoAutoShapeType.msoShapeVerticalScroll:
                    break;
                case office.MsoAutoShapeType.msoShapeHorizontalScroll:
                    break;
                case office.MsoAutoShapeType.msoShapeWave:
                    break;
                case office.MsoAutoShapeType.msoShapeDoubleWave:
                    break;
                case office.MsoAutoShapeType.msoShapeRectangularCallout:
                    RectangularCalloutInfo rectangularCalloutInfo = new RectangularCalloutInfo();
                    rectangularCalloutInfo.ControlPoint = new Core.Model.EPoint(shape.Adjustments[1] * shape.Width * Utils.tlw, shape.Adjustments[2] * shape.Height * Utils.tlh);
                    return rectangularCalloutInfo;
                case office.MsoAutoShapeType.msoShapeRoundedRectangularCallout:
                    RoundedRectangularCalloutInfo roundedRectangularCalloutInfo = new RoundedRectangularCalloutInfo();
                    roundedRectangularCalloutInfo.ControlPoint = new Core.Model.EPoint(shape.Adjustments[1] * shape.Width * Utils.tlw, shape.Adjustments[2] * shape.Height * Utils.tlh);
                    return roundedRectangularCalloutInfo;
                case office.MsoAutoShapeType.msoShapeOvalCallout:
                    OvalCalloutInfo ovalCalloutInfo = new OvalCalloutInfo();
                    ovalCalloutInfo.ControlPoint = new Core.Model.EPoint(shape.Adjustments[1] * shape.Width * Utils.tlw, shape.Adjustments[2] * shape.Height * Utils.tlh);
                    return ovalCalloutInfo;
                case office.MsoAutoShapeType.msoShapeCloudCallout:
                    break;
                case office.MsoAutoShapeType.msoShapeLineCallout1:
                    LineCalloutShape1NoAccentBarNoBorderInfo lineCalloutInfo = new LineCalloutShape1NoAccentBarNoBorderInfo(false, false, new Core.Model.EPoint(), new Core.Model.EPoint(), new Core.Model.EPoint());
                    lineCalloutInfo.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutInfo.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutInfo.StartPointTail = lineCalloutInfo.Point1;
                    return lineCalloutInfo;
                case office.MsoAutoShapeType.msoShapeLineCallout2:
                    //LineCalloutShape2NoAccentBarHasBorderInfo lineCalloutShape2 = new LineCalloutShape2NoAccentBarHasBorderInfo();
                    //lineCalloutShape2.IsHasAccentBar = false;
                    //lineCalloutShape2.IsHasStroke = true;
                    //lineCalloutShape2.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    //lineCalloutShape2.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    //lineCalloutShape2.Point3 = new Core.Model.EPoint(shape.Adjustments[6] * shape.Width * Utils.tlw, shape.Adjustments[5] * shape.Height * Utils.tlh);
                    //lineCalloutShape2.StartPointTail = lineCalloutShape2.Point1;
                    //return lineCalloutShape2;
                    break;
                case office.MsoAutoShapeType.msoShapeLineCallout3:
                    LineCalloutShape2NoAccentBarNoBorderInfo lineCalloutShape2 = new LineCalloutShape2NoAccentBarNoBorderInfo();
                    lineCalloutShape2.IsHasAccentBar = false;
                    lineCalloutShape2.IsHasStroke = false;
                    lineCalloutShape2.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutShape2.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutShape2.Point3 = new Core.Model.EPoint(shape.Adjustments[6] * shape.Width * Utils.tlw, shape.Adjustments[5] * shape.Height * Utils.tlh);
                    lineCalloutShape2.StartPointTail = lineCalloutShape2.Point1;
                    return lineCalloutShape2;
                case office.MsoAutoShapeType.msoShapeLineCallout4:
                    LineCalloutShape3NoAccentBarNoBorderInfo lineCalloutShape3 = new LineCalloutShape3NoAccentBarNoBorderInfo();
                    lineCalloutShape3.IsHasAccentBar = false;
                    lineCalloutShape3.IsHasStroke = false;
                    lineCalloutShape3.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutShape3.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutShape3.Point3 = new Core.Model.EPoint(shape.Adjustments[6] * shape.Width * Utils.tlw, shape.Adjustments[5] * shape.Height * Utils.tlh);
                    lineCalloutShape3.Point4 = new Core.Model.EPoint(shape.Adjustments[8] * shape.Width * Utils.tlw, shape.Adjustments[7] * shape.Height * Utils.tlh);
                    lineCalloutShape3.StartPointTail = lineCalloutShape3.Point1;
                    return lineCalloutShape3;
                case office.MsoAutoShapeType.msoShapeLineCallout1AccentBar:
                    break;
                case office.MsoAutoShapeType.msoShapeLineCallout2AccentBar:
                    LineCalloutShape1HasAccentBarNoBorderInfo lineCalloutInfoAccentBar = new LineCalloutShape1HasAccentBarNoBorderInfo(true, false, new Core.Model.EPoint(), new Core.Model.EPoint(), new Core.Model.EPoint());
                    lineCalloutInfoAccentBar.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutInfoAccentBar.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutInfoAccentBar.StartPointTail = lineCalloutInfoAccentBar.Point1;
                    return lineCalloutInfoAccentBar;
                case office.MsoAutoShapeType.msoShapeLineCallout3AccentBar:
                    LineCalloutShape2HasAccentBarNoBorderInfo lineCalloutShape2AccentBar = new LineCalloutShape2HasAccentBarNoBorderInfo();
                    lineCalloutShape2AccentBar.IsHasAccentBar = true;
                    lineCalloutShape2AccentBar.IsHasStroke = false;
                    lineCalloutShape2AccentBar.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutShape2AccentBar.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutShape2AccentBar.Point3 = new Core.Model.EPoint(shape.Adjustments[6] * shape.Width * Utils.tlw, shape.Adjustments[5] * shape.Height * Utils.tlh);
                    lineCalloutShape2AccentBar.StartPointTail = lineCalloutShape2AccentBar.Point1;
                    return lineCalloutShape2AccentBar;
                case office.MsoAutoShapeType.msoShapeLineCallout4AccentBar:
                    LineCalloutShape3HasAccentBarNoBorderInfo lineCalloutShape3AccentBar = new LineCalloutShape3HasAccentBarNoBorderInfo();
                    lineCalloutShape3AccentBar.IsHasAccentBar = true;
                    lineCalloutShape3AccentBar.IsHasStroke = false;
                    lineCalloutShape3AccentBar.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutShape3AccentBar.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutShape3AccentBar.Point3 = new Core.Model.EPoint(shape.Adjustments[6] * shape.Width * Utils.tlw, shape.Adjustments[5] * shape.Height * Utils.tlh);
                    lineCalloutShape3AccentBar.Point4 = new Core.Model.EPoint(shape.Adjustments[8] * shape.Width * Utils.tlw, shape.Adjustments[7] * shape.Height * Utils.tlh);
                    lineCalloutShape3AccentBar.StartPointTail = lineCalloutShape3AccentBar.Point1;
                    return lineCalloutShape3AccentBar;
                case office.MsoAutoShapeType.msoShapeLineCallout1NoBorder:
                    break;
                case office.MsoAutoShapeType.msoShapeLineCallout2NoBorder:
                    LineCalloutShape1NoAccentBarHasBorderInfo lineCalloutInfoBorder = new LineCalloutShape1NoAccentBarHasBorderInfo(false, true, new Core.Model.EPoint(), new Core.Model.EPoint(), new Core.Model.EPoint());
                    lineCalloutInfoBorder.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutInfoBorder.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutInfoBorder.StartPointTail = lineCalloutInfoBorder.Point1;
                    return lineCalloutInfoBorder;
                case office.MsoAutoShapeType.msoShapeLineCallout3NoBorder:
                    LineCalloutShape2NoAccentBarHasBorderInfo lineCalloutShape2Border = new LineCalloutShape2NoAccentBarHasBorderInfo();
                    lineCalloutShape2Border.IsHasAccentBar = false;
                    lineCalloutShape2Border.IsHasStroke = true;
                    lineCalloutShape2Border.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutShape2Border.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutShape2Border.Point3 = new Core.Model.EPoint(shape.Adjustments[6] * shape.Width * Utils.tlw, shape.Adjustments[5] * shape.Height * Utils.tlh);
                    lineCalloutShape2Border.StartPointTail = lineCalloutShape2Border.Point1;
                    return lineCalloutShape2Border;
                case office.MsoAutoShapeType.msoShapeLineCallout4NoBorder:
                    LineCalloutShape3NoAccentBarHasBorderInfo lineCalloutShape3Border = new LineCalloutShape3NoAccentBarHasBorderInfo();
                    lineCalloutShape3Border.IsHasAccentBar = false;
                    lineCalloutShape3Border.IsHasStroke = true;
                    lineCalloutShape3Border.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutShape3Border.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutShape3Border.Point3 = new Core.Model.EPoint(shape.Adjustments[6] * shape.Width * Utils.tlw, shape.Adjustments[5] * shape.Height * Utils.tlh);
                    lineCalloutShape3Border.Point4 = new Core.Model.EPoint(shape.Adjustments[8] * shape.Width * Utils.tlw, shape.Adjustments[7] * shape.Height * Utils.tlh);
                    lineCalloutShape3Border.StartPointTail = lineCalloutShape3Border.Point1;
                    return lineCalloutShape3Border;
                case office.MsoAutoShapeType.msoShapeLineCallout1BorderandAccentBar:
                    break;
                case office.MsoAutoShapeType.msoShapeLineCallout2BorderandAccentBar:
                    LineCalloutShape1HasAccentBarHasBorderInfo lineCalloutInfoAccentBarBorder = new LineCalloutShape1HasAccentBarHasBorderInfo(true, true, new Core.Model.EPoint(), new Core.Model.EPoint(), new Core.Model.EPoint());
                    lineCalloutInfoAccentBarBorder.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutInfoAccentBarBorder.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutInfoAccentBarBorder.StartPointTail = lineCalloutInfoAccentBarBorder.Point1;
                    return lineCalloutInfoAccentBarBorder;
                case office.MsoAutoShapeType.msoShapeLineCallout3BorderandAccentBar:
                    LineCalloutShape2HasAccentBarHasBorderInfo lineCalloutShape2AccentBarBorder = new LineCalloutShape2HasAccentBarHasBorderInfo();
                    lineCalloutShape2AccentBarBorder.IsHasAccentBar = true;
                    lineCalloutShape2AccentBarBorder.IsHasStroke = true;
                    lineCalloutShape2AccentBarBorder.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutShape2AccentBarBorder.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutShape2AccentBarBorder.Point3 = new Core.Model.EPoint(shape.Adjustments[6] * shape.Width * Utils.tlw, shape.Adjustments[5] * shape.Height * Utils.tlh);
                    lineCalloutShape2AccentBarBorder.StartPointTail = lineCalloutShape2AccentBarBorder.Point1;
                    return lineCalloutShape2AccentBarBorder;
                case office.MsoAutoShapeType.msoShapeLineCallout4BorderandAccentBar:
                    LineCalloutShape3HasAccentBarHasBorderInfo lineCalloutShape3AccentBarBorder = new LineCalloutShape3HasAccentBarHasBorderInfo();
                    lineCalloutShape3AccentBarBorder.IsHasAccentBar = true;
                    lineCalloutShape3AccentBarBorder.IsHasStroke = true;
                    lineCalloutShape3AccentBarBorder.Point1 = new Core.Model.EPoint(shape.Adjustments[2] * shape.Width * Utils.tlw, shape.Adjustments[1] * shape.Height * Utils.tlh);
                    lineCalloutShape3AccentBarBorder.Point2 = new Core.Model.EPoint(shape.Adjustments[4] * shape.Width * Utils.tlw, shape.Adjustments[3] * shape.Height * Utils.tlh);
                    lineCalloutShape3AccentBarBorder.Point3 = new Core.Model.EPoint(shape.Adjustments[6] * shape.Width * Utils.tlw, shape.Adjustments[5] * shape.Height * Utils.tlh);
                    lineCalloutShape3AccentBarBorder.Point4 = new Core.Model.EPoint(shape.Adjustments[8] * shape.Width * Utils.tlw, shape.Adjustments[7] * shape.Height * Utils.tlh);
                    lineCalloutShape3AccentBarBorder.StartPointTail = lineCalloutShape3AccentBarBorder.Point1;
                    return lineCalloutShape3AccentBarBorder;
                case office.MsoAutoShapeType.msoShapeActionButtonCustom:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonHome:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonHelp:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonInformation:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonBackorPrevious:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonForwardorNext:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonBeginning:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonEnd:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonReturn:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonDocument:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonSound:
                    break;
                case office.MsoAutoShapeType.msoShapeActionButtonMovie:
                    break;
                case office.MsoAutoShapeType.msoShapeBalloon:
                    break;
                case office.MsoAutoShapeType.msoShapeNotPrimitive:
                    break;
                case office.MsoAutoShapeType.msoShapeFlowchartOfflineStorage:
                    break;
                case office.MsoAutoShapeType.msoShapeLeftRightRibbon:
                    break;
                case office.MsoAutoShapeType.msoShapeDiagonalStripe:
                    DiagonalStripeInfo diagonalStripeInfo = new DiagonalStripeInfo();
                    diagonalStripeInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    diagonalStripeInfo.FactorBodySize = diagonalStripeInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return diagonalStripeInfo;
                case office.MsoAutoShapeType.msoShapePie:
                    PieInfo pieInfo = new PieInfo();
                    pieInfo.EndAngle = shape.Adjustments[1] > 0 ? 360 - shape.Adjustments[1] : -shape.Adjustments[1];
                    pieInfo.StartAngle = shape.Adjustments[2] > 0 ? 360 - shape.Adjustments[2] : -shape.Adjustments[2];
                    return pieInfo;
                case office.MsoAutoShapeType.msoShapeNonIsoscelesTrapezoid:
                    break;
                case office.MsoAutoShapeType.msoShapeDecagon:
                    DecagonInfo decagonInfo = new DecagonInfo(10);
                    return decagonInfo;
                case office.MsoAutoShapeType.msoShapeHeptagon:
                    HelptagonInfo helptagonInfo = new HelptagonInfo(7);
                    return helptagonInfo;
                case office.MsoAutoShapeType.msoShapeDodecagon:
                    DodecagonInfo dodecagonInfo = new DodecagonInfo(12);
                    return dodecagonInfo;
                case office.MsoAutoShapeType.msoShape6pointStar:
                    Star6Info star6Info = new Star6Info();
                    star6Info.CountPoint = 6;
                    star6Info.Radius = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    star6Info.Factor = star6Info.Radius / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return star6Info;
                case office.MsoAutoShapeType.msoShape7pointStar:
                    Star7Info star7Info = new Star7Info();
                    star7Info.CountPoint = 7;
                    star7Info.Radius = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    star7Info.Factor = star7Info.Radius / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return star7Info;
                case office.MsoAutoShapeType.msoShape10pointStar:
                    Star10Info star10Info = new Star10Info();
                    star10Info.CountPoint = 10;
                    star10Info.Radius = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    star10Info.Factor = star10Info.Radius / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return star10Info;
                case office.MsoAutoShapeType.msoShape12pointStar:
                    Star12Info star12Info = new Star12Info();
                    star12Info.CountPoint = 10;
                    star12Info.Radius = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    star12Info.Factor = star12Info.Radius / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return star12Info;
                case office.MsoAutoShapeType.msoShapeRound1Rectangle:
                    RoundSingleCornerRectangleInfo roundSingle = new RoundSingleCornerRectangleInfo();
                    roundSingle.CornerSize = shape.Adjustments[1] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                    roundSingle.Factor = roundSingle.CornerSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return roundSingle;
                case office.MsoAutoShapeType.msoShapeRound2SameRectangle:
                    RoundSameSideCornerRectangleInfo roundSame = new RoundSameSideCornerRectangleInfo();
                    roundSame.CornerSize1 = shape.Adjustments[1] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                    roundSame.CornerSize2 = shape.Adjustments[2] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                    roundSame.Factor1 = roundSame.CornerSize1 / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    roundSame.Factor2 = roundSame.CornerSize2 / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return roundSame;
                case office.MsoAutoShapeType.msoShapeRound2DiagRectangle:
                    RoundDiagonalCornerRectangleInfo roundDiagonal = new RoundDiagonalCornerRectangleInfo();
                    roundDiagonal.CornerSize1 = shape.Adjustments[1] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                    roundDiagonal.CornerSize2 = shape.Adjustments[2] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                    roundDiagonal.Factor1 = roundDiagonal.CornerSize1 / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    roundDiagonal.Factor2 = roundDiagonal.CornerSize2 / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return roundDiagonal;
                case office.MsoAutoShapeType.msoShapeSnipRoundRectangle:
                    SnipAndRoundSingleCornerRectangleInfo snipAndRound = new SnipAndRoundSingleCornerRectangleInfo();
                    snipAndRound.CornerSize1 = shape.Adjustments[1] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                    snipAndRound.CornerSize2 = shape.Adjustments[2] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                    snipAndRound.Factor1 = snipAndRound.CornerSize1 / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    snipAndRound.Factor2 = snipAndRound.CornerSize2 / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return snipAndRound;
                case office.MsoAutoShapeType.msoShapeSnip1Rectangle:
                    SnipSingleCornerRectangleInfo snipSingleCornerRectangleInfo = new SnipSingleCornerRectangleInfo();
                    snipSingleCornerRectangleInfo.CornerSize = shape.Adjustments[1] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                    snipSingleCornerRectangleInfo.Factor = snipSingleCornerRectangleInfo.CornerSize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return snipSingleCornerRectangleInfo;
                case office.MsoAutoShapeType.msoShapeSnip2SameRectangle:
                    SnipSameSideCornerRectangleInfo snipSame = new SnipSameSideCornerRectangleInfo();
                    for (int i = 1; i <= shape.Adjustments.Count; i++)
                    {
                        if (i == 1)
                        {
                            snipSame.CornerSize1 = shape.Adjustments[i] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                            snipSame.Factor1 = snipSame.CornerSize1 / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                        }
                        else
                        {
                            snipSame.CornerSize2 = shape.Adjustments[i] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                            snipSame.Factor2 = snipSame.CornerSize2 / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                        }
                    }
                    return snipSame;
                case office.MsoAutoShapeType.msoShapeSnip2DiagRectangle:
                    SnipDiagonalCornerRectangleInfo snipDiagonal = new SnipDiagonalCornerRectangleInfo();
                    snipDiagonal.CornerSize1 = shape.Adjustments[1] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                    snipDiagonal.CornerSize2 = shape.Adjustments[2] * Math.Min((shape.Width * Utils.tlw), (shape.Height * Utils.tlh));
                    snipDiagonal.Factor1 = snipDiagonal.CornerSize1 / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    snipDiagonal.Factor2 = snipDiagonal.CornerSize2 / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return snipDiagonal;
                case office.MsoAutoShapeType.msoShapeFrame:
                    FrameInfo frameInfo = new FrameInfo();
                    frameInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    frameInfo.FactorBodySize = frameInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return frameInfo;
                case office.MsoAutoShapeType.msoShapeHalfFrame:
                    HalfFrameInfo halfFrameInfo = new HalfFrameInfo();
                    halfFrameInfo.BodySizeH = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    halfFrameInfo.FactorBodySizeH = halfFrameInfo.BodySizeH / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    halfFrameInfo.BodySizeW = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    halfFrameInfo.FactorBodySizeW = halfFrameInfo.BodySizeW / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return halfFrameInfo;
                case office.MsoAutoShapeType.msoShapeTear:
                    break;
                case office.MsoAutoShapeType.msoShapeChord:
                    ChordInfo chordInfo = new ChordInfo();
                    chordInfo.EndAngle = shape.Adjustments[1] > 0 ? 360 - shape.Adjustments[1] : -shape.Adjustments[1];
                    chordInfo.StartAngle = shape.Adjustments[2] > 0 ? 360 - shape.Adjustments[2] : -shape.Adjustments[2];
                    return chordInfo;
                case office.MsoAutoShapeType.msoShapeCorner:
                    break;
                case office.MsoAutoShapeType.msoShapeMathPlus:
                    PlusInfo plusInfo = new PlusInfo();
                    plusInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    plusInfo.FactorBodySize = plusInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return plusInfo;
                case office.MsoAutoShapeType.msoShapeMathMinus:
                    MinusInfo minusInfo = new MinusInfo();
                    minusInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    minusInfo.FactorBodySize = minusInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return minusInfo;
                case office.MsoAutoShapeType.msoShapeMathMultiply:
                    MultiplyInfo multiplyInfo = new MultiplyInfo();
                    multiplyInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    multiplyInfo.FactorBodySize = multiplyInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return multiplyInfo;
                case office.MsoAutoShapeType.msoShapeMathDivide:
                    DivisionInfo divisionInfo = new DivisionInfo();
                    divisionInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    divisionInfo.FactorBodySize = divisionInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return divisionInfo;
                case office.MsoAutoShapeType.msoShapeMathEqual:
                    EqualInfo equalInfo = new EqualInfo();
                    equalInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    equalInfo.FactorBodySize = equalInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    equalInfo.Space = shape.Adjustments[2] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    equalInfo.FactorSpace = equalInfo.Space / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    return equalInfo;
                case office.MsoAutoShapeType.msoShapeMathNotEqual:
                    NotEqualInfo notEqualInfo = new NotEqualInfo();
                    notEqualInfo.BodySize = shape.Adjustments[1] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    notEqualInfo.FactorBodySize = notEqualInfo.BodySize / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    notEqualInfo.Space = shape.Adjustments[3] * Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    notEqualInfo.FactorSpace = notEqualInfo.Space / Math.Min(shape.Width * Utils.tlw, shape.Height * Utils.tlh);
                    notEqualInfo.Angle = shape.Adjustments[2] > 0 ? 90 - (360 - shape.Adjustments[2]) : 90 - (-shape.Adjustments[2]);
                    return notEqualInfo;
                case office.MsoAutoShapeType.msoShapeCornerTabs:
                    break;
                case office.MsoAutoShapeType.msoShapeSquareTabs:
                    break;
                case office.MsoAutoShapeType.msoShapePlaqueTabs:
                    break;
                case office.MsoAutoShapeType.msoShapeGear6:
                    break;
                case office.MsoAutoShapeType.msoShapeGear9:
                    break;
                case office.MsoAutoShapeType.msoShapeFunnel:
                    break;
                case office.MsoAutoShapeType.msoShapePieWedge:
                    break;
                case office.MsoAutoShapeType.msoShapeLeftCircularArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeLeftRightCircularArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeSwooshArrow:
                    break;
                case office.MsoAutoShapeType.msoShapeCloud:
                    break;
                case office.MsoAutoShapeType.msoShapeChartX:
                    break;
                case office.MsoAutoShapeType.msoShapeChartStar:
                    break;
                case office.MsoAutoShapeType.msoShapeChartPlus:
                    break;
                case office.MsoAutoShapeType.msoShapeLineInverse:
                    break;
                default:
                    break;
            }
            return null;
        }

        public GroupElement GetGroupElement(pp.Shape shape, pp.TimeLine timeLine, OpenXmlPart openXmlPart)
        {
            GroupElement groupElement = new GroupElement();
            SetBasicProperties(shape, groupElement, timeLine);
            foreach (pp.Shape item in shape.GroupItems)
            {
                groupElement.Members.Add(GetShape(item, item.Type, timeLine, openXmlPart));
            }
            return groupElement;
        }

        /// <summary>
        /// Lấy ObjectElement
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="_idShape"></param>
        /// <param name="_animation"></param>
        /// <returns></returns>
        public ObjectElementBase GetObjectElementBase(pp.Shape shape, string _idShape, EAnimation _animation)
        {
            return null;
        }

        /// <summary>
        /// Get StandardElement
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="_idShape"></param>
        /// <param name="_animation"></param>
        /// <returns></returns>
        public EStandardElement GetEStandardElement(pp.Shape shape, TimeLine timeLine, P.Shape shapeTag, OpenXmlPart openXmlPart)
        {
            EStandardElement eStandardElement = new EStandardElement();
            SetBasicProperties(shape, eStandardElement, timeLine);
            SetStandardElementProperties(shape, shapeTag, eStandardElement, openXmlPart);
            SetEffectObjectElement(shape, eStandardElement, shapeTag);
            return eStandardElement;
        }



        /// <summary>
        /// Cài đặt thuộc tính StandardElement
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="openXmlPart"></param>
        /// <param name="eStandardElement"></param>
        public void SetStandardElementProperties(pp.Shape shape, OpenXmlCompositeElement shapeTag, EStandardElement eStandardElement, OpenXmlPart openXmlPart)
        {
            try
            {
                eStandardElement.ShapePresent = GetShapePresent(shape, shapeTag);
            }
            catch { }
            eStandardElement.Border = GetBorderShape(shape, shapeTag, openXmlPart);
            eStandardElement.IsLikeSlideFill = eStandardElement.Border.Fill == null;
        }

        /// <summary>
        /// Cài đặt hiệu ứng cho StandardElement
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="eStandardElement"></param>
        private void SetEffectObjectElement(pp.Shape shape, EStandardElement eStandardElement, P.Shape effectTag)
        {
            eStandardElement.FrameEffects = new System.Collections.ObjectModel.ObservableCollection<FrameEffectBase>();
            if (effectTag == null) return;
            foreach (var item in effectTag.ShapeProperties)
            {
                if (item is EffectList effects)
                {
                    //Cài đặt outterShadow
                    if (effects.OuterShadow != null)
                    {
                        ShadowEffect shadowEffect = new ShadowEffect();
                        if (effects.OuterShadow.BlurRadius != null) shadowEffect.Blur = EmuToPixels(effects.OuterShadow.BlurRadius.Value);
                        if (effects.OuterShadow.Distance != null) shadowEffect.Distance = EmuToPixels(effects.OuterShadow.Distance.Value);
                        if (effects.OuterShadow.Direction != null) shadowEffect.Angle = effects.OuterShadow.Direction / 60000;

                        shadowEffect.Size = shape.Shadow.Size / 100;
                        shadowEffect.FakeSize = shape.Shadow.Size / 100;
                        shadowEffect.Transparency = shape.Shadow.Transparency;
                        shadowEffect.Color = ConvertColor(shape.Shadow.ForeColor.RGB);
                        shadowEffect.IsSize = true;
                        if (effects.OuterShadow.HorizontalRatio == null || effects.OuterShadow.VerticalRatio == null)
                        {
                            shadowEffect.IsSize = false;
                        }

                        shadowEffect.Type = ShadowType.Perspective;
                        if (effects.OuterShadow.VerticalSkew == null && effects.OuterShadow.HorizontalSkew == null && effects.OuterShadow.HorizontalRatio == null && effects.OuterShadow.VerticalRatio == null)
                            shadowEffect.Type = ShadowType.Outer;
                        else
                        {
                            shadowEffect.Type = ShadowType.Perspective;
                            if (effects.OuterShadow.HorizontalRatio != null && effects.OuterShadow.VerticalRatio != null && effects.OuterShadow.HorizontalSkew == null && effects.OuterShadow.VerticalSkew == null)
                            {
                                shadowEffect.PerspectiveShadowType = Core.Helper.Effect.PerspectiveShadowEnum.Below;
                                if (effects.OuterShadow.VerticalRatio > 0) shadowEffect.IsSize = true;
                                else shadowEffect.IsSize = false;
                            }
                            else if (effects.OuterShadow.VerticalRatio == null)
                            {
                                shadowEffect.PerspectiveShadowType = effects.OuterShadow.HorizontalRatio > 0 ? Core.Helper.Effect.PerspectiveShadowEnum.UpperLeft : Core.Helper.Effect.PerspectiveShadowEnum.UpperRight;
                            }
                            else
                            {
                                if (effects.OuterShadow.VerticalRatio < 0)
                                    shadowEffect.PerspectiveShadowType = effects.OuterShadow.HorizontalSkew > 0 ? Core.Helper.Effect.PerspectiveShadowEnum.LowerLeft : Core.Helper.Effect.PerspectiveShadowEnum.LowerRight;
                                else
                                {
                                    shadowEffect.PerspectiveShadowType = effects.OuterShadow.HorizontalSkew > 0 ? Core.Helper.Effect.PerspectiveShadowEnum.UpperLeft : Core.Helper.Effect.PerspectiveShadowEnum.UpperRight;
                                }
                            }
                        }
                        eStandardElement.FrameEffects.Add(shadowEffect);
                    }
                    //Cài đặt Inner Shadow
                    if (effects.InnerShadow != null)
                    {
                        ShadowEffect shadowEffect = new ShadowEffect();
                        if (effects.InnerShadow.BlurRadius != null) shadowEffect.Blur = EmuToPixels(effects.InnerShadow.BlurRadius.Value);
                        if (effects.InnerShadow.Distance != null) shadowEffect.Distance = EmuToPixels(effects.InnerShadow.Distance.Value);
                        shadowEffect.Size = shape.Shadow.Size;
                        shadowEffect.Transparency = shape.Shadow.Transparency;
                        shadowEffect.Color = ConvertColor(shape.Shadow.ForeColor.RGB);
                        shadowEffect.Type = ShadowType.Inner;
                        eStandardElement.FrameEffects.Add(shadowEffect);
                    }
                    if (effects.Reflection != null)
                    {
                        ReflectionEffect reflectionEffect = new ReflectionEffect();
                        reflectionEffect.Blur = shape.Reflection.Blur;
                        reflectionEffect.Transparency = shape.Reflection.Transparency;
                        reflectionEffect.Size = shape.Reflection.Size;
                        if (effects.Reflection.Distance != null) reflectionEffect.Distance = EmuToPixels(effects.Reflection.Distance.Value);
                        eStandardElement.FrameEffects.Add(reflectionEffect);
                    }
                    if (effects.Glow != null)
                    {
                        GlowEffect glowEffect = new GlowEffect();
                        glowEffect.Color = ConvertColor(shape.Glow.Color.RGB);
                        glowEffect.Transparency = shape.Glow.Transparency;
                        glowEffect.Size = PtToPx(shape.Glow.Radius);
                        eStandardElement.FrameEffects.Add(glowEffect);
                    }
                    if (effects.SoftEdge != null)
                    {
                        SoftEdgeEffect softEdgeEffect = new SoftEdgeEffect();
                        softEdgeEffect.Size = PtToPx(shape.SoftEdge.Radius);
                        eStandardElement.FrameEffects.Add(softEdgeEffect);
                    }
                }
            }


            //shadowEffect.Angle = shape.Shadow.
        }


        /// <summary>
        /// Hàm lấy thông tin cơ bản của ObjectElement
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="objectElementBase"></param>
        private void SetBasicProperties(pp.Shape shape, ObjectElementBase objectElementBase, TimeLine timeLine)
        {
            HelperClass helperClass = new HelperClass();
            objectElementBase.ID = ObjectElementsHelper.RandomString(10) + shape.Id.ToString();
            objectElementBase.Width = shape.Width * Utils.tlw;
            objectElementBase.Height = shape.Height * Utils.tlh;
            objectElementBase.Top = shape.Top * Utils.tlh;
            objectElementBase.Left = shape.Left * Utils.tlw;
            objectElementBase.Angle = shape.Rotation;
            objectElementBase.ZOder = (int)shape.ZOrderPosition;
            objectElementBase.Animations = GetAnimation(shape, timeLine, objectElementBase.ID);
            objectElementBase.IsScaleX = shape.HorizontalFlip == MsoTriState.msoTrue;
            objectElementBase.IsScaleY = shape.VerticalFlip == MsoTriState.msoTrue;
            objectElementBase.Name = shape.Name;
        }

        /// <summary>
        /// Hàm lấy viền của StandardElement
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private EBorder GetBorderShape(pp.Shape shape, OpenXmlCompositeElement shapeTag, OpenXmlPart openXmlPart)
        {
            EBorder _border = new EBorder();
            _border.Fill = GetFillColor(shape, shapeTag, openXmlPart);
            _border.BorderBrush = GetBorderBrushColor(shape, shapeTag, openXmlPart);
            //1707
            try
            {
                if (shape.Line.Visible == MsoTriState.msoFalse)
                    _border.BorderThickness = 0;
                 else
                _border.BorderThickness = (double)shape.Line.Weight;
                _border.DashType = DashTypeConverter(shape.Line.DashStyle);
                SetBorderShape(shape, ref _border, shapeTag);
            }
            catch (Exception ex)
            {
                _border = new EBorder();
            }
            return _border;
        }

        /// <summary>
        /// get CapType and Jointype
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="eBorder"></param>
        private void SetBorderShape(pp.Shape shape, ref EBorder eBorder, OpenXmlCompositeElement s)
        {
            if (s == null) return;
            if (s is P.Shape _shape)
            {
                foreach (var item in _shape.ShapeProperties)
                {
                    if (item is D.Outline outline)
                    {
                        if (outline.CapType != null)
                            eBorder.CapType = CapTypeConverter(outline.CapType);
                        foreach (var type in outline)
                        {
                            if (type is D.Round)
                            {
                                eBorder.JoinType = PenLineJoin.Round;
                                break;
                            }
                            if (type is D.Miter)
                            {
                                eBorder.JoinType = PenLineJoin.Miter;
                                break;
                            }
                            if (type is D.Bevel)
                            {
                                eBorder.JoinType = PenLineJoin.Bevel;
                                break;
                            }
                        }
                        break;
                    }
                }
            }
            else if (s is P.Picture pic)
            {
                foreach (var item in pic.ShapeProperties)
                {
                    if (item is D.Outline outline)
                    {
                        if (outline.CapType != null)
                            eBorder.CapType = CapTypeConverter(outline.CapType);
                        foreach (var type in outline)
                        {
                            if (type is D.Round)
                            {
                                eBorder.JoinType = PenLineJoin.Round;
                                break;
                            }
                            if (type is D.Miter)
                            {
                                eBorder.JoinType = PenLineJoin.Miter;
                                break;
                            }
                            if (type is D.Bevel)
                            {
                                eBorder.JoinType = PenLineJoin.Bevel;
                                break;
                            }
                        }
                        break;
                    }
                }
            }

        }



        /// <summary>
        /// Hàm chuyển đổi PenLineCap từ pp sang Avina
        /// </summary>
        /// <param name="lineCapValues"></param>
        /// <returns></returns>
        private PenLineCap CapTypeConverter(D.LineCapValues lineCapValues)
        {
            switch (lineCapValues)
            {
                case D.LineCapValues.Round:
                    return PenLineCap.Round;
                case D.LineCapValues.Square:
                    return PenLineCap.Square;
                case D.LineCapValues.Flat:
                    return PenLineCap.Flat;
                default:
                    break;
            }
            return PenLineCap.Triangle;
        }

        /// <summary>
        /// Hàm Convert từ DashType của PP sang DashType của Avina
        /// </summary>
        /// <param name="msoLineDashStyle"></param>
        /// <returns></returns>
        private DashType DashTypeConverter(MsoLineDashStyle msoLineDashStyle)
        {
            switch (msoLineDashStyle)
            {
                case MsoLineDashStyle.msoLineDashStyleMixed:
                    break;
                case MsoLineDashStyle.msoLineSolid:
                    return DashType.Solid;
                case MsoLineDashStyle.msoLineSquareDot:
                    return DashType.SquareDot;
                case MsoLineDashStyle.msoLineRoundDot:
                    return DashType.RoundDot;
                case MsoLineDashStyle.msoLineDash:
                    return DashType.Dash;
                case MsoLineDashStyle.msoLineDashDot:
                    return DashType.DashDot;
                case MsoLineDashStyle.msoLineDashDotDot:
                    return DashType.LongDashDotDot;
                case MsoLineDashStyle.msoLineLongDash:
                    return DashType.LongDash;
                case MsoLineDashStyle.msoLineLongDashDot:
                    return DashType.LongDashDot;
                case MsoLineDashStyle.msoLineLongDashDotDot:
                    return DashType.LongDashDotDot;
                case MsoLineDashStyle.msoLineSysDash:
                    return DashType.Dash;
                case MsoLineDashStyle.msoLineSysDot:
                    return DashType.DashDot;
                case MsoLineDashStyle.msoLineSysDashDot:
                    return DashType.DashDot;
                default:
                    break;
            }
            return DashType.Solid;
        }

        /// <summary>
        /// hàm convert từ ArrowTypes của PP sang Arrowtype của Avina
        /// </summary>
        /// <param name="msoArrowheadStyle"></param>
        /// <returns></returns>
        private ArrowTypes ArrowTypesConverter(MsoArrowheadStyle msoArrowheadStyle)
        {
            switch (msoArrowheadStyle)
            {
                case MsoArrowheadStyle.msoArrowheadStyleMixed:
                    return ArrowTypes.None;
                case MsoArrowheadStyle.msoArrowheadNone:
                    return ArrowTypes.None;
                case MsoArrowheadStyle.msoArrowheadTriangle:
                    return ArrowTypes.TriangleArrow;
                case MsoArrowheadStyle.msoArrowheadOpen:
                    return ArrowTypes.OpenArrow;
                case MsoArrowheadStyle.msoArrowheadStealth:
                    return ArrowTypes.StealthArrow;
                case MsoArrowheadStyle.msoArrowheadDiamond:
                    return ArrowTypes.DiamondArrow;
                case MsoArrowheadStyle.msoArrowheadOval:
                    return ArrowTypes.OvalArrow;
                default:
                    return ArrowTypes.None;
            }
        }

    
        /// <summary>
        /// Get Index Ans
        /// </summary>
        /// <param name="timeLine"></param>
        /// <param name="_shape"></param>
        /// <returns></returns>
        private int GetIndexAns(pp.TimeLine timeLine, pp.Shape _shape)
        {
            for (int i = 1; i <= timeLine.MainSequence.Count; i++)
            {
                if (timeLine.MainSequence[i].Shape == _shape)
                    return i;
            }
            return 0;
        }



        #region GetPathData
        public string GetPathDataCustomGeometry(pp.Shape shape, OpenXmlCompositeElement shapeTag)
        {
            string pathData = string.Empty;
            D.CustomGeometry Geometry = GetGeometry(shapeTag);
            if (Geometry == null) return pathData;
            foreach (D.Path path in Geometry.PathList)
            {
                foreach (var points in path)
                {
                    if (points is D.MoveTo moveTo)
                    {
                        pathData += "M " + EmuToPixels(int.Parse(moveTo.Point.X)) + ", " + EmuToPixels(int.Parse(moveTo.Point.Y)) + " ";
                    }
                    else if (points is D.CubicBezierCurveTo cubic)
                    {
                        pathData += "C";
                        foreach (D.Point point in cubic)
                        {
                            pathData += EmuToPixels(int.Parse(point.X)) + "," + EmuToPixels(int.Parse(point.Y)) + " ";
                        }
                    }
                    else if (points is D.LineTo lineTo)
                    {
                        pathData += "L " + EmuToPixels(int.Parse(lineTo.Point.X)) + ", " + EmuToPixels(int.Parse(lineTo.Point.Y)) + " ";
                    }
                    else if (points is D.QuadraticBezierCurveTo quadTo)
                    {
                        pathData += "Q";
                        foreach (D.Point point in quadTo)
                        {
                            pathData += EmuToPixels(int.Parse(point.X)) + "," + EmuToPixels(int.Parse(point.Y)) + " ";
                        }
                    }
                    else if (points is D.ArcTo arc)
                    {
                        double endX = Math.Cos(EmuToPixels(int.Parse(arc.SwingAngle))) * EmuToPixels(int.Parse(arc.WidthRadius));
                        double endY = Math.Sin(EmuToPixels(int.Parse(arc.SwingAngle))) * EmuToPixels(int.Parse(arc.HeightRadius));
                        pathData += "A" + EmuToPixels(int.Parse(arc.WidthRadius)) + "," + EmuToPixels(int.Parse(arc.HeightRadius)) + " " + EmuToPixels(int.Parse(arc.StartAngle)) + (EmuToPixels(int.Parse(arc.SwingAngle)) > 180 ? "1" : "0") + ((EmuToPixels(int.Parse(arc.SwingAngle)) > 0 ? "1" : "0")) + " " + endX + "," + endY + " ";
                    }
                    else if (points is D.CloseShapePath)
                    {
                        pathData += "z";
                    }
                }
            }
            return pathData;
        }


        public int PixelsToEmu(int pixels)
        {
            return (int)Math.Round((decimal)pixels * 9525);
        }

        /// <summary>
        /// Convert EMU to pixels
        /// </summary>
        /// <param name="emu">Size in EMU</param>
        /// <returns>Size in pixels</returns>
        public double EmuToPixels(long emu)
        {
            if (emu != 0)
            {
                return (double)Math.Round((decimal)emu / 9525, 5);
            }
            else
            {
                return 0.0;
            }
        }

        public double PtToPx(float pt)
        {
            return pt * 96 / 72;
        }

        /// <summary>
        /// Get Tab Shape
        /// </summary>
        /// <param name="_idShape"></param>
        /// <returns></returns>
        private P.Shape GetShapeTag(string _idShape, OpenXmlPart openXmlPart)
        {
            IEnumerable<P.Shape> listShape = null;
            if (openXmlPart is SlidePart slidePart)
            {
                listShape = slidePart?.Slide?.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().Select(p => p);
            }
            else if (openXmlPart is SlideMasterPart slideMasterPart)
            {
                listShape = slideMasterPart?.SlideMaster?.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().Select(p => p);
            }
            else if (openXmlPart is SlideLayoutPart slideLayoutPart)
            {
                listShape = slideLayoutPart?.SlideLayout?.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().Select(p => p);
            }


            if (listShape == null) return null;
            foreach (P.Shape shape in listShape)
            {
                if (shape.NonVisualShapeProperties.NonVisualDrawingProperties.Id == _idShape)
                    return shape;
            }
            return null;
        }
        /// <summary>
        /// Get Custom Geometry
        /// </summary>
        /// <param name="_shape"></param>
        /// <returns></returns>
        private D.CustomGeometry GetGeometry(OpenXmlCompositeElement shape)
        {
            if (shape is P.Shape)
            {
                foreach (OpenXmlElement element in (shape as P.Shape).ShapeProperties.ChildElements)
                {
                    if (element is D.CustomGeometry)
                        return (D.CustomGeometry)element;
                }
            }
            else if (shape is P.Picture)
            {
                foreach (OpenXmlElement element in (shape as P.Picture).ShapeProperties.ChildElements)
                {
                    if (element is D.CustomGeometry)
                        return (D.CustomGeometry)element;
                }
            }
            return null;
        }
        private bool MathOrWordArt(pp.Shape _shape, P.Shape _shapeTag)
        {
            //-- Khung chứa công thức toán
            foreach (pp.TextRange run in _shape.TextFrame.TextRange.Paragraphs().Runs())
            {
                if (run?.Font?.NameAscii == "Cambria Math" || run?.Font?.Name == "Cambria Math" || run?.Font?.NameOther == "Cambria Math")
                    return true;
            }
            //-- Nếu khung chứa chữ nghệ thuật
            var listRun = _shapeTag.Descendants<D.Run>();
            foreach (D.Run run in listRun)
            {
                if (run?.RunProperties.Outline != null)
                    return true;
            }
            return false;
        }
        #endregion


    }
}
