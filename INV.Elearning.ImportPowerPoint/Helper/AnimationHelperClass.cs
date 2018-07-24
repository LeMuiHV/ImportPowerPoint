using INV.Elearning.Core.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using pp = Microsoft.Office.Interop.PowerPoint;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using INV.Elearning.Animations;
using INV.Elearning.Animations.Enums;
using INV.Elearning.Animations.UI;
using Microsoft.Office.Interop.PowerPoint;
using office = Microsoft.Office.Core;
using System.Windows;
using System.Windows.Media;

namespace INV.Elearning.ImportPowerPoint.Helper
{
    public partial class HelperClass
    {
        /// <summary>
        /// Hàm lấy animation
        /// </summary>
        /// <returns></returns>
        public EAnimation GetAnimation(pp.Shape shape, pp.TimeLine timeLine, string idShape)
        {
            EAnimation animation = new EAnimation();
            animation.Entrance = new NoneAnimation("None", TimeSpan.FromSeconds(0));
            animation.Exit = new NoneAnimation("None", TimeSpan.FromSeconds(0));
            animation.MotionPaths = new List<DataMotionPath>();
            for (int i = 1; i <= timeLine.MainSequence.Count; i++)
            {
                if (timeLine.MainSequence[i].Shape == shape)
                    GetAnimation(timeLine.MainSequence[i], animation, shape, idShape);
            }
            return animation;
        }
        //---Mặc định sẽ là hiệu ứng FlyIn và Flyout
        private void GetAnimation(pp.Effect shapeAnimation, EAnimation animation, pp.Shape shape, string idShape)
        {
            double durAns = (double)shapeAnimation.Timing.Duration;
            if (durAns < 0.25) durAns = 0.25;
            try
            {
                switch (shapeAnimation.EffectType)
                {
                    case MsoAnimEffect.msoAnimEffectCustom:
                        break;
                    case MsoAnimEffect.msoAnimEffectAppear:
                        break;
                    case MsoAnimEffect.msoAnimEffectFly:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {

                            GetFlyInEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                        else
                        {
                            GetFlyExitAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectBlinds:
                        break;
                    case MsoAnimEffect.msoAnimEffectBox:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)

                        {
                            GetShapeEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation), KeyLanguage.Box, shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                        else
                        {
                            GetShapeExitAns(animation, durAns, GetAnimationDirection(shapeAnimation), KeyLanguage.Box, shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectCheckerboard:
                        break;
                    case MsoAnimEffect.msoAnimEffectCircle:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetShapeEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation), KeyLanguage.Circle, shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                        else
                        {
                            GetShapeExitAns(animation, durAns, GetAnimationDirection(shapeAnimation), KeyLanguage.Circle, shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectCrawl:
                        break;
                    case MsoAnimEffect.msoAnimEffectDiamond:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetShapeEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation), KeyLanguage.Diamond, shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                        else
                        {
                            GetShapeExitAns(animation, durAns, GetAnimationDirection(shapeAnimation), KeyLanguage.Diamond, shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectDissolve:
                        break;
                    case MsoAnimEffect.msoAnimEffectFade:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetFadeEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                        else
                        {
                            GetFadeExitAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectFlashOnce:
                        break;
                    case MsoAnimEffect.msoAnimEffectPeek:
                        break;
                    case MsoAnimEffect.msoAnimEffectPlus:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetShapeEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation), KeyLanguage.Plus, shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                        else
                        {
                            GetShapeExitAns(animation, durAns, GetAnimationDirection(shapeAnimation), KeyLanguage.Plus, shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectRandomBars:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetRandomBarsEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                        else
                        {
                            GetRandomBarsExitAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectSpiral:
                        break;
                    case MsoAnimEffect.msoAnimEffectSplit:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetSplitEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                        else
                        {
                            GetSplitExitAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectStretch:
                        break;
                    case MsoAnimEffect.msoAnimEffectStrips:
                        break;
                    case MsoAnimEffect.msoAnimEffectFadedSwivel:
                    case MsoAnimEffect.msoAnimEffectSwivel:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetSwivelEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                        else
                        {
                            GetSwivelExitAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectWedge:
                        break;
                    case MsoAnimEffect.msoAnimEffectWheel:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetWheelEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectParameters.Amount, shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                        else
                        {
                            GetWheelExitAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectParameters.Amount, shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectWipe:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetWipeEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                        else
                        {
                            GetWipeExitAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectZoom:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetZoomEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                        else
                        {
                            GetZoomExitAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectRandomEffects:
                        break;
                    case MsoAnimEffect.msoAnimEffectBoomerang:
                        break;
                    case MsoAnimEffect.msoAnimEffectBounce:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetBounceEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                        else
                        {
                            GetBounceExitAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectColorReveal:
                        break;
                    case MsoAnimEffect.msoAnimEffectCredits:
                        break;
                    case MsoAnimEffect.msoAnimEffectEaseIn:
                        break;
                    case MsoAnimEffect.msoAnimEffectDescend:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetFloatEntrAns(animation, durAns, KeyLanguage.FloatDown);
                            return;
                        }
                        else
                        {
                            GetFloatExitAns(animation, durAns, KeyLanguage.FloatUp);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectFloat:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetFloatEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                        else
                        {
                            GetFloatExitAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectGrowAndTurn:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetGrowEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                        else
                        {
                            GetShrinkTurnExitAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectLightSpeed:
                        break;
                    case MsoAnimEffect.msoAnimEffectPinwheel:
                        break;
                    case MsoAnimEffect.msoAnimEffectRiseUp:
                        break;
                    case MsoAnimEffect.msoAnimEffectSwish:
                        break;
                    case MsoAnimEffect.msoAnimEffectThinLine:
                        break;
                    case MsoAnimEffect.msoAnimEffectUnfold:
                        break;
                    case MsoAnimEffect.msoAnimEffectWhip:
                        break;
                    case MsoAnimEffect.msoAnimEffectAscend:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetFloatEntrAns(animation, durAns, KeyLanguage.FloatUp);
                            return;
                        }
                        else
                        {
                            GetFloatExitAns(animation, durAns, KeyLanguage.FloatDown);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectCenterRevolve:
                        break;

                    case MsoAnimEffect.msoAnimEffectSling:
                        break;
                    case MsoAnimEffect.msoAnimEffectSpinner:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetSpinEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                        else
                        {
                            GetSpinExitAns(animation, durAns, GetAnimationDirection(shapeAnimation));
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectStretchy:
                        break;
                    case MsoAnimEffect.msoAnimEffectZip:
                        break;
                    case MsoAnimEffect.msoAnimEffectArcUp:
                        break;
                    case MsoAnimEffect.msoAnimEffectFadedZoom:
                        if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
                        {
                            GetZoomEntrAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                        else
                        {
                            GetZoomExitAns(animation, durAns, GetAnimationDirection(shapeAnimation), shapeAnimation.EffectInformation.BuildByLevelEffect);
                            return;
                        }
                    case MsoAnimEffect.msoAnimEffectGlide:
                        break;
                    case MsoAnimEffect.msoAnimEffectExpand:
                        break;
                    case MsoAnimEffect.msoAnimEffectFlip:
                        break;
                    case MsoAnimEffect.msoAnimEffectShimmer:
                        break;
                    case MsoAnimEffect.msoAnimEffectFold:
                        break;
                    case MsoAnimEffect.msoAnimEffectChangeFillColor:
                        break;
                    case MsoAnimEffect.msoAnimEffectChangeFont:
                        break;
                    case MsoAnimEffect.msoAnimEffectChangeFontColor:
                        break;
                    case MsoAnimEffect.msoAnimEffectChangeFontSize:
                        break;
                    case MsoAnimEffect.msoAnimEffectChangeFontStyle:
                        break;
                    case MsoAnimEffect.msoAnimEffectGrowShrink:
                        break;
                    case MsoAnimEffect.msoAnimEffectChangeLineColor:
                        break;
                    case MsoAnimEffect.msoAnimEffectSpin:
                        break;
                    case MsoAnimEffect.msoAnimEffectTransparency:
                        break;
                    case MsoAnimEffect.msoAnimEffectBoldFlash:
                        break;
                    case MsoAnimEffect.msoAnimEffectBlast:
                        break;
                    case MsoAnimEffect.msoAnimEffectBoldReveal:
                        break;
                    case MsoAnimEffect.msoAnimEffectBrushOnColor:
                        break;
                    case MsoAnimEffect.msoAnimEffectBrushOnUnderline:
                        break;
                    case MsoAnimEffect.msoAnimEffectColorBlend:
                        break;
                    case MsoAnimEffect.msoAnimEffectColorWave:
                        break;
                    case MsoAnimEffect.msoAnimEffectComplementaryColor:
                        break;
                    case MsoAnimEffect.msoAnimEffectComplementaryColor2:
                        break;
                    case MsoAnimEffect.msoAnimEffectContrastingColor:
                        break;
                    case MsoAnimEffect.msoAnimEffectDarken:
                        break;
                    case MsoAnimEffect.msoAnimEffectDesaturate:
                        break;
                    case MsoAnimEffect.msoAnimEffectFlashBulb:
                        break;
                    case MsoAnimEffect.msoAnimEffectFlicker:
                        break;
                    case MsoAnimEffect.msoAnimEffectGrowWithColor:
                        break;
                    case MsoAnimEffect.msoAnimEffectLighten:
                        break;
                    case MsoAnimEffect.msoAnimEffectStyleEmphasis:
                        break;
                    case MsoAnimEffect.msoAnimEffectTeeter:
                        break;
                    case MsoAnimEffect.msoAnimEffectVerticalGrow:
                        break;
                    case MsoAnimEffect.msoAnimEffectWave:
                        break;
                    case MsoAnimEffect.msoAnimEffectMediaPlay:
                        break;
                    case MsoAnimEffect.msoAnimEffectMediaPause:
                        break;
                    case MsoAnimEffect.msoAnimEffectMediaStop:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathCircle:
                        GetCirclePath(animation, shape, durAns, idShape, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathRightTriangle:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathDiamond:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathHexagon:
                        break;
                    case MsoAnimEffect.msoAnimEffectPath5PointStar:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathCrescentMoon:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathSquare:
                        GetSquarePath(animation, shape, durAns, idShape, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathTrapezoid:
                        GetTrapeZoidPath(animation, shape, durAns, idShape, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathHeart:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathOctagon:
                        break;
                    case MsoAnimEffect.msoAnimEffectPath6PointStar:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathFootball:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathEqualTriangle:
                        GetEqualTrianglePath(animation, shape, durAns, idShape, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathParallelogram:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathPentagon:
                        break;
                    case MsoAnimEffect.msoAnimEffectPath4PointStar:
                        break;
                    case MsoAnimEffect.msoAnimEffectPath8PointStar:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathTeardrop:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathPointyStar:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathCurvedSquare:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathCurvedX:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathVerticalFigure8:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathCurvyStar:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathLoopdeLoop:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathBuzzsaw:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathHorizontalFigure8:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathPeanut:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathFigure8Four:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathNeutron:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathSwoosh:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathBean:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathPlus:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathInvertedTriangle:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathInvertedSquare:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathLeft:
                        GetLinePath(animation, shape, durAns, idShape, KeyLanguage.Left, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathTurnRight:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathArcDown:
                        GetArcPath(animation, shape, durAns, idShape, KeyLanguage.Down, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathZigzag:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathSCurve2:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathSineWave:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathBounceLeft:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathDown:
                        GetLinePath(animation, shape, durAns, idShape, KeyLanguage.Down, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathTurnUp:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathArcUp:
                        GetArcPath(animation, shape, durAns, idShape, KeyLanguage.Up, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathHeartbeat:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathSpiralRight:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathWave:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathCurvyLeft:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathDiagonalDownRight:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathTurnDown:
                        GetTurnsPath(animation, shape, durAns, idShape, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathArcLeft:
                        GetArcPath(animation, shape, durAns, idShape, KeyLanguage.Left, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathFunnel:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathSpring:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathBounceRight:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathSpiralLeft:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathDiagonalUpRight:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathTurnUpRight:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathArcRight:
                        GetArcPath(animation, shape, durAns, idShape, KeyLanguage.Right, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathSCurve1:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathDecayingWave:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathCurvyRight:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathStairsDown:
                        break;
                    case MsoAnimEffect.msoAnimEffectPathUp:
                        GetLinePath(animation, shape, durAns, idShape, KeyLanguage.Up, GetAnimationLock(shapeAnimation));
                        return;
                    case MsoAnimEffect.msoAnimEffectPathRight:
                        GetLinePath(animation, shape, durAns, idShape, KeyLanguage.Right, GetAnimationLock(shapeAnimation));
                        return;
                    default:
                        break;
                }
            }
            catch { }
            if (shapeAnimation.Exit == office.MsoTriState.msoFalse)
            {
                GetFlyInEntrAns(animation, durAns, KeyLanguage.Top);
                return;
            }
            else
            {
                GetFlyExitAns(animation, durAns, KeyLanguage.Top);
                return;
            }
        }

        #region Lấy Animation

        /// <summary>
        /// Lấy tùy chọn Animation
        /// </summary>
        /// <returns></returns>
        private string GetAnimationDirection(pp.Effect shapeAnimation)
        {
            string direction = "None";
            try
            {
                if (shapeAnimation?.EffectParameters?.Direction == null) return direction;
            }
            catch
            {
                return direction;
            }

            switch (shapeAnimation.EffectParameters?.Direction)
            {
                case MsoAnimDirection.msoAnimDirectionNone:
                    direction = KeyLanguage.None;
                    break;
                case MsoAnimDirection.msoAnimDirectionUp:
                    direction = KeyLanguage.Top;
                    break;
                case MsoAnimDirection.msoAnimDirectionRight:
                    direction = KeyLanguage.Right;
                    break;
                case MsoAnimDirection.msoAnimDirectionDown:
                    direction = KeyLanguage.Bottom;
                    break;
                case MsoAnimDirection.msoAnimDirectionLeft:
                    direction = KeyLanguage.Left;
                    break;
                case MsoAnimDirection.msoAnimDirectionUpLeft:
                    direction = KeyLanguage.TopLeft;
                    break;
                case MsoAnimDirection.msoAnimDirectionUpRight:
                    direction = KeyLanguage.TopRight;
                    break;
                case MsoAnimDirection.msoAnimDirectionDownRight:
                    direction = KeyLanguage.BottomRight;
                    break;
                case MsoAnimDirection.msoAnimDirectionDownLeft:
                    direction = KeyLanguage.BottomLeft;
                    break;
                case MsoAnimDirection.msoAnimDirectionTop:
                    direction = KeyLanguage.Top;
                    break;
                case MsoAnimDirection.msoAnimDirectionBottom:
                    direction = KeyLanguage.Bottom;
                    break;
                case MsoAnimDirection.msoAnimDirectionTopLeft:
                    direction = KeyLanguage.TopLeft;
                    break;
                case MsoAnimDirection.msoAnimDirectionTopRight:
                    direction = KeyLanguage.TopRight;
                    break;
                case MsoAnimDirection.msoAnimDirectionBottomRight:
                    direction = KeyLanguage.BottomRight;
                    break;
                case MsoAnimDirection.msoAnimDirectionBottomLeft:
                    direction = KeyLanguage.BottomLeft;
                    break;
                case MsoAnimDirection.msoAnimDirectionHorizontal:
                    direction = KeyLanguage.Horizontal;
                    break;
                case MsoAnimDirection.msoAnimDirectionVertical:
                    direction = KeyLanguage.Vertical;
                    break;
                case MsoAnimDirection.msoAnimDirectionIn:
                    direction = KeyLanguage.In;
                    break;
                case MsoAnimDirection.msoAnimDirectionOut:
                    direction = KeyLanguage.Out;
                    break;
                case MsoAnimDirection.msoAnimDirectionHorizontalIn:
                    direction = KeyLanguage.HorizontalIn;
                    break;
                case MsoAnimDirection.msoAnimDirectionHorizontalOut:
                    direction = KeyLanguage.HorizontalOut;
                    break;
                case MsoAnimDirection.msoAnimDirectionVerticalIn:
                    direction = KeyLanguage.VerticalIn;
                    break;
                case MsoAnimDirection.msoAnimDirectionVerticalOut:
                    direction = KeyLanguage.VerticalOut;
                    break;
                case MsoAnimDirection.msoAnimDirectionInCenter:
                    direction = KeyLanguage.SlideCenter;
                    break;
                default:
                    break;
            }
            return direction;
        }
        private string GetAnimationLock(pp.Effect shapeAnimation)
        {
            if (shapeAnimation?.EffectParameters?.Relative == null) return KeyLanguage.UnLocked;
            else
            {
                if (shapeAnimation.EffectParameters.Relative == office.MsoTriState.msoTrue)
                    return KeyLanguage.UnLocked;
                else if (shapeAnimation.EffectParameters.Relative == office.MsoTriState.msoFalse)
                    return KeyLanguage.Locked;
            }
            return KeyLanguage.UnLocked;
        }
        //-- Nhóm hiệu ứng Entrance
        private void GetFadeEntrAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Entrance = new FadeAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "Fade";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetGrowEntrAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Entrance = new GrowAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "Grow";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetFlyInEntrAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Entrance = new FlyAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "FlyIn";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetFloatEntrAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Entrance = new FloatAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "FloatIn";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetSplitEntrAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Entrance = new SplitAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "Split";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetWipeEntrAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Entrance = new WipeAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "Wipe";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetShapeEntrAns(EAnimation _animation, double _dur, string _effectDirecion, string _effectShapes, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Entrance = new ShapeAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "Shape";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    if (groupOptions.GroupName == "Direction")
                        foreach (var option in groupOptions.Values)
                        {
                            if (option.Name == _effectDirecion)
                                option.IsSelected = true;
                            else option.IsSelected = false;
                        }
                    else if (groupOptions.GroupName == "Shapes")
                    {
                        foreach (var option in groupOptions.Values)
                        {
                            if (option.Name == _effectShapes)
                                option.IsSelected = true;
                            else option.IsSelected = false;
                        }
                    }
                    else if (groupOptions.GroupName == "Sequence")
                    {
                        if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                        {
                            groupOptions.Values[0].IsSelected = false;
                            groupOptions.Values[1].IsSelected = true;
                        }
                        else
                        {
                            groupOptions.Values[0].IsSelected = true;
                            groupOptions.Values[1].IsSelected = false;
                        }
                    }
                }
            }
        }
        private void GetWheelEntrAns(EAnimation _animation, double _dur, string _optionName, float amount, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Entrance = new WheelAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "Wheel";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (groupOptions.GroupName == "Spokes")
                        {
                            if (amount == 1)
                            {
                                groupOptions.Values[0].IsSelected = true;
                                groupOptions.Values[1].IsSelected = false;
                                groupOptions.Values[2].IsSelected = false;
                                groupOptions.Values[3].IsSelected = false;
                                groupOptions.Values[4].IsSelected = false;
                            }
                            else
                            {
                                foreach (var spoke in groupOptions.Values)
                                {
                                    groupOptions.Values[0].IsSelected = false;
                                    if (spoke.Name == amount + " Spokes")
                                    {
                                        spoke.IsSelected = true;
                                    }
                                    else spoke.IsSelected = false;
                                }
                            }
                        }
                        else if (groupOptions.GroupName == "Sequence")
                        {
                            if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                            {
                                groupOptions.Values[0].IsSelected = false;
                                groupOptions.Values[1].IsSelected = true;
                            }
                            else
                            {
                                groupOptions.Values[0].IsSelected = true;
                                groupOptions.Values[1].IsSelected = false;
                            }
                        }
                    }
                }
            }
        }
        private void GetRandomBarsEntrAns(EAnimation _animation, double _dur, string _optionName, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Entrance = new RandomBarsAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "RandomBars";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (groupOptions.GroupName == "Direction")
                        {
                            if (option.Name == _optionName)
                                option.IsSelected = true;
                            else option.IsSelected = false;
                        }
                        else if (groupOptions.GroupName == "Sequence")
                        {
                            if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                            {
                                groupOptions.Values[0].IsSelected = false;
                                groupOptions.Values[1].IsSelected = true;
                            }
                            else
                            {
                                groupOptions.Values[0].IsSelected = true;
                                groupOptions.Values[1].IsSelected = false;
                            }
                        }
                    }
                }
            }
        }
        private void GetSpinEntrAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Entrance = new SpinAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "Spin";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetSpinGrowEntrAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Entrance = new SpinGrowAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "SpinGrow";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetGrowSpinEntrAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Entrance = new GrowSpinAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "GrowSpin";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetZoomEntrAns(EAnimation _animation, double _dur, string _optionName, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Entrance = new ZoomAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "Zoom";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (groupOptions.GroupName == "Zoom")
                        {
                            if (_optionName == KeyLanguage.In)
                            {
                                groupOptions.Values[0].IsSelected = true;
                                groupOptions.Values[1].IsSelected = false;
                            }
                            else
                            {
                                groupOptions.Values[0].IsSelected = false;
                                groupOptions.Values[1].IsSelected = true;
                            }
                        }
                        else if (groupOptions.GroupName == "Sequence")
                        {
                            if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                            {
                                groupOptions.Values[0].IsSelected = false;
                                groupOptions.Values[1].IsSelected = true;
                            }
                            else
                            {
                                groupOptions.Values[0].IsSelected = true;
                                groupOptions.Values[1].IsSelected = false;
                            }
                        }
                    }
                }
            }
        }
        private void GetSwivelEntrAns(EAnimation _animation, double _dur, string _optionName, MsoAnimateByLevel msoAnimateByLevel)
        {

            _animation.Entrance = new SwivelAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "Swivel";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    if (groupOptions.GroupName == "Sequence")
                    {
                        if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                        {
                            groupOptions.Values[0].IsSelected = false;
                            groupOptions.Values[1].IsSelected = true;
                        }
                        else
                        {
                            groupOptions.Values[0].IsSelected = true;
                            groupOptions.Values[1].IsSelected = false;
                        }
                    }
                }
            }
        }
        private void GetBounceEntrAns(EAnimation _animation, double _dur, string _optionName, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Entrance = new BounceAnimation();
            if (_animation.Entrance.EffectOptions?.Count > 0)
            {
                if (!(_animation.Entrance.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Entrance.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Entrance.AnimationType = eAnimationType.Entrance;
            _animation.Entrance.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Entrance.Name = "Bounce";
            if (_animation.Entrance.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Entrance.EffectOptions)
                {
                    if (groupOptions.GroupName == "Sequence")
                    {
                        if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                        {
                            groupOptions.Values[0].IsSelected = false;
                            groupOptions.Values[1].IsSelected = true;
                        }
                        else
                        {
                            groupOptions.Values[0].IsSelected = true;
                            groupOptions.Values[1].IsSelected = false;
                        }
                    }
                }
            }
        }
        //--- Nhóm hiệu ứng Exit
        private void GetFadeExitAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Exit = new FadeAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Entrance;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Fade";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetShrinkExitAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Exit = new FadeAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Entrance;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Shrink";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetFlyExitAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Exit = new FlyAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Entrance;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "FlyOut";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetFloatExitAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Exit = new FloatAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Exit;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "FloatOut";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetSplitExitAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Exit = new SplitAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Entrance;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Split";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetWipeExitAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Exit = new WipeAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Entrance;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Wipe";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetShapeExitAns(EAnimation _animation, double _dur, string _optionName, string name, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Exit = new ShapeAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Entrance;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Shape";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    if (groupOptions.GroupName == "Direction")
                        foreach (var option in groupOptions.Values)
                        {
                            if (option.Name == _optionName)
                                option.IsSelected = true;
                            else option.IsSelected = false;
                        }
                    else if (groupOptions.GroupName == "Shapes")
                    {
                        foreach (var option in groupOptions.Values)
                        {
                            if (option.Name == name)
                                option.IsSelected = true;
                            else option.IsSelected = false;
                        }
                    }
                    else if (groupOptions.GroupName == "Sequence")
                    {
                        if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                        {
                            groupOptions.Values[0].IsSelected = false;
                            groupOptions.Values[1].IsSelected = true;
                        }
                        else
                        {
                            groupOptions.Values[0].IsSelected = true;
                            groupOptions.Values[1].IsSelected = false;
                        }
                    }
                }
            }
        }
        private void GetWheelExitAns(EAnimation _animation, double _dur, string _optionName, float amount, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Exit = new WheelAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Entrance;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Wheel";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (groupOptions.GroupName == "Spokes")
                        {
                            if (amount == 1)
                            {
                                groupOptions.Values[0].IsSelected = true;
                                groupOptions.Values[1].IsSelected = false;
                                groupOptions.Values[2].IsSelected = false;
                                groupOptions.Values[3].IsSelected = false;
                                groupOptions.Values[4].IsSelected = false;
                            }
                            else
                            {
                                foreach (var spoke in groupOptions.Values)
                                {
                                    groupOptions.Values[0].IsSelected = false;
                                    if (spoke.Name == amount + " Spokes")
                                    {
                                        spoke.IsSelected = true;
                                    }
                                    else spoke.IsSelected = false;
                                }
                            }

                        }
                        else if (groupOptions.GroupName == "Sequence")
                        {
                            if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                            {
                                groupOptions.Values[0].IsSelected = false;
                                groupOptions.Values[1].IsSelected = true;
                            }
                            else
                            {
                                groupOptions.Values[0].IsSelected = true;
                                groupOptions.Values[1].IsSelected = false;
                            }
                        }
                    }
                }
            }
        }
        private void GetRandomBarsExitAns(EAnimation _animation, double _dur, string _optionName, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Exit = new RandomBarsAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Exit;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "RandomBars";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (groupOptions.GroupName == "Direction")
                        {
                            if (option.Name == _optionName)
                                option.IsSelected = true;
                            else option.IsSelected = false;
                        }
                        else if (groupOptions.GroupName == "Sequence")
                        {
                            if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                            {
                                groupOptions.Values[0].IsSelected = false;
                                groupOptions.Values[1].IsSelected = true;
                            }
                            else
                            {
                                groupOptions.Values[0].IsSelected = true;
                                groupOptions.Values[1].IsSelected = false;
                            }
                        }
                    }
                }
            }
        }
        private void GetSpinExitAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Exit = new SpinAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Entrance;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Spin";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetShrinkSpinExitAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Exit = new FadeAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Entrance;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Fade";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetShrinkTurnExitAns(EAnimation _animation, double _dur, string _optionName)
        {
            _animation.Exit = new SpinGrowAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Entrance;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Fade";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _optionName)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
        }
        private void GetZoomExitAns(EAnimation _animation, double _dur, string _optionName, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Exit = new ZoomAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Exit;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Zoom";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (groupOptions.GroupName == "Zoom")
                        {
                            if (_optionName == KeyLanguage.Out)
                            {
                                groupOptions.Values[0].IsSelected = true;
                                groupOptions.Values[1].IsSelected = false;
                            }
                            else
                            {
                                groupOptions.Values[0].IsSelected = false;
                                groupOptions.Values[1].IsSelected = true;
                            }
                        }
                        else if (groupOptions.GroupName == "Sequence")
                        {
                            if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                            {
                                groupOptions.Values[0].IsSelected = false;
                                groupOptions.Values[1].IsSelected = true;
                            }
                            else
                            {
                                groupOptions.Values[0].IsSelected = true;
                                groupOptions.Values[1].IsSelected = false;
                            }
                        }
                    }
                }
            }
        }
        private void GetSwivelExitAns(EAnimation _animation, double _dur, string _optionName, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Exit = new SwivelAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Exit;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Swivel";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    if (groupOptions.GroupName == "Sequence")
                    {
                        if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                        {
                            groupOptions.Values[0].IsSelected = false;
                            groupOptions.Values[1].IsSelected = true;
                        }
                        else
                        {
                            groupOptions.Values[0].IsSelected = true;
                            groupOptions.Values[1].IsSelected = false;
                        }
                    }
                }
            }
        }
        private void GetBounceExitAns(EAnimation _animation, double _dur, string _optionName, MsoAnimateByLevel msoAnimateByLevel)
        {
            _animation.Exit = new BounceAnimation();
            if (_animation.Exit.EffectOptions?.Count > 0)
            {
                if (!(_animation.Exit.EffectOptions.Last() is SequenceOption))
                {
                    _animation.Exit.EffectOptions.Add(new SequenceOption());
                }
            }
            _animation.Exit.AnimationType = eAnimationType.Exit;
            _animation.Exit.Duration = TimeSpan.FromSeconds(_dur);
            _animation.Exit.Name = "Bounce";
            if (_animation.Exit.EffectOptions != null)
            {
                foreach (var groupOptions in _animation.Exit.EffectOptions)
                {
                    if (groupOptions.GroupName == "Sequence")
                    {
                        if (msoAnimateByLevel == MsoAnimateByLevel.msoAnimateTextByFirstLevel && groupOptions.Values?.Count == 2)
                        {
                            groupOptions.Values[0].IsSelected = false;
                            groupOptions.Values[1].IsSelected = true;
                        }
                        else
                        {
                            groupOptions.Values[0].IsSelected = true;
                            groupOptions.Values[1].IsSelected = false;
                        }
                    }
                }
            }
        }
        //------Hiệu ứng theo đường path
        private void GetLinePath(EAnimation _animation, pp.Shape _shape, double _dur, string _idShape, string _direction, string _origin)
        {
            DataMotionPath data = new DataMotionPath();
            data.ZOder = 999;
            data.AnimationPath = new LinesAnimation(KeyLanguage.LinesAnimation, TimeSpan.FromSeconds(_dur));
            if (_direction == KeyLanguage.Down)
            {
                data.Width = 1;
                data.Height = _shape.Height;
                data.Top = _shape.Top + _shape.Height;
                data.Left = _shape.Left + _shape.Width / 2;
            }
            else if (_direction == KeyLanguage.Up)
            {
                data.Width = 1;
                data.Height = _shape.Height;
                data.Top = _shape.Top - _shape.Height;
                data.Left = _shape.Left + _shape.Width / 2;
            }
            else if (_direction == KeyLanguage.Left)
            {
                data.Width = _shape.Width;
                data.Height = 1;
                data.Top = _shape.Top + _shape.Height / 2;
                data.Left = _shape.Left - _shape.Width;

            }
            else if (_direction == KeyLanguage.Right)
            {
                data.Width = _shape.Width;
                data.Height = 1;
                data.Top = _shape.Top + _shape.Height / 2;
                data.Left = _shape.Left + _shape.Width;
            }
            data.ShapesBase = new LinePathInfo();
            data.XOwner = _idShape;

            if (data.AnimationPath.EffectOptions != null)
            {
                foreach (var groupOptions in data.AnimationPath.EffectOptions)
                {
                    if (groupOptions.GroupName == "Direction")
                    {
                        foreach (var option in groupOptions.Values)
                        {

                            if (option.Name == _direction)
                                option.IsSelected = true;
                            else option.IsSelected = false;
                        }
                    }
                    else if (groupOptions.GroupName == "Origin")
                    {
                        foreach (var option in groupOptions.Values)
                        {
                            if (option.Name == _origin)
                                option.IsSelected = true;
                            else option.IsSelected = false;
                        }
                    }
                }
            }
            _animation.MotionPaths.Add(data);
        }
        //,string _origin,string _speed
        private void GetArcPath(EAnimation _animation, pp.Shape _shape, double _dur, string _idShape, string _direction, string _origin)
        {
            DataMotionPath data = new DataMotionPath();
            data.IsLocked = false;
            data.ZOder = 999;
            data.AnimationPath = new ArcsAnimation(KeyLanguage.ArcsAnimation, TimeSpan.FromSeconds(_dur));
            data.Width = 200;
            data.Height = 200;
            data.Top = _shape.Top + _shape.Height / 2;
            data.Left = _shape.Left + _shape.Width / 2;
            data.ShapesBase = new ArcsPathInfo();
            data.XOwner = _idShape;
            if (data.AnimationPath.EffectOptions != null)
            {
                foreach (var groupOptions in data.AnimationPath.EffectOptions)
                {
                    if (groupOptions.GroupName == "Direction")
                    {
                        foreach (var option in groupOptions.Values)
                        {

                            if (option.Name == _direction)
                                option.IsSelected = true;
                            else option.IsSelected = false;
                        }
                    }
                    else if (groupOptions.GroupName == "Origin")
                    {
                        foreach (var option in groupOptions.Values)
                        {
                            if (option.Name == _origin)
                                option.IsSelected = true;
                            else option.IsSelected = false;
                        }
                    }
                }
            }
            _animation.MotionPaths.Add(data);
        }
        private void GetTurnsPath(EAnimation _animation, pp.Shape _shape, double _dur, string _idShape, string _origin)
        {
            DataMotionPath data = new DataMotionPath();
            data.ZOder = 999;
            data.AnimationPath = new TurnsAnimation(KeyLanguage.TurnsAnimation, TimeSpan.FromSeconds(_dur));
            data.Width = 1;
            data.Height = _shape.Height;
            data.Top = _shape.Top + _shape.Height / 2;
            data.Left = _shape.Left + _shape.Width / 2;
            data.ShapesBase = new TurnsPathInfo();
            data.XOwner = _idShape;
            foreach (var groupOptions in data.AnimationPath.EffectOptions)
            {
                if (groupOptions.GroupName == "Origin")
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _origin)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
            _animation.MotionPaths.Add(data);
        }
        //---Shapes
        private void GetCirclePath(EAnimation _animation, pp.Shape _shape, double _dur, string _idShape, string _origin)
        {
            DataMotionPath data = new DataMotionPath();
            data.ZOder = 999;
            data.AnimationPath = new CircleAnimation(KeyLanguage.CircleAnimation, TimeSpan.FromSeconds(_dur));
            data.Width = _shape.Width;
            data.Height = _shape.Height;
            data.Top = _shape.Top + _shape.Height / 2;
            data.Left = _shape.Left;
            data.ShapesBase = new CircleInfo();
            data.XOwner = _idShape;
            foreach (var groupOptions in data.AnimationPath.EffectOptions)
            {
                if (groupOptions.GroupName == "Origin")
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _origin)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
            _animation.MotionPaths.Add(data);
        }
        private void GetSquarePath(EAnimation _animation, pp.Shape _shape, double _dur, string _idShape, string _origin)
        {
            DataMotionPath data = new DataMotionPath();
            data.ZOder = 999;
            data.AnimationPath = new SquareAnimation(KeyLanguage.SquareAnimation, TimeSpan.FromSeconds(_dur));
            data.Width = _shape.Width;
            data.Height = _shape.Height;
            data.Top = _shape.Top + _shape.Height / 2;
            data.Left = _shape.Left;
            data.ShapesBase = new SquareInfo();
            data.XOwner = _idShape;
            foreach (var groupOptions in data.AnimationPath.EffectOptions)
            {
                if (groupOptions.GroupName == "Origin")
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _origin)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
            _animation.MotionPaths.Add(data);
        }
        private void GetEqualTrianglePath(EAnimation _animation, pp.Shape _shape, double _dur, string _idShape, string _origin)
        {
            DataMotionPath data = new DataMotionPath();
            data.ZOder = 999;
            data.AnimationPath = new EqualTriangleAnimation(KeyLanguage.EqualTriangleAnimation, TimeSpan.FromSeconds(_dur));
            data.Width = _shape.Width;
            data.Height = _shape.Height;
            data.Top = _shape.Top + _shape.Height / 2;
            data.Left = _shape.Left;
            data.ShapesBase = new EqualTrianglePathInfo();
            data.XOwner = _idShape;
            foreach (var groupOptions in data.AnimationPath.EffectOptions)
            {
                if (groupOptions.GroupName == "Origin")
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _origin)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
            _animation.MotionPaths.Add(data);
        }
        private void GetTrapeZoidPath(EAnimation _animation, pp.Shape _shape, double _dur, string _idShape,string _origin)
        {
            DataMotionPath data = new DataMotionPath();
            data.ZOder = 999;
            data.AnimationPath = new TrapezoidAnimation(KeyLanguage.TrapezoidAnimation, TimeSpan.FromSeconds(_dur));
            data.Width = _shape.Width;
            data.Height = _shape.Height;
            data.Top = _shape.Top + _shape.Height / 2;
            data.Left = _shape.Left;
            data.ShapesBase = new TrapesoidPathInfo();
            data.XOwner = _idShape;
            foreach (var groupOptions in data.AnimationPath.EffectOptions)
            {
                if (groupOptions.GroupName == "Origin")
                {
                    foreach (var option in groupOptions.Values)
                    {
                        if (option.Name == _origin)
                            option.IsSelected = true;
                        else option.IsSelected = false;
                    }
                }
            }
            _animation.MotionPaths.Add(data);
        }
        #endregion

        #region GetTransition
        /// <summary>
        /// hàm Lấy hiệu ứng chuyển trang
        /// </summary>
        /// <param name="_slide"></param>
        /// <param name="_page"></param>
        public void GetTransition(pp.Slide _slide, NormalPage _page)
        {
            TimeSpan duration = new TimeSpan();
            switch (_slide.SlideShowTransition.Speed)
            {
                case Microsoft.Office.Interop.PowerPoint.PpTransitionSpeed.ppTransitionSpeedFast:
                    duration = TimeSpan.FromSeconds(1);
                    break;
                case Microsoft.Office.Interop.PowerPoint.PpTransitionSpeed.ppTransitionSpeedMedium:
                    duration = TimeSpan.FromSeconds(2);
                    break;
                case Microsoft.Office.Interop.PowerPoint.PpTransitionSpeed.ppTransitionSpeedMixed:
                    duration = TimeSpan.FromSeconds(2);
                    break;
                case Microsoft.Office.Interop.PowerPoint.PpTransitionSpeed.ppTransitionSpeedSlow:
                    duration = TimeSpan.FromSeconds(3);
                    break;
            }

            //Những kiểu trans không có trong enum ppEntryEffect
            if ((int)_slide.SlideShowTransition.EntryEffect == 3888)
            {
                GetZoomTrans(_page, duration, KeyLanguage.In);
                return;
            }
            else if ((int)_slide.SlideShowTransition.EntryEffect == 3889)
            {
                GetZoomTrans(_page, duration, KeyLanguage.Out);
                return;
            }
            else if ((int)_slide.SlideShowTransition.EntryEffect == 3909)
            {
                GetNewsFlashTrans(_page, duration);
                return;
            }
            else
            {
                switch (_slide.SlideShowTransition.EntryEffect)
                {
                    case PpEntryEffect.ppEffectNone:
                        GetNoneTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectBlindsHorizontal:
                        GetBlindsTrans(_page, duration, KeyLanguage.FromHorizontal);
                        return;
                    case PpEntryEffect.ppEffectBlindsVertical:
                        GetBlindsTrans(_page, duration, KeyLanguage.FromVertical);
                        return;
                    case PpEntryEffect.ppEffectCheckerboardAcross:
                        GetCheckerboardTrans(_page, duration, KeyLanguage.FromLeft);
                        return;
                    case PpEntryEffect.ppEffectCheckerboardDown:
                        GetCheckerboardTrans(_page, duration, KeyLanguage.FromTop);
                        return;
                    case PpEntryEffect.ppEffectCoverLeft:
                        GetCoverTrans(_page, duration, KeyLanguage.FromRight);
                        return;
                    case PpEntryEffect.ppEffectCoverUp:
                        GetCoverTrans(_page, duration, KeyLanguage.FromBottom);
                        return;
                    case PpEntryEffect.ppEffectCoverRight:
                        GetCoverTrans(_page, duration, KeyLanguage.FromLeft);
                        return;
                    case PpEntryEffect.ppEffectCoverDown:
                        GetCoverTrans(_page, duration, KeyLanguage.FromTop);
                        return;
                    case PpEntryEffect.ppEffectCoverLeftUp:
                        GetCoverTrans(_page, duration, KeyLanguage.FromBottomRight);
                        return;
                    case PpEntryEffect.ppEffectCoverRightUp:
                        GetCoverTrans(_page, duration, KeyLanguage.FromBottomLeft);
                        return;
                    case PpEntryEffect.ppEffectCoverLeftDown:
                        GetCoverTrans(_page, duration, KeyLanguage.FromTopRight);
                        return;
                    case PpEntryEffect.ppEffectCoverRightDown:
                        GetCoverTrans(_page, duration, KeyLanguage.FromTopLeft);
                        return;
                    case PpEntryEffect.ppEffectDissolve:
                        GetDissolveTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectFade:
                        GetFadeTrans(_page, duration, KeyLanguage.ThroughBack);
                        return;
                    case PpEntryEffect.ppEffectUncoverLeft:
                        GetUnCoverTrans(_page, duration, KeyLanguage.FromRight);
                        return;
                    case PpEntryEffect.ppEffectUncoverUp:
                        GetUnCoverTrans(_page, duration, KeyLanguage.FromBottom);
                        return;
                    case PpEntryEffect.ppEffectUncoverRight:
                        GetUnCoverTrans(_page, duration, KeyLanguage.FromLeft);
                        return;
                    case PpEntryEffect.ppEffectUncoverDown:
                        GetUnCoverTrans(_page, duration, KeyLanguage.FromTop);
                        return;
                    case PpEntryEffect.ppEffectUncoverLeftUp:
                        GetUnCoverTrans(_page, duration, KeyLanguage.FromBottomRight);
                        return;
                    case PpEntryEffect.ppEffectUncoverRightUp:
                        GetUnCoverTrans(_page, duration, KeyLanguage.FromBottomLeft);
                        return;
                    case PpEntryEffect.ppEffectUncoverLeftDown:
                        GetUnCoverTrans(_page, duration, KeyLanguage.FromTopRight);
                        return;
                    case PpEntryEffect.ppEffectUncoverRightDown:
                        GetUnCoverTrans(_page, duration, KeyLanguage.FromTopLeft);
                        return;
                    case PpEntryEffect.ppEffectRandomBarsHorizontal:
                        GetRandomBarsTrans(_page, duration, KeyLanguage.Horizontal);
                        return;
                    case PpEntryEffect.ppEffectRandomBarsVertical:
                        GetRandomBarsTrans(_page, duration, KeyLanguage.Vertical);
                        return;
                    case PpEntryEffect.ppEffectBoxOut:
                        GetOutTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectBoxIn:
                        GetInTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectSplitHorizontalOut:
                        GetSplitTrans(_page, duration, KeyLanguage.HorizontalOut);
                        return;
                    case PpEntryEffect.ppEffectSplitHorizontalIn:
                        GetSplitTrans(_page, duration, KeyLanguage.HorizontalIn);
                        return;
                    case PpEntryEffect.ppEffectSplitVerticalOut:
                        GetSplitTrans(_page, duration, KeyLanguage.VerticalOut);
                        return;
                    case PpEntryEffect.ppEffectSplitVerticalIn:
                        GetSplitTrans(_page, duration, KeyLanguage.VerticalIn);
                        return;
                    case PpEntryEffect.ppEffectCircleOut:
                        GetCircleTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectDiamondOut:
                        GetDiamondTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectFadeSmoothly:
                        GetFadeTrans(_page, duration, KeyLanguage.Smoothly);
                        return;
                    case PpEntryEffect.ppEffectPlusOut:
                        GetPlusTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectPushDown:
                        GetPushTrans(_page, duration, KeyLanguage.FromTop);
                        return;
                    case PpEntryEffect.ppEffectPushLeft:
                        GetPushTrans(_page, duration, KeyLanguage.FromRight);
                        return;
                    case PpEntryEffect.ppEffectPushRight:
                        GetPushTrans(_page, duration, KeyLanguage.FromLeft);
                        return;
                    case PpEntryEffect.ppEffectPushUp:
                        GetPushTrans(_page, duration, KeyLanguage.FromBottom);
                        return;
                    case PpEntryEffect.ppEffectWheel1Spoke:
                        GetClockTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectWheel2Spokes:
                        GetClockTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectWheel3Spokes:
                        GetClockTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectWheel4Spokes:
                        GetClockTrans(_page, duration);
                        return;
                    case PpEntryEffect.ppEffectWheel8Spokes:
                        GetClockTrans(_page, duration);
                        return;
                }
            }
            ////---Mặc định nếu hiệu ứng Elearning không có sẽ quy về hiệu ứng push
            GetPushTrans(_page, duration, KeyLanguage.FromTop);
            return;
        }
        #endregion

        #region Danh sách hàm tạo hiệu ứng trang
        private void GetNoneTrans(NormalPage _page, TimeSpan _dur)
        {
            _page.Transition = new NoneTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = KeyLanguage.None;
        }
        private void GetFadeTrans(NormalPage _page, TimeSpan _dur, string _option)
        {
            _page.Transition = new FadeTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Fade";
            if (_page.Transition.EffectOptions?.Count > 0)
            {
                foreach (var item in _page.Transition.EffectOptions[0].Values)
                {
                    if (item.Name == _option)

                        item.IsSelected = true;
                    else
                        item.IsSelected = false;

                }
            }
        }
        private void GetPushTrans(NormalPage _page, TimeSpan _dur, string _option)
        {
            _page.Transition = new PushTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Push";
            if (_page.Transition.EffectOptions?.Count > 0)
            {
                foreach (var item in _page.Transition.EffectOptions[0].Values)
                {
                    if (item.Name == _option)

                        item.IsSelected = true;
                    else
                        item.IsSelected = false;

                }
            }
        }
        private void GetRandomBarsTrans(NormalPage _page, TimeSpan _dur, string _option)
        {
            _page.Transition = new RandomBarsTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Random Bars";
            _page.Transition.GroupName = "Sublte";
            if (_page.Transition.EffectOptions?.Count > 0)
            {
                foreach (var item in _page.Transition.EffectOptions[0].Values)
                {
                    if (item.Name == _option)

                        item.IsSelected = true;
                    else
                        item.IsSelected = false;

                }
            }
        }
        private void GetSplitTrans(NormalPage _page, TimeSpan _dur, string _option)
        {
            _page.Transition = new SplitTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Split";
            if (_page.Transition.EffectOptions?.Count > 0)
            {
                foreach (var item in _page.Transition.EffectOptions[0].Values)
                {
                    if (item.Name == _option)

                        item.IsSelected = true;
                    else
                        item.IsSelected = false;

                }
            }
        }
        private void GetCircleTrans(NormalPage _page, TimeSpan _dur)
        {
            _page.Transition = new CircleTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Circle";
        }
        private void GetDiamondTrans(NormalPage _page, TimeSpan _dur)
        {
            _page.Transition = new DiamondTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Diamond";
        }
        private void GetPlusTrans(NormalPage _page, TimeSpan _dur)
        {
            _page.Transition = new PlusTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Plus";
        }
        private void GetInTrans(NormalPage _page, TimeSpan _dur)
        {
            _page.Transition = new InTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "In";
        }
        private void GetOutTrans(NormalPage _page, TimeSpan _dur)
        {
            _page.Transition = new OutTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Out";
        }
        private void GetUnCoverTrans(NormalPage _page, TimeSpan _dur, string _option)
        {
            _page.Transition = new UnCoverTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "UnCover";
            if (_page.Transition.EffectOptions?.Count > 0)
            {
                foreach (var item in _page.Transition.EffectOptions[0].Values)
                {
                    if (item.Name == _option)

                        item.IsSelected = true;
                    else
                        item.IsSelected = false;

                }
            }
        }
        private void GetCoverTrans(NormalPage _page, TimeSpan _dur, string _option)
        {
            _page.Transition = new CoverTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Cover";
            if (_page.Transition.EffectOptions?.Count > 0)
            {
                foreach (var item in _page.Transition.EffectOptions[0].Values)
                {
                    if (item.Name == _option)

                        item.IsSelected = true;
                    else
                        item.IsSelected = false;

                }
            }
        }
        private void GetNewsFlashTrans(NormalPage _page, TimeSpan _dur)
        {
            _page.Transition = new NewsFlashTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = KeyLanguage.Newsflash;
        }
        private void GetDissolveTrans(NormalPage _page, TimeSpan _dur)
        {
            _page.Transition = new DissolveTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Dissolve";
        }
        private void GetCheckerboardTrans(NormalPage _page, TimeSpan _dur, string _option)
        {
            _page.Transition = new CheckerboardTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Checkerboard";
            if (_page.Transition.EffectOptions?.Count > 0)
            {
                foreach (var item in _page.Transition.EffectOptions[0].Values)
                {
                    if (item.Name == _option)

                        item.IsSelected = true;
                    else
                        item.IsSelected = false;

                }
            }
        }
        private void GetBlindsTrans(NormalPage _page, TimeSpan _dur, string _option)
        {
            _page.Transition = new BlindsTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Blinds";
            if (_page.Transition.EffectOptions?.Count > 0)
            {
                foreach (var item in _page.Transition.EffectOptions[0].Values)
                {
                    if (item.Name == _option)

                        item.IsSelected = true;
                    else
                        item.IsSelected = false;

                }
            }
        }
        private void GetClockTrans(NormalPage _page, TimeSpan _dur)
        {
            _page.Transition = new ClockTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Clock";
        }
        private void GetZoomTrans(NormalPage _page, TimeSpan _dur, string _option)
        {
            _page.Transition = new ZoomTransition();
            _page.Transition.Duration = _dur;
            _page.Transition.ID = _page.ID;
            _page.Transition.Name = "Zoom";
            if (_page.Transition.EffectOptions?.Count > 0)
            {
                foreach (var item in _page.Transition.EffectOptions[0].Values)
                {
                    if (item.Name == _option)

                        item.IsSelected = true;
                    else
                        item.IsSelected = false;

                }
            }
        }
        #endregion
    }

}
