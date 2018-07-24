using DocumentFormat.OpenXml.Packaging;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Trigger;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using pp = Microsoft.Office.Interop.PowerPoint;


namespace INV.Elearning.ImportPowerPoint.Helper
{
    public partial class HelperClass
    {

        public NormalPage GetNormalPage(pp.Slide slide, SlidePart slidePart)
        {
            NormalPage normalPage = new NormalPage();
            normalPage.CanShowInMenu = true;
            normalPage.PageConfig = new PageConfig();
            normalPage.PageConfig.PreviousButtonEnable = true;
            normalPage.Name = slide.Name;
            normalPage.ID = ObjectElementsHelper.RandomString(13);
            normalPage.MainLayer = new PageLayer();
            normalPage.MainLayer.ID = ObjectElementsHelper.RandomString(13);
            normalPage.MainLayer.ThemeLayoutOwnerID = ObjectElementsHelper.RandomString(13);
            normalPage.MainLayer.Background = GetFillColor(slide.Background.Fill, slidePart.Slide.CommonSlideData?.Background, slidePart);
            foreach (pp.Shape shape in slide.Shapes)
            {
                ObjectElementBase element = GetShape(shape, shape.Type, slide.TimeLine, slidePart);
                foreach (Animations.DataMotionPath motionPath in element.Animations.MotionPaths)
                {
                    string pathId = motionPath.ID;
                    TriggerDataObject trigger = new TriggerDataObject()
                    {
                        Type = ETriggerType.Layer,
                        Action = EAction.Move,
                        Source = element.ID,
                        SourceDetail = new SourceDetailData()
                        {
                            ObjectId = pathId
                        },
                        Event = EEvent.TimelineStarts,
                        Target = normalPage.MainLayer.ID
                    };
                    normalPage.MainLayer.TriggerData.Add(trigger);
                }
                normalPage.MainLayer.Children.Add(element);
            }
            normalPage.ID = ObjectElementsHelper.RandomString(13);

            GetTransition(slide, normalPage);
            return normalPage;
        }
    }
}
