using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace INV.Elearning.ImportPowerPoint.Helper
{
    public partial class HelperClass
    {
        /// <summary>
        /// GetSlideMasterPart
        /// </summary>
        /// <param name="_slideIndex"></param>
        /// <returns></returns>
        public SlideMasterPart GetSlideMasterPart(int _slideIndex)
        {
            if (Utils.presentationPart != null && Utils.presentationPart.Presentation != null)
            {
                Presentation presentation = Utils.presentationPart.Presentation;
                if (presentation.SlideMasterIdList != null)
                {
                    DocumentFormat.OpenXml.OpenXmlElementList slideIds = presentation.SlideMasterIdList.ChildElements;

                    // If the slide ID is in range...
                    if (_slideIndex < slideIds.Count)
                    {
                        // Get the relationship ID of the slide.
                        string slidePartRelationshipId = (slideIds[_slideIndex] as SlideMasterId).RelationshipId;

                        // Get the specified slide part from the relationship ID.
                        return (SlideMasterPart)Utils.presentationPart.GetPartById(slidePartRelationshipId);
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Hàm lấy SlidePart
        /// </summary>
        /// <param name="_slideIndex"></param>
        /// <returns></returns>
        public SlidePart GetSlidePart(int _slideIndex)
        {
            if (Utils.presentationPart != null && Utils.presentationPart.Presentation != null)
            {
                Presentation presentation = Utils.presentationPart.Presentation;
                if (presentation.SlideIdList != null)
                {
                    DocumentFormat.OpenXml.OpenXmlElementList slideIds = presentation.SlideIdList.ChildElements;

                    // If the slide ID is in range...
                    if (_slideIndex < slideIds.Count)
                    {
                        // Get the relationship ID of the slide.
                        string slidePartRelationshipId = (slideIds[_slideIndex] as SlideId).RelationshipId;

                        // Get the specified slide part from the relationship ID.
                        return (SlidePart)Utils.presentationPart.GetPartById(slidePartRelationshipId);
                    }
                }
            }
            return null;
        }

        public SlideLayoutPart GetSlideLayoutPart(int _slideLayoutIndex,SlideMasterPart slideMasterPart)
        {
            if (slideMasterPart != null && slideMasterPart.SlideMaster != null)
            {
                SlideMaster slideMaster = slideMasterPart.SlideMaster;
                if (slideMaster.SlideLayoutIdList != null)
                {
                    DocumentFormat.OpenXml.OpenXmlElementList slideIds = slideMaster.SlideLayoutIdList.ChildElements;
                    // If the slide ID is in range...
                    if (_slideLayoutIndex < slideIds.Count)
                    {
                        // Get the relationship ID of the slide.
                        string slidePartRelationshipId = (slideIds[_slideLayoutIndex] as SlideLayoutId).RelationshipId;
                        return (SlideLayoutPart)slideMasterPart.GetPartById(slidePartRelationshipId);
                        // Get the specified slide part from the relationship ID.
                    }
                }
            }
            return null;
        }
    }
}
