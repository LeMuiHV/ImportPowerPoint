using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using pp = Microsoft.Office.Interop.PowerPoint;
using office = Microsoft.Office.Core;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using draw = DocumentFormat.OpenXml.Drawing;

namespace INV.Elearning.ImportPowerPoint.Helper
{
    public class Utils
    {
        /// <summary>
        /// Path thư mục chứa các file tạm của module
        /// </summary>
        /// (Application.Current as IAppGlobal).DocumentTempFolder
        public static string InvicoPath
        {
            get
            {
                string AppPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Invico", "ImportPP");
                if (!Directory.Exists(AppPath))
                    Directory.CreateDirectory(AppPath);
                return AppPath;
            }
        }
        public static pp.Slides LstSlide;
        public static List<pp.Slide> LstSlideSelected = new List<pp.Slide>();
        public static PresentationPart presentationPart;
        public static SlidePart slidePart;
        public static SlideMasterPart themePart;
        public static void GetPresentationPart(PresentationDocument _presentationDocument)
        {
            presentationPart = _presentationDocument.PresentationPart;
        }
        public static void GetSlidePart(int _slideIndex)
        {
            if (presentationPart != null && presentationPart.Presentation != null)
            {
                Presentation presentation = presentationPart.Presentation;
                if (presentation.SlideIdList != null)
                {
                    DocumentFormat.OpenXml.OpenXmlElementList slideIds = presentation.SlideIdList.ChildElements;

                    // If the slide ID is in range...
                    if (_slideIndex < slideIds.Count)
                    {
                        // Get the relationship ID of the slide.
                        string slidePartRelationshipId = (slideIds[_slideIndex] as SlideId).RelationshipId;

                        // Get the specified slide part from the relationship ID.
                        slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);
                    }
                }
            }
        }

        public static void GetThemePart(int _slideIndex)
        {
            if (presentationPart != null && presentationPart.Presentation != null)
            {
                Presentation presentation = presentationPart.Presentation;
                if (presentation.SlideMasterIdList != null)
                {
                    DocumentFormat.OpenXml.OpenXmlElementList slideIds = presentation.SlideMasterIdList.ChildElements;

                    // If the slide ID is in range...
                    if (_slideIndex < slideIds.Count)
                    {
                        // Get the relationship ID of the slide.
                        string slidePartRelationshipId = (slideIds[_slideIndex] as SlideMasterId).RelationshipId;

                        // Get the specified slide part from the relationship ID.
                        themePart = (SlideMasterPart)presentationPart.GetPartById(slidePartRelationshipId);
                    }
                }
            }
        }
        public static void CopyStream(Stream stream, string destPath)
        {
            using (var fileStream = new FileStream(destPath, FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(fileStream);
            }
        }
        public static double tlw = 0;
        public static double tlh = 0;
        public static Slide GetSlide(string _presentationFile, int _slideIndex)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(_presentationFile, false))
            {
                return GetDetailSlide(presentationDocument, _slideIndex);
            }
        }
        public static Slide GetDetailSlide(PresentationDocument _presentationDocument, int _slideIndex)
        {
            // Get the presentation part of the presentation document.
            presentationPart = _presentationDocument.PresentationPart;
            if (presentationPart != null && presentationPart.Presentation != null)
            {
                Presentation presentation = presentationPart.Presentation;
                if (presentation.SlideIdList != null)
                {
                    DocumentFormat.OpenXml.OpenXmlElementList slideIds = presentation.SlideIdList.ChildElements;

                    // If the slide ID is in range...
                    if (_slideIndex < slideIds.Count)
                    {
                        // Get the relationship ID of the slide.
                        string slidePartRelationshipId = (slideIds[_slideIndex] as SlideId).RelationshipId;

                        // Get the specified slide part from the relationship ID.
                        slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);
                        return slidePart.Slide;
                    }
                }
            }
            return null;
        }
    }
}
