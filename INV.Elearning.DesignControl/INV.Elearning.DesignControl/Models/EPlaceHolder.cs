using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.DesignControl.Views;
using INV.Elearning.Text.Models;
using INV.Elearning.Text.ViewModels.Text;
using System;
using System.Xml.Serialization;

namespace INV.Elearning.DesignControl.Models
{
    [Serializable]
    [XmlRoot("placeholder")]

    public class EPlaceHolder : EStandardElement
    {

        private PlaceHolderType _type;
        [XmlElement("type")]
        public PlaceHolderType Type
        {
            get { return _type; }
            set { _type = value; OnPropertyChanged("Type"); }
        }

        private bool _isMasterSlide;
        [XmlAttribute("ismaster")]
        public bool IsMasterSlide
        {
            get { return _isMasterSlide; }
            set { _isMasterSlide = value; OnPropertyChanged("IsMasterSlide"); }
        }

        private DataDocument _document;
        [XmlElement("doc")]
        public DataDocument Document
        {
            get { return _document; }
            set { _document = value; OnPropertyChanged("Document"); }
        }


        public override IElearningElement Clone()
        {
            EPlaceHolder ePlaceHolder = new EPlaceHolder();
            base.Copy(ePlaceHolder);
            ePlaceHolder.Type = Type;
            ePlaceHolder.IsMasterSlide = IsMasterSlide;
            ePlaceHolder.Document = Document;
            return ePlaceHolder;
        }
    }
}
