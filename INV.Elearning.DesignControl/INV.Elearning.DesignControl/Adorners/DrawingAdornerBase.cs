using INV.Elearning.Core.Model.Theme;
using INV.Elearning.DesignControl.Views;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Shapes;

namespace INV.Elearning.DesignControl.Adorners
{
    public class DrawingAdornerBase : Adorner
    {
        public DrawingAdornerBase(UIElement adornedElement) : base(adornedElement)
        {
            Fill = (Brush)(new BrushConverter().ConvertFrom("#4472C4"));
            Stroke = (Brush)(new BrushConverter().ConvertFrom("#2F528F"));
            StrokeThissness = 2.5f;
        }
        #region Property
        private Brush _fill;
        // màu hình vẽ
        public Brush Fill
        {
            get { return _fill; }
            set { _fill = value; }
        }

        private Brush _stroke;
        // màu viền
        public Brush Stroke
        {
            get { return _stroke; }
            set { _stroke = value; }
        }

        private double _strokeThissness;

        public double StrokeThissness
        {
            get { return _strokeThissness; }
            set { _strokeThissness = value; }
        }


        #endregion
        #region Events
        public delegate void DrawingBaseEventHanler(object sender, PlaceHolderEventArg e);

        /// <summary>
        /// Sự kiện thay đổi vị trí
        /// </summary>
        public static readonly RoutedEvent CompletedEvent = EventManager.RegisterRoutedEvent(
            "ActionCompleted",
            RoutingStrategy.Bubble,
            typeof(DrawingBaseEventHanler),
            typeof(DrawingAdornerBase));

        /// <summary>
        /// Sự kiện thay đổi vị trí
        /// </summary>
        public event DrawingBaseEventHanler ActionCompleted
        {
            add
            {
                base.AddHandler(DrawingAdornerBase.CompletedEvent, value);
            }
            remove
            {
                base.RemoveHandler(DrawingAdornerBase.CompletedEvent, value);
            }
        }
        #endregion

    }

    public class PlaceHolderEventArg : RoutedEventArgs
    {
        /// <summary>
        /// Hàm khởi tạo mặc định
        /// </summary>
        public PlaceHolderEventArg()
        {
        }

        /// <summary>
        /// Hàm khởi tạo 1 tham số
        /// </summary>
        /// <param name="routedEvent"></param>
        public PlaceHolderEventArg(RoutedEvent routedEvent)
        {
            this.RoutedEvent = routedEvent;
        }

        private Rect _position;
        // chứa thông tim L,T,W,H
        public Rect Position
        {
            get { return _position; }
            set { _position = value; }
        }

        private Rectangle rectangle;

        public Rectangle Rectangle
        {
            get { return rectangle; }
            set { rectangle = value; }
        }

        private PlaceHolderType _placeHolderType;
        /// <summary>
        /// Kiểu placeholder
        /// </summary>
        public PlaceHolderType PlaceHolderType
        {
            get { return _placeHolderType; }
            set { _placeHolderType = value; }
        }




    }
}
