using INV.Elearning.Core.Model.Theme;
using INV.Elearning.DesignControl.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace INV.Elearning.DesignControl.Adorners
{
    public class DrawingPlaceHolderAdorner : DrawingAdornerBase
    {
        #region Fields

        private PlaceHolderType _placeHolderType;
        /// <summary>
        /// Kiểu place holder
        /// </summary>
        public PlaceHolderType PlaceHolderType
        {
            get { return _placeHolderType; }
            set { _placeHolderType = value; }
        }


        #endregion

        Point _startPoint; // điểm bắt đầu vẽ
        Point _endPoint; // diểm kết thúc vẽ
        // top left width height 
        double _width;
        double _height;
        double _top;
        double _left;

        // giá trị sau khi nhấn Control hoặc Shilft
        double _newHeight;
        double _newWidth;
        double _newTop;
        double _newLeft;

        // Cờ kiểm tra có giữ Shift hay control ko?
        bool _isShiftCapture = false;
        bool _isControlCapture = false;

        Rectangle rectangle;

        private VisualCollection _visuals;
        private Canvas _canvas;
        #region Contructor
        public DrawingPlaceHolderAdorner(UIElement adornedElement) : base(adornedElement)
        {
            this._visuals = new VisualCollection(this);
            _canvas = new Canvas();
            _canvas.Background = Brushes.Transparent;
            this._visuals.Add(_canvas);
            this.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
            {
                this.Focusable = true;
                this.Focus();
            }));
        }
        #endregion

        #region override Methods
        /// <summary>
        /// Sự kiện mouse down lấy tọa độ bắt đầu khi vẽ
        /// </summary>
        /// <param name="e"></param>
        protected override void OnMouseDown(MouseButtonEventArgs e)
        {
            base.OnMouseDown(e);
            _startPoint = e.GetPosition(this);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            this.Focus();
            this.Cursor = Cursors.Cross;
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                if (!IsMouseCaptured)
                    CaptureMouse();
                // Lấy điểm cuối
                _endPoint = e.GetPosition(this);
                // Set Width Height
                _width = Math.Abs(_startPoint.X - _endPoint.X);
                _height = Math.Abs(_startPoint.Y - _endPoint.Y);
                // Set Top Left
                _top = Math.Min(_startPoint.Y, _endPoint.Y);
                _left = Math.Min(_startPoint.X, _endPoint.X);

                if (rectangle == null)
                {
                    rectangle = new Rectangle();
                    rectangle.Fill = Brushes.Transparent;
                    rectangle.Stroke = Brushes.Gray;
                    rectangle.StrokeThickness = 1;
                    rectangle.Width = _width;
                    rectangle.Height = _height;
                    Canvas.SetLeft(rectangle, _left);
                    Canvas.SetTop(rectangle, _top);
                    _canvas.Children.Add(rectangle);
                }
                else
                {
                    if (_isShiftCapture)
                    {
                        if (_height < _width)
                        {
                            double _max = Math.Max(_width, _height);
                            if (_endPoint.X > _startPoint.X && _endPoint.Y > _startPoint.Y)
                            {
                                _newTop = _top;
                                _newLeft = _left;
                                Canvas.SetLeft(rectangle, _left);
                                Canvas.SetTop(rectangle, _top);
                            }
                            else if (_endPoint.X < _startPoint.X && _endPoint.Y > _startPoint.Y)
                            {
                                _newTop = _top;
                                _newLeft = _left;
                                Canvas.SetLeft(rectangle, _newLeft);
                                Canvas.SetTop(rectangle, _newTop);
                            }
                            else if (_endPoint.X > _startPoint.X && _endPoint.Y < _startPoint.Y)
                            {
                                _newLeft = _left;
                                _newTop = _top - (_max - _height);
                                Canvas.SetLeft(rectangle, _newLeft);
                                Canvas.SetTop(rectangle, _newTop);
                            }
                            else if (_endPoint.X < _startPoint.X && _endPoint.Y < _startPoint.Y)
                            {
                                _newTop = _top - (_max - _height);
                                _newLeft = _left - (_max - _width);
                                Canvas.SetLeft(rectangle, _newLeft);
                                Canvas.SetTop(rectangle, _newTop);

                            }
                            _newWidth = _width;
                            _newHeight = _width;
                            rectangle.Width = _newWidth;
                            rectangle.Height = _newHeight;
                        }
                        else
                        {
                            double _max = Math.Max(_width, _height);
                            if (_endPoint.X > _startPoint.X && _endPoint.Y > _startPoint.Y)
                            {
                                _newTop = _top;
                                _newLeft = _left;
                                Canvas.SetLeft(rectangle, _left);
                                Canvas.SetTop(rectangle, _top);
                            }
                            else if (_endPoint.X < _startPoint.X && _endPoint.Y > _startPoint.Y)
                            {
                                _newTop = _top;
                                _newLeft = _left - (_max - _width);
                                Canvas.SetLeft(rectangle, _newLeft);
                                Canvas.SetTop(rectangle, _newTop);
                            }
                            _newWidth = _height;
                            _newHeight = _height;
                            rectangle.Width = _newWidth;
                            rectangle.Height = _newHeight;
                        }
                    }
                    else if (_isControlCapture)
                    {
                        if (_endPoint.X < _startPoint.X && _endPoint.Y > _startPoint.Y)
                        {
                            _newLeft = _left;
                            _newTop = _top - 0.5 * _height;
                            Canvas.SetLeft(rectangle, _newLeft);
                            Canvas.SetTop(rectangle, _newTop);
                        }
                        if (_endPoint.X < _startPoint.X && _endPoint.Y < _startPoint.Y)
                        {
                            _newLeft = _left;
                            _newTop = _top;
                            Canvas.SetLeft(rectangle, _newLeft);
                            Canvas.SetTop(rectangle, _newTop);
                        }
                        if (_endPoint.X > _startPoint.X && _endPoint.Y < _startPoint.Y)
                        {
                            _newLeft = _left - 0.5 * _width;
                            _newTop = _top - 0.5 * _height;
                            Canvas.SetLeft(rectangle, _newLeft);
                            Canvas.SetTop(rectangle, _newTop);
                        }
                        if (_endPoint.X > _startPoint.X && _endPoint.Y > _startPoint.Y)
                        {
                            _newLeft = _left - 0.5 * _width;
                            _newTop = _top - 0.5 * _height;
                            Canvas.SetLeft(rectangle, _newLeft);
                            Canvas.SetTop(rectangle, _newTop);
                        }
                        _newWidth = 1.5 * _width;
                        _newHeight = 1.5 * _height;
                        rectangle.Height = _newHeight;
                        rectangle.Width = _newWidth;
                    }
                    else
                    {
                        rectangle.Width = _width;
                        rectangle.Height = _height;
                        Canvas.SetLeft(rectangle, _left);
                        Canvas.SetTop(rectangle, _top);
                    }

                }
            }
        }

        /// <summary>
        /// Xử lý sự kiện khi thả chuột
        /// </summary>
        /// <param name="e"></param>
        protected override void OnMouseUp(MouseButtonEventArgs e)
        {
            base.OnMouseUp(e);

            if (IsMouseCaptured)
                ReleaseMouseCapture();
            _canvas.Children.Remove(rectangle);
            //Khi nhấn control hoặc shift thì truyền width height mới
            if (_isShiftCapture || _isControlCapture)
            {
                _width = _newWidth;
                _height = _newHeight;
                _top = _newTop;
                _left = _newLeft;
            }
            PlaceHolderEventArg _arg = new PlaceHolderEventArg(CompletedEvent);
            if (rectangle == null)
            {
                _arg.Position = new Rect(_startPoint.X, _startPoint.Y, 50, 50);
            }
            else
            {
                _arg.Position = new Rect(_left, _top, _width, _height);

            }
            _arg.Rectangle = rectangle;
            _arg.PlaceHolderType = PlaceHolderType;
            this.RaiseEvent(_arg);
            rectangle = null;

        }

        /// <summary>
        /// Sự kiện khi nhấn phím Shift hoặc Control
        /// </summary>
        /// <param name="e"></param>
        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            base.OnPreviewKeyDown(e);
            if (rectangle != null)
            {
                //Nếu nhấn phím Shift
                if (e.Key == Key.LeftShift && _isShiftCapture == false || e.Key == Key.RightShift && _isShiftCapture == false)
                {
                    _isShiftCapture = true;
                    //Nếu height < witth thì set height = width
                    if (_height < _width)
                    {
                        if (_endPoint.Y > _startPoint.Y)
                        {
                            _newTop = _top;
                            _newLeft = _left;
                        }
                        else
                        {
                            _newLeft = _left;
                            _newTop = _top - (_width - _height);
                            Canvas.SetTop(rectangle, _newTop);
                        }
                        _newWidth = _width;
                        _newHeight = _width;
                        rectangle.Height = _newHeight;
                    }
                    //Nếu height > width thì set height = width
                    if (_height > _width)
                    {
                        if (_endPoint.X > _startPoint.X && _endPoint.Y > _startPoint.Y)
                        {
                            _newTop = _top;
                            _newLeft = _left;
                        }
                        else if (_endPoint.X < _startPoint.X && _endPoint.Y > _startPoint.Y)
                        {
                            _newTop = _top;
                            _newLeft = _left - (_height - _width);
                            Canvas.SetLeft(rectangle, _newLeft);
                        }
                        _newWidth = _height;
                        _newHeight = _height;
                        rectangle.Width = _newWidth;
                    }
                }
                //Nếu nhấn control
                if (e.Key == Key.LeftCtrl && _isControlCapture == false || e.Key == Key.RightCtrl && _isControlCapture == false)
                {
                    _isControlCapture = true;
                    if (_endPoint.X > _startPoint.X)
                    {
                        if (_endPoint.Y > _startPoint.Y)
                        {
                            _newTop = _top - 0.5 * _height;
                            _newLeft = _left - 0.5 * _width;
                            Canvas.SetTop(rectangle, _newTop);
                            Canvas.SetLeft(rectangle, _newLeft);
                        }
                        else
                        {
                            _newLeft = _left - (0.5 * _width);
                            _newTop = _top;
                            Canvas.SetLeft(rectangle, _newLeft);
                        }
                    }
                    else
                    {
                        if (_endPoint.Y > _startPoint.Y)
                        {
                            _newTop = _top - (0.5 * _height);
                            _newLeft = _left;
                            Canvas.SetTop(rectangle, _newTop);
                        }
                    }
                    _newWidth = 1.5 * _width;
                    _newHeight = 1.5 * _height;
                    rectangle.Width = _newWidth;
                    rectangle.Height = _newHeight;
                }
            }
        }

        /// <summary>
        /// Sự kiện khi nhả phím
        /// </summary>
        /// <param name="e"></param>
        protected override void OnKeyUp(KeyEventArgs e)
        {
            base.OnKeyUp(e);
            if (rectangle != null)
            {
                rectangle.Width = _width;
                rectangle.Height = _height;
                Canvas.SetTop(rectangle, _top);
                Canvas.SetLeft(rectangle, _left);
            }
            _isShiftCapture = false;
            _isControlCapture = false;
        }
        #endregion

        protected override int VisualChildrenCount
        {
            get
            {
                return this._visuals.Count;
            }
        }

        protected override Size ArrangeOverride(Size arrangeBounds)
        {
            this._canvas.Arrange(new Rect(arrangeBounds));
            return arrangeBounds;
        }

        protected override Visual GetVisualChild(int index)
        {
            return this._visuals[index];
        }
    }
}
