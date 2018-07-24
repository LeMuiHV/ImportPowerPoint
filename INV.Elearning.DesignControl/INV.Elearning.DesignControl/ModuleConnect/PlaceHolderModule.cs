using INV.Elearning.Controls;
using INV.Elearning.Core;
using INV.Elearning.Core.View;
using INV.Elearning.Core.ViewModel;
using INV.Elearning.DesignControl.Models;
using INV.Elearning.DesignControl.ViewModel;
using INV.Elearning.DesignControl.Views;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Windows;

namespace INV.Elearning.DesignControl.ModuleConnect
{
    /// <summary>
    /// Lớp mô tả các thông tin ghép nối của mô đun hình ảnh lên chương trình chính
    /// </summary>
    [ModuleInfo("Mô đun PlaceHolder", "Hương Việt Group", "Cung cấp, quản lý, định dạng đối tượng PlaceHolder", "1.0.0.0")]
    [Export(typeof(ModuleBase))]
    [ExportMetadata("ModuleMetadata", typeof(PlaceHolder))]
    public class PlaceHolderModule : ModuleBase
    {
        public override RibbonContextualTabGroup ContextualGroup => null;

        public override Type DataType => typeof(EPlaceHolder);

        public override Type ObjectType => typeof(PlaceHolder);

        public override ModuleInfo ModuleInfo => throw new NotImplementedException();

        public override IEnumerable<PanelItem> Panels => throw new NotImplementedException();

        public override IEnumerable<RibbonTabItem> RibbonTabItems => null;

        public override IEnumerable<System.Windows.Controls.MenuItem> MenuItems => null;

        public override string InsertGroupBoxName => "Text";

        private List<FrameworkElement> insertButtons;

        public override IEnumerable<FrameworkElement> InsertButtons
        {
            get
            {
                if (insertButtons == null)
                {
                    insertButtons = new List<FrameworkElement>();

                    Button button = new Button();
                    button.Header = "Header & Footer";
                    button.ToolTip = "Header & Footer";
                    button.Command = DesignTabControlViewModel.HeaderAndFooterCommand;
                    button.LargeIcon = @"pack://application:,,,/INV.Elearning.DesignControl;Component/Images/HeaderAndFooter32.png";
                    insertButtons.Add(button);

                    button = new Button();
                    button.Header = "Date Time";
                    button.ToolTip = "Date Time";
                    button.Command = DesignTabControlViewModel.HeaderAndFooterCommand;
                    button.LargeIcon = @"pack://application:,,,/INV.Elearning.DesignControl;Component/Images/DateTime32.png";
                    insertButtons.Add(button);

                    button = new Button();
                    button.Header = "Slide Number";
                    button.ToolTip = "Slide Number";
                    button.Command = DesignTabControlViewModel.HeaderAndFooterCommand;
                    button.LargeIcon = @"pack://application:,,,/INV.Elearning.DesignControl;Component/Images/SlideNumber32.png";
                    insertButtons.Add(button);

                }
                return insertButtons;
            }
        }
        public override string LanguagePackageUri => throw new NotImplementedException();

        public override void Dispose()
        {
            throw new NotImplementedException();
        }

        public override void Initialize()
        {
            throw new NotImplementedException();
        }

        public override ObjectElement Load(object data)
        {
            if (data is EPlaceHolder ePlaceHolder)
            {
                var newPlayer = new PlaceHolder();
                newPlayer.UpdateUI(ePlaceHolder);
                return newPlayer;
            }
            return null;
        }
    }
}
