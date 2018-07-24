using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.Core.ViewModel;
using INV.Elearning.DesignControl.Models;
using INV.Elearning.DesignControl.ViewModel;
using INV.Elearning.Text.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace INV.Elearning.DesignControl.UndoRedo
{
    public class UpdateLayoutMasterStep : StepBase
    {
        //public SlideBase SlideBase { get; set; }
        //public ELayoutMaster NewLayout { get; set; }
        //public PageElementBase Data { get; set; }
        //public EThemes EThemes { get; set; }
        //public LayoutType OldType { get; set; }
        //public string OldId { get; set; }
        //public PageLayer ThemeData { get; set; }
        //public UpdateLayoutMasterStep(PageElementBase data, ELayoutMaster newLayout, SlideBase slideBase, EThemes ethemes, LayoutType oldType, string oldID, PageLayer themeData)
        //{
        //    SlideBase = slideBase;
        //    NewLayout = newLayout;
        //    Data = data;
        //    EThemes = ethemes;
        //    OldType = oldType;
        //    OldId = oldID;
        //    if (themeData != null)
        //    {
        //        ThemeData = themeData;
        //    }
        //}

        //public override void UndoExcute()
        //{
        //    Global.BeginInit();
        //    //SlideBase.LoadData(Data);
        //    SlideBase.IsSelected = true;
        //    SlideBase.SelectLayoutType = OldType;
        //    SlideBase.MainLayout.ThemeLayoutOwnerID = OldId;
        //    if (SlideBase.ThemeLayout == null) SlideBase.ThemeLayout = new ThemeLayout();
        //    if (ThemeData != null)
        //    {
        //        SlideBase.ThemeLayout.UpdateUI(ThemeData);
        //    }
        //    for (int i = 0; i < SlideBase.MainLayout.Elements.Count; i++)
        //    {
        //        SlideBase.MainLayout.Elements.Remove(SlideBase.MainLayout.Elements[i]);
        //        i--;
        //    }
        //    foreach (var item in Data.MainLayer.Children)
        //    {
        //        var _objectElement = ObjectElementsHelper.LoadData(item);
        //        if (_objectElement != null)
        //        {
        //            if (_objectElement is IPlaceHolder place) place.IsMasterSlide = false;

        //            SlideBase.MainLayout.Elements.Add(_objectElement);
        //        }
        //    }
        //    foreach (ELayoutMaster item in (Application.Current as IAppGlobal).SelectedThemes.MainTheme.SlideMasters[0].LayoutMasters)
        //    {
        //        item.LayountCount = DesignTabControlViewModel.CountLayoutMaster(item);
        //    }
        //    Global.EndInit();
        //}

        //public override void RedoExcute()
        //{
        //    Global.BeginInit();
        //    SlideBase.IsSelected = true;
        //    SlideBase.SelectLayoutType = NewLayout.LayoutType;
        //    SlideBase.MainLayout.ThemeLayoutOwnerID = NewLayout.ID;
        //    if (SlideBase.ThemeLayout == null) SlideBase.ThemeLayout = new ThemeLayout();
        //    if (NewLayout.IsHideBackground)
        //    {
        //        SlideBase.ThemeLayout.UpdateUI(NewLayout.MainLayer);
        //    }
        //    else
        //    {
        //        ESlideMaster eSlideMaster = EThemes.MainTheme.SlideMasters[0];
        //        SlideBase.ThemeLayout.UpdateWithParent(NewLayout.MainLayer, eSlideMaster.MainLayer);
        //    }


        //    foreach (var item in NewLayout.MainLayer.Children)
        //    {
        //        if (item is DataDocument || item is EPlaceHolder)
        //        {
        //            ObjectElement _element = SlideBase.MainLayout.Elements.FirstOrDefault(x => x.ID == item.ID);
        //            if (_element != null)
        //            {
        //                _element.UpdateUI(item);
        //            }
        //            else
        //            {
        //                _element = ObjectElementsHelper.LoadData(item);
        //                if (_element != null)
        //                    SlideBase.MainLayout.Elements.Add(_element);
        //                if (_element is IPlaceHolder place) place.IsMasterSlide = false;

        //            }
        //        }

        //        for (int i = 0; i < SlideBase.MainLayout.Elements.Count; i++)
        //        {
        //            if (NewLayout.MainLayer.Children.FirstOrDefault(x => x.ID == SlideBase.MainLayout.Elements[i].ID) == null)
        //            {
        //                SlideBase.MainLayout.Elements.RemoveAt(i);
        //                i--;
        //            }
        //        }
        //    }
        //    foreach (ELayoutMaster item in (Application.Current as IAppGlobal).SelectedThemes.MainTheme.SlideMasters[0].LayoutMasters)
        //    {
        //        item.LayountCount = DesignTabControlViewModel.CountLayoutMaster(item);
        //    }
        //    Global.EndInit();
        //}
    }
}
