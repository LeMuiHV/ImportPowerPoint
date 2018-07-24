using INV.Elearning.Core;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.Model.Theme;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.Core.ViewModel;
using INV.Elearning.DesignControl.Helper;
using INV.Elearning.DesignControl.ViewModel;
using INV.Elearning.Text.Models;
using INV.Elearning.Text.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace INV.Elearning.DesignControl.UndoRedo
{
    public class SlideMasterOfficeStep : StepBase
    {
    //    public ETheme OldETheme { get; set; }
    //    public ETheme NewETheme { get; set; }

    //    public SlideMasterOfficeStep(ETheme oldETheme, ETheme newETheme)
    //    {
    //        OldETheme = oldETheme;
    //        NewETheme = newETheme;
    //    }

    //    public override void UndoExcute()
    //    {
    //        Global.BeginInit();
    //        //(Application.Current as IAppGlobal).SelectedThemes = OldEThemes;
    //        SlideMaster slide = null;
    //        if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster)
    //        {
    //            slide = (Application.Current as IAppGlobal).SelectedSlide as SlideMaster;
    //        }
    //        else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster)
    //        {
    //            slide = ((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).SlideParent;
    //        }
    //        ChangeThemeBackground(OldETheme);
    //        Global.BeginInit();
    //        slide.MainLayout.Fill = new ColorSolidBrush() { Color = Colors.Transparent, ColorSpecialName = "Transparent" };
    //        Global.EndInit();
    //        foreach (EColorManagment _color in (Application.Current as IAppGlobal).OfficeColors)
    //        {
    //            if (_color.Name == OldETheme.Colors.Name)
    //            {
    //                (Application.Current as IAppGlobal).SelectedTheme.Colors = _color;
    //                (Application.Current as IAppGlobal).SelectedThemeView.Colors = _color;
    //            }
    //        }
    //        Global.EndInit();
    //    }

    //    public override void RedoExcute()
    //    {
    //        Global.BeginInit();
    //        //(Application.Current as IAppGlobal).SelectedThemes = NewEThemes;
    //        SlideMaster slide = null;
    //        if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster)
    //        {
    //            slide = (Application.Current as IAppGlobal).SelectedSlide as SlideMaster;
    //        }
    //        else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster)
    //        {
    //            slide = ((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).SlideParent;
    //        }
    //        ChangeThemeBackground(NewETheme);
    //        slide.MainLayout.Fill = new ColorSolidBrush() { Color = Colors.Transparent, ColorSpecialName = "Transparent" };
    //        foreach (EColorManagment _color in (Application.Current as IAppGlobal).OfficeColors)
    //        {
    //            if (_color.Name == NewETheme.Colors.Name)
    //            {
    //                (Application.Current as IAppGlobal).SelectedTheme.Colors = _color;
    //                (Application.Current as IAppGlobal).SelectedThemeView.Colors = _color;
    //            }
    //        }

    //        Global.EndInit();
    //    }

    //    /// <summary>
    //    /// Thay đổi background khi lựa chọn theme
    //    /// </summary>
    //    /// <param name="eTheme"></param>
    //    public void ChangeThemeBackground(ETheme eTheme)
    //    {
    //        Global.BeginInit();
    //        SlideMaster slide = null;
    //        if ((Application.Current as IAppGlobal).SelectedSlide is SlideMaster)
    //        {
    //            slide = (Application.Current as IAppGlobal).SelectedSlide as SlideMaster;
    //        }
    //        else if ((Application.Current as IAppGlobal).SelectedSlide is LayoutMaster)
    //        {
    //            slide = ((Application.Current as IAppGlobal).SelectedSlide as LayoutMaster).SlideParent;
    //        }

    //        RemoveBackgroundSlideMasterShapes(slide);

    //        ObservableCollection<StandardElement> StandardElements = new ObservableCollection<StandardElement>();
    //        foreach (var item in eTheme.SlideMasters[0].MainLayer.Children)
    //        {
    //            if (item is EStandardElement eThemeShape && eThemeShape.IsGraphicBackground)
    //            {
    //                StandardElement StandardElement = new StandardElement();
    //                StandardElement.UpdateUI(eThemeShape);
    //                StandardElements.Add(StandardElement);
    //            }
    //        }

    //        Global.BeginInit();
    //        slide.MainLayout.Fill = new ColorSolidBrush() { Color = Colors.Transparent, ColorSpecialName = "Transparent" };
    //        Global.EndInit();

    //        int maxZindex = GetMaxZIndexShapeBackground(StandardElements);
    //        foreach (ObjectElement objectelement in slide.MainLayout.Elements)
    //        {
    //            objectelement.ZIndex += (maxZindex + 1);
    //        }
    //        foreach (StandardElement item in StandardElements)
    //        {

    //            slide.MainLayout.Elements.Add(item as StandardElement);
    //        }

    //        slide.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background, new Action(() =>
    //        {
    //            (Application.Current as IAppGlobal).SelectedTheme.ThumbnailBitmap = FileHelper.SaveToThumbnailIcon(slide.MainLayout, (Application.Current as IAppGlobal).SelectedTheme.Colors, new Size(341, 192));
    //            (Application.Current as IAppGlobal).SelectedTheme.ThumbnailBitmapSmall = FileHelper.SaveToThumbnailIconSmall(slide.MainLayout, (Application.Current as IAppGlobal).SelectedTheme.Colors);
    //        }));
    //        slide.MainLayout.UpdateThumbnail();


    //        foreach (LayoutMaster layout in slide.LayoutMasters)
    //        {
    //            RemoveBackgroundLayoutMasterShapes(layout);
    //            if (layout.ThemeLayout == null) layout.ThemeLayout = new ThemeLayout();
    //            foreach (ELayoutMaster item in eTheme.SlideMasters[0].LayoutMasters)
    //            {
    //                if (item.SlideName == layout.SlideName)
    //                {
    //                    layout.IsHideBackground = item.IsHideBackground;

    //                    foreach (EStandardElement eStandardElement in item.MainLayer.Children)
    //                    {
    //                        if (eStandardElement is DataDocument data)
    //                        {
    //                            TextEditor textEditor = new TextEditor();
    //                            textEditor.UpdateUI(data);
    //                            layout.MainLayout.Elements.Add(textEditor);
    //                        }
    //                        else
    //                        {
    //                            StandardElement layoutElement = new StandardElement();
    //                            layoutElement.UpdateUI(eStandardElement);
    //                            layout.MainLayout.Elements.Add(layoutElement);
    //                        }
    //                    }
    //                }
    //            }

    //            Global.BeginInit();
    //            layout.MainLayout.Fill = new ColorSolidBrush() { Color = Colors.Transparent, ColorSpecialName = "Transparent" };
    //            Global.EndInit();

    //            layout.UpdateLayoutTheme();
    //            layout.MainLayout.ThumbnailBitmap = DesignTabControlViewModel.GetThumbnail(layout.MainLayout, (Application.Current as IAppGlobal).SelectedTheme.Colors, new Size(150, 67));

    //        }
    //        Global.EndInit();
    //    }


    //    /// <summary>
    //    /// Get max Zindex của hình nền
    //    /// </summary>
    //    /// <param name="shapeBackgrounds"></param>
    //    /// <returns></returns>
    //    private int GetMaxZIndexShapeBackground(ObservableCollection<StandardElement> shapeBackgrounds)
    //    {
    //        int max = 0;
    //        foreach (StandardElement item in shapeBackgrounds)
    //        {
    //            if (item.ZIndex > max)
    //            {
    //                max = item.ZIndex;
    //            }
    //        }
    //        return max;
    //    }

    //    /// <summary>
    //    /// Xóa toàn bộ hình nền của Slide Master
    //    /// </summary>
    //    /// <param name="slide"></param>
    //    private void RemoveBackgroundSlideMasterShapes(SlideMaster slide)
    //    {
    //        Global.BeginInit();
    //        for (int i = 0; i < slide.MainLayout.Elements.Count; i++)
    //        {
    //            if (slide.MainLayout.Elements[i] is StandardElement && ((slide.MainLayout.Elements[i] as StandardElement).IsGraphicBackground))
    //            {
    //                slide.MainLayout.Elements.Remove(slide.MainLayout.Elements[i] as StandardElement);
    //                i--;
    //            }
    //        }
    //        Global.EndInit();
    //    }


    //    /// <summary>
    //    /// Xóa toàn bộ hình nền của Layout Master
    //    /// </summary>
    //    /// <param name="slide"></param>
    //    private static void RemoveBackgroundLayoutMasterShapes(LayoutMaster slide)
    //    {
    //        Global.BeginInit();
    //        for (int i = 0; i < slide.MainLayout.Elements.Count; i++)
    //        {
    //            slide.MainLayout.Elements.Remove(slide.MainLayout.Elements[i] as StandardElement);
    //            i--;
    //        }
    //        Global.EndInit();
    //    }
    }
}
