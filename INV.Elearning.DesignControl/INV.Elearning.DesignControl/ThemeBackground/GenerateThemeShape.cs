using INV.Elearing.Controls.Shapes;
using INV.Elearning.Core;
using INV.Elearning.Core.Helper;
using INV.Elearning.Core.Model;
using INV.Elearning.Core.View;
using INV.Elearning.Core.View.Theme;
using INV.Elearning.DesignControl.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media;

namespace INV.Elearning.DesignControl.ThemeBackground
{
    public static class GenerateThemeShape
    {
        public static ObservableCollection<StandardElement> Theme01(EColorManagment color)
        {
            SolidColor colorAccent5 = color.Accent5;
            SolidColor colorAccent4 = color.Accent4;
            SolidColor colorAccent3 = color.Accent3;
            SolidColor colorAccent6 = color.Accent6;


            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();

            StaticShape staticShape = new StaticShape() { PathData = "M499,583.5 L730,584.5 C730,584.5 698.5,551 685.25,532.5 662.51089,500.75106 630.82785,454.87885 585,352 560.5,297 524,189 514,145 504,101 499.125,81.75 499.125,81.75 ", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#8297b0") }, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 133, Height = 252, Top = 325, Left = 0, Thickness = 0, ZIndex = 2 };
            standardElement.ShapePresent = staticShape;
            standardElement.Opacity = 0.7;
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M1276,260.54167 L1275.75,121 1163.75,121.25 C1163.75,121.25 1199.5,158.37489 1205.5,165.87489 1211.5,173.37489 1242.6667,212.91644 1245,215.99977 1250.1492,222.80405 1268.0833,249.41635 1270.0833,252.12468 1270.8541,253.1684 1276,260.54167 1276,260.54167 z ", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#1e4471") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 133, Height = 258, Top = 0, Left = 893, Thickness = 0, ZIndex = 2 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent6.Color), ColorSpecialName = colorAccent6.Name };
            listshape.Add(standardElement);


            staticShape = new StaticShape() { PathData = "M85,815.97917 L129,815.33352 C129,815.33352 143.66667,686.39555 148.33333,663.73973 153,641.08391 184.00051,459.17025 232.00076,331.73099 280.00101,204.29172 312.32219,134.5594 338.33467,84.682066 346.5014,69.022861 376.33487,17.38081 376.33487,17.38081 L85,15.381762 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#1e4471") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 133, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent5.Color), ColorSpecialName = colorAccent5.Name };
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0,0 L731,0 L731,13 L0,13 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#8297b0") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 530, Height = 13, Top = 300, Left = 117, Thickness = 0, ZIndex = 3 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent5.Color), ColorSpecialName = colorAccent5.Name };
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M127.338,50.458 C127.338,49.905715 127.78572,49.458 128.338,49.458 L212.671,49.458 C213.22328,49.458 213.671,49.905715 213.671,50.458 L213.671,61.792001 C213.671,62.344285 213.22328,62.792001 212.671,62.792001 L128.338,62.792001 C127.78572,62.792001 127.338,62.344285 127.338,61.792001 z M127.338,102.457 C127.338,101.90472 127.78572,101.457 128.338,101.457 L212.671,101.457 C213.22328,101.457 213.671,101.90472 213.671,102.457 L213.671,113.791 C213.671,114.34329 213.22328,114.791 212.671,114.791 L128.338,114.791 C127.78572,114.791 127.338,114.34329 127.338,113.791 z M127.338,77.457 C127.338,76.904715 127.78572,76.457 128.338,76.457 L212.671,76.457 C213.22328,76.457 213.671,76.904715 213.671,77.457 L213.671,88.791001 C213.671,89.343285 213.22328,89.791001 212.671,89.791001 L128.338,89.791001 C127.78572,89.791001 127.338,89.343285 127.338,88.791001 z M31.170812,236.62499 L192.92105,235.87499 C192.92105,235.87499 212.671,234.29166 230.50433,222.29166 239.44535,216.27527 242.92113,207.12496 242.92113,205.62495 242.92113,204.12495 244.17113,194.37494 244.17113,193.12494 244.17113,191.87494 245.67113,50.79142 245.67113,50.79142 L321.42073,50.874993 C321.42073,50.874993 322.0042,37.791412 317.67087,27.124732 313.33753,16.458052 302.92614,7.8955584 302.00407,7.1246971 296.92075,2.8750001 285.2541,1.831055E-07 281.92076,1.831055E-07 278.58743,1.831055E-07 126.67079,1.1250001 126.67079,1.1250001 L126.67089,13.25 242.7958,13.125 C242.7958,13.125 234.33742,15.458333 233.67076,39.458332 233.00409,63.458332 233.00409,197.79166 233.00409,197.79166 233.00409,197.79166 229.92081,207.62499 221.33743,213.79166 212.46983,220.16252 206.29582,223.24999 195.42083,221.62499 186.74849,220.32912 178.67085,213.24999 175.33744,208.45833 172.00754,203.6717 170.50419,200.95833 168.67077,194.79166 167.94732,192.35834 167.33744,189.45833 167.33744,189.45833 L0.0041537429,189.12499 C0.0041537429,189.12499 -0.31625681,201.68479 4.6711339,210.12499 9.5461342,218.37499 12.920979,222.87499 17.004458,226.45833 21.465295,230.3728 31.170812,236.62499 31.170812,236.62499 z M246.29599,38.87501 L308.92101,38.999995 C308.92101,38.999995 305.796,25.624755 295.67105,17.125343 290.61567,12.881593 284.73057,11.719875 276.33779,11.792084 266.67099,11.875253 264.29599,14.125233 259.33785,16.791875 255.81337,18.687449 255.69907,19.14882 251.17099,24.500138 248.42099,27.750108 246.79623,32.375065 246.546,34.250048 246.32116,35.934769 246.29599,38.87501 246.29599,38.87501 z M16.003964,202.79166 L155.17092,202.62499 C155.17092,202.62499 159.0043,212.12499 165.17097,218.62499 168.02032,221.62836 169.671,224.62499 169.671,224.62499 L36.337346,225.12499 C36.337346,225.12499 29.670664,221.79166 21.670277,212.62499 16.772879,207.01366 16.003964,202.79166 16.003964,202.79166 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#8297b0") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 200, Height = 200, Top = 300, Left = 630, Thickness = 0, ZIndex = 3 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent4.Color), ColorSpecialName = colorAccent4.Name };
            listshape.Add(standardElement);
            return listshape;
        }

        public static void GenerateLayoutTheme01(SlideMaster slide, int maxZindex, EColorManagment color)
        {
            SolidColor colorAccent3 = color.Accent3;
            SolidColor colorAccent6 = color.Accent6;

            foreach (LayoutMaster layoutMaster in slide.LayoutMasters)
            {
                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();

                foreach (ObjectElement objectlayout in layoutMaster.MainLayout.Elements)
                {
                    objectlayout.ZIndex += (maxZindex + 3);
                }

                if (!((layoutMaster.LayoutName == "Blank Layout") || (layoutMaster.LayoutName == "Title Slide Layout")))
                {
                    layoutMaster.IsHideBackground = true;

                    StaticShape staticShape = new StaticShape() { PathData = "M499,583.5 L730,584.5 C730,584.5 698.5,551 685.25,532.5 662.51089,500.75106 630.82785,454.87885 585,352 560.5,297 524,189 514,145 504,101 499.125,81.75 499.125,81.75 ", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#8297b0") }, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 133, Height = 252, Top = 325, Left = 0, Thickness = 0, ZIndex = 2 };
                    standardElement.ShapePresent = staticShape;
                    standardElement.Opacity = 0.7;
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M1276,260.54167 L1275.75,121 1163.75,121.25 C1163.75,121.25 1199.5,158.37489 1205.5,165.87489 1211.5,173.37489 1242.6667,212.91644 1245,215.99977 1250.1492,222.80405 1268.0833,249.41635 1270.0833,252.12468 1270.8541,253.1684 1276,260.54167 1276,260.54167 z ", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#1e4471") }, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 133, Height = 258, Top = 0, Left = 893, Thickness = 0, ZIndex = 2 };
                    standardElement.ShapePresent = staticShape;
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent6.Color), ColorSpecialName = colorAccent6.Name };
                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }

            }
        }

        public static ObservableCollection<StandardElement> Theme_01()
        {
            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();

            StaticShape staticShape = new StaticShape() { PathData = "M499,583.5 L730,584.5 C730,584.5 698.5,551 685.25,532.5 662.51089,500.75106 630.82785,454.87885 585,352 560.5,297 524,189 514,145 504,101 499.125,81.75 499.125,81.75 ", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#8297b0") }, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 133, Height = 252, Top = 325, Left = 0, Thickness = 0, ZIndex = 2 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = Colors.Red };
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M1276,260.54167 L1275.75,121 1163.75,121.25 C1163.75,121.25 1199.5,158.37489 1205.5,165.87489 1211.5,173.37489 1242.6667,212.91644 1245,215.99977 1250.1492,222.80405 1268.0833,249.41635 1270.0833,252.12468 1270.8541,253.1684 1276,260.54167 1276,260.54167 z ", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#1e4471") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 133, Height = 258, Top = 0, Left = 893, Thickness = 0, ZIndex = 2 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = Colors.LightGreen };
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M127.338,50.458 C127.338,49.905715 127.78572,49.458 128.338,49.458 L212.671,49.458 C213.22328,49.458 213.671,49.905715 213.671,50.458 L213.671,61.792001 C213.671,62.344285 213.22328,62.792001 212.671,62.792001 L128.338,62.792001 C127.78572,62.792001 127.338,62.344285 127.338,61.792001 z M127.338,102.457 C127.338,101.90472 127.78572,101.457 128.338,101.457 L212.671,101.457 C213.22328,101.457 213.671,101.90472 213.671,102.457 L213.671,113.791 C213.671,114.34329 213.22328,114.791 212.671,114.791 L128.338,114.791 C127.78572,114.791 127.338,114.34329 127.338,113.791 z M127.338,77.457 C127.338,76.904715 127.78572,76.457 128.338,76.457 L212.671,76.457 C213.22328,76.457 213.671,76.904715 213.671,77.457 L213.671,88.791001 C213.671,89.343285 213.22328,89.791001 212.671,89.791001 L128.338,89.791001 C127.78572,89.791001 127.338,89.343285 127.338,88.791001 z M31.170812,236.62499 L192.92105,235.87499 C192.92105,235.87499 212.671,234.29166 230.50433,222.29166 239.44535,216.27527 242.92113,207.12496 242.92113,205.62495 242.92113,204.12495 244.17113,194.37494 244.17113,193.12494 244.17113,191.87494 245.67113,50.79142 245.67113,50.79142 L321.42073,50.874993 C321.42073,50.874993 322.0042,37.791412 317.67087,27.124732 313.33753,16.458052 302.92614,7.8955584 302.00407,7.1246971 296.92075,2.8750001 285.2541,1.831055E-07 281.92076,1.831055E-07 278.58743,1.831055E-07 126.67079,1.1250001 126.67079,1.1250001 L126.67089,13.25 242.7958,13.125 C242.7958,13.125 234.33742,15.458333 233.67076,39.458332 233.00409,63.458332 233.00409,197.79166 233.00409,197.79166 233.00409,197.79166 229.92081,207.62499 221.33743,213.79166 212.46983,220.16252 206.29582,223.24999 195.42083,221.62499 186.74849,220.32912 178.67085,213.24999 175.33744,208.45833 172.00754,203.6717 170.50419,200.95833 168.67077,194.79166 167.94732,192.35834 167.33744,189.45833 167.33744,189.45833 L0.0041537429,189.12499 C0.0041537429,189.12499 -0.31625681,201.68479 4.6711339,210.12499 9.5461342,218.37499 12.920979,222.87499 17.004458,226.45833 21.465295,230.3728 31.170812,236.62499 31.170812,236.62499 z M246.29599,38.87501 L308.92101,38.999995 C308.92101,38.999995 305.796,25.624755 295.67105,17.125343 290.61567,12.881593 284.73057,11.719875 276.33779,11.792084 266.67099,11.875253 264.29599,14.125233 259.33785,16.791875 255.81337,18.687449 255.69907,19.14882 251.17099,24.500138 248.42099,27.750108 246.79623,32.375065 246.546,34.250048 246.32116,35.934769 246.29599,38.87501 246.29599,38.87501 z M16.003964,202.79166 L155.17092,202.62499 C155.17092,202.62499 159.0043,212.12499 165.17097,218.62499 168.02032,221.62836 169.671,224.62499 169.671,224.62499 L36.337346,225.12499 C36.337346,225.12499 29.670664,221.79166 21.670277,212.62499 16.772879,207.01366 16.003964,202.79166 16.003964,202.79166 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#8297b0") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 133, Height = 258, Top = 146, Left = 850, Thickness = 0, ZIndex = 5 };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M85,815.97917 L129,815.33352 C129,815.33352 143.66667,686.39555 148.33333,663.73973 153,641.08391 184.00051,459.17025 232.00076,331.73099 280.00101,204.29172 312.32219,134.5594 338.33467,84.682066 346.5014,69.022861 376.33487,17.38081 376.33487,17.38081 L85,15.381762 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#1e4471") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 133, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = Colors.LightPink };
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0,0 L731,0 L731,13 L0,13 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#8297b0") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 731, Height = 13, Top = 146, Left = 117, Thickness = 0, ZIndex = 3 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = Colors.LightSkyBlue };
            listshape.Add(standardElement);
            return listshape;
        }

        public static ObservableCollection<StandardElement> Theme02(EColorManagment color)
        {
            SolidColor colorAccent1 = color.Accent1;

            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();

            StaticShape staticShape = new StaticShape() { PathData = "M264,591.25 C264,591.25 263,567.75 264,562.5 265,557.25 270.00619,500.25 270.50671,498.25 271.00722,496.25 289.02582,544 289.02582,544 L309.29673,532.24992 322.5604,573.74998 C322.5604,573.74998 319.05678,506.24988 318.30601,505.24988 317.55523,504.24988 345.33388,574.74998 345.33388,574.74998 L357.59654,506.24988 382.87259,575.74998 389.87982,529.74992 404.14453,568.99997 430.92215,553.74995 444.68634,569.74997 483.97686,508.24989 C483.97686,508.24989 495.48874,571.24997 497.24055,571.24997 498.99235,571.24997 506.75035,540.49993 506.75035,540.49993 L526.02021,562.74996 549.79472,541.24993 567.06255,579.99998 606.8536,541.99993 626.62397,574.49998 654.65287,543.49993 665.91451,572.74997 654.90313,512.24989 688.18746,564.74996 701.70138,552.74995 C701.70138,552.74995 707.45729,583.49999 708.45832,583.49999 709.45935,583.49999 726.22664,529.74992 726.22664,529.74992 L747.2483,577.49998 761.01248,554.24995 781.28342,568.99997 C781.28342,568.99997 786.28855,534.99992 785.53777,531.74992 L810.31295,566.87497 826.57969,556.24995 826.9551,576.99998 841.34491,531.49992 854.35832,586.24999 887.39238,568.74997 912.9187,586.87499 903.4089,536.62493 919.05002,567.87497 933.43988,536.49993 936.19272,568.37497 945.3271,550.62494 976.73445,580.87499 983.86679,536.37493 1004.6382,589.875 1028.0373,519.7499 1052.0621,573.74998 1069.3299,536.37493 1087.5988,574.99998 1098.485,568.24997 1097.2337,566.12497 1111.6235,495.49987 1130.3929,541.99994 1150.2883,530.99992 1163.552,571.62497 1159.5479,504.87489 1186.826,572.62498 1199.3388,505.12489 1224.1561,573.83331 1230.6625,528.49992 1246.5118,566.8333 1271.7045,552.16662 1285.7186,567.66664 1324.5915,507.16656 1338.1055,568.3333 1347.6153,538.99993 1366.6348,560.83329 1391.16,539.99993 1409.3454,578.33332 1447.5511,540.83327 1467.0716,571.83331 1494.9339,541.99994 1507.7801,571.99998 1496.2682,510.83323 1529.8027,561.3333 1542.8161,551.66662 1545.1519,558.83329 1545.1519,591 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#7dc587") }, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 1026, Height = 77, Top = 500, Left = 0, Thickness = 0, ZIndex = 1 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M192.17165,861.69994 C192.17165,861.69994 260.94976,789.34541 324.69298,840.29743 324.69298,840.29743 344.15017,854.64736 321.42165,864.44995 321.42165,864.44995 303.17196,866.94995 308.17197,850.6999 308.17197,850.6999 315.89325,830.8457 352.14335,842.34586 388.39345,853.84603 433.34444,867.55672 507.34464,867.05671 507.34464,867.05671 609.92338,871.86493 639.17346,831.61418 639.17346,831.61418 666.66211,802.36861 630.15258,780.99884 630.15258,780.99884 598.92863,765.10899 594.09529,791.9425 594.09529,791.9425 594.0654,810.99913 624.56549,815.99918 624.56549,815.99918 655.07974,817.31836 706.71226,799.28111 706.71226,799.28111 876.12962,748.36343 942.87982,759.55121 942.87982,759.55121 971.88366,763.745 1004.6742,777.44922 1038.5944,791.62552 1066.2945,832.14034 1023.2943,860.64067 980.29418,889.14099 908.81603,827.94238 972.62146,809.15265 1066.8289,781.41 1161.2756,835.27888 1172.122,837.15296 1216.9224,844.89373 1265.3412,856.70012 1291.6745,846.7 1316.9763,837.09157 1345.1749,813.19963 1329.1746,784.1993 L1329.3361,784.12633 C1329.3361,784.12633 1353.6742,822.45796 1291.8669,847.00818 1291.8669,847.00818 1268.8395,863.98845 1171.2783,839.10376 1171.2783,839.10376 1048.5248,783.06322 974.84646,811.50347 911.41246,835.98934 985.02637,883.66119 1021.9007,859.60741 1063.3392,832.57635 1038.8832,793.41081 1005.3324,779.4743 1005.3324,779.4743 982.63012,769.75475 942.6107,764.05117 942.6107,764.05117 896.48803,748.64195 706.95742,801.68946 706.95742,801.68946 668.51881,817.08501 622.60447,819.23162 622.60447,819.23162 593.30561,816.91969 590.8304,792.14215 590.8304,792.14215 591.65661,764.72319 631.58933,778.59806 631.58933,778.59806 666.81102,798.76469 641.85307,833.45187 641.85307,833.45187 625.33778,871.3102 508.52436,870.89748 508.52436,870.89748 441.84512,874.59415 351.99649,844.86228 351.99649,844.86228 319.86032,832.52513 309.16906,850.83453 309.16906,850.83453 302.69556,864.66998 321.29086,863.18208 321.29086,863.18208 343.57351,854.03001 322.60127,838.63178 312.50133,831.2162 296.32498,826.75039 277.92425,825.95518 251.02882,824.79286 222.21997,837.90857 192.57748,862.27632 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#7dc587") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 1000, Height = 100, Top = 50, Left = 0, Thickness = 0, ZIndex = 2 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
            listshape.Add(standardElement);

            return listshape;
        }

        public static void GenerateLayoutTheme02(SlideMaster slide, int maxZindex, EColorManagment color)
        {
            SolidColor colorAccent1 = color.Accent1;

            foreach (LayoutMaster layoutMaster in slide.LayoutMasters)
            {
                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();

                foreach (ObjectElement objectlayout in layoutMaster.MainLayout.Elements)
                {
                    objectlayout.ZIndex += (maxZindex + 3);
                }

                if (!((layoutMaster.LayoutName == "Blank Layout") || (layoutMaster.LayoutName == "Title Slide Layout")))
                {
                    layoutMaster.IsHideBackground = true;

                    StaticShape staticShape = new StaticShape() { PathData = "M264,591.25 C264,591.25 263,567.75 264,562.5 265,557.25 270.00619,500.25 270.50671,498.25 271.00722,496.25 289.02582,544 289.02582,544 L309.29673,532.24992 322.5604,573.74998 C322.5604,573.74998 319.05678,506.24988 318.30601,505.24988 317.55523,504.24988 345.33388,574.74998 345.33388,574.74998 L357.59654,506.24988 382.87259,575.74998 389.87982,529.74992 404.14453,568.99997 430.92215,553.74995 444.68634,569.74997 483.97686,508.24989 C483.97686,508.24989 495.48874,571.24997 497.24055,571.24997 498.99235,571.24997 506.75035,540.49993 506.75035,540.49993 L526.02021,562.74996 549.79472,541.24993 567.06255,579.99998 606.8536,541.99993 626.62397,574.49998 654.65287,543.49993 665.91451,572.74997 654.90313,512.24989 688.18746,564.74996 701.70138,552.74995 C701.70138,552.74995 707.45729,583.49999 708.45832,583.49999 709.45935,583.49999 726.22664,529.74992 726.22664,529.74992 L747.2483,577.49998 761.01248,554.24995 781.28342,568.99997 C781.28342,568.99997 786.28855,534.99992 785.53777,531.74992 L810.31295,566.87497 826.57969,556.24995 826.9551,576.99998 841.34491,531.49992 854.35832,586.24999 887.39238,568.74997 912.9187,586.87499 903.4089,536.62493 919.05002,567.87497 933.43988,536.49993 936.19272,568.37497 945.3271,550.62494 976.73445,580.87499 983.86679,536.37493 1004.6382,589.875 1028.0373,519.7499 1052.0621,573.74998 1069.3299,536.37493 1087.5988,574.99998 1098.485,568.24997 1097.2337,566.12497 1111.6235,495.49987 1130.3929,541.99994 1150.2883,530.99992 1163.552,571.62497 1159.5479,504.87489 1186.826,572.62498 1199.3388,505.12489 1224.1561,573.83331 1230.6625,528.49992 1246.5118,566.8333 1271.7045,552.16662 1285.7186,567.66664 1324.5915,507.16656 1338.1055,568.3333 1347.6153,538.99993 1366.6348,560.83329 1391.16,539.99993 1409.3454,578.33332 1447.5511,540.83327 1467.0716,571.83331 1494.9339,541.99994 1507.7801,571.99998 1496.2682,510.83323 1529.8027,561.3333 1542.8161,551.66662 1545.1519,558.83329 1545.1519,591 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#7dc587") }, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 1026, Height = 77, Top = 500, Left = 0, Thickness = 0, ZIndex = 1 };
                    standardElement.ShapePresent = staticShape;
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                }

            }
        }


        public static ObservableCollection<StandardElement> Theme03(EColorManagment color)
        {

            SolidColor colorAccent1 = color.Accent1;
            SolidColor colorAccent2 = color.Accent2;

            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();
            StaticShape staticShape = new StaticShape() { PathData = "M442,1038 L376.5,1038 376.5,434.5 312.5,434.5 312.5,370.5 376.5,370.5 376.5,240.5 439.5,240.5 439.5,303.5 1589.5,303.5 1589.5,368.5 440.5,368.5 z", Fill = Brushes.LightGray, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 1026, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 52, Height = 47, Top = 0, Left = 0, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
            standardElement = new StandardElement() { Width = 52, Height = 47, Top = 46, Left = 51, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
            standardElement = new StandardElement() { Width = 52, Height = 47, Top = 93, Left = 103, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
            standardElement = new StandardElement() { Width = 52, Height = 47, Top = 46, Left = 156, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
            standardElement = new StandardElement() { Width = 52, Height = 47, Top = 530, Left = 974, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
            standardElement = new StandardElement() { Width = 52, Height = 47, Top = 483, Left = 922, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
            standardElement = new StandardElement() { Width = 52, Height = 47, Top = 530, Left = 870, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
            standardElement = new StandardElement() { Width = 52, Height = 47, Top = 436, Left = 974, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            return listshape;
        }
        public static void GenerateLayoutTheme03(SlideMaster slide, int maxZindex, EColorManagment color)
        {
            SolidColor colorAccent1 = color.Accent1;
            SolidColor colorAccent2 = color.Accent2;
            foreach (LayoutMaster layoutMaster in slide.LayoutMasters)
            {
                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();

                foreach (ObjectElement objectlayout in layoutMaster.MainLayout.Elements)
                {
                    objectlayout.ZIndex += (maxZindex + 3);
                }

                if (!((layoutMaster.LayoutName == "Blank Layout") || (layoutMaster.LayoutName == "Title Slide Layout")))
                {
                    layoutMaster.IsHideBackground = true;

                    StaticShape staticShape = new StaticShape() { PathData = "M394,902.33333 L331.5,902.33333 331.5,105.16698 1608.8333,105.16698 1608.8333,169.16667 396.83335,169.16667 z", Fill = Brushes.LightGray, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 1026, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
                    standardElement = new StandardElement() { Width = 52, Height = 47, Top = 0, Left = 0, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
                    standardElement = new StandardElement() { Width = 52, Height = 47, Top = 46, Left = 51, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
                    standardElement = new StandardElement() { Width = 52, Height = 47, Top = 0, Left = 103, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M0.5,0.5 L64.5,0.5 L64.5,63.5 L0.5,63.5 z", Fill = Brushes.LightGreen, StrokeThickness = 0, Tag = "Accent1" };
                    standardElement = new StandardElement() { Width = 52, Height = 47, Top = 530, Left = 974, Thickness = 0, ZIndex = 1, Tag = "Accent1" };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                }
            }
        }

        public static ObservableCollection<StandardElement> Theme04(EColorManagment color)
        {
            SolidColor colorAccent5 = color.Accent5;
            SolidColor colorAccent4 = color.Accent4;
            SolidColor colorAccent3 = color.Accent3;
            SolidColor colorAccent6 = color.Accent6;

            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();

            StaticShape staticShape = new StaticShape() { PathData = "M255.66667,910.58333 L179.83301,910.58333 179.83301,784.50031 237.62475,684.5003 237.62475,145.50029 265.49963,111.00029 292.99951,145.00029 292.99951,684.50031 254.83301,782.50031 z", Fill = Brushes.Green, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 100, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M624.66667,1000.3333 L698.83302,1000.3333 698.83302,874.00031 714.00007,775.5003 714.00007,340.18779 685.33327,307.6675 657.33316,340.16717 657.33316,774.83371 624.83302,873.83368 z", Fill = Brushes.LightGreen, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 91, Height = 480, Top = 97, Left = 77, Thickness = 0, ZIndex = 2 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent4.Color), ColorSpecialName = colorAccent4.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M1150.0833,824.5 L1225.667,824.5 1225.667,698.16667 1216.3334,604.83279 1216.3334,306.07757 1188.0471,277.20268 1160.8337,306.12468 1160.8337,604.49987 1149.7504,698.24992 z", Fill = Brushes.Orange, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 91, Height = 353, Top = 224, Left = 166, Thickness = 0, ZIndex = 2 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent5.Color), ColorSpecialName = colorAccent5.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M1496.3333,599.66667 L1496.3333,725.16635 1570.8337,725.16635 1570.8337,599.49969 1537.1672,499.16661 1537.1672,331.24981 1510.0004,307.74984 1481.8336,331.16649 1481.8336,499.49963 z", Fill = Brushes.LightGray, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 91, Height = 270, Top = 305, Left = 252, Thickness = 0, ZIndex = 2 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent6.Color), ColorSpecialName = colorAccent6.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            return listshape;
        }

        public static void GenerateLayoutTheme04(SlideMaster slide, int maxZindex, EColorManagment color)
        {
            SolidColor colorAccent5 = color.Accent5;
            SolidColor colorAccent4 = color.Accent4;
            SolidColor colorAccent3 = color.Accent3;
            foreach (LayoutMaster layoutMaster in slide.LayoutMasters)
            {
                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();

                foreach (ObjectElement objectlayout in layoutMaster.MainLayout.Elements)
                {
                    objectlayout.ZIndex += (maxZindex + 3);
                }

                if (!((layoutMaster.LayoutName == "Blank Layout") || (layoutMaster.LayoutName == "Title Slide Layout")))
                {
                    layoutMaster.IsHideBackground = true;

                    StaticShape staticShape = new StaticShape() { PathData = "M151.66667,851.33333 L151.66667,939.16698 226.49967,939.16698 226.49967,850.83365 241.083,801.79263 241.083,176.17005 212.91656,139.50319 185.49979,176.41971 185.49979,800.50055 z", StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 100, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = maxZindex + 2 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M443.16667,821.16667 L443.16667,910.37469 519.24966,910.37469 519.24966,820.87469 510.1663,772.16636 510.1663,333.43589 481.99964,300.49778 454.49965,332.99787 454.49965,771.91593 z", StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 91, Height = 480, Top = 97, Left = 97, Thickness = 0, ZIndex = maxZindex + 2 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent4.Color), ColorSpecialName = colorAccent4.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M859,778.66667 L859,865.16635 933.5,865.16635 933.5,777.83302 899.16667,727.16704 899.16667,425.87652 872.125,396.12667 843.875,425.62653 843.875,726.62506 z", StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 91, Height = 353, Top = 224, Left = 186, Thickness = 0, ZIndex = maxZindex + 2 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent5.Color), ColorSpecialName = colorAccent5.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }
            }
        }

        public static ObservableCollection<StandardElement> Theme05(EColorManagment color)
        {
            SolidColor colorAccent1 = color.Accent1;
            SolidColor colorAccent2 = color.Accent2;
            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();

            StaticShape staticShape = new StaticShape() { PathData = "M129.5,225 L129.5,735.5 1146.833,735.5 1146.833,225 z M0.5,0.5 L1279.5,0.5 1279.5,799.5 0.5,799.5 z", Fill = Brushes.Green, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 1026, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M344.25,627.75 L392.25,627.75 392.25,707.83333 368.54182,684.12506 344.33333,708.33365 z", Fill = Brushes.LightGray, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 50, Height = 82, Top = 162, Left = 871, Thickness = 0, ZIndex = 2 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            return listshape;
        }

        public static void GenerateLayoutTheme05(SlideMaster slide, int maxZindex, EColorManagment color)
        {
            SolidColor colorAccent1 = color.Accent1;
            SolidColor colorAccent2 = color.Accent2;

            foreach (LayoutMaster layoutMaster in slide.LayoutMasters)
            {
                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();

                foreach (ObjectElement objectlayout in layoutMaster.MainLayout.Elements)
                {
                    objectlayout.ZIndex += (maxZindex + 3);
                }
                if (layoutMaster.LayoutName == "Two Content Layout" || layoutMaster.LayoutName == "Comparison Layout")
                {
                    layoutMaster.IsHideBackground = true;
                    StaticShape staticShape = new StaticShape() { PathData = "M653.6015,265.55322 L653.6015,742.66498 1159.918,742.66498 1159.918,265.55322 z M116.50099,265.55322 L116.50099,742.66498 626.36816,742.66498 626.36816,265.55322 z M116.50099,58.165998 L116.50099,234.26009 1159.918,234.26009 1159.918,58.165998 z M0.5,0.5 L1279.5,0.5 1279.5,799.5 0.5,799.5 z", Fill = Brushes.Green, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 1026, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M344.25,627.75 L392.25,627.75 392.25,707.83333 368.54182,684.12506 344.33333,708.33365 z", Fill = Brushes.LightGray, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 50, Height = 82, Top = 42, Left = 853, Thickness = 0, ZIndex = 2 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }
                else if (layoutMaster.LayoutName == "Content With Caption Layout" || layoutMaster.LayoutName == "Picture With Caption Layout")
                {
                    layoutMaster.IsHideBackground = true;

                    StaticShape staticShape = new StaticShape() { PathData = "M125.968,327.62622 L125.968,731.33301 548.90057,731.33301 548.90057,327.62622 z M598.92218,189.40457 L598.92218,731.33301 1163.833,731.33301 1163.833,189.40457 z M125.968,77.333009 L125.968,305.95905 548.90057,305.95905 548.90057,77.333009 z M0.5,0.5 L1279.5,0.5 1279.5,799.5 0.5,799.5 z", Fill = Brushes.Green, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 1026, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;

                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M344.25,627.75 L392.25,627.75 392.25,707.83333 368.54182,684.12506 344.33333,708.33365 z", Fill = Brushes.LightGray, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 50, Height = 82, Top = 42, Left = 853, Thickness = 0, ZIndex = 2 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
                    standardElement.ShapePresent = staticShape;

                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }
                else if (layoutMaster.LayoutName == "Title Content Layout " || layoutMaster.LayoutName == "Black Layout")
                {
                    layoutMaster.IsHideBackground = true;

                    StaticShape staticShape = new StaticShape() { PathData = "M176.5,94.500009 L176.5,705.5 1105.5,705.5 1105.5,94.500009 z M0.5,0.5 L1279.5,0.5 1279.5,799.5 0.5,799.5 z", Fill = Brushes.Green, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 1026, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;

                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M344.25,627.75 L392.25,627.75 392.25,707.83333 368.54182,684.12506 344.33333,708.33365 z", Fill = Brushes.LightGray, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 50, Height = 82, Top = 42, Left = 853, Thickness = 0, ZIndex = 2 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
                    standardElement.ShapePresent = staticShape;

                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }
                else if (layoutMaster.LayoutName == "Title Only Layout")
                {
                    layoutMaster.IsHideBackground = true;

                    StaticShape staticShape = new StaticShape() { PathData = "M123.99999,94.000009 L123.99999,272 1165,272 1165,94.000009 z M0.5,0.5 L1279.5,0.5 1279.5,799.5 0.5,799.5 z", Fill = Brushes.Green, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 1026, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;

                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M344.25,627.75 L392.25,627.75 392.25,707.83333 368.54182,684.12506 344.33333,708.33365 z", Fill = Brushes.LightGray, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 50, Height = 82, Top = 42, Left = 853, Thickness = 0, ZIndex = 2 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
                    standardElement.ShapePresent = staticShape;

                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }
            }

        }

        public static ObservableCollection<StandardElement> Theme06(EColorManagment color)
        {
            SolidColor colorAccent1 = color.Accent1;
            SolidColor colorAccent2 = color.Accent2;
            SolidColor colorAccent3 = color.Accent3;
            SolidColor colorAccent4 = color.Accent4;
            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();

            StaticShape staticShape = new StaticShape() { PathData = "M932.625,471.25 L939.3667,478.07501 1057.7836,371.36649 1032.7005,350.69938 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#ffaaaa") }, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 127, Height = 129, Top = 0, Left = 0, Thickness = 0, ZIndex = 10 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString("#4CD39B") };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M800.16667,594.5 C800.16695,594.49968 803.17787,590.42477 806.33302,591.5 810.36407,592.87371 811.19359,599.52075 809.83446,604.83301 808.72044,609.18723 803.66725,613.4998 803.66725,613.4998 803.66725,613.4998 814.83471,606.9998 823.83567,617.66646 826.16924,620.16612 820.0018,624.66627 820.0018,624.66627 820.0018,624.66627 824.6691,627.1665 823.5023,630.16643 822.33551,633.16636 814.11092,636.75083 805.60998,634.91755 805.60998,634.91755 812.83479,645.16632 806.0007,654.99945 799.50011,655.99944 797.49993,652.3328 797.49993,652.3328 797.49993,652.3328 795.16642,658.49993 793.1662,658.1666 791.16598,657.83327 782.33188,654.99993 782.83192,644.16659 782.83192,644.16659 774.66447,651.33324 763.83019,648.33326 763.83019,648.33326 762.24758,642.68211 764.58112,639.68213 764.58112,639.68213 757.8293,640.66657 757.66262,638.66658 757.49594,636.66659 758.16374,624.33324 769.66457,623.99991 769.66457,623.99991 754.9967,616.16657 757.99695,606.83323 759.01708,603.65973 767.75451,606.15992 767.75451,606.15992 767.75451,606.15992 763.33041,598.83267 765.66394,597.66601 767.99748,596.49935 780.49857,596.33318 786.99913,609.49981 786.99913,609.49981 779.16512,593.66649 795.16651,587.99985 795.16651,587.99985 799.49966,590.50032 800.16667,594.5 Z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#ff7575") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 50, Height = 82, Top = 42, Left = 853, Thickness = 0, ZIndex = 2 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString("#4CD39B") };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M800.16667,594.5 C800.16695,594.49968 803.17787,590.42477 806.33302,591.5 810.36407,592.87371 811.19359,599.52075 809.83446,604.83301 808.72044,609.18723 803.66725,613.4998 803.66725,613.4998 803.66725,613.4998 814.83471,606.9998 823.83567,617.66646 826.16924,620.16612 820.0018,624.66627 820.0018,624.66627 820.0018,624.66627 824.6691,627.1665 823.5023,630.16643 822.33551,633.16636 814.11092,636.75083 805.60998,634.91755 805.60998,634.91755 812.83479,645.16632 806.0007,654.99945 799.50011,655.99944 797.49993,652.3328 797.49993,652.3328 797.49993,652.3328 795.16642,658.49993 793.1662,658.1666 791.16598,657.83327 782.33188,654.99993 782.83192,644.16659 782.83192,644.16659 774.66447,651.33324 763.83019,648.33326 763.83019,648.33326 762.24758,642.68211 764.58112,639.68213 764.58112,639.68213 757.8293,640.66657 757.66262,638.66658 757.49594,636.66659 758.16374,624.33324 769.66457,623.99991 769.66457,623.99991 754.9967,616.16657 757.99695,606.83323 759.01708,603.65973 767.75451,606.15992 767.75451,606.15992 767.75451,606.15992 763.33041,598.83267 765.66394,597.66601 767.99748,596.49935 780.49857,596.33318 786.99913,609.49981 786.99913,609.49981 779.16512,593.66649 795.16651,587.99985 795.16651,587.99985 799.49966,590.50032 800.16667,594.5 Z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#ff7575") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 50, Height = 82, Top = 42, Left = 853, Thickness = 0, ZIndex = 2 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString("#4CD39B") };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M800.16667,594.5 C800.16695,594.49968 803.17787,590.42477 806.33302,591.5 810.36407,592.87371 811.19359,599.52075 809.83446,604.83301 808.72044,609.18723 803.66725,613.4998 803.66725,613.4998 803.66725,613.4998 814.83471,606.9998 823.83567,617.66646 826.16924,620.16612 820.0018,624.66627 820.0018,624.66627 820.0018,624.66627 824.6691,627.1665 823.5023,630.16643 822.33551,633.16636 814.11092,636.75083 805.60998,634.91755 805.60998,634.91755 812.83479,645.16632 806.0007,654.99945 799.50011,655.99944 797.49993,652.3328 797.49993,652.3328 797.49993,652.3328 795.16642,658.49993 793.1662,658.1666 791.16598,657.83327 782.33188,654.99993 782.83192,644.16659 782.83192,644.16659 774.66447,651.33324 763.83019,648.33326 763.83019,648.33326 762.24758,642.68211 764.58112,639.68213 764.58112,639.68213 757.8293,640.66657 757.66262,638.66658 757.49594,636.66659 758.16374,624.33324 769.66457,623.99991 769.66457,623.99991 754.9967,616.16657 757.99695,606.83323 759.01708,603.65973 767.75451,606.15992 767.75451,606.15992 767.75451,606.15992 763.33041,598.83267 765.66394,597.66601 767.99748,596.49935 780.49857,596.33318 786.99913,609.49981 786.99913,609.49981 779.16512,593.66649 795.16651,587.99985 795.16651,587.99985 799.49966,590.50032 800.16667,594.5 Z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#ff7575") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 50, Height = 82, Top = 42, Left = 853, Thickness = 0, ZIndex = 2 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString("#4CD39B") };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M140.07766,188.45337 L168.66667,182.33301 175.0159,349.43262 163.86908,349.24078 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#ff7575") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 35, Height = 169, Top = 42, Left = 853, Thickness = 0, ZIndex = 2 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString("#4CD39B") };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M1820.1875,697.25 L1820.0745,710.2 1973.1872,733.57503 1967.6234,705.075 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#ff7575") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 154, Height = 37, Top = 42, Left = 853, Thickness = 0, ZIndex = 18 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString("#4CD39B") };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M1732.75,559.75 L1734.2,569.7 1886.4448,561.53334 1878.8617,533.19997 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#ff7575") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 155, Height = 37, Top = 42, Left = 853, Thickness = 0, ZIndex = 16 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString("#4CD39B") };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            return listshape;
        }

        public static ObservableCollection<StandardElement> Theme07(EColorManagment color)
        {
            SolidColor colorAccent1 = color.Accent1;
            SolidColor colorAccent2 = color.Accent2;
            SolidColor colorAccent3 = color.Accent3;
            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();

            StaticShape staticShape = new StaticShape() { PathData = "M216.66667,707.66667 L356.49968,847.49968 619.83298,584.16635 353.83276,318.16597 219.16633,452.83247 z", Fill = Brushes.Green, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 419, Height = 353, Top = 0, Left = 0, Thickness = 0, ZIndex = 1, Opacity = 0.4 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M216.66667,707.66667 L356.49968,847.49968 619.83298,584.16635 353.83276,318.16597 219.16633,452.83247 z", Fill = Brushes.Green, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 289, Height = 237, Top = 54, Left = 0, Thickness = 0, ZIndex = 2, Opacity = 0.7 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.66666667,388 L159.16633,229.83333 317.666,388.333 158.49933,547.49968 z", Fill = Brushes.Green, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 185, Height = 116, Top = 115, Left = 101, Thickness = 0, ZIndex = 3, Opacity = 1 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.66666667,388 L159.16633,229.83333 317.666,388.333 158.49933,547.49968 z", Fill = Brushes.Orange, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 255, Height = 170, Top = 93, Left = 164, Thickness = 0, ZIndex = 2, Opacity = 0.7 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.66666667,388 L159.16633,229.83333 317.666,388.333 158.49933,547.49968 z", Fill = Brushes.Green, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 324, Height = 237, Top = 54, Left = 285, Thickness = 0, ZIndex = 2, Opacity = 0.7 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.66666667,388 L159.16633,229.83333 317.666,388.333 158.49933,547.49968 z", Fill = Brushes.Orange, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 180, Height = 130, Top = 107, Left = 609, Thickness = 0, ZIndex = 3, Opacity = 1 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.66666667,388 L159.16633,229.83333 317.666,388.333 158.49933,547.49968 z", Fill = Brushes.Green, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 116, Height = 86, Top = 80, Left = 550, Thickness = 0, ZIndex = 3, Opacity = 1 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.66666667,388 L159.16633,229.83333 317.666,388.333 158.49933,547.49968 z", Fill = Brushes.Orange, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 116, Height = 86, Top = 178, Left = 550, Thickness = 0, ZIndex = 3, Opacity = 1 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);


            return listshape;
        }

        public static void GenerateLayoutTheme07(SlideMaster slide, int maxZindex, EColorManagment color)
        {
            SolidColor colorAccent1 = color.Accent1;
            SolidColor colorAccent2 = color.Accent2;
            SolidColor colorAccent3 = color.Accent3;
            foreach (LayoutMaster layoutMaster in slide.LayoutMasters)
            {
                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();

                foreach (ObjectElement objectlayout in layoutMaster.MainLayout.Elements)
                {
                    objectlayout.ZIndex += (maxZindex + 3);
                }

                if (!((layoutMaster.LayoutName == "Blank Layout") || (layoutMaster.LayoutName == "Title Slide Layout")))
                {
                    layoutMaster.IsHideBackground = true;
                    StaticShape staticShape = new StaticShape() { PathData = "M648,1024.3333 L648,1341.8337 488.75018,1182.5837 z", Fill = Brushes.Green, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 160, Height = 320, Top = 200, Left = 864, Thickness = 0, ZIndex = 1, Opacity = 0.4 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M973.875,481.25 L1106.5625,613.9375 1239.501,481 1106.5838,348.08382 z", Fill = Brushes.Green, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 200, Height = 200, Top = 250, Left = 824, Thickness = 0, ZIndex = 2 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M973.875,481.25 L1106.5625,613.9375 1239.501,481 1106.5838,348.08382 z", Fill = Brushes.Green, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 100, Height = 100, Top = 300, Left = 724, Thickness = 0, ZIndex = 1, Opacity = 0.4 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M973.875,481.25 L1106.5625,613.9375 1239.501,481 1106.5838,348.08382 z", Fill = Brushes.Green, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 50, Height = 50, Top = 290, Left = 799, Thickness = 0, ZIndex = 1, Opacity = 0.7 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M973.875,481.25 L1106.5625,613.9375 1239.501,481 1106.5838,348.08382 z", Fill = Brushes.Green, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 50, Height = 50, Top = 360, Left = 799, Thickness = 0, ZIndex = 1, Opacity = 0.7 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }
            }

        }

        public static ObservableCollection<StandardElement> Theme08(EColorManagment eColor)
        {
            SolidColor colorAccent3 = eColor.Accent3;
            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();

            StaticShape staticShape = new StaticShape() { PathData = "M184.45799,219.99999 L197.67469,578.29997 294.12496,578.29997 303.12498,363.99998 208.12473,395.74998 205.12473,389.74998 300.62498,347.74997 302.87498,269.74997 z M166.35512,0.05 L179.36588,0.05 185.70448,188.33004 304.80391,247.92955 305.47149,230.99787 312.81089,230.99787 314.14567,253.34771 341.83511,266.89304 336.4976,278.74522 315.14645,270.27935 317.81535,344.44014 412.22744,391.8489 409.55822,398.96022 317.81511,361.37184 325.15454,578.77464 1279.95,578.77464 1279.95,621.10389 0.05,621.95002 0.05,579.79023 63.270901,579.79023 66.940717,486.66615 27.907269,499.53395 25.90556,497.50212 66.27316,479.21586 67.941273,431.46878 70.943394,431.46878 72.94473,478.5392 111.97693,498.51857 110.97574,502.24354 73.277791,486.32771 75.946683,580.29857 150.67536,580.29857 163.01931,215.75954 5.5548981,269.60226 0.55074768,257.07286 160.35011,190.36203 z", Fill = Brushes.Green, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 1026, Height = 353, Top = 224, Left = 0, Thickness = 0, ZIndex = 2 };
            standardElement.ShapePresent = staticShape;
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
            listshape.Add(standardElement);

            return listshape;
        }

        public static void GenerateLayoutTheme08(SlideMaster slide, int maxZindex, EColorManagment color)
        {
            SolidColor colorAccent3 = color.Accent3;
            foreach (LayoutMaster layoutMaster in slide.LayoutMasters)
            {
                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();

                foreach (ObjectElement objectlayout in layoutMaster.MainLayout.Elements)
                {
                    objectlayout.ZIndex += (maxZindex + 3);
                }
                if (!((layoutMaster.LayoutName == "Blank Layout") || (layoutMaster.LayoutName == "Title Slide Layout")))
                {
                    layoutMaster.IsHideBackground = true;

                    StaticShape staticShape = new StaticShape() { PathData = "M347.5,1099.75 L347.5,1058 411,1058 414.5,966.25 375.5,979.25 373,976.75 413.5,960 415.33333,912 418.50036,912 420.33371,959.33333 459.83443,979.5 458.50143,983.16667 420.37594,965.91667 423.50098,1058.5 485.83511,1058.5 496.00227,753.66667 380.00049,793.49972 376.00043,783.99971 496.3356,733.99967 501.16935,592.33291 511.33583,592.33291 515.75289,732.99992 634.12989,791.81243 630.00481,800.37493 513.75283,751.49992 524.33635,1057.8333 1627.3551,1057.8333 1627.3551,1099.25 z", Fill = Brushes.Green, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 1026, Height = 353, Top = 224, Left = 0, Thickness = 0, ZIndex = 2 };
                    standardElement.ShapePresent = staticShape;
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }
            }
        }



        public static ObservableCollection<StandardElement> Theme09(EColorManagment color)
        {
            SolidColor colorAccent1 = color.Accent1;
            SolidColor colorAccent2 = color.Accent2;
            SolidColor colorAccent3 = color.Accent3;
            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();

            StaticShape staticShape = new StaticShape() { PathData = "M1142.403,3.9999962 L1139.2729,4.6666908 C1139.963,6.4444599 1141.613,8.1793203 1141.333,10 1141.193,10.9385 1139.203,11.3334 1138.1329,12.000008 L1131.7329,7.5416946 C1127.6429,9.6608896 1121.8529,12.0735 1116.803,11.041698 1114.2529,10.5215 1112.533,9.0139198 1110.403,7.9999924 L1108.2729,9.3333626 1106.1329,7.9999924 1105.0729,8.666687 1107.2029,10 C1102.2029,11.173 1096.3829,9.8326998 1091.2029,10.666695 L1089.0729,7.958374 C1084.0128,9.5878296 1077.6729,8.3044395 1072.0028,7.9999924 1064.9229,7.6195102 1056.9928,10.0369 1050.6727,7.9999924 L1048.5327,9.3333626 1046.4027,7.9999924 1044.2727,9.3333626 1042.1327,7.9999924 C1040.1028,10.8205 1033.6327,11.5473 1028.7327,12.000008 1027.8727,10.8889 1027.4827,9.5719604 1026.1327,8.666687 1024.7227,7.71771 1022.5827,7.3333702 1020.8027,6.6666985 L1019.7327,8.666687 C1013.0226,9.1623497 1005.8826,7.9963398 999.4696,9.3333626 L997.33563,7.3749924 C990.05762,8.2497597 983.30859,10.5124 976.00256,11.291695 973.86157,11.5201 971.73553,10.7778 969.60254,10.520897 967.46954,10.2639 965.37451,9.7731304 963.20251,9.7500038 961.39154,9.73071 959.64752,10.1806 957.86951,10.395908 956.09149,10.6111 954.25751,11.3932 952.53552,11.041698 949.98846,10.5215 948.26947,9.0139198 946.1355,7.9999924 L941.86945,10.666695 C940.80249,10.3056 939.75848,9.91675 938.66943,9.5833778 935.49048,8.6107798 932.54742,7.15765 929.06946,6.7499924 927.53442,6.57019 926.20746,7.5328999 924.80243,7.958374 923.36243,8.39429 922.13043,9.2477398 920.5354,9.3333626 912.03339,9.7899799 903.44934,8.9675303 894.93536,9.3333626 L892.80237,6.708374 888.53534,7.9999924 886.40234,5.3333664 885.33533,6.4583778 C882.13531,6.9860802 879.04431,7.99194 875.73529,8.041687 874.13629,8.06567 872.8913,7.125 871.4693,6.6666985 L869.33527,9.3333626 867.20227,7.6250076 860.80225,10.666695 C857.80225,8.0297899 850.59521,6.6799302 845.8692,7.9999924 843.77423,8.58496 843.76221,10.8915 841.60223,11.375008 831.50818,13.6347 819.13214,6.2622099 809.60211,9.3333626 L807.46912,7.9999924 C805.33508,8.6735802 803.20209,9.34729 801.06909,10.020905 798.93512,10.6945 797.00507,12.3578 794.66907,12.041702 791.18109,11.5697 788.98004,9.34729 786.13507,7.9999924 785.42407,8.59729 785.11304,9.52386 784.00208,9.7916985 781.19006,10.4694 777.83002,11.4218 775.06903,10.666695 772.62701,9.9989004 773.25201,7.3566298 771.20203,6.2916946 769.51501,5.4152799 766.935,5.65277 764.802,5.3333664 L762.66901,7.9999924 760.53497,6.25 750.935,8.666687 753.06897,10 752.00195,11.041698 C742.93896,6.7049599 728.01392,6.0639601 717.8689,9.3333626 L715.73486,7.9999924 C713.60187,8.59729 711.67084,9.8568697 709.33484,9.7916985 706.25983,9.7059298 703.88086,7.6033902 700.80182,7.5833702 698.98383,7.5715299 698.18585,9.2312002 696.53485,9.7083664 694.54382,10.2839 692.2688,10.3473 690.13483,10.666695 L688.00183,7.9999924 683.7348,10.708408 679.46875,9.3333626 677.33478,10.666695 C672.26978,8.8200102 665.48071,6.3204298 660.26874,7.9999924 L658.1347,6.6666985 654.86871,8.666687 657.06873,11.333408 653.86871,13.333397 651.73468,12.000008 649.60168,12.79171 C647.82367,12.6598 646.04669,12.5278 644.26868,12.395897 642.49066,12.2639 640.23865,12.7672 638.93469,12.000008 637.29865,11.0363 637.5127,9.3333702 636.8017,7.9999924 L634.66864,10.666695 C628.47766,9.5583496 621.73065,8.39569 615.46863,9.3333626 L614.40161,11.333408 606.93457,12.000008 604.80157,9.7083664 598.40155,10.666695 C596.65857,13.5939 590.26257,16.0588 585.40155,15.333405 585.62354,14 585.21753,12.5638 586.06854,11.333408 587.13153,9.7958403 589.31256,8.6666899 590.93457,7.333374 L589.86853,6.6250038 C582.9425,7.7265 575.67249,7.7352901 568.53448,7.958374 562.14246,8.1581402 555.02344,9.83319 549.33441,7.9999924 L547.20142,9.3333626 C541.91541,8.5332003 536.36139,6.25842 531.20142,7.333374 L533.33441,8.666687 C532.26837,8.90277 531.25537,9.4750404 530.1344,9.375 528.55035,9.2335796 527.46136,7.90979 525.86835,7.9999924 523.49738,8.1342802 521.60138,9.3194599 519.46838,9.9791908 517.33435,10.6389 515.44135,11.8491 513.06836,11.958408 510.14932,12.0927 507.05133,11.6005 504.53433,10.666695 503.16733,10.1593 503.1123,8.8889198 502.40131,7.9999924 L500.26831,9.375 496.00128,7.9999924 493.86829,9.7916985 C491.02328,8.3056002 488.70529,6.29877 485.33426,5.3333664 483.54926,4.8220201 481.42325,5.5556002 479.46826,5.6666946 477.51224,5.77777 475.55725,5.8889198 473.60123,6.0000038 L475.73425,7.333374 C465.3562,8.1666899 454.96518,9.1146202 444.80115,10.666695 L442.66815,7.9999924 440.53415,9.3333626 436.26813,6.6666985 434.13412,7.9999924 427.73413,5.3333664 427.53412,6.0000038 428.80112,12.666702 C422.16211,13.5485 414.91309,10.6711 408.46808,12.000008 L404.26807,7.9999924 397.86804,10.666695 396.80103,9.6249962 391.46802,7.9999924 390.40103,10 385.06799,10.666695 384.00101,8.666687 C382.22299,8.4306002 380.474,7.8329501 378.668,7.958374 377.077,8.0687904 375.82297,8.875 374.40097,9.3333626 L372.26797,6.6666985 370.13397,7.9999924 368.00095,6.2916946 361.60095,9.3333626 357.33392,7.5 C354.48993,9.0139198 352.53293,11.9323 348.8009,12.041702 345.2359,12.1461 343.11191,9.34729 340.26788,7.9999924 L338.13388,9.3333626 336.00089,7.9999924 C327.51086,8.5638399 318.70383,9.7432899 310.40082,8.5000038 307.3898,8.0491896 304.9588,6.3148799 301.8678,6.2083626 298.55179,6.09412 295.48178,7.3519301 292.26776,7.8750038 289.08075,8.3937397 285.94376,9.5441303 282.66772,9.3333626 279.51373,9.1304903 277.27972,6.9541001 274.1337,6.708374 271.22971,6.4814501 268.4447,7.5694599 265.60068,7.9999924 265.24469,8.6666899 265.20868,9.43787 264.53369,10 263.0257,11.2571 261.62668,12.9199 259.20068,13.333397 257.03168,13.7031 254.39066,12.9938 252.80066,12.000008 251.21065,11.0062 251.37865,9.3333702 250.66765,7.9999924 247.82265,8.9583702 245.36064,10.9853 242.13362,10.874996 238.74861,10.7593 236.98961,7.3599901 233.6006,7.416687 230.6926,7.4653301 230.07359,10.8405 227.20059,11.124992 224.15558,11.4266 221.51157,9.6389198 218.66757,8.895874 215.82256,8.15277 212.97855,7.40979 210.13354,6.6666985 L209.06754,8.666687 C207.28954,8.90277 205.54053,9.5004301 203.73352,9.375 202.14352,9.2645903 200.88953,8.4583702 199.4675,7.9999924 L197.33351,8.8750076 C194.4895,8.5694599 191.68149,7.8551602 188.80049,7.958374 187.20248,8.0155602 186.12148,9.4574003 184.53348,9.3333626 182.75647,9.19452 181.99448,7.5885601 180.26747,7.2916985 178.20746,6.9376798 175.89246,7.0748301 173.86746,7.5 171.48744,7.9995098 169.60043,9.15277 167.46744,9.9791908 165.33342,10.8056 163.54942,12.6887 161.06741,12.458401 158.20441,12.1928 157.19841,9.6681499 154.66739,8.7916946 153.10139,8.2494497 151.11139,8.52777 149.33337,8.3958626 147.55638,8.2639198 145.77837,8.1319599 144.00037,7.9999924 L141.86737,10.249996 C134.16835,9.3311796 125.36632,7.0885 118.40031,9.3333626 L116.2673,7.9999924 C109.96828,8.8366098 102.79026,7.4889498 97.066956,9.3333626 L94.933548,7.9999924 92.800232,8.9583778 86.400223,7.9999924 84.266907,10.666695 82.133499,9.3333626 80.000191,10.666695 C76.677689,7.74652 69.63047,7.5022001 65.066872,5.3333664 L62.933464,7.9999924 C57.326839,7.3977098 50.923122,6.3705401 45.866814,7.9999924 L43.733406,6.6250038 26.666756,12.000008 C21.714447,8.9047899 13.041624,9.3161602 7.4666786,6.6666985 L6.3999939,7.333374 C7.3999991,8.22229 9.3782845,8.9135103 9.4000053,10 9.429244,11.4609 6.7370973,12.5446 6.5333748,13.999996 6.3972564,14.9722 7.7860203,15.7885 8.4666824,16.666698 9.1642342,17.566799 10.971019,18.3515 10.666714,19.333401 10.179016,20.906601 7.8222504,22 6.3999939,23.333397 10.656618,26.3204 9.3330545,31.7782 6.3999939,35.333405 L10.666714,37.999992 9.4000053,39.333401 C9.7021351,46.8867 10.576218,54.493401 9.2000198,62.000008 8.9904633,63.143101 8.2444611,64.222298 7.7666664,65.333405 7.288919,66.444504 5.9482155,67.541603 6.3333511,68.666702 7.0028782,70.622597 10.720918,71.999901 10.733318,74 10.745818,75.993797 7.8444605,77.555603 6.3999939,79.333405 L8.6000252,80.666695 C7.8666801,81.555603 6.0956454,82.351501 6.3999939,83.333397 6.8877077,84.906601 10.089016,85.772102 10.666714,87.333389 11.014419,88.2733 9.0054731,89.0392 8.8666916,90 8.0904312,95.374001 8.8433933,101.204 12.800026,106 L11.066723,106.66701 C10.377817,108.222 9.7703056,109.793 9.0000153,111.33301 8.3249617,112.683 6.7976174,113.919 6.7333412,115.33301 6.5807571,118.69 12.744023,121.309 12.800026,124.66699 12.889323,130.022 3.6726992,134.787 5.3333473,140.04199 L9.600029,138.66701 11.400013,139.33301 C11.866721,141.11099 13.377725,142.90199 12.800026,144.66702 12.468223,145.681 9.9936056,145.839 9.0000153,146.66701 7.8133302,147.65601 7.2666688,148.88901 6.3999939,150 7.4666791,150.222 8.8449831,150.146 9.600029,150.66699 10.928119,151.58299 12.009921,152.77 12.200012,154 12.410422,155.362 11.444819,156.70599 10.733318,158 10.213217,158.946 8.6248627,159.668 8.5333633,160.66699 8.3682718,162.468 9.5111246,164.222 10.000019,166 10.488917,167.778 11.955621,169.556 11.466713,171.33299 10.566117,174.608 5.3704538,177.35899 5.8666801,180.66699 6.628027,185.742 13.589325,190.259 12.800026,195.33298 12.600523,196.616 10.711118,197.556 9.6666718,198.66701 8.6222324,199.778 6.8791876,200.73 6.5333748,202 6.0249953,203.867 6.9888978,205.778 7.2166824,207.66701 8.0624208,214.67999 11.137519,223.26401 3.9333153,228.66701 8.2578211,232.886 11.083219,238.463 9.2000198,243.33299 8.5576019,244.995 4.5484915,245.651 4.0666771,247.33299 3.7063594,248.592 6.0889058,249.556 7.1000099,250.66699 8.1111212,251.778 9.7482452,252.745 10.133324,254 11.889721,259.72601 10.674818,266.15201 6.3999939,271.33301 L8.5333633,272.66699 7.0666885,274 C7.3999991,276 7.7333498,278 8.0666924,280 8.4000015,282 9.5558949,284.013 9.0666962,286 8.7983828,287.09 6.0084553,287.565 5.9333229,288.66699 5.8209848,290.315 7.68891,291.77802 8.5666847,293.33301 9.4444447,294.88901 11.31232,296.35199 11.200027,298 11.124919,299.10199 8.3342419,299.577 8.0666924,300.66699 7.0959282,304.62201 11.467319,309.10999 8.5333633,312.66699 L9.7333527,314 C9.6081047,319.63599 9.6157646,325.302 8.4222221,330.88901 7.5437994,335.00101 5.4916139,339.646 8.5333633,343.33301 L6.3999939,344.66699 8.4000015,346 C7.5630293,351.754 7.049418,357.63699 8.6000252,363.33301 8.9007435,364.43799 8.1333408,365.556 7.9000092,366.66699 7.6666698,367.77802 7.2000184,368.879 7.2000122,370 7.2000184,371.121 7.6666698,372.22198 7.9000092,373.33301 8.1333408,374.444 8.8768129,375.55899 8.6000252,376.66699 8.3530111,377.655 6.1820459,378.34299 6.3999939,379.33301 6.7349272,380.85599 9.5850048,381.82599 10.066719,383.33301 12.302822,390.32901 7.4479795,397.53299 7.3333549,404.66699 7.3118386,406.005 7.68891,407.33301 7.8666878,408.66699 8.0444708,410 7.8444004,411.375 8.4000015,412.66699 8.8207226,413.64499 10.991718,414.33301 10.733318,415.33301 10.223216,417.30801 6.8754678,418.69501 6.3333511,420.66699 6.0622954,421.65201 8.4335423,422.33499 8.5333633,423.33301 8.977273,427.77301 10.119116,432.241 9.4000053,436.66699 9.1694136,438.086 7.7777801,439.33301 6.9666862,440.66699 6.1555557,442 4.5214515,443.23999 4.5333672,444.66699 4.5452614,446.09799 6.200016,447.33301 7.0333672,448.66699 7.8666801,450 9.3608847,451.23901 9.5333672,452.66699 9.9817057,456.37701 6.9951582,460.026 4.2666626,463.33301 9.3970747,465.776 6.8681378,471.30399 6.6666603,475.33301 6.606997,476.52701 8.0666904,477.556 8.7666702,478.66699 9.4666843,479.77802 10.939618,480.80701 10.866718,482 10.804018,483.026 8.4257021,483.64001 8.4000015,484.66699 8.3746719,485.67999 10.708017,486.32001 10.733318,487.33301 10.758318,488.33301 9.266674,489.11099 8.5333633,490 L11.000023,491.33301 C10.433317,492.444 9.8666859,493.556 9.3000412,494.66699 8.7333527,495.77802 7.303319,496.849 7.6000214,498 8.0980606,499.93301 11.181719,501.388 11.533413,503.33301 11.75172,504.54099 9.9555855,505.556 9.1666794,506.66699 8.3777914,507.77802 7.0045881,508.79099 6.8000221,510 6.0519156,514.42102 9.6595354,519.38202 6.3999939,523.33301 L8.5333633,524.66699 7.3999977,526 C7.5158195,529.85999 7.8269401,533.729 8.6444473,537.55603 8.9925928,539.185 9.340744,540.815 9.6889114,542.44397 10.037016,544.07397 11.56112,545.77301 10.733318,547.33301 9.6102753,549.45099 6.4222164,550.88898 4.2666626,552.66699 12.114821,556.495 8.8292027,565.80298 4.2666626,571.33301 8.6030426,573.677 9.0274534,578.815 6.3999939,582 L9.2000198,583.33301 C8.1555614,588.22198 4.6948519,593.14301 6.0666847,598 6.4321365,599.29401 8.311141,600.22198 9.4333267,601.33301 10.555617,602.44397 12.568522,603.36102 12.800026,604.66699 13.046924,606.05902 11.61642,607.37903 10.733318,608.66699 9.4766645,610.49902 7.8444605,612.22198 6.3999939,614 L8.5333633,615.33301 C7.5111194,617.11102 5.5434041,618.77802 5.466671,620.66699 5.2843337,625.15503 11.825821,630.00897 8.5333633,634 L10.666714,635.33301 C8.4244213,638.65503 9.7736053,643.26599 13.533421,646 L6.3999939,651.33301 C10.396917,653.617 11.931421,657.32599 12.800026,660.66699 L5.3333473,661.33301 4.2666626,666 5.3333473,666.66699 8.5333633,668.66699 6.4666748,670 12.800026,675.33301 11.133327,676 C9.2337742,680.11603 6.0221953,684.508 7.6666641,688.66699 8.1333408,689.84698 9.2444639,690.88898 10.033417,692 10.822218,693.11102 12.278722,694.12 12.400017,695.33301 12.851624,699.849 8.8298731,704.15302 8.3333588,708.66699 7.8436103,713.11902 9.6723146,718.034 6.3999939,722 7.288919,722.44397 8.5449924,722.70099 9.0666962,723.33301 10.079617,724.56097 11.107819,725.96503 10.800018,727.33301 10.460717,728.84198 7.80055,729.84698 7.266674,731.33301 4.6066017,738.73798 9.2844639,746.461 10.666714,754 11.233319,757.091 8.3854218,760.56799 10.666714,763.33301 L6.7333412,764.66699 C7.1111183,766 7.2373486,767.37097 7.8666878,768.66699 8.7561131,770.49799 11.56592,772.09497 11.266727,774 11.004719,775.66699 8.0222607,776.66699 6.3999939,778 L10.666714,780.66699 C7.7119899,783.83197 9.0874138,788.07599 6.3999939,791.33301 L9.600029,793.29199 C12.444522,792.43103 15.013229,791.02301 18.133335,790.70801 19.693441,790.55103 20.977844,791.56897 22.400055,792 L24.533463,788 30.93338,789.33301 32.000065,791.33301 C33.777878,791.55603 35.858185,791.341 37.333393,792 38.627392,792.578 38.755692,793.77802 39.466801,794.66699 41.600098,794.20801 43.610905,793.29199 45.866814,793.29199 47.466915,793.29199 48.71122,794.20801 50.133419,794.66699 L52.266827,793.33301 58.66684,794.66699 60.800152,792 65.066872,795 C70.029671,792.72198 76.018089,790.13501 82.133499,790.58301 84.543213,790.76001 86.157616,792.44299 88.533516,792.75 90.978828,793.06598 93.511337,792.5 96.000252,792.375 98.489151,792.25 100.97826,792.125 103.46725,792 L104.53327,794 304.00079,794.66699 305.06778,792.66699 C307.55679,792.33301 310.0098,791.461 312.53381,791.66699 314.28381,791.81 315.37881,792.97198 316.80081,793.625 321.76685,791.44299 328.90585,793.01001 334.93387,792.56299 336.71188,792.43103 338.61987,792.60498 340.26788,792.16699 343.3429,791.349 345.47791,788.64801 348.8009,788.875 352.48593,789.12701 354.48993,791.84698 357.33392,793.33301 L359.46793,791.625 C361.60095,792.63898 363.25793,794.28699 365.86795,794.66699 367.73196,794.93799 369.42297,793.77802 371.20096,793.33301 372.97897,792.88898 374.75699,792.44397 376.534,792 L377.60098,794 381.401,794.66699 C382.26801,793.55603 382.73898,792.28497 384.00101,791.33301 384.76801,790.755 386.134,790.63898 387.20102,790.29199 L393.60101,793.33301 C393.95703,792.66699 393.93204,791.86499 394.66803,791.33301 395.46103,790.76001 396.80103,790.61102 397.86804,790.25 L402.13406,792.375 C408.78207,789.172 418.14108,789.07098 425.6011,786.66699 L427.26813,787.33301 429.86813,794.66699 C434.90714,793.427 439.98816,792.21802 444.80115,790.66699 L446.93417,792 449.06818,790.66699 451.20117,793.33301 453.3342,792 C460.09119,793.38599 467.58023,792.36298 474.66824,792 L476.80124,796.375 C481.58627,795.60901 487.57129,794.52301 489.60129,791.70801 496.41031,793.21503 503.71432,789.45398 510.93433,789.33301 517.00836,789.23199 522.60138,791.59302 528.0014,793.33301 L530.1344,792 532.26837,794.66699 534.40137,793.29199 538.6684,794.66699 C543.55841,790.27301 556.65747,790.88098 564.26849,793.33301 L566.40149,792 568.53448,793.66699 C574.5545,790.90399 583.84955,795.47601 590.93457,794 589.66852,792.44397 588.38354,790.89398 587.13452,789.33301 586.25055,788.22803 585.40155,787.11102 584.53455,786 L587.73456,784 588.80151,785.04199 C590.57953,785.58301 592.66956,785.83698 594.13458,786.66699 594.96857,787.13898 594.84656,788 595.20154,788.66699 596.26855,788.76398 597.45953,788.63098 598.40155,788.95801 600.74457,789.77197 602.66858,790.98602 604.80157,792 605.15759,791.33301 605.29858,790.60602 605.86859,790 606.74158,789.07001 608.00159,788.30603 609.0686,787.45801 614.20361,790.29999 621.47363,792.66699 628.26862,792.08301 629.87463,791.94501 630.91467,790.64099 632.53467,790.625 634.13464,790.60901 635.37964,791.54199 636.8017,792 L637.86865,790.91699 C643.77869,789.20801 650.59668,788.21698 657.06873,788.66699 L654.93469,792.66699 C662.26172,795.90997 672.79779,789.19098 681.60181,790.29199 684.2298,790.62 685.86877,792.31897 688.00183,793.33301 L690.13483,791.04199 C710.69885,792.02301 732.02393,793.86603 752.00195,790.66699 753.21295,792.93701 759.26501,794.42999 762.66901,793.33301 L764.802,796 C770.03802,794.02301 774.828,791.61603 779.73505,789.33301 780.44604,789.90302 780.83502,790.71899 781.86908,791.04199 783.83105,791.65503 786.13507,791.68103 788.26904,792 L790.4021,789.33301 792.5351,792 C792.89105,791.33301 792.97809,790.58502 793.60205,790 794.43909,789.216 795.73511,788.66699 796.80206,788 L797.86908,790.54199 C805.27112,791.15399 813.63214,792.80499 820.26917,790.66699 L822.40216,791.875 C828.43018,790.28601 835.39618,790.62598 841.60223,789.33301 841.95819,789.98602 841.74921,790.91101 842.66919,791.29199 845.55225,792.48499 848.81323,793.66699 852.26923,793.75 854.08624,793.79401 854.88525,792.10199 856.53522,791.625 858.52722,791.04901 860.80225,790.98602 862.93524,790.66699 L865.06927,793.29199 869.33527,792 871.4693,794.33301 C874.31329,793.84698 877.06628,792.703 880.00232,792.875 881.84131,792.98297 882.84729,794.34698 884.26935,795.08301 893.25934,790.60602 906.95636,792.67102 918.4024,792 920.56342,791.87299 922.66943,792.5 924.80243,792.75 926.93542,793 929.08643,793.80103 931.20245,793.5 934.35345,793.052 936.55945,790.99701 939.73547,790.625 941.30847,790.44098 942.40948,791.909 944.0025,792 946.85651,792.16302 949.69147,791.56897 952.53552,791.354 955.38049,791.13898 958.2915,791.14697 961.06952,790.70801 962.62054,790.46301 963.7395,789.26099 965.33551,789.33301 968.35852,789.47101 970.89154,790.94098 973.86957,791.29199 976.66455,791.62097 979.55859,791.32501 982.40259,791.34198 995.2666,791.41699 1010.8926,788.62201 1020.8027,793.75 1022.5827,792.86102 1024.3627,791.97198 1026.1327,791.08301 1027.9127,790.19397 1029.3228,788.89203 1031.4727,788.41699 1032.8328,788.11401 1034.4628,788.52899 1035.7327,788.95801 1038.0928,789.75098 1039.5728,791.50403 1042.1327,792 1044.5227,792.46198 1047.1128,791.76398 1049.6028,791.646 1059.9227,791.15601 1070.1829,793.33301 1080.5328,793.33301 1082.1328,793.33301 1083.2029,791.95801 1084.8029,791.95801 1086.4028,791.95801 1087.6428,792.875 1089.0729,793.33301 L1091.2029,790.91699 1101.8729,792 1104.0029,789.33301 1108.2729,792.375 1114.673,789.33301 1116.803,790.83301 C1121.3729,789.09399 1127.7229,791.64398 1132.333,793.33301 1133.203,792.22198 1133.553,790.883 1134.933,790 1135.823,789.43201 1137.3829,789.55603 1138.603,789.33301 1139.533,790.44397 1141.2729,791.414 1141.403,792.66699 1141.5031,793.66498 1139.933,794.44397 1139.203,795.33301 L1140.2729,796 1145.603,795.33301 1143.473,794 1144.533,792 C1153.543,793.58301 1163.0431,790.28003 1172.2731,789.33301 L1174.4031,791.29199 1178.6731,788.75 C1183.0431,791.26898 1189.8231,791.63397 1195.7332,792 1196.0931,791.33301 1196.2131,790.59698 1196.8031,790 1197.6631,789.13898 1198.9331,788.47198 1200.0032,787.70801 L1210.6732,792.41699 1214.9332,790.29199 C1217.0732,791.31897 1218.7231,792.95697 1221.3333,793.375 1222.8832,793.62299 1224.0132,792.12402 1225.6033,792 1239.7133,790.89697 1254.0533,792.26202 1268.2733,792 1268.6433,790.88898 1269.6034,789.79602 1269.4033,788.66699 1269.2234,787.67297 1266.9033,786.98199 1267.2034,786 1267.6934,784.427 1271.1833,783.59302 1271.4734,782 1271.8033,780.15302 1268.9333,778.52502 1268.8733,776.66699 1268.8234,775.25299 1271.2533,774.07898 1271.1333,772.66699 1270.9734,770.78003 1268.6133,769.19202 1268.0734,767.33301 1267.1533,764.18799 1271.5233,761.19702 1271.6034,758 1271.6333,756.987 1269.2933,756.34601 1269.2733,755.33301 1269.2433,754.33301 1270.7333,753.55603 1271.4734,752.66699 L1269.3334,751.33301 1271.4734,750 C1269.9333,748.66699 1267.5233,747.591 1266.8733,746 1265.0834,741.64099 1271.9933,737.03699 1270.2733,732.66699 1269.8033,731.49298 1266.5333,731.203 1266.3334,730 1266.1033,728.54602 1268.2433,727.33301 1269.2034,726 1270.1633,724.66699 1272.1533,723.46002 1272.0734,722 1272.0033,720.88098 1269.8933,720.22198 1268.8033,719.33301 1270.3633,718 1272.8134,716.93201 1273.4734,715.33301 1273.8534,714.388 1272.2134,713.54498 1271.5333,712.66699 1270.8434,711.76703 1269.1233,710.99103 1269.3334,710 1269.5734,708.88202 1272.0734,708.39697 1272.6733,707.33301 1273.1433,706.492 1272.9634,705.46899 1272.3334,704.66699 1271.5033,703.59497 1269.3134,703.08801 1268.5333,702 1267.6633,700.76801 1267.9133,699.33301 1267.6033,698 1266.6533,693.92297 1273.7134,690.02802 1272.3334,686 1271.7334,684.25 1268.5133,683.33301 1266.6033,682 1268.2433,680.66699 1270.8633,679.63098 1271.5333,678 1271.9333,677.03101 1269.4333,676.33099 1269.3334,675.33301 1269.0233,672.21198 1267.8633,669.12701 1267.8033,666 1267.7833,664.63202 1268.5634,663.29401 1269.2733,662 1269.7933,661.05402 1271.6233,660.32898 1271.4734,659.33301 1271.3033,658.25098 1268.7234,657.742 1268.4734,656.66699 1268.2034,655.51898 1269.5333,654.44397 1270.0734,653.33301 1270.6034,652.22198 1271.8334,651.15503 1271.6733,650 1271.2833,647.29999 1271.3234,644.409 1269.3334,642 L1270.6034,640.66699 C1269.7234,635.35901 1271.3833,630.00299 1271.4734,624.66699 1271.5033,622.83801 1270.3334,621.06897 1269.4033,619.33301 1268.8933,618.38501 1267.2733,617.66602 1267.2034,616.66699 1267.1133,615.28497 1268.3833,614 1268.9734,612.66699 1269.5634,611.33301 1270.6934,610.04999 1270.7333,608.66699 1270.7933,606.86298 1269.1034,605.13397 1269.2733,603.33301 1269.3633,602.33502 1270.7333,601.55603 1271.4734,600.66699 L1269.3334,599.33301 1271.4734,598 1270.2034,596.66699 C1270.7933,593.12598 1271.9434,589.164 1269.3334,586 L1271.5333,584.66699 1269.3334,582 1270.7333,580.66699 C1269.7633,577.13501 1271.8033,572.974 1268.6033,570 1272.2334,567.26099 1268.5433,562.88702 1268.7333,559.33301 1268.8334,557.49799 1270.6733,555.828 1270.9333,554 1271.4534,550.45801 1268.8633,546.495 1271.4734,543.33301 L1269.3334,542 C1270.2034,537.55603 1272.5333,533.12903 1271.9333,528.66699 1271.7833,527.49103 1270.6733,526.44397 1270.0333,525.33301 1269.4033,524.22198 1268.3234,523.17401 1268.1333,522 1266.8134,513.495 1270.6733,504.94101 1270.2733,496.39999 1270.0033,490.79999 1269.7333,485.20001 1269.4734,479.60001 1269.2933,475.957 1272.0133,471.90601 1269.3334,468.66699 L1271.4734,467.33301 C1266.4333,464.185 1270.8534,458.43399 1270.2034,454 1269.9933,452.54501 1267.4833,451.45801 1267.3334,450 1267.1533,448.181 1268.3733,446.401 1269.2733,444.66699 1269.7633,443.715 1270.7333,442.88901 1271.4734,442 L1270.4033,441.33301 C1270.1733,429.99899 1269.6133,418.60901 1271.4734,407.33301 1271.6934,405.98801 1270.7333,404.66699 1270.3733,403.33301 1270.0033,402 1268.8933,400.66599 1269.2733,399.33301 1269.6833,397.85199 1272.6034,396.83701 1272.6034,395.33301 1272.6034,393.82999 1269.5233,392.828 1269.2733,391.33301 1269.1034,390.33899 1271.4734,389.66699 1271.4734,388.66699 1271.4734,386.83099 1269.2733,385.16901 1269.2733,383.33301 1269.2733,382.33301 1270.7333,381.556 1271.4734,380.66699 L1267.2034,378 1268.2733,377.33301 C1268.4333,371.31299 1270.2933,365.35001 1270.6733,359.33301 1270.7234,358.436 1270.4933,357.52499 1270.0734,356.66699 1269.3933,355.28601 1267.6433,354.103 1267.4033,352.66699 1267.1733,351.31201 1268.1133,349.97198 1268.7333,348.66699 1269.3833,347.297 1271.3434,346.09299 1271.2034,344.66699 1271.0133,342.75699 1268.2034,341.23099 1267.8033,339.33301 1267.5133,337.979 1268.7633,336.66699 1269.2333,335.33301 1269.7134,334 1270.5533,332.698 1270.6733,331.33301 1271.3334,323.77399 1267.8234,316.23099 1268.3334,308.66699 1268.4233,307.31 1269.1633,306 1269.5734,304.66699 1269.9833,303.33301 1270.7333,302.02399 1270.8033,300.66699 1271.0233,296.66699 1268.5333,292.224 1271.4734,288.66699 L1270.4033,288 C1270.3134,286 1271.0133,283.92499 1270.1333,282 1269.6233,280.87201 1267.6934,280.22198 1266.4733,279.33301 1270.3234,276.423 1272.7733,272.44299 1272.6034,268.66699 1272.5433,267.289 1271.4933,266 1270.9333,264.66699 1270.3833,263.33301 1269.1633,262.043 1269.2733,260.66699 1269.3534,259.668 1271.5734,258.99799 1271.4734,258 1271.3633,256.94601 1268.9133,256.383 1268.7333,255.33298 1268.5133,253.965 1269.8234,252.66701 1270.3733,251.33299 1270.9133,250 1272.3134,248.696 1272.0033,247.33299 1271.7533,246.241 1269.1333,245.757 1268.8733,244.66702 1268.4833,243.105 1269.7633,241.556 1270.2034,240 1270.6433,238.444 1271.7234,236.909 1271.5333,235.33298 1270.9634,230.45799 1267.8833,225.01199 1271.4734,220.66699 L1269.3334,219.33301 1271.4734,218 C1268.9634,213.81799 1272.6133,208.63699 1269.3334,204.66702 L1271.4734,203.33299 1269.3334,202 1270.5333,200.66699 C1270.3234,195.743 1268.6333,190.88901 1267.6733,186 1267.0433,182.81 1271.4534,179.88 1271.5333,176.66701 1271.7034,170.00101 1266.5834,162.592 1271.4734,156.66701 L1268.6033,155.33298 C1269.7833,154.444 1272.1133,153.821 1272.1333,152.66701 1272.1733,150.75301 1268.9534,149.242 1268.7333,147.33299 1268.5333,145.502 1270.7633,143.83299 1270.9333,142 1271.5133,135.787 1266.9033,128.86501 1271.4734,123.33299 L1269.8033,122.66701 1267.2034,119.333 1270.4033,118.66699 C1272.3833,112.307 1268.0533,105.79 1267.2034,99.333405 L1270.4033,98.666702 C1270.7333,98 1270.9634,97.308701 1271.4033,96.666702 1272.0333,95.747398 1272.8733,94.888901 1273.6034,94 L1269.3334,91.333389 1273.6034,88.666702 1271.4033,87.333389 1273.6034,84.66671 1269.3334,83.333397 C1270.1833,81.555603 1272.4033,79.823997 1271.8733,77.999992 1271.2533,75.9011 1268.0934,74.444504 1266.2034,72.666702 1272.5834,68.481903 1265.2733,60.996498 1267.1333,55.333405 1267.7833,53.374001 1270.9933,51.971298 1271.5333,50 1271.8033,49.014301 1270.0734,48.222301 1269.3334,47.333393 L1271.4734,46.000004 1269.3334,44.666691 1270.7333,43.333397 C1269.2633,38.287399 1268.0934,33.118698 1268.6333,27.999992 1268.8434,26 1268.6233,23.9638 1269.2733,22.000008 1269.5934,21.020201 1271.6733,20.325701 1271.4734,19.333401 1271.1633,17.822001 1268.2833,16.841101 1267.9333,15.333405 1267.5333,13.5499 1268.8733,11.7778 1269.3334,10 1267.5634,9.8194599 1265.7034,9.8275805 1264.0033,9.4583702 1262.4534,9.1215801 1261.3833,7.9327402 1259.7333,7.9166985 1258.1133,7.9008799 1257.0933,9.3592501 1255.4733,9.375 1253.8733,9.3906298 1252.6233,8.4583702 1251.2032,7.9999924 L1249.0732,9.3333626 1246.9332,7.7499962 C1238.5732,8.9342699 1228.9633,9.0338697 1221.2732,6.6666985 L1217.0732,10.666695 1212.8032,7.9999924 C1211.0232,8.84729 1209.5232,9.9913301 1207.4731,10.541706 1204.4431,11.3525 1201.0731,11.5139 1197.8732,12.000008 L1196.8031,10 C1195.7332,9.7916899 1194.6332,9.6501503 1193.6031,9.375 1192.1332,8.9810801 1190.9131,7.8356299 1189.3331,7.9999924 1187.1831,8.2243004 1185.8732,9.6731005 1184.0031,10.375004 1182.3031,11.011 1180.4431,11.4584 1178.6731,12.000008 L1176.5331,8.9166832 C1172.8931,10.7311 1165.963,12.0587 1162.6731,10 L1169.0731,6.0000038 1168.0031,5.3333664 C1164.8131,8.0707397 1157.8831,7.7821698 1153.073,9.3333626 L1150.933,8.3749962 1144.533,9.3333626 z M1.0666847,0 L1280,0.66669464 1278.9301,800 1.0666847,800 0,799.33301 z", Fill = Brushes.Orange, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 1026, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 5 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M36.625,591.375 C36.625,591.375 147.5,459.83333 231.5,52.75 231.5,52.75 283.5,380.5 426.625,591.25 z", Fill = Brushes.OrangeRed, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 248, Height = 302, Top = 275, Left = 3, Thickness = 0, ZIndex = 1 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M36.625,591.375 C36.625,591.375 147.5,459.83333 231.5,52.75 231.5,52.75 283.5,380.5 426.625,591.25 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#1e1e1e") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 161, Height = 178, Top = 398, Left = 43, Thickness = 0, ZIndex = 3, Opacity = 0.7 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M452,777.66667 C452,777.66667 386.00022,672.9998 352.83333,491.83302 352.83333,491.83302 434.00006,713.49976 474.49992,706.49977 474.49992,706.49977 458.97392,769.08824 452.01561,777.67156 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#1e1e1e") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 95, Height = 178, Top = 392, Left = 126, Thickness = 0, ZIndex = 4, Opacity = 0.4 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M551.16698,958.41667 C551.16698,958.41667 587.83391,976.83273 673.83497,744.16542 673.83497,744.16542 637.68888,936.43644 574.68787,1029.4372 574.68787,1029.4372 561.08365,1007.3333 551.16698,958.41667 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#1e1e1e") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 95, Height = 178, Top = 392, Left = 28, Thickness = 0, ZIndex = 4, Opacity = 0.4 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent2.Color), ColorSpecialName = colorAccent2.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            return listshape;
        }

        public static void GenerateLayoutTheme09(SlideMaster slide, int maxZindex, EColorManagment color)
        {
            SolidColor colorAccent1 = color.Accent1;
            SolidColor colorAccent2 = color.Accent2;
            SolidColor colorAccent3 = color.Accent3;

            foreach (LayoutMaster layoutMaster in slide.LayoutMasters)
            {
                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();

                foreach (ObjectElement objectlayout in layoutMaster.MainLayout.Elements)
                {
                    objectlayout.ZIndex += (maxZindex + 3);
                }
                if (!((layoutMaster.LayoutName == "Blank Layout") || (layoutMaster.LayoutName == "Title Slide Layout")))
                {
                    layoutMaster.IsHideBackground = true;

                    StaticShape staticShape = new StaticShape() { PathData = "M1142.403,3.9999962 L1139.2729,4.6666908 C1139.963,6.4444599 1141.613,8.1793203 1141.333,10 1141.193,10.9385 1139.203,11.3334 1138.1329,12.000008 L1131.7329,7.5416946 C1127.6429,9.6608896 1121.8529,12.0735 1116.803,11.041698 1114.2529,10.5215 1112.533,9.0139198 1110.403,7.9999924 L1108.2729,9.3333626 1106.1329,7.9999924 1105.0729,8.666687 1107.2029,10 C1102.2029,11.173 1096.3829,9.8326998 1091.2029,10.666695 L1089.0729,7.958374 C1084.0128,9.5878296 1077.6729,8.3044395 1072.0028,7.9999924 1064.9229,7.6195102 1056.9928,10.0369 1050.6727,7.9999924 L1048.5327,9.3333626 1046.4027,7.9999924 1044.2727,9.3333626 1042.1327,7.9999924 C1040.1028,10.8205 1033.6327,11.5473 1028.7327,12.000008 1027.8727,10.8889 1027.4827,9.5719604 1026.1327,8.666687 1024.7227,7.71771 1022.5827,7.3333702 1020.8027,6.6666985 L1019.7327,8.666687 C1013.0226,9.1623497 1005.8826,7.9963398 999.4696,9.3333626 L997.33563,7.3749924 C990.05762,8.2497597 983.30859,10.5124 976.00256,11.291695 973.86157,11.5201 971.73553,10.7778 969.60254,10.520897 967.46954,10.2639 965.37451,9.7731304 963.20251,9.7500038 961.39154,9.73071 959.64752,10.1806 957.86951,10.395908 956.09149,10.6111 954.25751,11.3932 952.53552,11.041698 949.98846,10.5215 948.26947,9.0139198 946.1355,7.9999924 L941.86945,10.666695 C940.80249,10.3056 939.75848,9.91675 938.66943,9.5833778 935.49048,8.6107798 932.54742,7.15765 929.06946,6.7499924 927.53442,6.57019 926.20746,7.5328999 924.80243,7.958374 923.36243,8.39429 922.13043,9.2477398 920.5354,9.3333626 912.03339,9.7899799 903.44934,8.9675303 894.93536,9.3333626 L892.80237,6.708374 888.53534,7.9999924 886.40234,5.3333664 885.33533,6.4583778 C882.13531,6.9860802 879.04431,7.99194 875.73529,8.041687 874.13629,8.06567 872.8913,7.125 871.4693,6.6666985 L869.33527,9.3333626 867.20227,7.6250076 860.80225,10.666695 C857.80225,8.0297899 850.59521,6.6799302 845.8692,7.9999924 843.77423,8.58496 843.76221,10.8915 841.60223,11.375008 831.50818,13.6347 819.13214,6.2622099 809.60211,9.3333626 L807.46912,7.9999924 C805.33508,8.6735802 803.20209,9.34729 801.06909,10.020905 798.93512,10.6945 797.00507,12.3578 794.66907,12.041702 791.18109,11.5697 788.98004,9.34729 786.13507,7.9999924 785.42407,8.59729 785.11304,9.52386 784.00208,9.7916985 781.19006,10.4694 777.83002,11.4218 775.06903,10.666695 772.62701,9.9989004 773.25201,7.3566298 771.20203,6.2916946 769.51501,5.4152799 766.935,5.65277 764.802,5.3333664 L762.66901,7.9999924 760.53497,6.25 750.935,8.666687 753.06897,10 752.00195,11.041698 C742.93896,6.7049599 728.01392,6.0639601 717.8689,9.3333626 L715.73486,7.9999924 C713.60187,8.59729 711.67084,9.8568697 709.33484,9.7916985 706.25983,9.7059298 703.88086,7.6033902 700.80182,7.5833702 698.98383,7.5715299 698.18585,9.2312002 696.53485,9.7083664 694.54382,10.2839 692.2688,10.3473 690.13483,10.666695 L688.00183,7.9999924 683.7348,10.708408 679.46875,9.3333626 677.33478,10.666695 C672.26978,8.8200102 665.48071,6.3204298 660.26874,7.9999924 L658.1347,6.6666985 654.86871,8.666687 657.06873,11.333408 653.86871,13.333397 651.73468,12.000008 649.60168,12.79171 C647.82367,12.6598 646.04669,12.5278 644.26868,12.395897 642.49066,12.2639 640.23865,12.7672 638.93469,12.000008 637.29865,11.0363 637.5127,9.3333702 636.8017,7.9999924 L634.66864,10.666695 C628.47766,9.5583496 621.73065,8.39569 615.46863,9.3333626 L614.40161,11.333408 606.93457,12.000008 604.80157,9.7083664 598.40155,10.666695 C596.65857,13.5939 590.26257,16.0588 585.40155,15.333405 585.62354,14 585.21753,12.5638 586.06854,11.333408 587.13153,9.7958403 589.31256,8.6666899 590.93457,7.333374 L589.86853,6.6250038 C582.9425,7.7265 575.67249,7.7352901 568.53448,7.958374 562.14246,8.1581402 555.02344,9.83319 549.33441,7.9999924 L547.20142,9.3333626 C541.91541,8.5332003 536.36139,6.25842 531.20142,7.333374 L533.33441,8.666687 C532.26837,8.90277 531.25537,9.4750404 530.1344,9.375 528.55035,9.2335796 527.46136,7.90979 525.86835,7.9999924 523.49738,8.1342802 521.60138,9.3194599 519.46838,9.9791908 517.33435,10.6389 515.44135,11.8491 513.06836,11.958408 510.14932,12.0927 507.05133,11.6005 504.53433,10.666695 503.16733,10.1593 503.1123,8.8889198 502.40131,7.9999924 L500.26831,9.375 496.00128,7.9999924 493.86829,9.7916985 C491.02328,8.3056002 488.70529,6.29877 485.33426,5.3333664 483.54926,4.8220201 481.42325,5.5556002 479.46826,5.6666946 477.51224,5.77777 475.55725,5.8889198 473.60123,6.0000038 L475.73425,7.333374 C465.3562,8.1666899 454.96518,9.1146202 444.80115,10.666695 L442.66815,7.9999924 440.53415,9.3333626 436.26813,6.6666985 434.13412,7.9999924 427.73413,5.3333664 427.53412,6.0000038 428.80112,12.666702 C422.16211,13.5485 414.91309,10.6711 408.46808,12.000008 L404.26807,7.9999924 397.86804,10.666695 396.80103,9.6249962 391.46802,7.9999924 390.40103,10 385.06799,10.666695 384.00101,8.666687 C382.22299,8.4306002 380.474,7.8329501 378.668,7.958374 377.077,8.0687904 375.82297,8.875 374.40097,9.3333626 L372.26797,6.6666985 370.13397,7.9999924 368.00095,6.2916946 361.60095,9.3333626 357.33392,7.5 C354.48993,9.0139198 352.53293,11.9323 348.8009,12.041702 345.2359,12.1461 343.11191,9.34729 340.26788,7.9999924 L338.13388,9.3333626 336.00089,7.9999924 C327.51086,8.5638399 318.70383,9.7432899 310.40082,8.5000038 307.3898,8.0491896 304.9588,6.3148799 301.8678,6.2083626 298.55179,6.09412 295.48178,7.3519301 292.26776,7.8750038 289.08075,8.3937397 285.94376,9.5441303 282.66772,9.3333626 279.51373,9.1304903 277.27972,6.9541001 274.1337,6.708374 271.22971,6.4814501 268.4447,7.5694599 265.60068,7.9999924 265.24469,8.6666899 265.20868,9.43787 264.53369,10 263.0257,11.2571 261.62668,12.9199 259.20068,13.333397 257.03168,13.7031 254.39066,12.9938 252.80066,12.000008 251.21065,11.0062 251.37865,9.3333702 250.66765,7.9999924 247.82265,8.9583702 245.36064,10.9853 242.13362,10.874996 238.74861,10.7593 236.98961,7.3599901 233.6006,7.416687 230.6926,7.4653301 230.07359,10.8405 227.20059,11.124992 224.15558,11.4266 221.51157,9.6389198 218.66757,8.895874 215.82256,8.15277 212.97855,7.40979 210.13354,6.6666985 L209.06754,8.666687 C207.28954,8.90277 205.54053,9.5004301 203.73352,9.375 202.14352,9.2645903 200.88953,8.4583702 199.4675,7.9999924 L197.33351,8.8750076 C194.4895,8.5694599 191.68149,7.8551602 188.80049,7.958374 187.20248,8.0155602 186.12148,9.4574003 184.53348,9.3333626 182.75647,9.19452 181.99448,7.5885601 180.26747,7.2916985 178.20746,6.9376798 175.89246,7.0748301 173.86746,7.5 171.48744,7.9995098 169.60043,9.15277 167.46744,9.9791908 165.33342,10.8056 163.54942,12.6887 161.06741,12.458401 158.20441,12.1928 157.19841,9.6681499 154.66739,8.7916946 153.10139,8.2494497 151.11139,8.52777 149.33337,8.3958626 147.55638,8.2639198 145.77837,8.1319599 144.00037,7.9999924 L141.86737,10.249996 C134.16835,9.3311796 125.36632,7.0885 118.40031,9.3333626 L116.2673,7.9999924 C109.96828,8.8366098 102.79026,7.4889498 97.066956,9.3333626 L94.933548,7.9999924 92.800232,8.9583778 86.400223,7.9999924 84.266907,10.666695 82.133499,9.3333626 80.000191,10.666695 C76.677689,7.74652 69.63047,7.5022001 65.066872,5.3333664 L62.933464,7.9999924 C57.326839,7.3977098 50.923122,6.3705401 45.866814,7.9999924 L43.733406,6.6250038 26.666756,12.000008 C21.714447,8.9047899 13.041624,9.3161602 7.4666786,6.6666985 L6.3999939,7.333374 C7.3999991,8.22229 9.3782845,8.9135103 9.4000053,10 9.429244,11.4609 6.7370973,12.5446 6.5333748,13.999996 6.3972564,14.9722 7.7860203,15.7885 8.4666824,16.666698 9.1642342,17.566799 10.971019,18.3515 10.666714,19.333401 10.179016,20.906601 7.8222504,22 6.3999939,23.333397 10.656618,26.3204 9.3330545,31.7782 6.3999939,35.333405 L10.666714,37.999992 9.4000053,39.333401 C9.7021351,46.8867 10.576218,54.493401 9.2000198,62.000008 8.9904633,63.143101 8.2444611,64.222298 7.7666664,65.333405 7.288919,66.444504 5.9482155,67.541603 6.3333511,68.666702 7.0028782,70.622597 10.720918,71.999901 10.733318,74 10.745818,75.993797 7.8444605,77.555603 6.3999939,79.333405 L8.6000252,80.666695 C7.8666801,81.555603 6.0956454,82.351501 6.3999939,83.333397 6.8877077,84.906601 10.089016,85.772102 10.666714,87.333389 11.014419,88.2733 9.0054731,89.0392 8.8666916,90 8.0904312,95.374001 8.8433933,101.204 12.800026,106 L11.066723,106.66701 C10.377817,108.222 9.7703056,109.793 9.0000153,111.33301 8.3249617,112.683 6.7976174,113.919 6.7333412,115.33301 6.5807571,118.69 12.744023,121.309 12.800026,124.66699 12.889323,130.022 3.6726992,134.787 5.3333473,140.04199 L9.600029,138.66701 11.400013,139.33301 C11.866721,141.11099 13.377725,142.90199 12.800026,144.66702 12.468223,145.681 9.9936056,145.839 9.0000153,146.66701 7.8133302,147.65601 7.2666688,148.88901 6.3999939,150 7.4666791,150.222 8.8449831,150.146 9.600029,150.66699 10.928119,151.58299 12.009921,152.77 12.200012,154 12.410422,155.362 11.444819,156.70599 10.733318,158 10.213217,158.946 8.6248627,159.668 8.5333633,160.66699 8.3682718,162.468 9.5111246,164.222 10.000019,166 10.488917,167.778 11.955621,169.556 11.466713,171.33299 10.566117,174.608 5.3704538,177.35899 5.8666801,180.66699 6.628027,185.742 13.589325,190.259 12.800026,195.33298 12.600523,196.616 10.711118,197.556 9.6666718,198.66701 8.6222324,199.778 6.8791876,200.73 6.5333748,202 6.0249953,203.867 6.9888978,205.778 7.2166824,207.66701 8.0624208,214.67999 11.137519,223.26401 3.9333153,228.66701 8.2578211,232.886 11.083219,238.463 9.2000198,243.33299 8.5576019,244.995 4.5484915,245.651 4.0666771,247.33299 3.7063594,248.592 6.0889058,249.556 7.1000099,250.66699 8.1111212,251.778 9.7482452,252.745 10.133324,254 11.889721,259.72601 10.674818,266.15201 6.3999939,271.33301 L8.5333633,272.66699 7.0666885,274 C7.3999991,276 7.7333498,278 8.0666924,280 8.4000015,282 9.5558949,284.013 9.0666962,286 8.7983828,287.09 6.0084553,287.565 5.9333229,288.66699 5.8209848,290.315 7.68891,291.77802 8.5666847,293.33301 9.4444447,294.88901 11.31232,296.35199 11.200027,298 11.124919,299.10199 8.3342419,299.577 8.0666924,300.66699 7.0959282,304.62201 11.467319,309.10999 8.5333633,312.66699 L9.7333527,314 C9.6081047,319.63599 9.6157646,325.302 8.4222221,330.88901 7.5437994,335.00101 5.4916139,339.646 8.5333633,343.33301 L6.3999939,344.66699 8.4000015,346 C7.5630293,351.754 7.049418,357.63699 8.6000252,363.33301 8.9007435,364.43799 8.1333408,365.556 7.9000092,366.66699 7.6666698,367.77802 7.2000184,368.879 7.2000122,370 7.2000184,371.121 7.6666698,372.22198 7.9000092,373.33301 8.1333408,374.444 8.8768129,375.55899 8.6000252,376.66699 8.3530111,377.655 6.1820459,378.34299 6.3999939,379.33301 6.7349272,380.85599 9.5850048,381.82599 10.066719,383.33301 12.302822,390.32901 7.4479795,397.53299 7.3333549,404.66699 7.3118386,406.005 7.68891,407.33301 7.8666878,408.66699 8.0444708,410 7.8444004,411.375 8.4000015,412.66699 8.8207226,413.64499 10.991718,414.33301 10.733318,415.33301 10.223216,417.30801 6.8754678,418.69501 6.3333511,420.66699 6.0622954,421.65201 8.4335423,422.33499 8.5333633,423.33301 8.977273,427.77301 10.119116,432.241 9.4000053,436.66699 9.1694136,438.086 7.7777801,439.33301 6.9666862,440.66699 6.1555557,442 4.5214515,443.23999 4.5333672,444.66699 4.5452614,446.09799 6.200016,447.33301 7.0333672,448.66699 7.8666801,450 9.3608847,451.23901 9.5333672,452.66699 9.9817057,456.37701 6.9951582,460.026 4.2666626,463.33301 9.3970747,465.776 6.8681378,471.30399 6.6666603,475.33301 6.606997,476.52701 8.0666904,477.556 8.7666702,478.66699 9.4666843,479.77802 10.939618,480.80701 10.866718,482 10.804018,483.026 8.4257021,483.64001 8.4000015,484.66699 8.3746719,485.67999 10.708017,486.32001 10.733318,487.33301 10.758318,488.33301 9.266674,489.11099 8.5333633,490 L11.000023,491.33301 C10.433317,492.444 9.8666859,493.556 9.3000412,494.66699 8.7333527,495.77802 7.303319,496.849 7.6000214,498 8.0980606,499.93301 11.181719,501.388 11.533413,503.33301 11.75172,504.54099 9.9555855,505.556 9.1666794,506.66699 8.3777914,507.77802 7.0045881,508.79099 6.8000221,510 6.0519156,514.42102 9.6595354,519.38202 6.3999939,523.33301 L8.5333633,524.66699 7.3999977,526 C7.5158195,529.85999 7.8269401,533.729 8.6444473,537.55603 8.9925928,539.185 9.340744,540.815 9.6889114,542.44397 10.037016,544.07397 11.56112,545.77301 10.733318,547.33301 9.6102753,549.45099 6.4222164,550.88898 4.2666626,552.66699 12.114821,556.495 8.8292027,565.80298 4.2666626,571.33301 8.6030426,573.677 9.0274534,578.815 6.3999939,582 L9.2000198,583.33301 C8.1555614,588.22198 4.6948519,593.14301 6.0666847,598 6.4321365,599.29401 8.311141,600.22198 9.4333267,601.33301 10.555617,602.44397 12.568522,603.36102 12.800026,604.66699 13.046924,606.05902 11.61642,607.37903 10.733318,608.66699 9.4766645,610.49902 7.8444605,612.22198 6.3999939,614 L8.5333633,615.33301 C7.5111194,617.11102 5.5434041,618.77802 5.466671,620.66699 5.2843337,625.15503 11.825821,630.00897 8.5333633,634 L10.666714,635.33301 C8.4244213,638.65503 9.7736053,643.26599 13.533421,646 L6.3999939,651.33301 C10.396917,653.617 11.931421,657.32599 12.800026,660.66699 L5.3333473,661.33301 4.2666626,666 5.3333473,666.66699 8.5333633,668.66699 6.4666748,670 12.800026,675.33301 11.133327,676 C9.2337742,680.11603 6.0221953,684.508 7.6666641,688.66699 8.1333408,689.84698 9.2444639,690.88898 10.033417,692 10.822218,693.11102 12.278722,694.12 12.400017,695.33301 12.851624,699.849 8.8298731,704.15302 8.3333588,708.66699 7.8436103,713.11902 9.6723146,718.034 6.3999939,722 7.288919,722.44397 8.5449924,722.70099 9.0666962,723.33301 10.079617,724.56097 11.107819,725.96503 10.800018,727.33301 10.460717,728.84198 7.80055,729.84698 7.266674,731.33301 4.6066017,738.73798 9.2844639,746.461 10.666714,754 11.233319,757.091 8.3854218,760.56799 10.666714,763.33301 L6.7333412,764.66699 C7.1111183,766 7.2373486,767.37097 7.8666878,768.66699 8.7561131,770.49799 11.56592,772.09497 11.266727,774 11.004719,775.66699 8.0222607,776.66699 6.3999939,778 L10.666714,780.66699 C7.7119899,783.83197 9.0874138,788.07599 6.3999939,791.33301 L9.600029,793.29199 C12.444522,792.43103 15.013229,791.02301 18.133335,790.70801 19.693441,790.55103 20.977844,791.56897 22.400055,792 L24.533463,788 30.93338,789.33301 32.000065,791.33301 C33.777878,791.55603 35.858185,791.341 37.333393,792 38.627392,792.578 38.755692,793.77802 39.466801,794.66699 41.600098,794.20801 43.610905,793.29199 45.866814,793.29199 47.466915,793.29199 48.71122,794.20801 50.133419,794.66699 L52.266827,793.33301 58.66684,794.66699 60.800152,792 65.066872,795 C70.029671,792.72198 76.018089,790.13501 82.133499,790.58301 84.543213,790.76001 86.157616,792.44299 88.533516,792.75 90.978828,793.06598 93.511337,792.5 96.000252,792.375 98.489151,792.25 100.97826,792.125 103.46725,792 L104.53327,794 304.00079,794.66699 305.06778,792.66699 C307.55679,792.33301 310.0098,791.461 312.53381,791.66699 314.28381,791.81 315.37881,792.97198 316.80081,793.625 321.76685,791.44299 328.90585,793.01001 334.93387,792.56299 336.71188,792.43103 338.61987,792.60498 340.26788,792.16699 343.3429,791.349 345.47791,788.64801 348.8009,788.875 352.48593,789.12701 354.48993,791.84698 357.33392,793.33301 L359.46793,791.625 C361.60095,792.63898 363.25793,794.28699 365.86795,794.66699 367.73196,794.93799 369.42297,793.77802 371.20096,793.33301 372.97897,792.88898 374.75699,792.44397 376.534,792 L377.60098,794 381.401,794.66699 C382.26801,793.55603 382.73898,792.28497 384.00101,791.33301 384.76801,790.755 386.134,790.63898 387.20102,790.29199 L393.60101,793.33301 C393.95703,792.66699 393.93204,791.86499 394.66803,791.33301 395.46103,790.76001 396.80103,790.61102 397.86804,790.25 L402.13406,792.375 C408.78207,789.172 418.14108,789.07098 425.6011,786.66699 L427.26813,787.33301 429.86813,794.66699 C434.90714,793.427 439.98816,792.21802 444.80115,790.66699 L446.93417,792 449.06818,790.66699 451.20117,793.33301 453.3342,792 C460.09119,793.38599 467.58023,792.36298 474.66824,792 L476.80124,796.375 C481.58627,795.60901 487.57129,794.52301 489.60129,791.70801 496.41031,793.21503 503.71432,789.45398 510.93433,789.33301 517.00836,789.23199 522.60138,791.59302 528.0014,793.33301 L530.1344,792 532.26837,794.66699 534.40137,793.29199 538.6684,794.66699 C543.55841,790.27301 556.65747,790.88098 564.26849,793.33301 L566.40149,792 568.53448,793.66699 C574.5545,790.90399 583.84955,795.47601 590.93457,794 589.66852,792.44397 588.38354,790.89398 587.13452,789.33301 586.25055,788.22803 585.40155,787.11102 584.53455,786 L587.73456,784 588.80151,785.04199 C590.57953,785.58301 592.66956,785.83698 594.13458,786.66699 594.96857,787.13898 594.84656,788 595.20154,788.66699 596.26855,788.76398 597.45953,788.63098 598.40155,788.95801 600.74457,789.77197 602.66858,790.98602 604.80157,792 605.15759,791.33301 605.29858,790.60602 605.86859,790 606.74158,789.07001 608.00159,788.30603 609.0686,787.45801 614.20361,790.29999 621.47363,792.66699 628.26862,792.08301 629.87463,791.94501 630.91467,790.64099 632.53467,790.625 634.13464,790.60901 635.37964,791.54199 636.8017,792 L637.86865,790.91699 C643.77869,789.20801 650.59668,788.21698 657.06873,788.66699 L654.93469,792.66699 C662.26172,795.90997 672.79779,789.19098 681.60181,790.29199 684.2298,790.62 685.86877,792.31897 688.00183,793.33301 L690.13483,791.04199 C710.69885,792.02301 732.02393,793.86603 752.00195,790.66699 753.21295,792.93701 759.26501,794.42999 762.66901,793.33301 L764.802,796 C770.03802,794.02301 774.828,791.61603 779.73505,789.33301 780.44604,789.90302 780.83502,790.71899 781.86908,791.04199 783.83105,791.65503 786.13507,791.68103 788.26904,792 L790.4021,789.33301 792.5351,792 C792.89105,791.33301 792.97809,790.58502 793.60205,790 794.43909,789.216 795.73511,788.66699 796.80206,788 L797.86908,790.54199 C805.27112,791.15399 813.63214,792.80499 820.26917,790.66699 L822.40216,791.875 C828.43018,790.28601 835.39618,790.62598 841.60223,789.33301 841.95819,789.98602 841.74921,790.91101 842.66919,791.29199 845.55225,792.48499 848.81323,793.66699 852.26923,793.75 854.08624,793.79401 854.88525,792.10199 856.53522,791.625 858.52722,791.04901 860.80225,790.98602 862.93524,790.66699 L865.06927,793.29199 869.33527,792 871.4693,794.33301 C874.31329,793.84698 877.06628,792.703 880.00232,792.875 881.84131,792.98297 882.84729,794.34698 884.26935,795.08301 893.25934,790.60602 906.95636,792.67102 918.4024,792 920.56342,791.87299 922.66943,792.5 924.80243,792.75 926.93542,793 929.08643,793.80103 931.20245,793.5 934.35345,793.052 936.55945,790.99701 939.73547,790.625 941.30847,790.44098 942.40948,791.909 944.0025,792 946.85651,792.16302 949.69147,791.56897 952.53552,791.354 955.38049,791.13898 958.2915,791.14697 961.06952,790.70801 962.62054,790.46301 963.7395,789.26099 965.33551,789.33301 968.35852,789.47101 970.89154,790.94098 973.86957,791.29199 976.66455,791.62097 979.55859,791.32501 982.40259,791.34198 995.2666,791.41699 1010.8926,788.62201 1020.8027,793.75 1022.5827,792.86102 1024.3627,791.97198 1026.1327,791.08301 1027.9127,790.19397 1029.3228,788.89203 1031.4727,788.41699 1032.8328,788.11401 1034.4628,788.52899 1035.7327,788.95801 1038.0928,789.75098 1039.5728,791.50403 1042.1327,792 1044.5227,792.46198 1047.1128,791.76398 1049.6028,791.646 1059.9227,791.15601 1070.1829,793.33301 1080.5328,793.33301 1082.1328,793.33301 1083.2029,791.95801 1084.8029,791.95801 1086.4028,791.95801 1087.6428,792.875 1089.0729,793.33301 L1091.2029,790.91699 1101.8729,792 1104.0029,789.33301 1108.2729,792.375 1114.673,789.33301 1116.803,790.83301 C1121.3729,789.09399 1127.7229,791.64398 1132.333,793.33301 1133.203,792.22198 1133.553,790.883 1134.933,790 1135.823,789.43201 1137.3829,789.55603 1138.603,789.33301 1139.533,790.44397 1141.2729,791.414 1141.403,792.66699 1141.5031,793.66498 1139.933,794.44397 1139.203,795.33301 L1140.2729,796 1145.603,795.33301 1143.473,794 1144.533,792 C1153.543,793.58301 1163.0431,790.28003 1172.2731,789.33301 L1174.4031,791.29199 1178.6731,788.75 C1183.0431,791.26898 1189.8231,791.63397 1195.7332,792 1196.0931,791.33301 1196.2131,790.59698 1196.8031,790 1197.6631,789.13898 1198.9331,788.47198 1200.0032,787.70801 L1210.6732,792.41699 1214.9332,790.29199 C1217.0732,791.31897 1218.7231,792.95697 1221.3333,793.375 1222.8832,793.62299 1224.0132,792.12402 1225.6033,792 1239.7133,790.89697 1254.0533,792.26202 1268.2733,792 1268.6433,790.88898 1269.6034,789.79602 1269.4033,788.66699 1269.2234,787.67297 1266.9033,786.98199 1267.2034,786 1267.6934,784.427 1271.1833,783.59302 1271.4734,782 1271.8033,780.15302 1268.9333,778.52502 1268.8733,776.66699 1268.8234,775.25299 1271.2533,774.07898 1271.1333,772.66699 1270.9734,770.78003 1268.6133,769.19202 1268.0734,767.33301 1267.1533,764.18799 1271.5233,761.19702 1271.6034,758 1271.6333,756.987 1269.2933,756.34601 1269.2733,755.33301 1269.2433,754.33301 1270.7333,753.55603 1271.4734,752.66699 L1269.3334,751.33301 1271.4734,750 C1269.9333,748.66699 1267.5233,747.591 1266.8733,746 1265.0834,741.64099 1271.9933,737.03699 1270.2733,732.66699 1269.8033,731.49298 1266.5333,731.203 1266.3334,730 1266.1033,728.54602 1268.2433,727.33301 1269.2034,726 1270.1633,724.66699 1272.1533,723.46002 1272.0734,722 1272.0033,720.88098 1269.8933,720.22198 1268.8033,719.33301 1270.3633,718 1272.8134,716.93201 1273.4734,715.33301 1273.8534,714.388 1272.2134,713.54498 1271.5333,712.66699 1270.8434,711.76703 1269.1233,710.99103 1269.3334,710 1269.5734,708.88202 1272.0734,708.39697 1272.6733,707.33301 1273.1433,706.492 1272.9634,705.46899 1272.3334,704.66699 1271.5033,703.59497 1269.3134,703.08801 1268.5333,702 1267.6633,700.76801 1267.9133,699.33301 1267.6033,698 1266.6533,693.92297 1273.7134,690.02802 1272.3334,686 1271.7334,684.25 1268.5133,683.33301 1266.6033,682 1268.2433,680.66699 1270.8633,679.63098 1271.5333,678 1271.9333,677.03101 1269.4333,676.33099 1269.3334,675.33301 1269.0233,672.21198 1267.8633,669.12701 1267.8033,666 1267.7833,664.63202 1268.5634,663.29401 1269.2733,662 1269.7933,661.05402 1271.6233,660.32898 1271.4734,659.33301 1271.3033,658.25098 1268.7234,657.742 1268.4734,656.66699 1268.2034,655.51898 1269.5333,654.44397 1270.0734,653.33301 1270.6034,652.22198 1271.8334,651.15503 1271.6733,650 1271.2833,647.29999 1271.3234,644.409 1269.3334,642 L1270.6034,640.66699 C1269.7234,635.35901 1271.3833,630.00299 1271.4734,624.66699 1271.5033,622.83801 1270.3334,621.06897 1269.4033,619.33301 1268.8933,618.38501 1267.2733,617.66602 1267.2034,616.66699 1267.1133,615.28497 1268.3833,614 1268.9734,612.66699 1269.5634,611.33301 1270.6934,610.04999 1270.7333,608.66699 1270.7933,606.86298 1269.1034,605.13397 1269.2733,603.33301 1269.3633,602.33502 1270.7333,601.55603 1271.4734,600.66699 L1269.3334,599.33301 1271.4734,598 1270.2034,596.66699 C1270.7933,593.12598 1271.9434,589.164 1269.3334,586 L1271.5333,584.66699 1269.3334,582 1270.7333,580.66699 C1269.7633,577.13501 1271.8033,572.974 1268.6033,570 1272.2334,567.26099 1268.5433,562.88702 1268.7333,559.33301 1268.8334,557.49799 1270.6733,555.828 1270.9333,554 1271.4534,550.45801 1268.8633,546.495 1271.4734,543.33301 L1269.3334,542 C1270.2034,537.55603 1272.5333,533.12903 1271.9333,528.66699 1271.7833,527.49103 1270.6733,526.44397 1270.0333,525.33301 1269.4033,524.22198 1268.3234,523.17401 1268.1333,522 1266.8134,513.495 1270.6733,504.94101 1270.2733,496.39999 1270.0033,490.79999 1269.7333,485.20001 1269.4734,479.60001 1269.2933,475.957 1272.0133,471.90601 1269.3334,468.66699 L1271.4734,467.33301 C1266.4333,464.185 1270.8534,458.43399 1270.2034,454 1269.9933,452.54501 1267.4833,451.45801 1267.3334,450 1267.1533,448.181 1268.3733,446.401 1269.2733,444.66699 1269.7633,443.715 1270.7333,442.88901 1271.4734,442 L1270.4033,441.33301 C1270.1733,429.99899 1269.6133,418.60901 1271.4734,407.33301 1271.6934,405.98801 1270.7333,404.66699 1270.3733,403.33301 1270.0033,402 1268.8933,400.66599 1269.2733,399.33301 1269.6833,397.85199 1272.6034,396.83701 1272.6034,395.33301 1272.6034,393.82999 1269.5233,392.828 1269.2733,391.33301 1269.1034,390.33899 1271.4734,389.66699 1271.4734,388.66699 1271.4734,386.83099 1269.2733,385.16901 1269.2733,383.33301 1269.2733,382.33301 1270.7333,381.556 1271.4734,380.66699 L1267.2034,378 1268.2733,377.33301 C1268.4333,371.31299 1270.2933,365.35001 1270.6733,359.33301 1270.7234,358.436 1270.4933,357.52499 1270.0734,356.66699 1269.3933,355.28601 1267.6433,354.103 1267.4033,352.66699 1267.1733,351.31201 1268.1133,349.97198 1268.7333,348.66699 1269.3833,347.297 1271.3434,346.09299 1271.2034,344.66699 1271.0133,342.75699 1268.2034,341.23099 1267.8033,339.33301 1267.5133,337.979 1268.7633,336.66699 1269.2333,335.33301 1269.7134,334 1270.5533,332.698 1270.6733,331.33301 1271.3334,323.77399 1267.8234,316.23099 1268.3334,308.66699 1268.4233,307.31 1269.1633,306 1269.5734,304.66699 1269.9833,303.33301 1270.7333,302.02399 1270.8033,300.66699 1271.0233,296.66699 1268.5333,292.224 1271.4734,288.66699 L1270.4033,288 C1270.3134,286 1271.0133,283.92499 1270.1333,282 1269.6233,280.87201 1267.6934,280.22198 1266.4733,279.33301 1270.3234,276.423 1272.7733,272.44299 1272.6034,268.66699 1272.5433,267.289 1271.4933,266 1270.9333,264.66699 1270.3833,263.33301 1269.1633,262.043 1269.2733,260.66699 1269.3534,259.668 1271.5734,258.99799 1271.4734,258 1271.3633,256.94601 1268.9133,256.383 1268.7333,255.33298 1268.5133,253.965 1269.8234,252.66701 1270.3733,251.33299 1270.9133,250 1272.3134,248.696 1272.0033,247.33299 1271.7533,246.241 1269.1333,245.757 1268.8733,244.66702 1268.4833,243.105 1269.7633,241.556 1270.2034,240 1270.6433,238.444 1271.7234,236.909 1271.5333,235.33298 1270.9634,230.45799 1267.8833,225.01199 1271.4734,220.66699 L1269.3334,219.33301 1271.4734,218 C1268.9634,213.81799 1272.6133,208.63699 1269.3334,204.66702 L1271.4734,203.33299 1269.3334,202 1270.5333,200.66699 C1270.3234,195.743 1268.6333,190.88901 1267.6733,186 1267.0433,182.81 1271.4534,179.88 1271.5333,176.66701 1271.7034,170.00101 1266.5834,162.592 1271.4734,156.66701 L1268.6033,155.33298 C1269.7833,154.444 1272.1133,153.821 1272.1333,152.66701 1272.1733,150.75301 1268.9534,149.242 1268.7333,147.33299 1268.5333,145.502 1270.7633,143.83299 1270.9333,142 1271.5133,135.787 1266.9033,128.86501 1271.4734,123.33299 L1269.8033,122.66701 1267.2034,119.333 1270.4033,118.66699 C1272.3833,112.307 1268.0533,105.79 1267.2034,99.333405 L1270.4033,98.666702 C1270.7333,98 1270.9634,97.308701 1271.4033,96.666702 1272.0333,95.747398 1272.8733,94.888901 1273.6034,94 L1269.3334,91.333389 1273.6034,88.666702 1271.4033,87.333389 1273.6034,84.66671 1269.3334,83.333397 C1270.1833,81.555603 1272.4033,79.823997 1271.8733,77.999992 1271.2533,75.9011 1268.0934,74.444504 1266.2034,72.666702 1272.5834,68.481903 1265.2733,60.996498 1267.1333,55.333405 1267.7833,53.374001 1270.9933,51.971298 1271.5333,50 1271.8033,49.014301 1270.0734,48.222301 1269.3334,47.333393 L1271.4734,46.000004 1269.3334,44.666691 1270.7333,43.333397 C1269.2633,38.287399 1268.0934,33.118698 1268.6333,27.999992 1268.8434,26 1268.6233,23.9638 1269.2733,22.000008 1269.5934,21.020201 1271.6733,20.325701 1271.4734,19.333401 1271.1633,17.822001 1268.2833,16.841101 1267.9333,15.333405 1267.5333,13.5499 1268.8733,11.7778 1269.3334,10 1267.5634,9.8194599 1265.7034,9.8275805 1264.0033,9.4583702 1262.4534,9.1215801 1261.3833,7.9327402 1259.7333,7.9166985 1258.1133,7.9008799 1257.0933,9.3592501 1255.4733,9.375 1253.8733,9.3906298 1252.6233,8.4583702 1251.2032,7.9999924 L1249.0732,9.3333626 1246.9332,7.7499962 C1238.5732,8.9342699 1228.9633,9.0338697 1221.2732,6.6666985 L1217.0732,10.666695 1212.8032,7.9999924 C1211.0232,8.84729 1209.5232,9.9913301 1207.4731,10.541706 1204.4431,11.3525 1201.0731,11.5139 1197.8732,12.000008 L1196.8031,10 C1195.7332,9.7916899 1194.6332,9.6501503 1193.6031,9.375 1192.1332,8.9810801 1190.9131,7.8356299 1189.3331,7.9999924 1187.1831,8.2243004 1185.8732,9.6731005 1184.0031,10.375004 1182.3031,11.011 1180.4431,11.4584 1178.6731,12.000008 L1176.5331,8.9166832 C1172.8931,10.7311 1165.963,12.0587 1162.6731,10 L1169.0731,6.0000038 1168.0031,5.3333664 C1164.8131,8.0707397 1157.8831,7.7821698 1153.073,9.3333626 L1150.933,8.3749962 1144.533,9.3333626 z M1.0666847,0 L1280,0.66669464 1278.9301,800 1.0666847,800 0,799.33301 z", Fill = Brushes.Orange, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 1026, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 10 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M36.625,591.375 C36.625,591.375 147.5,459.83333 231.5,52.75 231.5,52.75 283.5,380.5 426.625,591.25 z", Fill = Brushes.OrangeRed, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 110, Height = 154, Top = 422, Left = 166, Thickness = 0, ZIndex = 1 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M36.625,591.375 C36.625,591.375 147.5,459.83333 231.5,52.75 231.5,52.75 283.5,380.5 426.625,591.25 z", Fill = Brushes.OrangeRed, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 218, Height = 304, Top = 272, Left = 3, Thickness = 0, ZIndex = 5 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent3.Color), ColorSpecialName = colorAccent3.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M36.625,591.375 C36.625,591.375 147.5,459.83333 231.5,52.75 231.5,52.75 283.5,380.5 426.625,591.25 z", Fill = Brushes.OrangeRed, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 56, Height = 80, Top = 496, Left = 196, Thickness = 0, ZIndex = 2, Opacity = 0.7 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M452,777.66667 C452,777.66667 386.00022,672.9998 352.83333,491.83302 352.83333,491.83302 434.00006,713.49976 474.49992,706.49977 474.49992,706.49977 458.97392,769.08824 452.01561,777.67156 z", Fill = Brushes.OrangeRed, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 33, Height = 74, Top = 495, Left = 226, Thickness = 0, ZIndex = 3, Opacity = 0.4 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M551.16698,958.41667 C551.16698,958.41667 587.83391,976.83273 673.83497,744.16542 673.83497,744.16542 637.68888,936.43644 574.68787,1029.4372 574.68787,1029.4372 561.08365,1007.3333 551.16698,958.41667 z", Fill = Brushes.OrangeRed, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 33, Height = 74, Top = 495, Left = 189, Thickness = 0, ZIndex = 3, Opacity = 0.4 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M36.625,591.375 C36.625,591.375 147.5,459.83333 231.5,52.75 231.5,52.75 283.5,380.5 426.625,591.25 z", Fill = Brushes.OrangeRed, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 109, Height = 160, Top = 416, Left = 55, Thickness = 0, ZIndex = 6, Opacity = 0.7 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M452,777.66667 C452,777.66667 386.00022,672.9998 352.83333,491.83302 352.83333,491.83302 434.00006,713.49976 474.49992,706.49977 474.49992,706.49977 458.97392,769.08824 452.01561,777.67156 z", Fill = Brushes.OrangeRed, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 68, Height = 160, Top = 400, Left = 111, Thickness = 0, ZIndex = 7, Opacity = 0.4 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M551.16698,958.41667 C551.16698,958.41667 587.83391,976.83273 673.83497,744.16542 673.83497,744.16542 637.68888,936.43644 574.68787,1029.4372 574.68787,1029.4372 561.08365,1007.3333 551.16698,958.41667 z", Fill = Brushes.OrangeRed, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 68, Height = 160, Top = 400, Left = 41, Thickness = 0, ZIndex = 8, Opacity = 0.4 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent1.Color), ColorSpecialName = colorAccent1.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }
            }
        }

        public static ObservableCollection<StandardElement> Theme10(EColorManagment color)
        {
            SolidColor colorAccent4 = color.Accent4;
            SolidColor colorAccent5 = color.Accent5;

            ObservableCollection<StandardElement> listshape = new ObservableCollection<StandardElement>();
            StaticShape staticShape = new StaticShape() { PathData = "M0.5,0.5 L81.5,0.5 L81.5,799.5 L0.5,799.5 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#c91810") }, StrokeThickness = 0 };
            StandardElement standardElement = new StandardElement() { Width = 50, Height = 577, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent4.Color), ColorSpecialName = colorAccent4.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M0.5,0.5 L37.5,0.5 L37.5,799.5 L0.5,799.5 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#fd5236") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 30, Height = 577, Top = 0, Left = 50, Thickness = 0, ZIndex = 3 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent5.Color), ColorSpecialName = colorAccent5.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            staticShape = new StaticShape() { PathData = "M1263.5,950.5 L1388.5,950.5 1388.5,825 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#fd5236") }, StrokeThickness = 0 };
            standardElement = new StandardElement() { Width = 100, Height = 100, Top = 477, Left = 926, Thickness = 0, ZIndex = 5 };
            standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent5.Color), ColorSpecialName = colorAccent5.Name };
            standardElement.ShapePresent = staticShape;
            listshape.Add(standardElement);

            return listshape;

        }

        public static void GenerateLayoutTheme10(SlideMaster slide, int maxZindex, EColorManagment color)
        {
            SolidColor colorAccent4 = color.Accent4;
            SolidColor colorAccent5 = color.Accent5;

            foreach (LayoutMaster layoutMaster in slide.LayoutMasters)
            {
                if (layoutMaster.ThemeLayout == null) layoutMaster.ThemeLayout = new ThemeLayout();

                foreach (ObjectElement objectlayout in layoutMaster.MainLayout.Elements)
                {
                    objectlayout.ZIndex += (maxZindex + 3);
                }

                if (layoutMaster.LayoutName == "Title Only Layout" || layoutMaster.LayoutName == "Title Content Layout")
                {
                    layoutMaster.IsHideBackground = true;
                    StaticShape staticShape = new StaticShape() { PathData = "M0,0 L1280,0 L1280,39 L0,39 z", StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 1026, Height = 39, Top = 0, Left = 0, Thickness = 0, ZIndex = 1 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent5.Color), ColorSpecialName = colorAccent5.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M0,0 L1280,0 L1280,39 L0,39 z", StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 1026, Height = 39, Top = 39, Left = 0, Thickness = 0, ZIndex = 1 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent4.Color), ColorSpecialName = colorAccent4.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M264.16667,862.16667 L264.5,736.5 390,862 z", StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 100, Height = 100, Top = 477, Left = 0, Thickness = 0, ZIndex = 1 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent5.Color), ColorSpecialName = colorAccent5.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }
                else if (layoutMaster.LayoutName == "Title Slide Layout")
                {
                    layoutMaster.IsHideBackground = true;
                    StaticShape staticShape = new StaticShape() { PathData = "M0.5,0.5 L81.5,0.5 L81.5,799.5 L0.5,799.5 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#c91810") }, StrokeThickness = 0 };
                    StandardElement standardElement = new StandardElement() { Width = 50, Height = 577, Top = 0, Left = 976, Thickness = 0 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent4.Color), ColorSpecialName = colorAccent4.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M0.5,0.5 L37.5,0.5 L37.5,799.5 L0.5,799.5 z", Fill = new SolidColorBrush() { Color = (Color)ColorConverter.ConvertFromString("#fd5236") }, StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 30, Height = 577, Top = 0, Left = 946, Thickness = 0 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent5.Color), ColorSpecialName = colorAccent5.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);

                    staticShape = new StaticShape() { PathData = "M264.16667,862.16667 L264.5,736.5 390,862 z", StrokeThickness = 0 };
                    standardElement = new StandardElement() { Width = 100, Height = 100, Top = 477, Left = 0, Thickness = 0, ZIndex = 1 };
                    standardElement.Fill = new ColorSolidBrush() { Color = (Color)ColorConverter.ConvertFromString(colorAccent5.Color), ColorSpecialName = colorAccent5.Name };
                    standardElement.ShapePresent = staticShape;
                    layoutMaster.MainLayout.Elements.Add(standardElement);
                }
            }
        }



    }
}
