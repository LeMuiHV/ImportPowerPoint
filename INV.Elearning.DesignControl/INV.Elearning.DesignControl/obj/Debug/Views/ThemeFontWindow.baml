   M S B A M L   `   `   `  ����Q  MINV.Elearning.Controls, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null_Dclr-namespace:INV.Elearning.Controls;assembly=INV.Elearning.ControlsINV.Elearning.Controls  5   	   M IINV.Elearning.Core, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullaCclr-namespace:INV.Elearning.Core.Helper;assembly=INV.Elearning.CoreINV.Elearning.Core.Helper 5   	    INV.Elearning.DesignControl7   1INV.Elearning.DesignControl.Views.ThemeFontWindow   5      W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   XcontrolDclr-namespace:INV.Elearning.Controls;assembly=INV.Elearning.Controls     5   	   RvmCclr-namespace:INV.Elearning.Core.Helper;assembly=INV.Elearning.Core     5   	     � UseLayoutRounding  .  � 
ResizeMode$ NoResize=�5   3    � WindowStartupLocation$ CenterOwner=�6I      	TFW_Title � Title# �   5	   	   $	��300��6-   $	��500��6:     TFW_InfoContent,   &INV.Elearning.Controls.ElearningWindow   InfoButtonContent# �  6F      IsInfoButtonEnabled . .��5
      � $
��5,10q�6    � ColumnDefinitions 5   
   �� 5      $	��200 �6   �� 5       � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      .+�5   
   ��   TFW_HeadingFont#���  6   $��Center=�6d    �� Target 6<   �� 	 �� ElementName$	 cbbMajor��6d   %   INV.Elearning.Controls.ComboBox  5   
   -   5        
  Name$
 cbbMajor��5      $��14(�6-   $��1��6;   $��-5,5,5,5q�6I    �� SelectedIndex$ 0��6�   $��28��6"   �� SelectedItem 6[   �� *	MajorFont+ �� Mode$ OneWayToSource=���6�   ��    System.Windows.Media.Fonts   SystemFontFamilies �� Source# Z *  $INV.Elearning.Core.Helper.Converters   FontFamiliesConverter �� 	Converter# Z ��5      5  -System.Windows.Controls.ItemContainerTemplate�  5      .R�5      �� $��Center=�6H   ��6    �� ��61   ��  � ToolTip 6c   �� 5      5      �� 5   
     TFW_BodyFont#���  6   $��2��6a   $��Center=�6n    69   �� $	 cbbMinor��6n     5   
   -   5        $
 cbbMinor��5      $��14(�6-   $��3��6;   $��-5,5,5,5q�6H   $ 0��6�   $��28��6"   �� SelectedValue 6Z   �� *	MinorFont+$ OneWayToSource=���6�   �� # Z # Z ��5         5!      .R�5"      �� $��Center=�6H   ��6    �� ��61   ��  6c   �� 5#      5%      �� 5&   
   $��1��6     
TFW_Sample#���  6    $��Center=�6G   �� 5'   
   $��1��6   $��1��6   $��3��6.   
�������$��1q�6W    �� Padding$ 5q�6k   $	��5,0q�6w   ._�5(      �� .+�5)      �� $��Center=�6     TFW_Heading#���  67   $
��Bold"�6}   $��16(�6�   ��6\   �� *	MajorFont+6�   �� 5*      $��Center=�6     TFW_BodyText#���  67   $��WrapWithOverflow=�6~   ��6]   �� *	MinorFont+6~   5+      5,      � 5-   
   $��4��6   $
��0, 5q�6   $��2��6*    5.      �� 5/      $
��Auto �6#   �� 50      .+�52      ��   TFW_Name#���  6;   $��Center=�6`    6   �� $	 txtName��6`   �� 53      -   5         ��Name$ txtName��53      $��1��6'   $	��2,3q�67   $	��5,0q�6E   $��Center=�6R   
���������6�   �� *Name+ �� UpdateSourceTrigger$ PropertyChanged=�54      55         Fotters 57   
   ��   TFW_Save#���  6   $��80��66   $��32��6A   $��Right=�6M   $��	0,0,120,0q�6i   ��6|   ��  �� RelativeSource ��  �� AncestorType$ WindowL� �� Path$ SaveCommand��� 58   
    	 
TFW_Cancel#��� 	 6   $��80��68   $��32��6C   $��Right=�6O   $��0,0,20,0q�6k   ��6}   ��  �� $ WindowL�$ CancelCommand�5:      