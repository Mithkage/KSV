   M S B A M L   `   `   `  ����  RTS    RT_Isolate.InputWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35H 9http://schemas.microsoft.com/winfx/2006/xaml/presentation      5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	     � Title$  Isolate Elements Input��5   	   $	��180��6(   $	��500��65    � WindowStartupLocation$ CenterScreen=�5   	    � 
ResizeMode$ NoResize=�5	   	    � ShowInTaskbar$ False��5
   	    � Topmost$
 True��5   	    � WindowStyle$ 
ToolWindow=�5   	    � AllowsTransparency$ False��5   	   ���	#FFE0F2F1 � 	Resources 5         ModernButtonStyle �� 
TargetType	 �� Property
 �� Value3  -System.Windows.Media.Effects.DropShadowEffect   Color   	Direction   ShadowDepth   
BlurRadius   Opacity �� 
TargetType �� CornerRadius �� Triggers P� Property P� Value��%{  &	        '�� W    �� 5   
   " ��6*   .�5      �� 	 � ��
 �	#FF26A69A6+   �� 5      	 � ��

 �����6+   �� 5      	 � ��$

 10,5q�6(   �� 5      	 � ��
 �	#FF00796B6,   �� 5      	 � ��$
 1q�60   �� 5      	 � ��$
 14(�6)   �� 5      	 � ��$
 SemiBold"�6+   �� 5      	 �@J�Effect
 5        5      $ 	#FF616161��6'   $	 270Q�69   $ 2Q�6I   $ 4Q�6Y   $	 0.4Q�6h   5      �� 5      	 � ��
 5      �� 5       " ��6&   .R�5!      �� $ 8��6!   #��z��62   #��z��5"   !   #��z��5#   !   ._�5$      �� $��Center=�6/   $��Center=�6L   5%      5&      5(       5)      P� 5*       �@J�IsMouseOver$
 True��61   . �5+      �� 	 � ��
 �	#FF4DB6AC63   �� 5,      	 � ��
 �	#FF00897B64   5-      P� 5.       � ��$
 True��6/   . �5/      �� 	 � ��
 �	#FF00897B63   �� 50      	 � ��
 �	#FF00695C64   �� 51      	 �@J�Effect
 52        53      $ 	#FF616161��6/   $	 270Q�6A   $ 1Q�6Q   $ 2Q�6a   $	 0.6Q�6p   55      56      58      �� 5;   
   " ��6   .�5<      �� 	 � ��
 �	#FF00796B6,   �� 5=      	 � ��$
 1q�60   �� 5?      	 � ��$
 5q�6(   �� 5@      	 � ��

 �����6+   �� 5A      	 � ��
 �	#FF3333336+   �� 5B      	 � ��$
 Center=�69   5C      5D      .��5F      � $��15q�6    � RowDefinitions 5G   
   �� 5H      $
��Auto �6   �� 5I      $
��Auto �6   �� 5J      $��* �6   �� 5K      $
��Auto �6    � ColumnDefinitions 5M   
   �� 5N      $��* �6   �� 5O      $
��Auto �6   �� 5P      $
��Auto �6   .+�5S   
   �� $��0��6   $��3��6!   $��SemiBold"�65   $��15(�6K   $��0,0,0,5q�6Y   ���	#FF333333$��Center=�6�   #!Enter values to isolate elements:5U      �� 5V   
   $��1��6   $��3��6!   $
��Wrap=�65   $��0,0,0,10q�6I   ���	#FF333333$��12(�6r   ��Enter one or more values, separated by commas, to isolate elements by 'RTS_ID' or 'RTS_Cable_XX' parameters. Example: 'C101, C102, Cable-A'5Y      �� 5[   
   -   5         ��Name$ InputTextBox��5[      $��2��6(   $��3��65   $��25��6I   �� 5]   
   $��3��6    �� Orientation$ 
Horizontal=�6"   $��Right=�6;   $��0,15,0,0q�6W   $	��309��6i   .+�5_      �� -   5        $��Clear Overrides��5_      $	��120��6/   $��30��6;   $��0,0,10,0q�6G   #��[  5`      �� 5c      -   5        $��OK��5c      $��80��6"   $��30��6-   $��0,0,10,0q�69    �� 	IsDefault$
 True��6\   #��[  5d      �� 5g      -   5        $��Cancel��5g      $��80��6&   $��30��61    �� IsCancel$
 True��6R   #��[  5h      5i      5j      5k      