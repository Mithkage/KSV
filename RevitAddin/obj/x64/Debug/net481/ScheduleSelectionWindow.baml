   M S B A M L   `   `   `  ����  RTS+    %RTS_Schedules.ScheduleSelectionWindow   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35H 9http://schemas.microsoft.com/winfx/2006/xaml/presentation      5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	     � Title$  Select RTS Schedules��5   	   $	��550��6&   $	��400��63    � WindowStartupLocation$ CenterScreen=�5	   	    � 
ResizeMode$ NoResize=�5
   	    � ShowInTaskbar$ False��5   	    � Topmost$
 True��5   	   ���	#FFE0F2F1 � 	Resources 5         ModernButtonStyle �� 
TargetType �� Property �� Value3  -System.Windows.Media.Effects.DropShadowEffect	   Color
   	Direction   ShadowDepth   
BlurRadius   Opacity �� 
TargetType �� CornerRadius �� Triggers P� Property P� Value��%a  &	        �� 5   
   " ��6*   .�5      ��  � �� �	#FF26A69A6+   �� 5       � ��
 �����6+   �� 5       � ��$
 10,5q�6(   �� 5       � �� �	#FF00796B6,   �� 5       � ��$ 1q�60   �� 5       � ��$ 14(�6)   �� 5       � ��$ SemiBold"�6+   �� 5       �@J�Effect 5        5      $	 	#FF616161��6'   $	
 270Q�69   $ 2Q�6I   $ 4Q�6Y   $	 0.4Q�6h   5      �� 5       � �� 5      �� 5      " ��6&   .R�5       �� $ 8��6!   #��z��62   #��z��5!   !   #��z��5"   !   ._�5#      �� $��Center=�6/   $��Center=�6L   5$      5%      5'       5(      P� 5)       �@J�IsMouseOver$
 True��61   . �5*      ��  � �� �	#FF4DB6AC63   �� 5+       � �� �	#FF00897B64   5,      P� 5-       � ��$
 True��6/   . �5.      ��  � �� �	#FF00897B63   �� 5/       � �� �	#FF00695C64   �� 50       �@J�Effect 51        52      $	 	#FF616161��6/   $	
 270Q�6A   $ 1Q�6Q   $ 2Q�6a   $	 0.6Q�6p   54      55      57      58      .��5:      � $��15q�6    � RowDefinitions 5;   
   �� 5<      $
��Auto �6   �� 5=      $
��Auto �6   �� 5>      $��* �6   �� 5?      $
��Auto �6   .+�5C   
   �� $"��Select Schedules to Process:��6   $��0��5D      $��15(�5E      $��SemiBold"�6"   $��0,0,0,20q�68   $��Center=�5F      ���	#FF33333361   �� 5I   
    �� Orientation$ 
Horizontal=�6   $��Center=�6.   $��0,0,0,10q�6K   $��1��6]   .+�5J      �� -   5        $��
Create All��5J      $	��100��6*   $��30��66   $��5q�6B   #��[  5K      �� 5L      -   5        $��
Remove All��5L      $	��100��6*   $��30��66   $��5q�6B   #��[  5M      5N      �� 5Q   
   $
��Auto=�6   $��0,0,0,10q�6:   $��2��6L   .��5R      �� -   5         ��Name$ SchedulesPanel��5R      $��5q�61   5T      5U      �� 5X   
   $ 
Horizontal=�6   $��Right=�6.   $��0q�6J   $��3��6U   .+�5Y      �� -   5        $��Generate Selected��5Y      $	��120��61   $��30��6=   $��5q�6I   #��[  5Z      �� 5[      -   5        $��Cancel��5[      $��80��6&   $��30��61   $��5q�6=   #��[  5\      5]      5^      5_      