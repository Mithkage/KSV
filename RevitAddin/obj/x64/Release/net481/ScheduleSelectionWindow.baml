   M S B A M L   `   `   `  ����   RTS+    %RTS_Schedules.ScheduleSelectionWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35H 9http://schemas.microsoft.com/winfx/2006/xaml/presentation      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006       � Title$  Select RTS Schedules��$	��550��$	��400�� � WindowStartupLocation$ CenterScreen=� � 
ResizeMode$ NoResize=� � ShowInTaskbar$ False�� � Topmost$
 True�����	#FFE0F2F1 � 	Resources    ModernButtonStyle �� 
TargetType �� Property �� Value3  -System.Windows.Media.Effects.DropShadowEffect	   Color
   	Direction   ShadowDepth   
BlurRadius   Opacity �� 
TargetType �� CornerRadius �� Triggers P� Property P� Value��%�  &	        �� " ��.���  � �� �	#FF26A69A��  � ��
 �������  � ��$
 10,5q���  � �� �	#FF00796B��  � ��$ 1q���  � ��$ 14(���  � ��$ SemiBold"���  �@J�Effect   $	 	#FF616161��$	
 270Q�$ 2Q�$ 4Q�$	 0.4Q���  � �� �� " ��.R��� $ 8��#��z��#��z��#��z��._��� $��Center=�$��Center=� P�  �@J�IsMouseOver$
 True��. ���  � �� �	#FF4DB6AC��  � �� �	#FF00897BP�  � ��$
 True��. ���  � �� �	#FF00897B��  � �� �	#FF00695C��  �@J�Effect   $	 	#FF616161��$	
 270Q�$ 1Q�$ 2Q�$	 0.6Q�.��� $��15q� � RowDefinitions �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto �.+��� $"��Select Schedules to Process:��$��0��$��15(�$��SemiBold"�$��0,0,0,20q�$��Center=����	#FF333333��  �� Orientation$ 
Horizontal=�$��Center=�$��0,0,0,10q�$��1��.+��� -   $��
Create All��$	��100��$��30��$��5q�#��[  �� -   $��
Remove All��$	��100��$��30��$��5q�#��[  �� $
��Auto=�$��0,0,0,10q�$��2��.���� -    ��Name$ SchedulesPanel��$��5q��� $ 
Horizontal=�$��Right=�$��0q�$��3��.+��� -   $��Generate Selected��$	��120��$��30��$��5q�#��[  �� -   $��Cancel��$��80��$��30��$��5q�#��[  