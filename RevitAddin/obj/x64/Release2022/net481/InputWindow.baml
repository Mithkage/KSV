   M S B A M L   `   `   `  ����   RTS    RT_Isolate.InputWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35^ ZWindowsFormsIntegration, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35H 9http://schemas.microsoft.com/winfx/2006/xaml/presentation      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006       � Title$  Isolate Elements Input��$	��180��$	��500�� � WindowStartupLocation$ CenterScreen=� � 
ResizeMode$ NoResize=� � ShowInTaskbar$ False�� � Topmost$
 True�� � WindowStyle$ 
ToolWindow=� � AllowsTransparency$ False�����	#FFE0F2F1 � 	Resources    ModernButtonStyle �� 
TargetType	 �� Property
 �� Value3  -System.Windows.Media.Effects.DropShadowEffect   Color   	Direction   ShadowDepth   
BlurRadius   Opacity �� 
TargetType �� CornerRadius �� Triggers P� Property P� Value��%;  &	        '�� �    �� " ��.��� 	 � ��
 �	#FF26A69A�� 	 � ��

 ������� 	 � ��$

 10,5q��� 	 � ��
 �	#FF00796B�� 	 � ��$
 1q��� 	 � ��$
 14(��� 	 � ��$
 SemiBold"��� 	 �@J�Effect
   $ 	#FF616161��$	 270Q�$ 2Q�$ 4Q�$	 0.4Q��� 	 � ��
 �� " ��.R��� $ 8��#��z��#��z��#��z��._��� $��Center=�$��Center=� P�  �@J�IsMouseOver$
 True��. ��� 	 � ��
 �	#FF4DB6AC�� 	 � ��
 �	#FF00897BP�  � ��$
 True��. ��� 	 � ��
 �	#FF00897B�� 	 � ��
 �	#FF00695C�� 	 �@J�Effect
   $ 	#FF616161��$	 270Q�$ 1Q�$ 2Q�$	 0.6Q��� " ��.��� 	 � ��
 �	#FF00796B�� 	 � ��$
 1q��� 	 � ��$
 5q��� 	 � ��

 ������� 	 � ��
 �	#FF333333�� 	 � ��$
 Center=�.��� $��15q� � RowDefinitions �� $
��Auto ��� $
��Auto ��� $��* ��� $
��Auto � � ColumnDefinitions �� $��* ��� $
��Auto ��� $
��Auto �.+��� $��0��$��3��$��SemiBold"�$��15(�$��0,0,0,5q����	#FF333333$��Center=�#!Enter values to isolate elements:�� $��1��$��3��$
��Wrap=�$��0,0,0,10q����	#FF333333$��12(���Enter one or more values, separated by commas, to isolate elements by 'RTS_ID' or 'RTS_Cable_XX' parameters. Example: 'C101, C102, Cable-A'�� -    ��Name$ InputTextBox��$��2��$��3��$��25���� $��3�� �� Orientation$ 
Horizontal=�$��Right=�$��0,15,0,0q�$	��309��.+��� -   $��Clear Overrides��$	��120��$��30��$��0,0,10,0q�#��[  �� -   $��OK��$��80��$��30��$��0,0,10,0q� �� 	IsDefault$
 True��#��[  �� -   $��Cancel��$��80��$��30�� �� IsCancel$
 True��#��[  