   M S B A M L   `   `   `  ����   GoogleTrendsCombinator'    !GoogleTrendsCombinator.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     -     �Name$  MyWindow��
������� � Icon$ /Resources/gt_icon.png�� � Title$ Google Trends Combinator��$	��449��$	��780��.��� $	��405��$	��750��.+��� -   $��MyTabControl��$
��Left=�$	��371��$	��Top=�$	��730��$��8,12,0,0q�.@��� -   $��MakeTab��$
��Make��.��� $	��730��$	��333��.+��� -   $��btnSelectDaily��$��Select Daily Files��$��23��$
��Left=�$��	12,46,0,0q�$	��Top=�$	��114���� -   $��label1��$��Step 1��$��32��$
��Left=�$��	12,16,0,0q�$	��Top=�$��16(��� -   $��	listDaily��$	��190��$
��Left=�$��	12,75,0,0q�$	��Top=�$	��333������  �� Path$ 	DailyList� �� ElementName$ MyWindow���� -   $��label2��$��Step 2��$��16(�$��32��$
��Left=�$��
376,16,0,0q�$	��Top=��� -   $��btnSelectWeekly��$��Select Weekly Files��$��23��$��376,46,231,0q�$	��Top=��� -	   $��
listWeekly��$	��190��$
��Left=�$��
376,75,0,0q�$	��Top=�$	��333������ $ 
WeeklyList�$ MyWindow���� -
   $��btnClear��$��
Start Over��$��90��$��23��$��525,287,0,0q�$	��Top=�$
��Left=��� -   $��	btnSubmit��$
��Save��$��90��$��23��$��
0,287,20,0q�$	��Top=�$��Right=��� -   $��	checkBox1��$#��Ask me where to save the file��$��16��$
��Left=�$��
12,271,0,0q�$	��Top=� d� 	IsChecked �� $ AskWhereToSave�$ MyWindow���� -   $��	checkBox2��$��Generate charts��$��16��$
��Left=�$��190,271,0,0q�$	��Top=� �� $ IsMakeChartsChecked�$ MyWindow���� -   $��	comboBox1��$��23��$
��Left=�$��297,269,0,0q�$	��Top=�$	��120���� -   $��
lbl_Status��$��32��$
��Left=�$��
12,309,0,0q�$	��Top=�$��12(�
���\\��$	��337������ $ 
StatusText�$ MyWindow�� �� NotifyOnTargetUpdated .  � Triggers :�  :� RoutedEvent$ Binding.TargetUpdated��.[��� .���� .��U� 	 �� TargetProperty$	 Opacity�.\�9  3System.Windows.Media.Animation.EasingDoubleKeyFrame  
 O� KeyTime$
 0�� O� Value$ 0Q�  $
 0:0:0.25��$ 1Q�  $
 0:0:5��$ 1Q�  $
 0:0:6��$ 0Q��� $��TODO: Search��