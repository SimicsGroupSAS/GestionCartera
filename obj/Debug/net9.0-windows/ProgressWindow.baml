   M S B A M L   `   `   `  ����  Gestión de Cartera"    MiProyectoWPF.ProgressWindow   5   	   W SPresentationCore, Version=9.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=9.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=9.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=9.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=9.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	     � Title$  Procesando...��5   	   $	��150��5   	   $	��350��6    � WindowStartupLocation$ CenterOwner=�5   	    � 
ResizeMode$ NoResize=�5   	    � WindowStyle$
 None=�5   	    � AllowsTransparency$
 True��5	   	   
������  � Topmost$
 True��5   	    � 	Resources 5      3  -System.Windows.Media.Effects.DropShadowEffect   DropShadowEffect   
BlurRadius   ShadowDepth	   	Direction
   Color   Opacity��%s   &	          5   
   $ 15Q�5      $ 0Q�5      $	 0Q�5      $

 Gray��5      $	 0.6Q�5      5      .��5      ��  �� CornerRadius$ 10��6   $��1q�5      ���#CCC
������� J� Effect# [  5      ._�5   
   � $��20q�6    � RowDefinitions 5      �� 5      $
��Auto �6    �� 5      $��* �6    .+�5"      �� -   5         ��Name$ 
txtMessage��5"      $��0��5#      $��Procesando...��5$      $��14(�5%      $��Center=�5&      $��0,0,0,15q�5'      � 5)      $��1��6    � IsIndeterminate$
 True��5*      $��15��5+      $��Center=�5,      5-      5.      5/      