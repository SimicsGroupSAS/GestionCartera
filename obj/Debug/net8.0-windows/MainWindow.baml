   M S B A M L   `   `   `  ����  MiProyectoWPF    MiProyectoWPF.MainWindow   5   	   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   _ [MaterialDesignThemes.Wpf, Version=4.9.0.0, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	      MaterialDesignPaper#���   5   	     MaterialDesignFont#���  5   	     MaterialDesignBody#���  5   	   $��Regular"�5   	   $��13(�5	   	     � Title$"  SIMICS - Gestión de Cartera��5
   	   $	��600��6-   $	��800��6:   .��5      �  � RowDefinitions 5   
   �� 5      $
��Auto �6   �� 5      $��* �6   .+�5   
   �� $��0��6   
�������$��60��61   ���#EEEEEE$��0,0,0,1q�6S   ._�5      � .+�5      ��  �� Orientation$ 
Horizontal=�6   $
��Left=�66   $
��10,0q�6Q   .+�5      �� -   5         ��Name$ 	logoImage��5      $��40��6/   $
��Auto��6;   $��0,0,10,0q�6H   $��Center=�6Z   5      �� 5      $"��SIMICS - Gestión de Cartera��6   $��20(�5      $��Medium"�6)     PrimaryGreenBrush#��[ 5      $��Center=�5      $��Center=�5      �� 5"      $ 
Horizontal=�6   $��Right=�66   $
��10,0q�6R   .+�5#      �� -   5          MaterialDesignFlatButton#��[ 5#       � ToolTip$ Configurar Credenciales��5$      #��[ 5%      .��5'      �� $ 
Horizontal=�6%   .+�5(      �� $��Credenciales��6(   $��8,0,0,0q�6<   5)      5*      �� 5+      -   5        #��[ 5+      $ 	Acerca de��5,      #��[ 5-      .��5/      �� $ 
Horizontal=�6%   .+�50      �� $��	Acerca de��6(   $��8,0,0,0q�69   51      52      53      54      55      � 58   
   $��1��6   .+�59      �� -   5         ��Name$ mainTabControl��59      
������ .@�5;      �� $��Inicio��6    �� 
IsSelected$
 True��6*   .��5<      �� $
��Auto=�6#   .��5=      �� $��32q�6%   $��Center=�61   .+�5>      �� $$��Sistema de Gestión de Cartera��6(     HeaderStyle#��[ 5?   '   $��Center=�5@   '   $��0,0,0,32q�6D   �� 5B        MaterialDesignDivider#���  6%   $��1q�6[   
������� �� CornerRadius$ 4��5C   7   $��0,8,0,16q�6H    �� Padding$ 16q�6Z   ._�5D   "   �� .+�5E   &   �� $'��!Paso 1: Generar Estados de Cuenta��60     SubHeaderStyle#��[ 5F   /   �� 5G   &   $
��Wrap=�60   $��0,0,0,16q�6D   nlEsta opción procesa el archivo Excel de cartera y genera los documentos de estado de cuenta en formato PDF.5I   '   �� 5J   &   -   5        $��Generar Estados de Cuenta��5J   -   #��[ 5K   -   
�������$��Right=�5N   -   5O   #   5P      �� 5R      #���  6%   $��1q�6[   
�������$ 4��5S   7   $��0,8,0,16q�6H   $ 16q�6Z   ._�5T   "   �� .+�5U   &   �� $1��+Paso 2: Enviar Estados de Cuenta por Correo��60   #��[ 5V   /   �� 5W   &   $
��Wrap=�60   $��0,0,0,16q�6D   geEsta opción permite seleccionar las empresas y enviar los estados de cuenta por correo electrónico.5Y   '   �� 5Z   &   -   5        $��Enviar Estados de Cuenta��5Z   -     SecondaryGreenBrush#��[ 5[   -   
�������$��Right=�5^   -   5_   #   5`      5a      5b      5c      �� 5f      -   5        	 ��Name$	 tabEnvioCorreos��5f   4   $��Envío de Correos��6   .��5g      � $��16q�6    5h      �� 5i      $
��Auto �6,   �� 5j      $��* �6,   �� 5k      $
��Auto �6,   �� 5l      $	��200 �6,   .+�5o      ]� $��0��6$   .+�5p      �� $-��'Seleccione las empresas para gestionar:��6(   $
��Bold"�5q   '   $��16(�69   $��Center=�5r   '   �� 5s      -   5        $��Seleccionar Archivo Excel��5s   %   #��[ 5t   %   
�������$��Right=�5w   %   5x      �� 5z      $��1��6'   $��	0,16,0,16q�64   $
��Auto=�5{   '   $
��Auto=�5|   '   .��5}      �� -	   5        
 ��Name$
 CheckboxContainer��5}   )   5~      �� 5�      $ 
Horizontal=�6%   $��2��6>   $��Center=�6K   $��0,0,0,16q�6h   .+�5�      �� -
   5        $��Seleccionar Todos��5�   %   $	��8,0q�5�   %   62   �� 5�      -   5        $��Deseleccionar Todos��5�   %   $	��8,0q�5�   %   62   �� 5�      -   5        $��Ver Detalles de Archivos��5�   %   $	��8,0q�5�   %   62   �� 5�      -   5        $��Enviar Correos��5�   %   #��[ 5�   %   
�������$	��8,0q�5�   %   62   5�      � 5�      $��3��6    5�      �� 5�   "   $
��Auto �60   �� 5�   "   $��* �60   .+�5�      �� $��Log de operaciones:��6(   $��SemiBold"�5�   '   $��0,0,0,8q�5�   '   $��0��68   �� 5�      $��1��6%   $��1q�5�   %   ���#CCCCCC���#F8F8F8._�5�   "   �� $
��Auto=�6/   $
��Auto=�5�   .   $��8q�5�   .   .��5�   &   �� -   5         ��Name$ txtEnvioLog��5�   .    �� 
IsReadOnly$
 True��5�   -    �� TextWrapping$
 Wrap=�5�   -    �� AcceptsReturn$
 True��5�   -   $��0q�5�   -   
������ $��Consolas)�5�   -   $��12(�6C   5�   #   5�      5�      5�      5�      �� 5�      -   5        $	 tabGeneracionDocumentos��5�   <   $��Generación de Documentos��6   .��5�      � $��16q�6    5�      �� 5�      $
��Auto �6,   �� 5�      $��* �6,   �� 5�      $
��Auto �6,   .+�5�      �� $��0��6%   $��0,0,0,16q�62   .+�5�      �� $&�� Generación de Estados de Cuenta��6(   $
��Bold"�5�   '   $��16(�69   $��0,0,0,16q�5�   '   �� 5�      -   5        $��Seleccionar Archivo Excel��5�   %   #��[ 5�   %   
�������$
��Left=�5�   %   �� 5�      -   5         ��Name$ txtExcelSeleccionado��5�   (   $"��Ningún archivo seleccionado��5�   '   $��0,8,0,0q�5�   '   5�      �� 5�      $��1��6!   $��1q�5�   !   ���#CCCCCC���#F8F8F8._�5�      �� $
��Auto=�6+   $
��Auto=�5�   *   $��8q�5�   *   .��5�   "   �� -   5        $ txtGeneracionLog��5�   *   $
 True��5�   )   $
 Wrap=�5�   )   $
 True��5�   )   $��0q�5�   )   
������ $��Consolas)�5�   )   $��12(�6?   5�      5�      �� 5�      $��2��6%   $��0,16,0,0q�62   .+�5�      �� -   5        $��Generar Estados de Cuenta��5�   %   #��[ 5�   %   
�������$��40��5�   %   $
��Left=�5�   %   �� 5�      -   5        $ txtEstadoGeneracion��5�   (   $��Italic$�5�   '   $��0,8,0,0q�5�   '   5�      5�      5�      5�      5�      5�      5�      