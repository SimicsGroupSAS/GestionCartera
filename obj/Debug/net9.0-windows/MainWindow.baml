   M S B A M L   `   `   `  ����  Gestión de Cartera    MiProyectoWPF.MainWindow   5   	   W SPresentationCore, Version=9.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=9.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=9.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=9.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=9.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   _ [MaterialDesignThemes.Wpf, Version=4.9.0.0, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	      MaterialDesignBody#���   5   	   $��Regular"�5   	   $��13(�5   	   &   System.Windows.Media.TextOptions    TextFormattingMode$  Ideal=�5   	      TextRenderingMode$
 Auto=�5	   	   
�������  MaterialDesignFont#���  5   	    � Title$ Gestión de Cartera��5   	   $	��600��6$   $	��800��61   .��5      � .+�5   
   ]� .+�5      �� $	��Top=�6   
������� �� Padding$ 16q�6=   ._�5      ]� .+�5      �� $B��<d:/Alerta Cartera/Background/MiProyectoWPF/Assets/Simics.png��6   $��40��6b   $��Center=�6n   o� 5      
������ 
���   �$��Right=�6G   .@�5      m� $��Herramientas��6#   .@�5      m� -   5        $��Configurar Credenciales��5   '   6H   5      m� 5      $��Ayuda��6#   .@�5      m� -   5        $��	Acerca de��5   '   6:   5      5      5      5      �� 5      -   5         ��Name$ mainTabControl��5        MaterialDesignTabControl#��[ 61   .@�5!      �� $��Inicio��6    �� 
IsSelected$
 True��6*   .��5"      � .+�5#      �� $��20q�6%   $��Center=�61   .+�5$      �� $$��Sistema de Gestión de Cartera��6(   $��24(�5%   '   $
��Bold"�65   $��Center=�5&   '   $��0,0,0,30q�6D    # MaterialDesignPrimaryColorBrush#���  5'   '   #  MaterialDesignThemes.Wpf.Card  5)      $��	0,10,0,20q�62   $��16q�6E   
�������.��5*   "   �� .+�5+   &   �� $'��!Paso 1: Generar Estados de Cuenta��60   $��18(�5,   /   $
��Bold"�6=   $��0,0,0,10q�6O     MaterialDesignBlackBrush#���  5-   /   �� 5.   &   $
��Wrap=�60   $��0,0,0,10q�6D     MaterialDesignDarkGrayBrush#���  5/   0   ><Genera documentos PDF a partir del archivo Excel de cartera.51   '   �� 52   &   -   5        $��Generar Estados de Cuenta��52   -     MaterialDesignRaisedButton#��[ 53   -   $��40��54   -   $��14(�69   #���  55   -   
�������$
��0,10q�57   P   58   #   59        5;      $��	0,10,0,20q�62   $��16q�6E   
�������.��5<   "   �� .+�5=   &   �� $1��+Paso 2: Enviar Estados de Cuenta por Correo��60   $��18(�5>   /   $
��Bold"�6=   $��0,0,0,10q�6O   #���  5?   /   �� 5@   &   $
��Wrap=�60   $��0,0,0,10q�6D   #���  5A   0   :8Envía los documentos generados por correo electrónico.5C   '   �� 5D   &   -   5        $��Enviar Estados de Cuenta��5D   -   #��[ 5E   -   #���  5F   -   
�������$��40��5H   -   $��14(�69   $
��0,10q�5I   O   5J   #   5K      5L      5M      5N      �� 5Q      -   5         ��Name$ tabGeneracionDocumentos��5Q   <   $��Generación de Documentos��6   .��5R      �  � RowDefinitions 5S      �� 5T      $
��Auto �6,   �� 5U      $��* �6,   �� 5V      $
��Auto �6,   .+�5Y      �� $��0��6%   $��16q�62   .+�5Z      �� $&�� Generación de Estados de Cuenta��6(   $��16(�5[   (   $��0,0,0,10q�5\   (   #���  5]   (   �� 5^      -   5        $��Seleccionar Archivo Excel��5^   %   #��[ 5_   %   $	��200��5`   %   $
��Left=�5a   %   $	��0,5q�5b   %   #���  5c   %   
�������5e   %   �� 5f      -   5         ��Name$ txtExcelSeleccionado��5f   (   $"��Ningún archivo seleccionado��5g   (   $	��0,5q�5h   (   #���  5i   (   5j      �� 5l      $��1��6'   $��16q�64   $
��Auto=�5m   '   $
��Auto=�5n   '   $
��True��5o   '   .��5p      �� -	   5        	 ��Name$	 txtGeneracionLog��5p   &   
 �� 
IsReadOnly$

 True��5q   &    �� TextWrapping$
 Wrap=�5r   &    �� AcceptsReturn$
 True��5s   &   " �� VerticalScrollBarVisibility$ Visible=�5t   &   $ �� HorizontalScrollBarVisibility$
 Auto=�5u   &   ���#F8F8F8$��Consolas)�5w   &   #���  5x   &   ��5y   &   ��  �� RelativeSource ��  �� AncestorType ScrollViewer �� Path$ ViewportWidth���5z   &   ��  ��  ScrollViewer$ ViewportWidth�5{      �� 5}      $��2��6%   $��16q�62   .+�5~      �� -
   5        $��Generar Estados de Cuenta��5~   %   #��[ 5   %   $	��200��5�   %   $��35��61   $��14(�6=   #���  5�   %   
�������5�   %   �� 5�      -   5        $ txtEstadoGeneracion��5�   (   $	��0,5q�6E   $
��Wrap=�5�   (   $��Italic$�6<   #���  5�   (   5�      5�      5�      �� 5�      -   5        $ tabEnvioCorreos��5�   4   $��Envío de Correos��6   .��5�      �  5�      �� 5�      $
��Auto �6,   �� 5�      $��* �6,   �� 5�      $
��Auto �6,   �� 5�      $	��200 �6,   .+�5�      ]� $��0��6$   $��
16,16,16,0q�61   .+�5�      �� $-��'Seleccione las empresas para gestionar:��6(   $��16(�5�   (   $��Center=�5�   (   $
��Left=�5�   (   #���  5�   (   �� 5�      -   5        $��Seleccionar Archivo Excel��5�   %   #��[ 5�   %   $	��200��5�   %   $��16,0,0,0q�61   #���  5�   %   
�������$��Right=�5�   %   $��Right=�5�   %   5�      �� 5�      $��1��6'   $��16q�64   .��5�      �� -   5         ��Name$ CheckboxContainer��5�   )   5�      �� 5�       �� Orientation$ 
Horizontal=�6%   $��Center=�6>   $��2��6[   $��16q�6h   .+�5�      �� -   5        $��Seleccionar Todos��5�   %   #��[ 5�   %   $	��150��5�   %   $��5q�61   #���  5�   %   
�������5�   %   �� 5�      -   5        $��Deseleccionar Todos��5�   %      MaterialDesignOutlinedButton#��[ 5�   %   $	��150��5�   %   $��5q�61    % !MaterialDesignSecondaryColorBrush#���  5�   %   #���  5�   %   5�   %   �� 5�      -   5        $��Ver Detalles de Archivos��5�   %   #��[ 5�   %   $	��180��5�   %   $��5q�61   #���  5�   %   
�������5�   %   �� 5�      -   5        $��Enviar Correos��5�   %   #��[ 5�   %   $	��150��5�   %   $��5q�61   #���  5�   %   
�������5�   %   5�      � 5�      $��3��6    5�      �� 5�   "   $
��Auto �60   �� 5�   "   $��* �60   .+�5�      �� $��Log de operaciones:��6(   $��12(�5�   (   $
��Bold"�5�   (   $��	16,8,16,0q�5�   (   $��0��5�   (   #���  5�   (   �� 5�      $��1��6+   $��
16,0,16,16q�5�   +   $
��Auto=�5�   +   $
��Auto=�5�   +   $
��True��5�   +   .��5�   "   �� -   5        $	 txtEnvioLog��5�   *   $

 True��5�   )   $
 Wrap=�5�   )   $
 True��5�   )   $ Visible=�5�   )   $
 Auto=�5�   )   ���#F8F8F8$��Consolas)�5�   )   #���  5�   )   ��5�   )   ��  ��  ScrollViewer$ ViewportWidth���5�   )   ��  ��  ScrollViewer$ ViewportWidth�5�      5�      5�      5�      5�      5�      5�      5�      