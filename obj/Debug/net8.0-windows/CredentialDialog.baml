   M S B A M L   `   `   `  ����  MiProyectoWPF$    MiProyectoWPF.CredentialDialog   5   	   W SPresentationCore, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=8.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=8.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	     � Title$.  (Configuración de Credenciales de Correo��5   	   $	��320��5   	   $	��500��6    � WindowStartupLocation$ CenterOwner=�5   	    � 
ResizeMode$ NoResize=�5   	   .��5      � $��20q�6    � RowDefinitions 5	   
   �� 5
      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $
��Auto �6   �� 5      $��* �6   �� 5      $
��Auto �6    � ColumnDefinitions 5   
   �� 5      $	��150 �6   �� 5      $��* �6   .+�5   
   �� $��0��6   $��2��6!   $4��.Configurar credenciales de correo electrónico��5      $��16(�5      $
��Bold"�6"   $��0,0,0,15q�5      �� 5   
   $��1��6   $��2��6!   $��Información de servidor:��5       $��SemiBold"�5!      $��0,0,0,5q�5"      �� 5$   
   $��2��6   $��2��6!   $+��%smtppro.zoho.com:587 (SSL habilitado)��5%      
�������$��0,0,0,15q�5'      �� 5*   
   $��3��6   $��Correo electrónico:��6!   $��Center=�6<   �� 5+   
   -   5         ��Name$ txtEmail��5+   /   $��3��6   $��1��6   $	��0,5q�6A   �� 5-   
   $��4��6   $��Contraseña:��6!   $��Center=�64   G� 5.   
   -   5         G�Name$ txtPassword��5.   3   $��4��6   $��1��6#   $	��0,5q�6H   �� 51   
   $��5��6   $��Copia oculta (BCC):��6!   $��Center=�6<   �� 52   
   -   5        $ txtBccEmail��52   /   $��5��6   $��1��6    �� 
IsReadOnly$
 True��6D   ���#F0F0F0$	��0,5q�53   '   �� 55   
   $��6��6   $��2��6"   $��0,15,0,0q�66   .+�56      �� -   5         ��Name$ chkGuardarCredenciales��56      $*��$Guardar credenciales de forma segura��67   	 d� 	IsChecked$
	 TrueY�57      58      �� 5:   
   $��7��6   $��2��6"   
 �� Orientation$
 
Horizontal=�5;      $��Right=�5<      $��0,10,0,0q�5=      .+�5>      �� -   5         ��Name$ 
btnValidar��5>      $��Validar��6)   $	��100��6;   $��30��6G   $��0,0,10,0q�5?      6&   �� 5@      -   5        $ 
btnGuardar��5@      $��Guardar��6)   $	��100��6;   $��30��6G    �� 	IsDefault$
 True��5A      $��0,0,10,0q�6%   67   �� 5B      -   5        $ btnCancelar��5B      $��Cancelar��6*   $	��100��6=   $��30��6I    �� IsCancel$
 True��5C      6$   5D      5E      5F      