   M S B A M L   `   `   `  ����  Gestión de Cartera%    MiProyectoWPF.DetallesPdfWindow   5   	   W SPresentationCore, Version=9.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=9.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35e aSystem.Windows.Controls.Ribbon, Version=9.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089R NWindowsBase, Version=9.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35F 9http://schemas.microsoft.com/winfx/2006/xaml/presentation     5   	   R NSystem.Xaml, Version=9.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	     � Title$  Detalle de Archivos PDF��5   	   $	��500��5   	   $	��700��6    � WindowStartupLocation$ CenterOwner=�5   	   .��5      � $��15q�6    � RowDefinitions 5   
   �� 5	      $
��Auto �6   �� 5
      $��* �6   �� 5      $
��Auto �6   .+�5   
   �� $��0��6   $*��$Relación de Archivos PDF y Empresas��5      $��18(�5      $
��Bold"�6"   $��0,0,0,15q�5      &   System.Windows.Controls.DataGrid  5   
   -   5        $��
dgDetalles��5      $��1��5         AutoGenerateColumns$ False��5         
IsReadOnly$
 True��5         AlternatingRowBackground �#F5F5F5   GridLinesVisibility$	 All=�5         Columns 5      0  *System.Windows.Controls.DataGridTextColumn  5      ,  &System.Windows.Controls.DataGridColumn   Header$ Empresa��6%   	   Width5  /System.Windows.Controls.DataGridLengthConverter$	 * 6R   1  +System.Windows.Controls.DataGridBoundColumn
   Binding
 66   �� *	Empresa+6R     5      $	 NIT��6%   $		 120 6J   
 62   �� *Nit+6J     5      $ Archivo��6%   $		 200 6X   
 66   �� *NombreArchivo+6X     5      $ Tipo Cartera��6%   $		 100 6[   
 6;   �� *TipoCartera+6[     5      $ Coincidencia��6%   $		 100 6`   
 6;   �� *TipoCoincidencia+6`   5       �� 5"   
   -   5        $��2��5"      $��Cerrar��5#      $��Right=�5$      $	��100��5%      $��30��5&      $��0,15,0,0q�5'      5(      5)      5*      