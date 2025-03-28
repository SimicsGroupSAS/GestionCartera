using System;
using System.Collections.Generic;
using System.Windows;
using System.IO;
using System.Linq;

namespace MiProyectoWPF
{
    public class ArchivoEmpresaInfo
    {
        public required string Empresa { get; set; }
        public required string Nit { get; set; }
        public required string NombreArchivo { get; set; }
        public required string TipoCartera { get; set; }
        public required string TipoCoincidencia { get; set; }
        public required string RutaCompleta { get; set; }
    }
    
    public partial class DetallesPdfWindow : Window
    {
        public DetallesPdfWindow(Dictionary<string, string> empresasArchivos, Dictionary<string, string> empresaNits)
        {
            InitializeComponent();
            CargarDatos(empresasArchivos, empresaNits);
        }
        
        private void CargarDatos(Dictionary<string, string> empresasArchivos, Dictionary<string, string> empresaNits)
        {
            var datos = new List<ArchivoEmpresaInfo>();
            
            foreach (var kvp in empresasArchivos)
            {
                string empresa = kvp.Key;
                string rutaArchivo = kvp.Value;
                
                // Solo procesar archivos que existen
                if (!File.Exists(rutaArchivo))
                    continue;
                
                string nombreArchivo = Path.GetFileName(rutaArchivo);
                string tipoCartera = ObtenerTipoCartera(nombreArchivo);
                string nit = empresaNits.ContainsKey(empresa) ? empresaNits[empresa] : "";
                
                // Determinar tipo de coincidencia aproximadamente
                string tipoCoincidencia = "Directa";
                string nombreBase = Path.GetFileNameWithoutExtension(rutaArchivo)
                    .Replace("_CarteraVencida", "").Replace("_CarteraPorVencer", "");
                
                if (!nombreBase.StartsWith(empresa, StringComparison.OrdinalIgnoreCase) && 
                    !empresa.StartsWith(nombreBase, StringComparison.OrdinalIgnoreCase))
                {
                    tipoCoincidencia = "Por NIT";
                }
                else if (nombreBase.Length != empresa.Length)
                {
                    tipoCoincidencia = "Parcial";
                }
                
                datos.Add(new ArchivoEmpresaInfo
                {
                    Empresa = empresa,
                    Nit = nit,
                    NombreArchivo = nombreArchivo,
                    TipoCartera = tipoCartera,
                    TipoCoincidencia = tipoCoincidencia,
                    RutaCompleta = rutaArchivo
                });
            }
            
            // Filtrar para mostrar solo los que tienen coincidencia
            var datosConCoincidencia = datos.Where(d => !string.IsNullOrEmpty(d.NombreArchivo)).ToList();
            
            // Actualizar el título de la ventana para indicar el número de coincidencias
            this.Title = $"Detalle de Archivos PDF - {datosConCoincidencia.Count} archivos encontrados";
            
            dgDetalles.ItemsSource = datosConCoincidencia;
        }
        
        private string ObtenerTipoCartera(string nombreArchivo)
        {
            if (nombreArchivo.Contains("_CarteraVencida"))
                return "Vencida";
            else if (nombreArchivo.Contains("_CarteraPorVencer"))
                return "Por Vencer";
            else
                return "Desconocido";
        }
        
        private void CerrarButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
