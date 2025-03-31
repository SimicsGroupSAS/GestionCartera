using System;
using System.Collections.Generic;
using System.IO; // Usar System.IO explícitamente
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Globalization;
using System.Text;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.Kernel.Pdf.Canvas;
using iText.IO.Image;
using iText.Kernel.Events;
using iText.Kernel.Geom; // Importar Rectangle explícitamente

namespace MiProyectoWPF
{
    // Clase para manejar el fondo de la plantilla simplificada para PNG
    public class BackgroundImageHandler : IEventHandler
    {
        private ImageData backgroundImage;
        private readonly Action<string> logCallback;

        public BackgroundImageHandler(string imagePath, Action<string> logCallback)
        {
            this.logCallback = logCallback;
            
            if (System.IO.File.Exists(imagePath))
            {
                try
                {
                    backgroundImage = ImageDataFactory.Create(imagePath);
                    logCallback($"Plantilla de fondo cargada correctamente: {imagePath}");
                }
                catch (Exception ex)
                {
                    logCallback($"Error al cargar la imagen de fondo: {ex.Message}");
                    throw;
                }
            }
            else
            {
                throw new FileNotFoundException($"No se encontró el archivo de plantilla: {imagePath}");
            }
        }

        public void HandleEvent(Event @event)
        {
            PdfDocumentEvent docEvent = (PdfDocumentEvent)@event;
            PdfDocument pdf = docEvent.GetDocument();
            PdfPage page = docEvent.GetPage();
            Rectangle pageSize = page.GetPageSize();
            
            PdfCanvas canvas = new PdfCanvas(page.NewContentStreamBefore(), page.GetResources(), pdf);
            
            try
            {
                // Dibujar la imagen de fondo para que se adapte a la página
                canvas.AddImageFittedIntoRectangle(backgroundImage, pageSize, false);
            }
            catch (Exception ex)
            {
                logCallback($"Error al dibujar fondo: {ex.Message}");
            }
            finally
            {
                canvas.Release();
            }
        }
    }

    public class SimplePdfGenerator
    {
        private readonly string rutaBase;
        private readonly string pdfOutputFolder;
        private readonly string tempFolderPath; // Añadida referencia a la carpeta temp
        private readonly Action<string> logCallback;
        private string templatePath;
        private bool templateFound = false;
        
        // Lista para rastrear las empresas a las que se les generaron PDFs
        private List<GeneratedPdfInfo> generatedPdfs = new List<GeneratedPdfInfo>();
        
        // Propiedad para acceder a la lista de PDFs generados
        public IReadOnlyList<GeneratedPdfInfo> GeneratedPdfFiles => generatedPdfs.AsReadOnly();

        public SimplePdfGenerator(string rutaBase, string pdfOutputFolder, Action<string> logCallback, string? tempFolderPath = null)
        {
            this.rutaBase = rutaBase ?? throw new ArgumentNullException(nameof(rutaBase));
            this.pdfOutputFolder = pdfOutputFolder ?? throw new ArgumentNullException(nameof(pdfOutputFolder));
            this.logCallback = logCallback ?? throw new ArgumentNullException(nameof(logCallback));
            this.tempFolderPath = tempFolderPath ?? string.Empty; // Asignar cadena vacía si es nulo

            // Modificar la ruta de la plantilla para usar un PNG directamente - ruta exacta que proporciona el usuario
            this.templatePath = @"D:\Alerta Cartera\Background\MiProyectoWPF\Archivos\PlantillaSIMICS.png";
            
            // Verificar que la carpeta de salida existe
            if (!System.IO.Directory.Exists(pdfOutputFolder))
            {
                System.IO.Directory.CreateDirectory(pdfOutputFolder);
                logCallback($"Carpeta creada: {pdfOutputFolder}");
            }
            
            // Verificar la existencia de la plantilla PNG
            if (System.IO.File.Exists(templatePath))
            {
                logCallback($"Plantilla PNG encontrada: {templatePath}");
                templateFound = true;
            }
            else
            {
                logCallback($"ADVERTENCIA: Plantilla PNG no encontrada en {templatePath}. Buscando alternativas...");
                
                // Probar con otras rutas alternativas
                string[] rutasAlternativas = {
                    System.IO.Path.Combine(rutaBase, "PlantillaSIMICS.png"),
                    System.IO.Path.Combine(rutaBase, "Archivos", "PlantillaSIMICS.png"),
                    System.IO.Path.Combine(rutaBase, "Background", "MiProyectoWPF", "Archivos", "PlantillaSIMICS.png")
                };
                
                foreach (var rutaAlt in rutasAlternativas)
                {
                    if (System.IO.File.Exists(rutaAlt))
                    {
                        this.templatePath = rutaAlt;
                        logCallback($"Se encontró plantilla PNG en ruta alternativa: {this.templatePath}");
                        templateFound = true;
                        break;
                    }
                }
                
                if (!templateFound)
                {
                    logCallback("No se encontró ninguna plantilla PNG. Se solicitará al usuario seleccionar una imagen.");
                }
            }
        }
        
        public bool IsTemplateFound => templateFound;
        
        public void SetTemplatePath(string path)
        {
            if (System.IO.File.Exists(path))
            {
                templatePath = path;
                templateFound = true;
                logCallback($"Plantilla PNG actualizada: {path}");
            }
            else
            {
                logCallback($"Error: La ruta especificada no existe: {path}");
            }
        }

        public async Task GenerateDocuments(string excelFilePath)
        {
            logCallback("Iniciando generación de documentos...");
            
            // Configurar cultura española
            try {
                CultureInfo.CurrentCulture = new CultureInfo("es-CO");
                logCallback("Cultura configurada: Español (Colombia)");
            }
            catch (Exception ex) {
                logCallback($"Error al configurar cultura: {ex.Message}");
            }

            // Limpiar carpetas de salida
            logCallback("Limpiando carpeta de salida...");
            CleanOutputFolder();

            // Procesar el archivo Excel
            logCallback($"Procesando archivo Excel: {excelFilePath}");
            await Task.Run(() => ProcessExcelFile(excelFilePath));

            logCallback("Proceso de generación de documentos completado.");
        }

        private void CleanOutputFolder()
        {
            try
            {
                // Eliminar archivos en la carpeta de PDF
                if (Directory.Exists(pdfOutputFolder))
                {
                    foreach (var file in System.IO.Directory.GetFiles(pdfOutputFolder))
                    {
                        System.IO.File.Delete(file);
                    }
                    logCallback($"Carpeta limpiada: {pdfOutputFolder}");
                }

                // Limpiar también la carpeta temp si existe
                if (!string.IsNullOrEmpty(tempFolderPath) && System.IO.Directory.Exists(tempFolderPath))
                {
                    foreach (var file in System.IO.Directory.GetFiles(tempFolderPath))
                    {
                        try
                        {
                            System.IO.File.Delete(file);
                        }
                        catch (Exception tempEx)
                        {
                            // Registrar error pero continuar con el proceso
                            logCallback($"Error al eliminar archivo temporal {System.IO.Path.GetFileName(file)}: {tempEx.Message}");
                        }
                    }
                    logCallback($"Carpeta temporal limpiada: {tempFolderPath}");
                }
            }
            catch (Exception ex)
            {
                logCallback($"Error al limpiar carpetas: {ex.Message}");
            }
        }

        private void ProcessExcelFile(string excelFilePath)
        {
            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    var worksheet = workbook.Worksheet(1);
                    logCallback($"Leyendo hoja: {worksheet.Name}");
                    
                    // Determinar hasta qué fila hay datos
                    var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0; // Verificar si LastRowUsed es nulo
                    if (lastRow == 0)
                    {
                        logCallback("No se encontraron filas en el archivo Excel.");
                        return;
                    }
                    
                    logCallback($"Última fila con datos: {lastRow}");
                    
                    // Inicia en la fila 5 (equivalente a Skip(4))
                    int startRow = 5;
                    logCallback($"Comenzando lectura desde la fila {startRow}");
                    
                    // Agrupar por cliente y NIT
                    var clientGroups = new Dictionary<string, ClienteInfo>();
                    int totalEmpresasEncontradas = 0;
                    
                    // Leer fila por fila desde startRow hasta la última
                    for (int rowNum = startRow; rowNum <= lastRow; rowNum++)
                    {
                        try 
                        {
                            var row = worksheet.Row(rowNum);
                            
                            // Si la fila está vacía, saltarla
                            if (row.IsEmpty())
                            {
                                logCallback($"Fila {rowNum} está vacía, saltando.");
                                continue;
                            }
                            
                            // Leer el nombre del cliente (columna B) y el NIT (columna C)
                            string nombre = string.Empty;
                            string nit = string.Empty;
                            try
                            {
                                nombre = TryGetStringValue(row.Cell("B"), $"Error al leer nombre en fila {rowNum}").Trim();
                                nit = TryGetStringValue(row.Cell("C"), $"Error al leer NIT en fila {rowNum}").Trim();
                            }
                            catch
                            {
                                logCallback($"Error al leer nombre o NIT en fila {rowNum}, saltando.");
                                continue;
                            }
                            
                            // Si no hay nombre o contiene "Total", saltamos esta fila
                            if (string.IsNullOrEmpty(nombre) || nombre.Contains("Total", StringComparison.OrdinalIgnoreCase))
                            {
                                logCallback($"Fila {rowNum} sin nombre válido o contiene 'Total', saltando.");
                                continue;
                            }
                            
                            totalEmpresasEncontradas++;
                            logCallback($"Procesando fila {rowNum} para cliente: {nombre}, NIT: {nit}");
                            
                            // Leer los datos de la fila con las columnas correctas: L, M, N, O, U
                            var clienteRow = new ClienteRow
                            {
                                Numero = TryGetStringValue(row.Cell("L"), $"Error al leer número en fila {rowNum} para cliente {nombre}"),
                                Fecha = TryGetDateTimeValue(row.Cell("M"), DateTime.Now, $"Error al leer fecha en fila {rowNum} para cliente {nombre}"),
                                FechaVence = TryGetDateTimeValue(row.Cell("N"), DateTime.Now, $"Error al leer fecha de vencimiento en fila {rowNum} para cliente {nombre}"),
                                ValorTotal = TryGetDoubleValue(row.Cell("O"), 0, $"Error al leer valor total en fila {rowNum} para cliente {nombre}"),
                                NumDias = TryGetIntValue(row.Cell("U"), 0, $"Error al leer días en fila {rowNum} para cliente {nombre}")
                            };
                            
                            // Crear una clave única combinando nombre y NIT
                            string key = string.IsNullOrEmpty(nit) ? nombre : $"{nombre}_{nit}";
                            
                            // Si el cliente no existe en el diccionario, crearlo
                            if (!clientGroups.ContainsKey(key))
                            {
                                clientGroups[key] = new ClienteInfo
                                {
                                    Nombre = nombre,
                                    Nit = nit,
                                    Rows = new List<ClienteRow>()
                                };
                            }
                            
                            // Agregar la fila al grupo correspondiente del cliente
                            clientGroups[key].Rows.Add(clienteRow);
                        }
                        catch (Exception ex)
                        {
                            logCallback($"Error al procesar fila {rowNum}: {ex.Message}");
                            // Continuamos con la siguiente fila
                        }
                    }
                    
                    logCallback($"Se encontraron un total de {totalEmpresasEncontradas} registros de empresas en el Excel");
                    logCallback($"Se agruparon en {clientGroups.Count} clientes únicos para procesar");
                    
                    // Lista para rastrear empresas sin facturas recientes
                    List<string> empresasSinFacturasRecientes = new List<string>();
                    int empresasProcesadas = 0;
                    
                    // Generar documentos para cada cliente
                    foreach (var clientGroup in clientGroups)
                    {
                        empresasProcesadas++;
                        string statusPrefix = $"[{empresasProcesadas}/{clientGroups.Count}]";
                        
                        // Verificar si tiene facturas recientes (días >= -8)
                        var filteredRows = clientGroup.Value.Rows.Where(r => r.NumDias >= -8).ToList();
                        
                        if (filteredRows.Count == 0)
                        {
                            empresasSinFacturasRecientes.Add(clientGroup.Value.Nombre);
                            logCallback($"{statusPrefix} Cliente {clientGroup.Value.Nombre} no tiene facturas con días >= -8. Omitiendo.");
                            continue;
                        }
                        
                        logCallback($"{statusPrefix} Generando documento para {clientGroup.Value.Nombre}");
                        GenerateDocumentForClient(clientGroup.Value);
                    }
                    
                    logCallback($"Resumen de procesamiento:");
                    logCallback($"Total de empresas encontradas en Excel: {totalEmpresasEncontradas}");
                    logCallback($"Total de clientes únicos: {clientGroups.Count}");
                    logCallback($"Total de clientes sin facturas recientes: {empresasSinFacturasRecientes.Count}");
                    logCallback($"Total de PDFs generados: {generatedPdfs.Count}");
                    
                    // Mostrar empresas sin facturas recientes
                    if (empresasSinFacturasRecientes.Count > 0)
                    {
                        logCallback("Clientes sin facturas recientes (días >= -8):");
                        foreach (var empresa in empresasSinFacturasRecientes)
                        {
                            logCallback($"  - {empresa}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logCallback($"Error al procesar el archivo Excel: {ex.Message}");
                if (ex.InnerException != null)
                {
                    logCallback($"Error interno: {ex.InnerException.Message}");
                }
            }
        }

        // Métodos auxiliares para extraer valores de celdas con manejo de errores
        private string TryGetStringValue(IXLCell? cell, string errorMessage)
        {
            try
            {
                if (cell == null || cell.IsEmpty()) return string.Empty;
                return cell.GetString()?.Trim() ?? string.Empty;
            }
            catch (Exception)
            {
                logCallback(errorMessage);
                return string.Empty;
            }
        }

        private DateTime TryGetDateTimeValue(IXLCell cell, DateTime defaultValue, string errorMessage)
        {
            try
            {
                if (cell == null || cell.IsEmpty()) return defaultValue;
                return cell.GetDateTime();
            }
            catch (Exception)
            {
                try
                {
                    // Intenta extraer como string y convertir
                    string dateStr = cell.GetString().Trim();
                    if (DateTime.TryParse(dateStr, out DateTime result))
                        return result;
                }
                catch {}
                
                logCallback(errorMessage);
                return defaultValue;
            }
        }

        private double TryGetDoubleValue(IXLCell cell, double defaultValue, string errorMessage)
        {
            try
            {
                if (cell == null || cell.IsEmpty()) return defaultValue;
                return cell.GetDouble();
            }
            catch (Exception)
            {
                try
                {
                    // Intenta extraer como string y convertir
                    string valueStr = cell.GetString().Trim();
                    if (double.TryParse(valueStr, out double result))
                        return result;
                }
                catch {}
                
                logCallback(errorMessage);
                return defaultValue;
            }
        }

        private int TryGetIntValue(IXLCell cell, int defaultValue, string errorMessage)
        {
            try
            {
                if (cell == null || cell.IsEmpty()) return defaultValue;
                
                // Intentar diferentes métodos para obtener un entero
                if (cell.DataType == XLDataType.Number)
                {
                    return Convert.ToInt32(cell.GetDouble());
                }
                else if (cell.DataType == XLDataType.Text)
                {
                    string valueStr = cell.GetString().Trim();
                    if (int.TryParse(valueStr, out int result))
                        return result;
                }
                
                // Si llegamos aquí, intenta una conversión general desde el valor
                if (!cell.Value.IsBlank)
                {
                    try { return Convert.ToInt32(cell.Value.GetNumber()); } 
                    catch {}
                }
                
                logCallback(errorMessage);
                return defaultValue;
            }
            catch (Exception)
            {
                logCallback(errorMessage);
                return defaultValue;
            }
        }

        // Modificar el método para aceptar ClienteInfo en lugar de nombre y lista
        private void GenerateDocumentForClient(ClienteInfo clienteInfo)
        {
            string clientName = clienteInfo.Nombre;
            List<ClienteRow> clientRows = clienteInfo.Rows;
            string nit = clienteInfo.Nit;
            
            logCallback($"Generando documento para cliente: {clientName}, NIT: {nit}");
            
            // Filtrar filas por días >= -8
            var filteredRows = clientRows.Where(r => r.NumDias >= -8).ToList();
            if (filteredRows.Count == 0)
            {
                logCallback($"Cliente {clientName} no tiene facturas con días >= -8. Omitiendo.");
                return;
            }
            
            try
            {
                // Determinar tipo de cartera
                bool tienePositivos = filteredRows.Any(r => r.NumDias >= 0);
                bool tienePendiente = filteredRows.Any(r => r.NumDias < 0 && r.NumDias >= -8);
                string sufijo;
                string tipoCartera;
                
                if (tienePositivos)
                {
                    sufijo = "_CarteraVencida";
                    tipoCartera = "Vencida";
                    logCallback($"Cliente {clientName}: Tipo de cartera = VENCIDA");
                }
                else if (tienePendiente)
                {
                    sufijo = "_CarteraPorVencer";
                    tipoCartera = "Por Vencer";
                    logCallback($"Cliente {clientName}: Tipo de cartera = POR VENCER");
                }
                else
                {
                    logCallback($"Cliente {clientName}: Sin cartera vencida o por vencer. Omitiendo.");
                    return;
                }
                
                // Generar nombre de archivo válido, incluyendo el NIT si está disponible
                string nombreBase = clientName;
                if (!string.IsNullOrEmpty(nit))
                {
                    nombreBase = $"{clientName}_{nit}";
                }
                
                string nombreArchivo = SanitizeFileName(nombreBase.Length > 50 ? nombreBase.Substring(0, 50) : nombreBase);
                string pdfFileName = $"{nombreArchivo}{sufijo}.pdf";
                string pdfFilePath = System.IO.Path.Combine(pdfOutputFolder, pdfFileName);
                
                // Crear PDF directamente con iText7
                CreatePdfDocument(clientName, nit, filteredRows, tienePositivos, pdfFilePath);
                
                // NUEVO: Registrar el PDF generado en la lista
                generatedPdfs.Add(new GeneratedPdfInfo
                {
                    NombreEmpresa = clientName,
                    Nit = nit,
                    RutaArchivo = pdfFilePath,
                    NombreArchivo = pdfFileName,
                    TipoCartera = tipoCartera
                });
                
                logCallback($"Documento generado para {clientName} (NIT: {nit}): {pdfFileName}");
            }
            catch (Exception ex)
            {
                logCallback($"Error al generar documento para {clientName}: {ex.Message}");
            }
        }

        // Modificar el método para incluir el NIT en el documento
        private void CreatePdfDocument(string clientName, string nit, List<ClienteRow> rows, bool tienePositivos, string outputPath)
        {
            // Fecha formateada
            string fechaFormateada = $"Barranquilla, {DateTime.Now.Day} de {CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(DateTime.Now.Month).ToLower()} de {DateTime.Now.Year}";
            
            try
            {
                // Crear PDF con iText7
                using (var writer = new PdfWriter(outputPath))
                {
                    using (var pdf = new PdfDocument(writer))
                    {
                        // Establecer explícitamente el tamaño de página
                        pdf.SetDefaultPageSize(PageSize.LETTER);
                        
                        // Configurar el evento de fondo si existe la plantilla PNG
                        if (templateFound && !string.IsNullOrEmpty(templatePath) && System.IO.File.Exists(templatePath))
                        {
                            try 
                            {
                                logCallback($"Aplicando plantilla PNG como fondo: {templatePath}");
                                pdf.AddEventHandler(PdfDocumentEvent.START_PAGE, new BackgroundImageHandler(templatePath, logCallback));
                                logCallback($"Plantilla PNG aplicada correctamente");
                            }
                            catch (Exception ex)
                            {
                                logCallback($"Error al aplicar plantilla PNG como fondo: {ex.Message}");
                                // Continuar sin fondo si hay error
                            }
                        }
                        else 
                        {
                            logCallback("No se aplicará fondo porque no se encontró la plantilla PNG.");
                        }
                        
                        using (var document = new Document(pdf))
                        {
                            // Ajustar márgenes superior e inferior de la página
                            document.SetMargins(80,36, 30, 36); // Top, Right, Bottom, Left (en puntos)

                            // Agregar fecha - alineado a la derecha
                            var parrafoFecha = new Paragraph(fechaFormateada)
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetMarginBottom(8); // Reducir espacio después de la fecha
                            document.Add(parrafoFecha);
                            
                            // Agregar encabezado para cliente - mantener alineación izquierda para encabezados
                            var paraCliente = new Paragraph("Señor(a):")
                                .SetTextAlignment(TextAlignment.LEFT)
                                .SetMarginTop(8) // Espacio antes del párrafo
                                .SetMarginBottom(2); // Espacio mínimo después del párrafo
                            paraCliente.SetBold();
                            document.Add(paraCliente);
                            
                            // Agregar nombre del cliente y NIT si está disponible
                            Paragraph paraClienteNombre;
                            if (!string.IsNullOrEmpty(nit))
                            {
                                paraClienteNombre = new Paragraph($"{clientName} (NIT: {nit})")
                                    .SetTextAlignment(TextAlignment.LEFT)
                                    .SetMarginBottom(8);
                            }
                            else
                            {
                                paraClienteNombre = new Paragraph(clientName)
                                    .SetTextAlignment(TextAlignment.LEFT)
                                    .SetMarginBottom(8);
                            }
                            document.Add(paraClienteNombre);
                            
                            // Preparar asunto
                            var paraAsunto = new Paragraph()
                                .SetTextAlignment(TextAlignment.LEFT)
                                .SetMarginBottom(8);
                            var textoAsunto = new Text("Asunto: ");
                            textoAsunto.SetBold();
                            paraAsunto.Add(textoAsunto);
                            
                            string asunto;
                            
                            if (tienePositivos)
                            {
                                asunto = "Estado de Cartera vencida";
                            }
                            else
                            {
                                asunto = "Aviso de proximidad de vencimiento de factura(s)";
                            }
                            
                            paraAsunto.Add(asunto);
                            document.Add(paraAsunto);
                            
                            // Agregar párrafo principal - JUSTIFICADO
                            var parrafoTexto = "Para SIMICS GROUP S.A.S. es muy importante contar con clientes como usted y mantenerlo informado sobre la situación actual de su cartera. Adjuntamos el estado de cuenta correspondiente; si tiene alguna observación, le agradecemos que nos la comunique por este medio para su pronta revisión.";
                            var paraPrincipal = new Paragraph(parrafoTexto)
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetMarginTop(8)
                                .SetMarginBottom(10);
                            document.Add(paraPrincipal);
                            
                            // Crear tabla para datos - ajustar al contenido y añadir márgenes
                            var table = new Table(UnitValue.CreatePercentArray(new float[] { 2, 2, 3, 2, 1 }))
                                .UseAllAvailableWidth()
                                .SetMarginTop(20) // Margen superior
                                .SetMarginBottom(20); // Margen inferior
                            
                            // Encabezados de tabla con alineación centrada
                            Cell[] headerCells = new Cell[] {
                                new Cell().Add(new Paragraph("Número de Documento")).SetBold().SetTextAlignment(TextAlignment.CENTER),
                                new Cell().Add(new Paragraph("Fecha de Emisión")).SetBold().SetTextAlignment(TextAlignment.CENTER),
                                new Cell().Add(new Paragraph("Fecha de Vencimiento del Documento")).SetBold().SetTextAlignment(TextAlignment.CENTER),
                                new Cell().Add(new Paragraph("Valor Total")).SetBold().SetTextAlignment(TextAlignment.CENTER),
                                new Cell().Add(new Paragraph("Días")).SetBold().SetTextAlignment(TextAlignment.CENTER)
                            };
                            
                            foreach (var cell in headerCells)
                            {
                                table.AddHeaderCell(cell);
                            }
                            
                            // Agregar datos a la tabla con alineación centrada
                            foreach (var row in rows)
                            {
                                table.AddCell(
                                    new Cell().Add(new Paragraph(row.Numero))
                                        .SetTextAlignment(TextAlignment.CENTER));
                                
                                table.AddCell(
                                    new Cell().Add(new Paragraph(row.Fecha.ToShortDateString()))
                                        .SetTextAlignment(TextAlignment.CENTER));
                                
                                table.AddCell(
                                    new Cell().Add(new Paragraph(row.FechaVence.ToShortDateString()))
                                        .SetTextAlignment(TextAlignment.CENTER));
                                
                                table.AddCell(
                                    new Cell().Add(new Paragraph(string.Format(CultureInfo.CurrentCulture, "{0:C}", row.ValorTotal)))
                                        .SetTextAlignment(TextAlignment.RIGHT)); // Alineación derecha para valores monetarios
                                
                                table.AddCell(
                                    new Cell().Add(new Paragraph(row.NumDias.ToString()))
                                        .SetTextAlignment(TextAlignment.CENTER));
                            }
                            
                            document.Add(table);
                            
                            // Agregar total - alineado a la derecha
                            double valorTotal = rows.Sum(r => r.ValorTotal);
                            var parrafoTotal = new Paragraph()
                                .SetTextAlignment(TextAlignment.RIGHT)
                                .SetMarginTop(10)
                                .SetMarginBottom(10);
                            var textoTotal = new Text("Total de Deuda: ");
                            textoTotal.SetBold();
                            parrafoTotal.Add(textoTotal);
                            parrafoTotal.Add(string.Format(CultureInfo.CurrentCulture, "{0:C}", valorTotal));
                            document.Add(parrafoTotal);

                            // Crear un Div para agrupar todo el contenido que debe mantenerse junto
                            Div footerGroup = new Div().SetKeepTogether(true);
                            
                            // Agregar texto estándar - JUSTIFICADO
                            var paraEstandar = new Paragraph("El pago de sus facturas nos ayuda a cumplir nuestros compromisos financieros.")
                                .SetTextAlignment(TextAlignment.JUSTIFIED)
                                .SetMarginBottom(8);
                            footerGroup.Add(paraEstandar);

                            // Agregar línea de despedida y firma - alineados a la izquierda
                            var despedida = new Paragraph("Cordialmente,")
                                .SetTextAlignment(TextAlignment.LEFT)
                                .SetMarginBottom(8);
                            footerGroup.Add(despedida);

                            // Agregar imagen de firma
                            string firmaPath = @"D:\Alerta Cartera\Background\MiProyectoWPF\Archivos\firma.png";
                            logCallback($"Buscando archivo de firma en: {firmaPath}");

                            if (System.IO.File.Exists(firmaPath))
                            {
                                try
                                {
                                    logCallback("Archivo de firma encontrado. Agregando al documento...");
                                    ImageData firmaImage = ImageDataFactory.Create(firmaPath);
                                    iText.Layout.Element.Image firma = new iText.Layout.Element.Image(firmaImage)
                                        .SetWidth(100) // Ajustar ancho
                                        .SetHorizontalAlignment(HorizontalAlignment.LEFT);
                                    footerGroup.Add(firma);
                                }
                                catch (Exception ex)
                                {
                                    logCallback($"Error al agregar firma: {ex.Message}");
                                }
                            }
                            else
                            {
                                logCallback($"Archivo de firma no encontrado en: {firmaPath}");
                            }

                            // Agregar datos de contacto
                            footerGroup.Add(new Paragraph("JUAN MANUEL CUERVO").SetBold().SetTextAlignment(TextAlignment.LEFT));
                            footerGroup.Add(new Paragraph("GERENTE FINANCIERO").SetBold().SetTextAlignment(TextAlignment.LEFT));
                            footerGroup.Add(new Paragraph("SIMICS GROUP S.A.S.").SetBold().SetTextAlignment(TextAlignment.LEFT));
                            
                            // Agregar todo el grupo al documento
                            document.Add(footerGroup);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logCallback($"Error al crear el documento PDF: {ex.Message}");
                throw;
            }
        }

        private string SanitizeFileName(string fileName)
        {
            var invalidChars = System.IO.Path.GetInvalidFileNameChars();
            return new string(fileName.Where(c => !invalidChars.Contains(c)).ToArray());
        }

        // Método para guardar la lista de PDFs generados en un archivo de registro
        public void SaveGeneratedPdfsList(string outputFilePath = null)
        {
            string path = outputFilePath ?? System.IO.Path.Combine(pdfOutputFolder, "pdfs_generados.txt");
            
            try
            {
                using (var writer = new StreamWriter(path))
                {
                    writer.WriteLine("Fecha de generación: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    writer.WriteLine($"Total de PDFs generados: {generatedPdfs.Count}");
                    writer.WriteLine("-------------------------------------------------------------");
                    
                    foreach (var pdf in generatedPdfs)
                    {
                        writer.WriteLine($"Empresa: {pdf.NombreEmpresa}");
                        writer.WriteLine($"NIT: {pdf.Nit}");
                        writer.WriteLine($"Archivo: {pdf.NombreArchivo}");
                        writer.WriteLine($"Tipo: {pdf.TipoCartera}");
                        writer.WriteLine($"Ruta: {pdf.RutaArchivo}");
                        writer.WriteLine("-------------------------------------------------------------");
                    }
                }
                
                logCallback($"Lista de PDFs generados guardada en: {path}");
            }
            catch (Exception ex)
            {
                logCallback($"Error al guardar lista de PDFs generados: {ex.Message}");
            }
        }

        // Método para identificar PDFs sin destinatario en el Excel
        public List<GeneratedPdfInfo> IdentifyOrphanedPdfs(List<string> empresasEnExcel)
        {
            List<GeneratedPdfInfo> orphanedPdfs = new List<GeneratedPdfInfo>();
            
            if (empresasEnExcel == null)
            {
                logCallback("Error: Lista de empresas nula al identificar PDFs huérfanos.");
                return orphanedPdfs;
            }
            
            foreach (var pdf in generatedPdfs)
            {
                bool found = false;
                
                // Buscar por nombre de empresa exacto
                if (empresasEnExcel.Contains(pdf.NombreEmpresa, StringComparer.OrdinalIgnoreCase))
                {
                    found = true;
                }
                // También buscar por nombre truncado (para nombres largos)
                else if (pdf.NombreEmpresa.Length > 30)
                {
                    string nombreTruncado = pdf.NombreEmpresa.Substring(0, 30);
                    if (empresasEnExcel.Contains(nombreTruncado, StringComparer.OrdinalIgnoreCase))
                    {
                        found = true;
                    }
                }
                
                if (!found)
                {
                    orphanedPdfs.Add(pdf);
                }
            }
            
            return orphanedPdfs;
        }
    }

    // Clase para almacenar información del cliente incluyendo su NIT
    public class ClienteInfo
    {
        public string Nombre { get; set; } = string.Empty;
        public string Nit { get; set; } = string.Empty;
        public List<ClienteRow> Rows { get; set; } = new List<ClienteRow>();
    }

    public class ClienteRow
    {
        public string Numero { get; set; } = string.Empty;
        public DateTime Fecha { get; set; }
        public DateTime FechaVence { get; set; }
        public double ValorTotal { get; set; }
        public int NumDias { get; set; }
    }

    // Clase para almacenar información de PDFs generados
    public class GeneratedPdfInfo
    {
        public string NombreEmpresa { get; set; } = string.Empty;
        public string Nit { get; set; } = string.Empty;
        public string RutaArchivo { get; set; } = string.Empty;
        public string NombreArchivo { get; set; } = string.Empty;
        public string TipoCartera { get; set; } = string.Empty;
    }
}
