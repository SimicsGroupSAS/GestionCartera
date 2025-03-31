using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace MiProyectoWPF
{
    public partial class MainWindow : Window
    {
        private List<CheckBox> checkboxes = new();
        private Dictionary<string, string> empresasArchivos = new(); // Empresa -> Ruta del archivo
        private Dictionary<string, string> empresasArchivosInfo = new Dictionary<string, string>(); // Info detallada de archivos
        private string? selectedExcelFilePath = string.Empty; // Excel para envío de correos
        private string? selectedGenerationExcelFilePath = string.Empty; // Excel para generación de documentos
        private readonly string tempFolderPath; // Carpeta temp
        private readonly string baseFolder; // Carpeta base para documentos y plantillas
        private string pdfOutputFolder; // Carpeta para guardar los PDFs generados
        private SimplePdfGenerator? documentGenerator; // Para mantener referencia al generador
        private string bccEmailAddress = "pofika1666@nokdot.com"; // Correo que recibirá copia de todos los envíos
        private string ccFinanzasEmail = ""; // Opcional: para enviar CC a otro departamento si se necesita
        private readonly string emailSubject = "ESTADO DE CARTERA - SIMICS GROUP SAS NIT 900853554-3"; // Asunto estándar para todos los correos
        private readonly Dictionary<string, string> nitNormalizadoCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase); // Caché para NITs normalizados

        public MainWindow()
        {
            try
            {
                string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string projectDirectory = Path.GetFullPath(Path.Combine(appDirectory, "..", ".."));
                baseFolder = Path.GetFullPath(Path.Combine(projectDirectory, "..")); // Sube un nivel más para llegar a d:\Alerta Cartera
                
                tempFolderPath = Path.Combine(projectDirectory, "temp");

                Console.WriteLine($"Ubicación del proyecto: {projectDirectory}");
                Console.WriteLine($"Carpeta base seleccionada: {baseFolder}");
                Console.WriteLine($"Carpeta temporal: {tempFolderPath}");

                if (!Directory.Exists(baseFolder))
                {
                    Directory.CreateDirectory(baseFolder);
                    MessageBox.Show($"Se ha creado la carpeta base: {baseFolder}", "Información");
                }

                InitializeComponent();

                this.Loaded += MainWindow_Loaded;

                CreateTempFolder();
                
                txtGeneracionLog.TextChanged += (s, e) => {
                    if (s is TextBox tb) {
                        tb.ScrollToEnd();
                        Dispatcher.InvokeAsync(() => tb.ScrollToEnd(), 
                            System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                    }
                };
                
                txtEnvioLog.TextChanged += (s, e) => {
                    if (s is TextBox tb) {
                        tb.ScrollToEnd();
                        Dispatcher.InvokeAsync(() => tb.ScrollToEnd(), 
                            System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                    }
                };
                
                string archivosPath = Path.Combine(baseFolder, "Archivos");
                if (!Directory.Exists(archivosPath))
                {
                    Directory.CreateDirectory(archivosPath);
                    Console.WriteLine($"Carpeta 'Archivos' creada en: {archivosPath}");
                }
                
                pdfOutputFolder = Path.Combine(archivosPath, "Clientes cartera pdf");
                if (!Directory.Exists(pdfOutputFolder))
                {
                    Directory.CreateDirectory(pdfOutputFolder);
                    Console.WriteLine($"Carpeta 'Clientes cartera pdf' creada en: {pdfOutputFolder}");
                }
                
                mainTabControl.SelectionChanged += (s, e) => {
                    if (e.Source is TabControl)
                    {
                        Console.WriteLine($"Cambiado a pestaña: {(e.Source as TabControl)?.SelectedIndex}");
                    }
                };
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error durante la inicialización: {ex.Message}\n\nDetalles: {ex.StackTrace}", "Error Crítico", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.Loaded -= MainWindow_Loaded;
            CheckAndLoadCredentials();
        }

        private void CheckAndLoadCredentials()
        {
            try
            {
                var credentials = CredentialManager.ReadSmtpCredentials();
                if (credentials == null)
                {
                    LogMessage("No se pudieron cargar las credenciales. Usando valores predeterminados.");
                    credentials = new CredentialManager.SmtpCredentials();
                }
                
                if (!credentials.IsComplete)
                {
                    LogMessage("No se encontraron credenciales SMTP. Solicitando al usuario...");
                    
                    var dialog = new CredentialDialog();
                    dialog.Owner = this;
                    
                    bool? result = dialog.ShowDialog();
                    
                    if (result == true)
                    {
                        credentials = dialog.Credentials;
                        LogMessage("Credenciales SMTP configuradas por el usuario.");
                    }
                    else
                    {
                        LogMessage("El usuario canceló la configuración de credenciales. La aplicación se cerrará.");
                        MessageBox.Show(
                            "Es necesario configurar las credenciales de correo para utilizar la aplicación. La aplicación se cerrará.",
                            "Configuración requerida",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                        
                        Application.Current.Shutdown();
                        return;
                    }
                }
                
                bccEmailAddress = credentials.BccEmail;
                ccFinanzasEmail = credentials.CcEmail;
                
                LogMessage($"Credenciales cargadas para: {credentials.Username}");
            }
            catch (Exception ex)
            {
                LogMessage($"Error al cargar credenciales: {ex.Message}");
                
                MessageBox.Show(
                    $"Error al cargar credenciales: {ex.Message}\nLa aplicación se cerrará.",
                    "Error de configuración",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                
                Application.Current.Shutdown();
            }
        }

        #region Métodos de Inicio

        private void CreateTempFolder()
        {
            try
            {
                if (!Directory.Exists(tempFolderPath))
                {
                    Directory.CreateDirectory(tempFolderPath);
                    Console.WriteLine($"Carpeta 'temp' creada en: {tempFolderPath}");
                }
                else
                {
                    Console.WriteLine($"La carpeta 'temp' ya existe en: {tempFolderPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al crear la carpeta 'temp': {ex.Message}");
                MessageBox.Show($"Error al crear la carpeta 'temp': {ex.Message}", "Error");
            }
        }
        
        private void GenerarEstadosCuenta_Click(object sender, RoutedEventArgs e)
        {
            mainTabControl.SelectedItem = tabGeneracionDocumentos;
        }
        
        private void EnviarEstadosCuenta_Click(object sender, RoutedEventArgs e)
        {
            mainTabControl.SelectedItem = tabEnvioCorreos;
        }

        #endregion

        #region Generación de Documentos

        private void SelectExcelFileForGeneration_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Archivos Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*",
                Title = "Seleccionar Archivo Excel de Cartera"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                selectedGenerationExcelFilePath = openFileDialog.FileName;
                txtExcelSeleccionado.Text = $"Archivo seleccionado: {Path.GetFileName(selectedGenerationExcelFilePath)}";
                
                txtExcelSeleccionado.Text += $"\nLos documentos se generarán en: {pdfOutputFolder}";
            }
        }

        private async void ProcessExcelAndGenerateDocuments_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedGenerationExcelFilePath))
            {
                MessageBox.Show("Por favor, seleccione un archivo Excel primero.", "Archivo no seleccionado");
                return;
            }
            
            try
            {
                documentGenerator = new SimplePdfGenerator(baseFolder, pdfOutputFolder, LogMessage, tempFolderPath);
                
                if (!documentGenerator.IsTemplateFound)
                {
                    MessageBox.Show("No se encontró la plantilla PNG requerida para generar los documentos.", "Error de plantilla");
                    return;
                }
                
                txtEstadoGeneracion.Text = "Generando documentos... Por favor espere.";
                txtGeneracionLog.Clear();
                
                await documentGenerator.GenerateDocuments(selectedGenerationExcelFilePath);
                
                documentGenerator.SaveGeneratedPdfsList();
                
                var generatedFiles = documentGenerator.GeneratedPdfFiles;
                
                txtEstadoGeneracion.Text = $"¡Proceso completado! Se generaron {generatedFiles.Count} documentos.";
                
                int vencida = generatedFiles.Count(p => p.TipoCartera == "Vencida");
                int porVencer = generatedFiles.Count(p => p.TipoCartera == "Por Vencer");
                
                MessageBox.Show(
                    $"Documentos generados exitosamente:\n\n" +
                    $"- Total: {generatedFiles.Count} documentos\n" +
                    $"- Cartera Vencida: {vencida}\n" +
                    $"- Cartera Por Vencer: {porVencer}\n\n" +
                    $"Ahora puede enviar los correos.",
                    "Proceso completado");
                
                var result = MessageBox.Show(
                    "¿Desea ir al paso de envío de correos?", 
                    "Siguiente paso", 
                    MessageBoxButton.YesNo);
                
                if (result == MessageBoxResult.Yes)
                {
                    mainTabControl.SelectedItem = tabEnvioCorreos;
                }
            }
            catch (Exception ex)
            {
                txtEstadoGeneracion.Text = "Error al generar documentos.";
                LogMessage($"Error: {ex.Message}");
                MessageBox.Show($"Error al generar los documentos: {ex.Message}", "Error");
            }
        }
        
        private void LogTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox textBox)
            {
                textBox.CaretIndex = textBox.Text.Length;
                textBox.ScrollToEnd();
            }
        }

        private void LogMessage(string message)
        {
            Dispatcher.Invoke(() => 
            {
                TextBox targetTextBox;
                
                if (mainTabControl.SelectedItem == tabGeneracionDocumentos)
                {
                    targetTextBox = txtGeneracionLog;
                }
                else if (mainTabControl.SelectedItem == tabEnvioCorreos)
                {
                    targetTextBox = txtEnvioLog;
                }
                else
                {
                    targetTextBox = txtGeneracionLog;
                }
                
                targetTextBox.AppendText(message + Environment.NewLine);
                
                targetTextBox.CaretIndex = targetTextBox.Text.Length;
                targetTextBox.ScrollToEnd();
                
                ForzarScrollToEnd(targetTextBox);
                
                Console.WriteLine(message);
            });
        }

        private void ForzarScrollToEnd(TextBox textBox)
        {
            textBox.ScrollToEnd();
            
            if (!textBox.IsFocused)
            {
                textBox.Focus();
                textBox.ScrollToEnd();
            }
            
            Dispatcher.InvokeAsync(() => {
                textBox.ScrollToEnd();
            }, System.Windows.Threading.DispatcherPriority.Normal);
            
            Dispatcher.InvokeAsync(() => {
                textBox.ScrollToEnd();
            }, System.Windows.Threading.DispatcherPriority.ApplicationIdle);
            
            Dispatcher.InvokeAsync(() => {
                textBox.ScrollToEnd();
            }, System.Windows.Threading.DispatcherPriority.Render);
            
            Dispatcher.InvokeAsync(() => {
                textBox.ScrollToEnd();
                textBox.UpdateLayout();
            }, System.Windows.Threading.DispatcherPriority.Background);
            
            Dispatcher.InvokeAsync(() => {
                System.Threading.Thread.Sleep(15);
                textBox.ScrollToEnd();
            });
        }

        #endregion

        #region Envío de Correos

        private string NormalizarNIT(string nit)
        {
            if (string.IsNullOrWhiteSpace(nit))
                return string.Empty;
            
            // Verificar si ya está en caché
            if (nitNormalizadoCache.TryGetValue(nit, out string? resultado))
            {
                return resultado;
            }
            
            string nitNormalizado = string.Empty;
            
            // Manejo para formato "NIT 900832816 - 8"
            if (nit.StartsWith("NIT", StringComparison.OrdinalIgnoreCase))
            {
                // Extraer solo los dígitos y el guion
                nit = nit.Substring(3).Trim();
            }
            
            // Eliminar espacios y caracteres no esenciales
            string soloDigitosGuion = new string(nit.Where(c => char.IsDigit(c) || c == '-').ToArray());
            
            // Si no tiene guion pero parece ser un NIT completo, añadir el guion
            if (!soloDigitosGuion.Contains('-') && soloDigitosGuion.Length >= 9)
            {
                // Extraer el último dígito como dígito de verificación
                string baseNit = soloDigitosGuion.Substring(0, soloDigitosGuion.Length - 1);
                string digitoVerificacion = soloDigitosGuion[soloDigitosGuion.Length - 1].ToString();
                soloDigitosGuion = $"{baseNit}-{digitoVerificacion}";
            }
            
            nitNormalizado = soloDigitosGuion;
            
            // Almacenar en caché
            nitNormalizadoCache[nit] = nitNormalizado;
            
            return nitNormalizado;
        }

        private async Task LoadAllPdfFilesAsync(Dictionary<string, string> pdfFiles, Dictionary<string, string> empresasNits)
        {
            LogMessage("Cargando archivos PDF...");
            
            await Task.Run(() => {
                try
                {
                    if (!Directory.Exists(pdfOutputFolder))
                    {
                        LogMessage($"⚠️ La carpeta de PDF no existe: {pdfOutputFolder}");
                        return;
                    }
                    
                    // Lista para almacenar las empresas encontradas
                    HashSet<string> empresasEncontradas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    
                    // Procesar el archivo de registro si existe
                    string pdfsGeneradosPath = Path.Combine(pdfOutputFolder, "pdfs_generados.txt");
                    bool loadedFromFile = false;
                    
                    if (File.Exists(pdfsGeneradosPath))
                    {
                        LogMessage("Cargando información desde pdfs_generados.txt");
                        
                        string currentEmpresa = "";
                        string currentNit = "";
                        string currentRuta = "";
                        
                        string[] lines = File.ReadAllLines(pdfsGeneradosPath);
                        foreach (var line in lines)
                        {
                            string trimmedLine = line.Trim();
                            
                            if (string.IsNullOrEmpty(trimmedLine) || trimmedLine.StartsWith("Fecha") || 
                                trimmedLine.StartsWith("Total") || trimmedLine.Contains("---------"))
                                continue;
                            
                            if (trimmedLine.StartsWith("Empresa:"))
                                currentEmpresa = trimmedLine.Replace("Empresa:", "").Trim();
                            else if (trimmedLine.StartsWith("NIT:"))
                                currentNit = trimmedLine.Replace("NIT:", "").Trim();
                            else if (trimmedLine.StartsWith("Ruta:"))
                            {
                                currentRuta = trimmedLine.Replace("Ruta:", "").Trim();
                                
                                if (!string.IsNullOrEmpty(currentEmpresa) && !string.IsNullOrEmpty(currentRuta) && 
                                    File.Exists(currentRuta))
                                {
                                    // Agregar a la lista de empresas encontradas
                                    empresasEncontradas.Add(currentEmpresa);
                                    
                                    pdfFiles[currentEmpresa] = currentRuta;
                                    
                                    if (!string.IsNullOrEmpty(currentNit))
                                    {
                                        empresasNits[currentEmpresa] = currentNit;
                                        // Normalizar el NIT sólo para empresas reales
                                        if (!string.IsNullOrEmpty(currentNit))
                                        {
                                            string nitNormalizado = NormalizarNIT(currentNit);
                                            if (!string.IsNullOrEmpty(nitNormalizado))
                                                empresasNits[nitNormalizado] = currentEmpresa;
                                        }
                                    }
                                    
                                    loadedFromFile = true;
                                }
                                
                                currentEmpresa = "";
                                currentNit = "";
                                currentRuta = "";
                            }
                        }
                        
                        LogMessage($"Se cargaron {pdfFiles.Count} PDFs desde el archivo de registro");
                    }
                    
                    if (!loadedFromFile)
                    {
                        LogMessage("Buscando PDFs directamente en la carpeta de salida");
                        
                        if (Directory.Exists(pdfOutputFolder))
                        {
                            string[] files = Directory.GetFiles(pdfOutputFolder, "*.pdf");
                            
                            foreach (var file in files)
                            {
                                string fileName = Path.GetFileNameWithoutExtension(file);
                                
                                string nombreBase = fileName
                                    .Replace("_CarteraVencida", "")
                                    .Replace("_CarteraPorVencer", "");
                                
                                string empresa = nombreBase;
                                string nit = "";
                                
                                if (nombreBase.Contains("_"))
                                {
                                    var parts = nombreBase.Split('_');
                                    if (parts.Length >= 2)
                                    {
                                        empresa = parts[0];
                                        nit = parts[1];
                                    }
                                }
                                
                                // Agregar a la lista de empresas encontradas
                                empresasEncontradas.Add(empresa);
                                
                                pdfFiles[empresa] = file;
                                
                                if (!string.IsNullOrEmpty(nit))
                                {
                                    empresasNits[empresa] = nit;
                                    
                                    // Normalizar el NIT sólo para empresas encontradas
                                    string nitNormalizado = NormalizarNIT(nit);
                                    if (!string.IsNullOrEmpty(nitNormalizado))
                                        empresasNits[nitNormalizado] = empresa;
                                }
                            }
                            
                            LogMessage($"Se encontraron {pdfFiles.Count} archivos PDF en la carpeta");
                        }
                    }
                    
                    LogMessage($"Total de NITs normalizados y cacheados: {nitNormalizadoCache.Count}");
                }
                catch (Exception ex)
                {
                    LogExceptionDetails(ex, "LoadAllPdfFilesAsync");
                    Dispatcher.Invoke(() => {
                        MessageBox.Show($"Error al cargar los archivos PDF: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    });
                }
            });
        }

        private async void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            LogMessage("Iniciando selección de archivo Excel");
            
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Archivos Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*",
                Title = "Seleccionar Archivo Excel con Correos"
            };

            LogMessage($"Mostrando diálogo de selección de archivo");
            bool? dialogResult = openFileDialog.ShowDialog();
            LogMessage($"Resultado del diálogo: {dialogResult}");

            if (dialogResult == true)
            {
                selectedExcelFilePath = openFileDialog.FileName;
                LogMessage($"Archivo seleccionado: {selectedExcelFilePath}");
                LogFileDetails(selectedExcelFilePath, "SelectExcelFile");
                
                var progressWindow = new ProgressWindow("Procesando archivo Excel...");
                progressWindow.Owner = this;
                progressWindow.Show();
                
                try
                {
                    await Task.Run(() => LoadCheckboxesAsync());
                }
                catch (Exception ex)
                {
                    LogExceptionDetails(ex, "SelectExcelFile");
                    MessageBox.Show($"Error al cargar datos del Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    progressWindow.Close();
                }
            }
            else
            {
                LogMessage("Selección de archivo cancelada por el usuario");
            }
        }

        private async Task LoadCheckboxesAsync()
        {
            LogMessage("Iniciando LoadCheckboxesAsync");
            
            try
            {
                Dispatcher.Invoke(() => {
                    CheckboxContainer.Children.Clear();
                    checkboxes.Clear();
                    empresasArchivos.Clear();
                });
                
                Dictionary<string, string> pdfFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, string> empresasNits = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                
                await LoadAllPdfFilesAsync(pdfFiles, empresasNits);
                
                if (pdfFiles.Count == 0)
                {
                    Dispatcher.Invoke(() => {
                        MessageBox.Show("No se encontraron PDFs generados. Por favor, genere los documentos primero.", 
                                      "Sin documentos", MessageBoxButton.OK, MessageBoxImage.Information);
                    });
                    return;
                }
                
                LogMessage($"Se encontraron {pdfFiles.Count} archivos PDF generados");
                
                Dictionary<string, string> empresaCorreos = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, string> nitCorreos = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                
                if (!string.IsNullOrEmpty(selectedExcelFilePath) && File.Exists(selectedExcelFilePath))
                {
                    await LoadEmailsFromExcelAsync(empresaCorreos, nitCorreos);
                    LogMessage($"Se encontraron correos para {empresaCorreos.Count} empresas en el Excel");
                }
                else
                {
                    LogMessage("⚠️ No se ha seleccionado un archivo Excel válido para buscar correos");
                }
                
                int totalDisplayed = 0;
                int withEmail = 0;
                int withoutEmail = 0;
                
                foreach (var kvp in pdfFiles)
                {
                    string nombreEmpresa = kvp.Key;
                    string rutaPdf = kvp.Value;
                    string nit = empresasNits.ContainsKey(nombreEmpresa) ? empresasNits[nombreEmpresa] : "";
                    
                    bool tieneCorreo = empresaCorreos.TryGetValue(nombreEmpresa, out string correo);
                    if (!tieneCorreo && !string.IsNullOrEmpty(nit))
                    {
                        string nitNormalizado = NormalizarNIT(nit);
                        if (!string.IsNullOrEmpty(nitNormalizado))
                        {
                            tieneCorreo = nitCorreos.TryGetValue(nitNormalizado, out correo);
                        }
                    }
                    
                    if (!tieneCorreo || string.IsNullOrWhiteSpace(correo))
                    {
                        correo = "No se encontró correo electrónico";
                        withoutEmail++;
                    }
                    else
                    {
                        withEmail++;
                    }
                    
                    Dispatcher.Invoke(() => {
                        CrearControlParaEmpresa(nombreEmpresa, nit, correo, rutaPdf, tieneCorreo);
                    });
                    
                    empresasArchivos[nombreEmpresa] = rutaPdf;
                    totalDisplayed++;
                }
                
                LogMessage($"Total de empresas mostradas: {totalDisplayed}");
                LogMessage($"   Con correo: {withEmail}");
                LogMessage($"   Sin correo: {withoutEmail}");
            }
            catch (Exception ex)
            {
                LogExceptionDetails(ex, "LoadCheckboxesAsync");
                Dispatcher.Invoke(() => {
                    MessageBox.Show($"Error al cargar datos: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private async Task LoadEmailsFromExcelAsync(Dictionary<string, string> empresaCorreos, Dictionary<string, string> nitCorreos)
        {
            LogMessage("Cargando correos electrónicos del archivo Excel");
            
            await Task.Run(() => {
                try
                {
                    using (var workbook = new XLWorkbook(selectedExcelFilePath))
                    {
                        foreach (var worksheet in workbook.Worksheets)
                        {
                            LogMessage($"Procesando hoja: {worksheet.Name}");
                            
                            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
                            if (lastRow <= 1) continue;
                            
                            // Columnas específicas según la estructura real del Excel
                            int tipoDocCol = 2;    // Columna B - Tipo de documento (NIT)
                            int numIdCol = 3;      // Columna C - Número de identificación
                            int digitoVerCol = 4;  // Columna D - Dígito de verificación
                            int correoCol = 6;     // Columna F - Correo electrónico (FIJO)
                            int nombreCol = 7;     // Columna G - Nombre de la empresa (FIJO)
                            
                            LogMessage($"Usando estructura de columnas específica:");
                            LogMessage($"- Tipo Documento: Columna B ({tipoDocCol})");
                            LogMessage($"- Número ID: Columna C ({numIdCol})");
                            LogMessage($"- Dígito Verificación: Columna D ({digitoVerCol})");
                            LogMessage($"- Correo Electrónico: Columna F ({correoCol})");
                            LogMessage($"- Nombre Empresa: Columna G ({nombreCol})");
                            
                            int empresasConNit = 0;
                            int empresasConCorreo = 0;
                            
                            for (int row = 2; row <= lastRow; row++)
                            {
                                try
                                {
                                    // CORRECCIÓN: Leer el nombre de empresa de la columna G
                                    string nombre = worksheet.Cell(row, nombreCol).GetString().Trim();
                                    
                                    // Leer tipo doc y componentes del NIT
                                    string tipoDoc = worksheet.Cell(row, tipoDocCol).GetString().Trim();
                                    string numId = worksheet.Cell(row, numIdCol).GetString().Trim();
                                    string digitoVer = worksheet.Cell(row, digitoVerCol).GetString().Trim();
                                    
                                    // CORRECCIÓN: Leer correo de la columna F
                                    string correo = worksheet.Cell(row, correoCol).GetString().Trim();
                                    
                                    // Saltar filas con nombres vacíos o inválidos
                                    if (string.IsNullOrEmpty(nombre) || 
                                        nombre.ToLower().Contains("total") || 
                                        nombre.Equals("nombre", StringComparison.OrdinalIgnoreCase))
                                        continue;
                                    
                                    // Construir NIT completo si es tipo "NIT"
                                    string nitCompleto = string.Empty;
                                    if (tipoDoc.Equals("NIT", StringComparison.OrdinalIgnoreCase) && 
                                        !string.IsNullOrEmpty(numId))
                                    {
                                        if (!string.IsNullOrEmpty(digitoVer))
                                            nitCompleto = $"{numId}-{digitoVer}";
                                        else
                                            nitCompleto = numId;
                                        
                                        LogMessage($"NIT construido para {nombre}: {nitCompleto}");
                                    }
                                    
                                    // Guardar información de correo por nombre y por NIT
                                    if (!string.IsNullOrEmpty(correo))
                                    {
                                        // Guardar por nombre de empresa
                                        empresaCorreos[nombre] = correo;
                                        empresasConCorreo++;
                                        
                                        // Si hay NIT, también guardar indexado por NIT
                                        if (!string.IsNullOrEmpty(nitCompleto))
                                        {
                                            string nitNormalizado = NormalizarNIT(nitCompleto);
                                            if (!string.IsNullOrEmpty(nitNormalizado))
                                            {
                                                nitCorreos[nitNormalizado] = correo;
                                                
                                                string nitSinGuion = nitNormalizado.Replace("-", "");
                                                if (!string.IsNullOrEmpty(nitSinGuion))
                                                    nitCorreos[nitSinGuion] = correo;
                                                
                                                empresasConNit++;
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex) 
                                { 
                                    LogMessage($"Error al procesar fila {row}: {ex.Message}");
                                }
                            }
                            
                            LogMessage($"Total de empresas con correo: {empresasConCorreo}");
                            LogMessage($"Total de NITs asociados a correos: {empresasConNit}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"Error al procesar Excel: {ex.Message}");
                }
            });
        }

        private List<string> GetEmailsForCompany(string companyName)
        {
            List<string> emails = new List<string>();
            
            if (string.IsNullOrEmpty(selectedExcelFilePath) || !File.Exists(selectedExcelFilePath))
            {
                LogMessage($"Error: No se ha seleccionado un archivo Excel válido.");
                return emails;
            }
            
            LogMessage($"Buscando correos para empresa: {companyName}");
            
            try
            {
                string nitEmpresa = "";
                
                Dispatcher.Invoke(() =>
                {
                    foreach (var checkbox in checkboxes)
                    {
                        if (checkbox.Content?.ToString() == companyName && checkbox.Tag != null)
                        {
                            nitEmpresa = checkbox.Tag.ToString() ?? "";
                            LogMessage($"NIT encontrado para {companyName}: {nitEmpresa}");
                            break;
                        }
                    }
                });
                
                using var workbook = new XLWorkbook(selectedExcelFilePath);
                
                string nitNormalizado = NormalizarNIT(nitEmpresa);
                string nitSinGuion = nitNormalizado.Replace("-", "");
                
                foreach (var worksheet in workbook.Worksheets)
                {
                    LogMessage($"Revisando hoja: {worksheet.Name}");
                    
                    // CORRECCIÓN: Actualizar índices de columnas
                    int tipoDocCol = 2;    // Columna B
                    int numIdCol = 3;      // Columna C
                    int digitoVerCol = 4;  // Columna D
                    int correoCol = 6;     // Columna F - CORREO
                    int nombreCol = 7;     // Columna G - NOMBRE
                    
                    LogMessage("Usando estructura de columnas específica para buscar correos:");
                    LogMessage($"- Correos en columna F ({correoCol})");
                    LogMessage($"- Nombres en columna G ({nombreCol})");
                    
                    var filas = worksheet.RowsUsed().Skip(1);
                    
                    foreach (var row in filas)
                    {
                        try {
                            // CORRECCIÓN: Leer nombre de columna G
                            string nombre = row.Cell(nombreCol).GetString().Trim();
                            
                            // Leer componentes del NIT
                            string tipoDoc = row.Cell(tipoDocCol).GetString().Trim();
                            string numId = row.Cell(numIdCol).GetString().Trim();
                            string digitoVer = row.Cell(digitoVerCol).GetString().Trim();
                            
                            // CORRECCIÓN: Leer correo de columna F
                            string correo = row.Cell(correoCol).GetString().Trim();
                            
                            // Construir NIT completo
                            string nitCompleto = string.Empty;
                            if (tipoDoc.Equals("NIT", StringComparison.OrdinalIgnoreCase) && 
                                !string.IsNullOrEmpty(numId))
                            {
                                if (!string.IsNullOrEmpty(digitoVer))
                                    nitCompleto = $"{numId}-{digitoVer}";
                                else
                                    nitCompleto = numId;
                            }
                            
                            // Normalizar el NIT para comparación
                            string nitFilaNormalizado = NormalizarNIT(nitCompleto);
                            string nitFilaSinGuion = nitFilaNormalizado.Replace("-", "");
                            
                            // Diferentes tipos de coincidencia
                            bool coincidePorNombre = string.Equals(nombre, companyName, StringComparison.OrdinalIgnoreCase);
                            
                            bool coincidePorNIT = !string.IsNullOrEmpty(nitNormalizado) && 
                                                !string.IsNullOrEmpty(nitFilaNormalizado) && 
                                                (nitNormalizado.Equals(nitFilaNormalizado, StringComparison.OrdinalIgnoreCase) || 
                                                nitSinGuion.Equals(nitFilaSinGuion, StringComparison.OrdinalIgnoreCase));
                            
                            bool nitEnNombre = !string.IsNullOrEmpty(nitFilaNormalizado) && 
                                            companyName.Contains(nitFilaNormalizado, StringComparison.OrdinalIgnoreCase);
                            
                            if (coincidePorNombre || coincidePorNIT || nitEnNombre)
                            {
                                LogMessage($"Coincidencia encontrada: {(coincidePorNombre ? "Nombre" : coincidePorNIT ? "NIT" : "NIT en nombre")}");
                                
                                if (!string.IsNullOrWhiteSpace(correo))
                                {
                                    string[] multipleEmails = correo.Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                                    
                                    foreach (string email in multipleEmails)
                                    {
                                        string cleanEmail = email.Trim();
                                        if (IsValidEmail(cleanEmail) && !emails.Contains(cleanEmail))
                                        {
                                            emails.Add(cleanEmail);
                                            LogMessage($"Correo válido encontrado: {cleanEmail}");
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex) {
                            LogMessage($"Error al procesar fila: {ex.Message}");
                        }
                    }
                    
                    if (emails.Count > 0)
                        break;
                }
                
                emails = emails.Distinct().ToList();
                
                if (emails.Count == 0)
                    LogMessage($"⚠️ No se encontraron correos para {companyName}. Verifique que el nombre o NIT coincide.");
                else
                    LogMessage($"Total de correos encontrados para {companyName}: {emails.Count}");
            }
            catch (Exception ex)
            {
                LogMessage($"Error al buscar correos para {companyName}: {ex.Message}");
            }
            
            return emails;
        }

        private bool IsValidEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return false;
                
            if (email.Contains("@") && email.Contains("."))
            {
                if (email.StartsWith("@") || email.EndsWith("@") || 
                    email.StartsWith(".") || email.EndsWith(".") ||
                    email.Contains("..") || email.Split('@').Length != 2)
                {
                    return false;
                }
                
                try {
                    var addr = new System.Net.Mail.MailAddress(email);
                    return addr.Address == email;
                }
                catch {
                    LogMessage($"Correo con formato incorrecto: {email}");
                    return false;
                }
            }
            
            return false;
        }

        private void CrearControlParaEmpresa(string nombre, string nit, string correo, string rutaPdf, bool tieneCorreo)
        {
            StackPanel empresaPanel = new StackPanel
            {
                Margin = new Thickness(0, 8, 0, 0)
            };
            
            CheckBox checkbox = new CheckBox
            {
                Content = nombre,
                Tag = nit,
                IsChecked = tieneCorreo,
                Foreground = tieneCorreo ? Brushes.Green : Brushes.Red,
                FontWeight = FontWeights.Bold
            };
            
            TextBlock infoText = new TextBlock
            {
                Text = $"NIT: {nit} | Correo: {correo}",
                Margin = new Thickness(24, 2, 0, 0),
                FontSize = 10,
                Foreground = tieneCorreo ? Brushes.Gray : Brushes.Red
            };
            
            empresaPanel.Children.Add(checkbox);
            empresaPanel.Children.Add(infoText);
            
            CheckboxContainer.Children.Add(empresaPanel);
            checkboxes.Add(checkbox);
            
            empresasArchivos[nombre] = rutaPdf;
        }

        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var checkbox in checkboxes)
            {
                checkbox.IsChecked = true;
            }
            
            LogMessage($"Se seleccionaron todas las empresas ({checkboxes.Count})");
        }

        private void DeselectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var checkbox in checkboxes)
            {
                checkbox.IsChecked = false;
            }
            
            LogMessage("Se deseleccionaron todas las empresas");
        }

        private async void ExecuteAction_Click(object sender, RoutedEventArgs e)
        {
            var seleccionados = checkboxes
                .Where(cb => cb.IsChecked == true && cb.Content != null)
                .Select(cb => cb.Content.ToString())
                .Where(content => !string.IsNullOrEmpty(content))
                .Cast<string>() 
                .ToList();

            if (seleccionados.Count == 0)
            {
                MessageBox.Show("No se seleccionaron empresas.", "Advertencia");
                return;
            }

            var result = MessageBox.Show(
                $"¿Está seguro de enviar correos a {seleccionados.Count} empresas seleccionadas?",
                "Confirmar envío",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
                
            if (result == MessageBoxResult.No)
            {
                return;
            }
            
            var credentials = CredentialManager.ReadSmtpCredentials();
            var emailSender = new EmailSender(
                CredentialManager.DEFAULT_SMTP_SERVER, 
                CredentialManager.DEFAULT_SMTP_PORT, 
                credentials.Username, 
                credentials.Password, 
                CredentialManager.DEFAULT_SMTP_SSL, 
                LogMessage);
            
            Dispatcher.Invoke(() => {
                txtEnvioLog.Clear();
                LogMessage("Iniciando proceso de envío de correos...");
                
                ForzarScrollToEnd(txtEnvioLog);
            });
            
            foreach (var empresa in seleccionados)
            {
                if (empresasArchivos.TryGetValue(empresa, out var archivoRuta))
                {
                    var correos = GetEmailsForCompany(empresa);
                    if (correos.Count > 0)
                    {
                        await emailSender.SendEmailWithAttachmentAsync(
                            correos, 
                            emailSubject, 
                            "Para SIMICS GROUP S.A.S. es muy importante contar con clientes como usted y mantenerlo informado sobre la situación actual de su cartera.\n\nAdjuntamos el estado de cuenta correspondiente; si tiene alguna observación, le agradecemos que nos la comunique por este medio para su pronta revisión.",
                            archivoRuta,
                            bccRecipient: bccEmailAddress);
                        
                        LogMessage($"Correo enviado a {empresa} ({string.Join(", ", correos)})");
                    }
                    else
                    {
                        LogMessage($"⚠️ No se encontraron correos para {empresa}. Archivo no enviado.");
                    }
                }
                else
                {
                    LogMessage($"⚠️ No se encontró archivo PDF para {empresa}. Correo no enviado.");
                }
            }
        }

        private void MostrarDetallesArchivos_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, string> empresaNits = new Dictionary<string, string>();
            
            foreach (var checkbox in checkboxes)
            {
                string empresa = checkbox.Content?.ToString() ?? "";
                string nit = checkbox.Tag?.ToString() ?? "";
                
                if (!string.IsNullOrEmpty(empresa) && !string.IsNullOrEmpty(nit))
                {
                    empresaNits[empresa] = nit;
                }
            }
            
            var ventanaDetalles = new DetallesPdfWindow(empresasArchivos, empresaNits)
            {
                Owner = this
            };
            
            ventanaDetalles.ShowDialog();
        }

        private void ConfigurarCredenciales()
        {
            var dialog = new CredentialDialog
            {
                Owner = this
            };
            
            if (dialog.ShowDialog() == true)
            {
                var credentials = dialog.Credentials;
                bccEmailAddress = credentials.BccEmail;
                ccFinanzasEmail = credentials.CcEmail;
                LogMessage($"Credenciales actualizadas para: {credentials.Username}");
            }
        }

        private void MenuItem_Exit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void MenuItem_ConfigCredentials_Click(object sender, RoutedEventArgs e)
        {
            ConfigurarCredenciales();
        }

        private void MenuItem_About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "SIMICS - Sistema de Gestión de Cartera\nVersión 1.0\n\nDesarrollado para SIMICS GROUP S.A.S.",
                "Acerca de",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void LogFileDetails(string filePath, string context = "")
        {
            try
            {
                if (!string.IsNullOrEmpty(context))
                {
                    LogMessage($"[{context}] Verificando archivo: {filePath}");
                }
                
                if (!File.Exists(filePath))
                {
                    LogMessage($"[{context}] El archivo NO existe: {filePath}");
                    return;
                }
                
                FileInfo fileInfo = new FileInfo(filePath);
                LogMessage($"[{context}] Archivo encontrado: {filePath}");
                LogMessage($"[{context}] Tamaño: {fileInfo.Length:N0} bytes");
                LogMessage($"[{context}] Fecha de creación: {fileInfo.CreationTime}");
                LogMessage($"[{context}] Fecha de última modificación: {fileInfo.LastWriteTime}");
            }
            catch (Exception ex)
            {
                LogMessage($"[{context}] Error al obtener detalles del archivo {filePath}: {ex.Message}");
            }
        }

        private void LogExceptionDetails(Exception ex, string context = "")
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"[{context}] EXCEPCIÓN: {ex.GetType().FullName}: {ex.Message}");
            
            if (ex.StackTrace != null)
                sb.AppendLine($"[{context}] StackTrace: {ex.StackTrace}");
            
            Exception innerEx = ex.InnerException;
            while (innerEx != null)
            {
                sb.AppendLine($"[{context}] Inner Exception: {innerEx.GetType().FullName}: {innerEx.Message}");
                if (innerEx.StackTrace != null)
                    sb.AppendLine($"[{context}] Inner StackTrace: {innerEx.StackTrace}");
                innerEx = innerEx.InnerException;
            }
            
            LogMessage(sb.ToString());
        }
        
        #endregion
    }
}