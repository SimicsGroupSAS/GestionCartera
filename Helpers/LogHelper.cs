using System;
using System.IO;
using System.Text;
using System.Threading;

namespace MiProyectoWPF.Helpers
{
    /// <summary>
    /// Proporciona funciones de registro detallado para ayudar a diagnosticar problemas de lectura/escritura de archivos
    /// </summary>
    public static class LogHelper
    {
        private static readonly object _lockObject = new object();
        private static string _logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file_operations.log");
        
        /// <summary>
        /// Configura la ruta del archivo de log
        /// </summary>
        public static void SetLogFilePath(string path)
        {
            _logFilePath = path;
            
            // Crear un encabezado al inicializar el log
            try
            {
                // Si el archivo ya existe y es mayor a 5MB, crear uno nuevo
                if (File.Exists(_logFilePath))
                {
                    FileInfo fileInfo = new FileInfo(_logFilePath);
                    if (fileInfo.Length > 5 * 1024 * 1024) // 5MB
                    {
                        string backupPath = Path.Combine(
                            Path.GetDirectoryName(_logFilePath),
                            $"{Path.GetFileNameWithoutExtension(_logFilePath)}_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(_logFilePath)}");
                        
                        File.Move(_logFilePath, backupPath);
                    }
                }
                
                // Crear o añadir al archivo de log
                using (StreamWriter writer = new StreamWriter(_logFilePath, true, Encoding.UTF8))
                {
                    writer.WriteLine($"========== INICIO DE LOG {DateTime.Now:yyyy-MM-dd HH:mm:ss} ==========");
                    writer.WriteLine($"Sistema operativo: {Environment.OSVersion}");
                    writer.WriteLine($"Directorio de la aplicación: {AppDomain.CurrentDomain.BaseDirectory}");
                    writer.WriteLine();
                }
            }
            catch
            {
                // No hacer nada si falla, para evitar bloquear la aplicación
            }
        }
        
        /// <summary>
        /// Registra un mensaje en el archivo de log con hora y nivel
        /// </summary>
        public static void Log(string message, LogLevel level = LogLevel.Info)
        {
            try
            {
                // Si no está configurado, establecer la ruta predeterminada
                if (string.IsNullOrEmpty(_logFilePath))
                {
                    SetLogFilePath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file_operations.log"));
                }
                
                string threadId = Thread.CurrentThread.ManagedThreadId.ToString();
                string formattedMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] [Thread {threadId}] {message}";
                
                lock (_lockObject)
                {
                    using (StreamWriter writer = new StreamWriter(_logFilePath, true, Encoding.UTF8))
                    {
                        writer.WriteLine(formattedMessage);
                    }
                }
                
                // También mostrar en consola para depuración inmediata
                Console.WriteLine(formattedMessage);
            }
            catch
            {
                // No hacer nada si falla, para evitar bloquear la aplicación
            }
        }
        
        /// <summary>
        /// Registra información detallada sobre un archivo
        /// </summary>
        public static void LogFileDetails(string filePath, string context = "")
        {
            try
            {
                if (!string.IsNullOrEmpty(context))
                {
                    Log($"[{context}] Verificando archivo: {filePath}", LogLevel.Debug);
                }
                
                if (!File.Exists(filePath))
                {
                    Log($"[{context}] El archivo NO existe: {filePath}", LogLevel.Warning);
                    return;
                }
                
                FileInfo fileInfo = new FileInfo(filePath);
                Log($"[{context}] Archivo encontrado: {filePath}", LogLevel.Debug);
                Log($"[{context}] Tamaño: {fileInfo.Length:N0} bytes", LogLevel.Debug);
                Log($"[{context}] Fecha de creación: {fileInfo.CreationTime}", LogLevel.Debug);
                Log($"[{context}] Fecha de última modificación: {fileInfo.LastWriteTime}", LogLevel.Debug);
                Log($"[{context}] Atributos: {fileInfo.Attributes}", LogLevel.Debug);
            }
            catch (Exception ex)
            {
                Log($"[{context}] Error al obtener detalles del archivo {filePath}: {ex.Message}", LogLevel.Error);
            }
        }
        
        /// <summary>
        /// Registra excepciones detalladamente
        /// </summary>
        public static void LogException(Exception ex, string context = "")
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"[{context}] EXCEPCIÓN: {ex.GetType().FullName}: {ex.Message}");
            sb.AppendLine($"[{context}] StackTrace: {ex.StackTrace}");
            
            Exception innerEx = ex.InnerException;
            while (innerEx != null)
            {
                sb.AppendLine($"[{context}] Inner Exception: {innerEx.GetType().FullName}: {innerEx.Message}");
                sb.AppendLine($"[{context}] Inner StackTrace: {innerEx.StackTrace}");
                innerEx = innerEx.InnerException;
            }
            
            Log(sb.ToString(), LogLevel.Error);
        }
    }
    
    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }
}
