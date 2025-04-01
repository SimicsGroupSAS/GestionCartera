using System.Configuration;
using System.Data;
using System.Windows;
using System;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO;

namespace MiProyectoWPF
{
    public partial class App : System.Windows.Application
    {
        private static string logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "error_log.txt");

        public App()
        {
            // Capturar todas las excepciones posibles
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException!; // Usar ! para suprimir advertencia
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            // Escribir información de inicio al log
            try
            {
                File.AppendAllText(logFile, $"[{DateTime.Now}] Gestión de Cartera iniciada\n");
            }
            catch { /* Ignorar errores de log */ }
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            LogError("DispatcherUnhandledException", e.Exception);
            e.Handled = true; // Prevenir que la app se cierre
            ShowErrorMessage(e.Exception);
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                LogError("UnhandledException", ex);
                ShowErrorMessage(ex);
            }
        }

        private void TaskScheduler_UnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
        {
            LogError("TaskSchedulerException", e.Exception);
            e.SetObserved(); // Prevenir que la app se cierre
            ShowErrorMessage(e.Exception);
        }

        private void ShowErrorMessage(Exception ex)
        {
            MessageBox.Show(
                $"Se ha producido una excepción no controlada:\n\n{ex.Message}\n\nDetalles del error han sido guardados en el archivo de registro.",
                "Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        private void LogError(string type, Exception ex)
        {
            try
            {
                string errorMessage = $"[{DateTime.Now}] {type}: {ex.Message}\n{ex.StackTrace}\n\n";
                File.AppendAllText(logFile, errorMessage);
                Debug.WriteLine(errorMessage);
            }
            catch { /* Ignorar errores de log */ }
        }
    }
}

