using System;
using System.Security.Cryptography;
using System.Text;
using System.IO;
using System.Security;
using System.Runtime.InteropServices;

namespace MiProyectoWPF
{
    public static class CredentialManager
    {
        private const string SMTP_USER_KEY = "SIMICS_SMTP_USER";
        private const string SMTP_PASSWORD_KEY = "SIMICS_SMTP_PASSWORD";
        private const string BCC_EMAIL_KEY = "SIMICS_BCC_EMAIL";
        private const string CC_EMAIL_KEY = "SIMICS_CC_EMAIL";
        private const string ENTROPY_KEY = "SIMICS_ENTROPY";
        
        // Valores predeterminados para el servidor SMTP (no son secretos)
        public const string DEFAULT_SMTP_SERVER = "smtppro.zoho.com";
        public const int DEFAULT_SMTP_PORT = 587;
        public const bool DEFAULT_SMTP_SSL = true;

        // Estructura para almacenar las credenciales
        public class SmtpCredentials
        {
            public string Username { get; set; } = string.Empty;
            public string Password { get; set; } = string.Empty;
            public string BccEmail { get; set; } = string.Empty;
            public string CcEmail { get; set; } = string.Empty;
            
            public bool IsComplete => !string.IsNullOrEmpty(Username) && !string.IsNullOrEmpty(Password);
        }
        
        // Leer las credenciales SMTP de las variables de entorno
        public static SmtpCredentials ReadSmtpCredentials()
        {
            var creds = new SmtpCredentials();
            
            // Intentar leer las credenciales cifradas
            try
            {
                string? entropyValue = Environment.GetEnvironmentVariable(ENTROPY_KEY, EnvironmentVariableTarget.User);
                
                if (!string.IsNullOrEmpty(entropyValue))
                {
                    byte[] entropy = Convert.FromBase64String(entropyValue);
                    
                    // Leer el nombre de usuario cifrado
                    string? encryptedUser = Environment.GetEnvironmentVariable(SMTP_USER_KEY, EnvironmentVariableTarget.User);
                    if (!string.IsNullOrEmpty(encryptedUser))
                    {
                        byte[] encryptedUserData = Convert.FromBase64String(encryptedUser);
                        creds.Username = Unprotect(encryptedUserData, entropy);
                    }
                    
                    // Leer la contraseña cifrada
                    string? encryptedPassword = Environment.GetEnvironmentVariable(SMTP_PASSWORD_KEY, EnvironmentVariableTarget.User);
                    if (!string.IsNullOrEmpty(encryptedPassword))
                    {
                        byte[] encryptedPasswordData = Convert.FromBase64String(encryptedPassword);
                        creds.Password = Unprotect(encryptedPasswordData, entropy);
                    }
                    
                    // Leer el correo BCC cifrado
                    string? encryptedBcc = Environment.GetEnvironmentVariable(BCC_EMAIL_KEY, EnvironmentVariableTarget.User);
                    if (!string.IsNullOrEmpty(encryptedBcc))
                    {
                        byte[] encryptedBccData = Convert.FromBase64String(encryptedBcc);
                        creds.BccEmail = Unprotect(encryptedBccData, entropy);
                    }
                    
                    // Leer el correo CC cifrado
                    string? encryptedCc = Environment.GetEnvironmentVariable(CC_EMAIL_KEY, EnvironmentVariableTarget.User);
                    if (!string.IsNullOrEmpty(encryptedCc))
                    {
                        byte[] encryptedCcData = Convert.FromBase64String(encryptedCc);
                        creds.CcEmail = Unprotect(encryptedCcData, entropy);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al leer credenciales: {ex.Message}");
                // Si hay un error, devolvemos credenciales vacías
                return new SmtpCredentials();
            }
            
            return creds;
        }
        
        // Guardar las credenciales SMTP en las variables de entorno
        public static bool SaveSmtpCredentials(SmtpCredentials creds)
        {
            try
            {
                // Generar un valor de entropía aleatorio (o reutilizar el existente)
                byte[] entropy;
                string? existingEntropy = Environment.GetEnvironmentVariable(ENTROPY_KEY, EnvironmentVariableTarget.User);
                
                if (string.IsNullOrEmpty(existingEntropy))
                {
                    // Generar un nuevo valor de entropía
                    entropy = new byte[16];
                    using (var rng = RandomNumberGenerator.Create())
                    {
                        rng.GetBytes(entropy);
                    }
                    
                    // Guardar la entropía
                    Environment.SetEnvironmentVariable(
                        ENTROPY_KEY, 
                        Convert.ToBase64String(entropy), 
                        EnvironmentVariableTarget.User);
                }
                else
                {
                    // Usar la entropía existente
                    entropy = Convert.FromBase64String(existingEntropy);
                }
                
                // Cifrar y guardar el nombre de usuario
                if (!string.IsNullOrEmpty(creds.Username))
                {
                    byte[] encryptedUser = Protect(creds.Username, entropy);
                    Environment.SetEnvironmentVariable(
                        SMTP_USER_KEY, 
                        Convert.ToBase64String(encryptedUser), 
                        EnvironmentVariableTarget.User);
                }
                
                // Cifrar y guardar la contraseña
                if (!string.IsNullOrEmpty(creds.Password))
                {
                    byte[] encryptedPassword = Protect(creds.Password, entropy);
                    Environment.SetEnvironmentVariable(
                        SMTP_PASSWORD_KEY, 
                        Convert.ToBase64String(encryptedPassword), 
                        EnvironmentVariableTarget.User);
                }
                
                // Cifrar y guardar el correo BCC
                if (!string.IsNullOrEmpty(creds.BccEmail))
                {
                    byte[] encryptedBcc = Protect(creds.BccEmail, entropy);
                    Environment.SetEnvironmentVariable(
                        BCC_EMAIL_KEY, 
                        Convert.ToBase64String(encryptedBcc), 
                        EnvironmentVariableTarget.User);
                }
                
                // Cifrar y guardar el correo CC
                if (!string.IsNullOrEmpty(creds.CcEmail))
                {
                    byte[] encryptedCc = Protect(creds.CcEmail, entropy);
                    Environment.SetEnvironmentVariable(
                        CC_EMAIL_KEY, 
                        Convert.ToBase64String(encryptedCc), 
                        EnvironmentVariableTarget.User);
                }
                
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al guardar credenciales: {ex.Message}");
                return false;
            }
        }
        
        // Validar las credenciales SMTP
        public static async Task<bool> ValidateSmtpCredentials(SmtpCredentials creds)
        {
            try
            {
                var emailSender = new EmailSender(
                    DEFAULT_SMTP_SERVER, 
                    DEFAULT_SMTP_PORT, 
                    creds.Username, 
                    creds.Password, 
                    DEFAULT_SMTP_SSL,
                    msg => Console.WriteLine(msg));
                
                // Enviar un correo de prueba al propio usuario
                bool result = await emailSender.SendTestEmailAsync(creds.Username);
                return result;
            }
            catch
            {
                return false;
            }
        }
        
        // Borrar las credenciales almacenadas
        public static void ClearCredentials()
        {
            Environment.SetEnvironmentVariable(SMTP_USER_KEY, null, EnvironmentVariableTarget.User);
            Environment.SetEnvironmentVariable(SMTP_PASSWORD_KEY, null, EnvironmentVariableTarget.User);
            Environment.SetEnvironmentVariable(BCC_EMAIL_KEY, null, EnvironmentVariableTarget.User);
            Environment.SetEnvironmentVariable(CC_EMAIL_KEY, null, EnvironmentVariableTarget.User);
            Environment.SetEnvironmentVariable(ENTROPY_KEY, null, EnvironmentVariableTarget.User);
        }
        
        // Método auxiliar para cifrar datos
        private static byte[] Protect(string data, byte[] entropy)
        {
            if (string.IsNullOrEmpty(data))
                return new byte[0];
                
            byte[] dataBytes = Encoding.UTF8.GetBytes(data);
            return ProtectedData.Protect(dataBytes, entropy, DataProtectionScope.CurrentUser);
        }
        
        // Método auxiliar para descifrar datos
        private static string Unprotect(byte[] data, byte[] entropy)
        {
            if (data == null || data.Length == 0)
                return string.Empty;
                
            byte[] decryptedData = ProtectedData.Unprotect(data, entropy, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(decryptedData);
        }
    }
}
