using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace MiProyectoWPF
{
    public class EmailSender
    {
        private readonly string smtpServer;
        private readonly int port;
        private readonly string username;
        private readonly string password;
        private readonly bool enableSsl;
        private readonly Action<string> logCallback;
        
        public EmailSender(
            string smtpServer, 
            int port, 
            string username, 
            string password, 
            bool enableSsl = true, 
            Action<string>? logCallback = null)
        {
            this.smtpServer = smtpServer;
            this.port = port;
            this.username = username;
            this.password = password;
            this.enableSsl = enableSsl;
            this.logCallback = logCallback ?? (message => Console.WriteLine(message));
        }
        
        // Propiedades públicas solo para lectura para diagnóstico
        public string SmtpServer => smtpServer;
        public int Port => port;
        public string Username => username;
        public bool EnableSsl => enableSsl;
        
        public async Task<bool> SendEmailWithAttachmentAsync(
            List<string> recipients, 
            string subject, 
            string body, 
            string attachmentPath, 
            string? ccRecipient = null,
            string? bccRecipient = null,
            bool isBodyHtml = true)
        {
            try
            {
                logCallback($"Enviando correo a: {string.Join(", ", recipients)}...");
                
                // Verificar que el archivo existe
                if (!File.Exists(attachmentPath))
                {
                    logCallback($"Error: El archivo adjunto no existe: {attachmentPath}");
                    return false;
                }
                
                // Información detallada sobre el tamaño del archivo
                var fileInfo = new FileInfo(attachmentPath);
                logCallback($"Información del archivo: {Path.GetFileName(attachmentPath)}, tamaño: {fileInfo.Length / 1024} KB");
                
                // Crear el mensaje con la misma estructura que SendTestEmailAsync
                using var message = new MailMessage
                {
                    From = new MailAddress(username),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = isBodyHtml
                };
                
                // Agregar destinatarios
                foreach (var recipient in recipients)
                {
                    if (!string.IsNullOrWhiteSpace(recipient))
                        message.To.Add(recipient);
                }
                
                // Mostrar cuántos destinatarios se agregaron realmente
                logCallback($"Destinatarios efectivamente agregados: {message.To.Count}");
                
                // Agregar CC y BCC 
                if (!string.IsNullOrWhiteSpace(ccRecipient))
                    message.CC.Add(ccRecipient);
                
                if (!string.IsNullOrWhiteSpace(bccRecipient))
                    message.Bcc.Add(bccRecipient);
                
                // Agregar adjunto
                var attachment = new Attachment(attachmentPath);
                message.Attachments.Add(attachment);
                
                // Configurar cliente SMTP exactamente igual que el método de prueba
                using var client = new SmtpClient(smtpServer)
                {
                    Port = port,
                    Credentials = new NetworkCredential(username, password),
                    EnableSsl = enableSsl,
                    Timeout = 120000
                };
                
                // Detallar la configuración del cliente SMTP
                logCallback($"Configurando cliente SMTP - Servidor: {smtpServer}, Puerto: {port}, SSL: {enableSsl}, Timeout: {client.Timeout}ms");
                
                try
                {
                    logCallback("Iniciando envío del mensaje...");
                    await client.SendMailAsync(message);
                    logCallback("¡Correo enviado con éxito!");
                    return true;
                }
                catch (Exception ex)
                {
                    logCallback($"Error durante el envío del mensaje: {ex.GetType().Name} - {ex.Message}");
                    if (ex.InnerException != null)
                        logCallback($"Detalles adicionales: {ex.InnerException.GetType().Name} - {ex.InnerException.Message}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                logCallback($"Error al enviar correo: {ex.Message}");
                if (ex.InnerException != null)
                    logCallback($"Detalles adicionales: {ex.InnerException.Message}");
                return false;
            }
        }
        
        public async Task<bool> SendTestEmailAsync(string recipient)
        {
            try
            {
                logCallback($"Enviando correo de prueba a {recipient}...");
                
                using var message = new MailMessage
                {
                    From = new MailAddress(username),
                    Subject = "Correo de prueba SIMICS - Estado de Cartera",
                    Body = "Este es un correo de prueba enviado desde la aplicación de estado de cartera de SIMICS GROUP S.A.S.",
                    IsBodyHtml = false
                };
                
                message.To.Add(recipient);
                
                using var client = new SmtpClient(smtpServer)
                {
                    Port = port,
                    Credentials = new NetworkCredential(username, password),
                    EnableSsl = enableSsl,
                    Timeout = 120000
                };
                
                await client.SendMailAsync(message);
                
                logCallback("Correo de prueba enviado con éxito.");
                return true;
            }
            catch (Exception ex)
            {
                logCallback($"Error al enviar correo de prueba: {ex.Message}");
                if (ex.InnerException != null)
                    logCallback($"Detalles adicionales: {ex.InnerException.Message}");
                return false;
            }
        }
        
        public async Task<bool> SendTestEmailWithBccAsync(string recipient, string? bccRecipient = null)
        {
            try
            {
                logCallback($"Enviando correo de prueba a {recipient}" + 
                    (string.IsNullOrEmpty(bccRecipient) ? "" : $" con copia oculta a {bccRecipient}") + "...");
                
                using var message = new MailMessage
                {
                    From = new MailAddress(username),
                    Subject = "Correo de prueba SIMICS - Estado de Cartera",
                    Body = "Este es un correo de prueba enviado desde la aplicación de estado de cartera de SIMICS GROUP S.A.S.",
                    IsBodyHtml = false
                };
                
                message.To.Add(recipient);
                
                // Agregar BCC si está especificado
                if (!string.IsNullOrWhiteSpace(bccRecipient))
                {
                    message.Bcc.Add(bccRecipient);
                }
                
                using var client = new SmtpClient(smtpServer)
                {
                    Port = port,
                    Credentials = new NetworkCredential(username, password),
                    EnableSsl = enableSsl,
                    Timeout = 120000
                };
                
                await client.SendMailAsync(message);
                
                logCallback("Correo de prueba enviado con éxito.");
                return true;
            }
            catch (Exception ex)
            {
                logCallback($"Error al enviar correo de prueba: {ex.Message}");
                if (ex.InnerException != null)
                    logCallback($"Detalles adicionales: {ex.InnerException.Message}");
                return false;
            }
        }
        
        public static string GetEmailHtmlTemplate(string textContent)
        {
            // Plantilla HTML para el correo con firma profesional
            string htmlTemplate = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>Estado de Cartera</title>
</head>
<body style='font-family: Arial, sans-serif; line-height: 1.6; color: #333;'>
    <div style='max-width: 600px; margin: 0 auto; padding: 20px;'>
        {textContent}
        <br/>
        <hr style='border: none; height: 1px; background-color: #ddd; margin: 20px 0;'>
        
        <table style='width: 100%; border-collapse: collapse;'>
            <tr>
                <td style='vertical-align: top; padding-right: 15px;'>
                    <img src='http://simicsgroup.com/wp-content/uploads/2023/08/Logo-v6_Icono2021Firma.png' alt='Logo' style='width: 60px;'>
                </td>
                <td>
                    <h3 style='margin: 0; font-size: 16px;'>JUAN MANUEL CUERVO PINILLA</h3>
                    <p style='margin: 0; font-weight: 500; font-size: 14px;'>Gerente Financiero</p>
                    <p style='margin: 5px 0 0; font-size: 12px;'>
                        <img src='http://simicsgroup.com/wp-content/uploads/2023/08/image002.png' style='width: 12px; vertical-align: middle;'> 
                        <a href='tel:+573163114545' style='color: #333; text-decoration: none;'>+57-3163114545</a>
                    </p>
                    <p style='margin: 3px 0; font-size: 12px;'>
                        <img src='http://simicsgroup.com/wp-content/uploads/2023/08/image003.png' style='width: 12px; vertical-align: middle;'> 
                        <a href='mailto:juan.cuervo@simicsgroup.com' style='color: #333; text-decoration: none;'>juan.cuervo@simicsgroup.com</a>
                    </p>
                    <p style='margin: 3px 0; font-size: 12px;'>
                        <img src='http://simicsgroup.com/wp-content/uploads/2023/08/image004.png' style='width: 12px; vertical-align: middle;'> 
                        CR 53 No. 96-24 Oficina 3D, Barranquilla, Colombia
                    </p>
                </td>
            </tr>
            <tr>
                <td colspan='2' style='text-align: center; padding-top: 20px;'>
                    <img src='http://simicsgroup.com/wp-content/uploads/2023/08/Logo-v6_2021-1Firma.png' style='width: 180px;'><br>
                    <a href='https://www.simicsgroup.com/' style='color: #333; text-decoration: none; font-size: 12px;'>www.simicsgroup.com</a>
                    <div style='margin-top: 10px;'>
                        <a href='https://www.linkedin.com/company/simicsgroupsas' style='margin: 0 5px;'>
                            <img src='http://simicsgroup.com/wp-content/uploads/2023/08/image006.png' alt='LinkedIn' style='width: 24px;'>
                        </a>
                        <a href='https://www.instagram.com/simicsgroupsas/' style='margin: 0 5px;'>
                            <img src='http://simicsgroup.com/wp-content/uploads/2023/08/image007.png' alt='Instagram' style='width: 24px;'>
                        </a>
                        <a href='https://www.facebook.com/SIMICSGroupSAS/' style='margin: 0 5px;'>
                            <img src='http://simicsgroup.com/wp-content/uploads/2023/08/image008.png' alt='Facebook' style='width: 24px;'>
                        </a>
                    </div>
                </td>
            </tr>
        </table>
    </div>
</body>
</html>";

            return htmlTemplate;
        }
    }
}
