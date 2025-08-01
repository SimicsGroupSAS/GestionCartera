using System;
using System.Threading.Tasks;
using System.Windows;

namespace MiProyectoWPF
{
    public partial class CredentialDialog : Window
    {
        private bool _validationSuccessful = false;
        private const string DEFAULT_BCC_EMAIL = "pofika1666@nokdot.com"; // Valor fijo para BCC
        
        public CredentialManager.SmtpCredentials Credentials { get; private set; } = 
            new CredentialManager.SmtpCredentials();

        public CredentialDialog()
        {
            InitializeComponent();
            
            // Cargar credenciales existentes si las hay
            Credentials = CredentialManager.ReadSmtpCredentials();
            
            // Mostrar valores existentes en los campos
            txtEmail.Text = Credentials.Username;
            if (!string.IsNullOrEmpty(Credentials.Password))
                txtPassword.Password = Credentials.Password;
            
            // Mostrar el BCC existente (editable ahora)
            txtBccEmail.Text = string.IsNullOrEmpty(Credentials.BccEmail) ? "pofika1666@nokdot.com" : Credentials.BccEmail;
        }
        
        private void btnCancelar_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        
        private void btnGuardar_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtEmail.Text))
            {
                MessageBox.Show("Por favor, ingrese un correo electrónico válido.", 
                              "Correo requerido", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtEmail.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtPassword.Password))
            {
                MessageBox.Show("Por favor, ingrese una contraseña.", 
                              "Contraseña requerida", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtPassword.Focus();
                return;
            }
            
            if (string.IsNullOrWhiteSpace(txtBccEmail.Text) || !IsValidEmail(txtBccEmail.Text))
            {
                MessageBox.Show("Por favor, ingrese un correo electrónico válido para el campo BCC.", 
                              "BCC requerido", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtBccEmail.Focus();
                return;
            }
            
            // Actualizar credenciales con los valores del formulario
            Credentials.Username = txtEmail.Text.Trim();
            Credentials.Password = txtPassword.Password;
            Credentials.BccEmail = txtBccEmail.Text.Trim(); // Guardar el BCC modificado
            Credentials.CcEmail = string.Empty; // No se usa CC
            
            // Guardar credenciales si está marcada la opción
            if (chkGuardarCredenciales.IsChecked == true)
            {
                if (CredentialManager.SaveSmtpCredentials(Credentials))
                {
                    MessageBox.Show("Credenciales guardadas correctamente.", 
                                  "Éxito", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("No se pudieron guardar las credenciales. Continuar de todos modos.", 
                                  "Error al guardar", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            
            DialogResult = true;
            Close();
        }
        
        private async void btnValidar_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtEmail.Text) || string.IsNullOrWhiteSpace(txtPassword.Password))
            {
                MessageBox.Show("Debe ingresar un correo electrónico y contraseña para validar.", 
                              "Datos incompletos", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            // Deshabilitar controles durante la validación
            SetControlsEnabled(false);
            btnValidar.Content = "Validando...";
            
            var testCredentials = new CredentialManager.SmtpCredentials
            {
                Username = txtEmail.Text.Trim(),
                Password = txtPassword.Password
            };
            
            // Realizar validación
            bool isValid = await CredentialManager.ValidateSmtpCredentials(testCredentials);
            
            // Mostrar resultado y habilitar controles
            if (isValid)
            {
                MessageBox.Show("Credenciales validadas correctamente. Se ha enviado un correo de prueba.", 
                              "Validación exitosa", MessageBoxButton.OK, MessageBoxImage.Information);
                _validationSuccessful = true;
                
                // Guardar automáticamente las credenciales validadas
                Credentials.Username = txtEmail.Text.Trim();
                Credentials.Password = txtPassword.Password;
                Credentials.BccEmail = txtBccEmail.Text.Trim();
                
                // Guardar en el sistema si está marcada la opción
                if (chkGuardarCredenciales.IsChecked == true)
                {
                    if (CredentialManager.SaveSmtpCredentials(Credentials))
                    {
                        MessageBox.Show("Credenciales guardadas correctamente.", 
                                      "Éxito", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                
                // Cerrar automáticamente el diálogo al validar correctamente
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show("No se pudieron validar las credenciales. Verifique su correo y contraseña.", 
                              "Error de validación", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
            SetControlsEnabled(true);
            btnValidar.Content = "Validar";
        }
        
        private void SetControlsEnabled(bool enabled)
        {
            txtEmail.IsEnabled = enabled;
            txtPassword.IsEnabled = enabled;
            txtBccEmail.IsEnabled = enabled;
            chkGuardarCredenciales.IsEnabled = enabled;
            btnGuardar.IsEnabled = enabled;
            btnCancelar.IsEnabled = enabled;
            btnValidar.IsEnabled = enabled;
        }

        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
    }
}
