using System.Windows;

namespace MiProyectoWPF
{
    public partial class InputDialog : Window
    {
        public string InputText { get; private set; }
        
        public InputDialog(string message, string defaultText = "")
        {
            InitializeComponent();
            MessageText.Text = message;
            InputTextBox.Text = defaultText;
            InputText = defaultText;
        }
        
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            InputText = InputTextBox.Text;
            DialogResult = true;
            Close();
        }
        
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        
        public static string Show(Window owner, string message, string defaultText = "")
        {
            var dialog = new InputDialog(message, defaultText)
            {
                Owner = owner
            };
            
            return dialog.ShowDialog() == true ? dialog.InputText : string.Empty;
        }
    }
}
