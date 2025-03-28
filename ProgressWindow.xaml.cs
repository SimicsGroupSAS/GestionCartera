using System.Windows;

namespace MiProyectoWPF
{
    public partial class ProgressWindow : Window
    {
        public ProgressWindow(string message = "Procesando...")
        {
            InitializeComponent();
            txtMessage.Text = message;
        }
    }
}
