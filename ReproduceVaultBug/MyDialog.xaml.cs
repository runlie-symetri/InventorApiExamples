using System.Windows;

namespace ReproduceVaultBug
{
    /// <summary>
    /// Interaction logic for MyDialog.xaml
    /// </summary>
    public partial class MyDialog : Window
    {
        public MyDialog()
        {
            InitializeComponent();
        }

        public void SetText(string text)
        {
            Text.Text = text;
        }
    }
}
