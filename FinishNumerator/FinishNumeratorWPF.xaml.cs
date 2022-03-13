using System.Windows;
using System.Windows.Input;

namespace FinishNumerator
{
    /// <summary>
    /// Логика взаимодействия для FinishNumeratorWPF.xaml
    /// </summary>
    public partial class FinishNumeratorWPF : Window
    {
        public FinishNumeratorWPF()
        {
            InitializeComponent();
        }

        private void btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        private void FinishNumeratorWPF_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Space)
            {
                DialogResult = true;
                Close();
            }

            else if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }
    }
}
