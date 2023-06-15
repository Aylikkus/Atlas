using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Atlas
{
    /// <summary>
    /// Логика взаимодействия для LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        private Options options = Options.Instance;
        public LoginWindow()
        {
            InitializeComponent();
        }

        private void field_GotFocus(object sender, RoutedEventArgs e)
        {
            (sender as TextBox).Text = string.Empty;
        }

        private void confirmBtn_Click(object sender, RoutedEventArgs e)
        {
            options.Name = nameField.Text;
            options.Surname = surnameField.Text;
            options.Patronymic = patronymicField.Text;
            Options.Save();
            Close();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            if (options.Name == null || options.Surname == null)
                Application.Current.Shutdown();
        }
    }
}
