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
using System.Data;
using System.Data.SqlClient;
using System.Collections.ObjectModel;

namespace RAI_WPF
{

    /// <summary>
    /// Interaction logic for UserWindow.xaml
    /// </summary>
    public partial class ChangePasswordWindow : Window
    {
        DataSet Users;
        DataSet roles;
        public ChangePasswordWindow()
        {
            InitializeComponent();
            
        }
       
        private void PasswordSaveButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
                if (txtCurrentPassword.Password != MyVariables.user.Password)
                {
                    ErrorLabel.Content = "Current Password Is Incorrect";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }
                if (txtNewPassword.Password != txtNewPassword2.Password)
                {
                    ErrorLabel.Content = "New Passwords Don't Match";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }

                MyVariables.user.Password = txtNewPassword.Password;
                MyVariables.user.Save();
                ErrorLabel.Content = "Saved Successfully";
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Green);
                

            }
            catch (Exception ex)
            {

                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }

      
    }
}
