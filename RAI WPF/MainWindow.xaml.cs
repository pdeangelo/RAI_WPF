using System;
using System.Configuration;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Data.SqlClient;
using System.Collections.ObjectModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Spire.Doc;
using Microsoft.Win32;

namespace RAI_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
       

        public MainWindow()
        {
          
            InitializeComponent();
            

        }

        private void BtnGo_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            { 
                string userName = txtUser.Text;
                string password = txtPassword.Password;

                AppUser user = new AppUser(userName);

                MyVariables.user = user;
                MyVariables.userRole = user.RoleName;
                MyVariables.userIsAdmin = user.IsADmin;

                if (user.WinUserID.ToLower() != userName.ToLower())
                {
                    ErrorLabel.Content = "User Name Does Not Exist";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                }
                else if (user.Password == password)
                {
                    Loans loanWindow = new Loans();
                    loanWindow.Show();
                    this.Close();
                }
                else
                {
                    ErrorLabel.Content = "Password Does Not Match";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                }
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
            }
        }

    }
}
