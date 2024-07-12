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
    public partial class UserWindow : Window
    {
        DataSet Users;
        DataSet roles;
        public UserWindow()
        {
            InitializeComponent();
            SetUsersDataSet();
            SetUserComboBox();
            SetRolesDataSet();
            SetRolesComboBox();
            cmbUserID.Visibility = System.Windows.Visibility.Visible;
            txtUserID.Visibility = System.Windows.Visibility.Hidden;
            //if (!MyVariables.userIsAdmin)
            //{

            //    //cmbUserName.Visibility = System.Windows.Visibility.Hidden;
            //    //txtUserName.Visibility = System.Windows.Visibility.Visible;
            //    txtUserName.Text = MyVariables.user.UserName;
            //    txtUserName.IsReadOnly = true;
            //    txtPassword.Password = MyVariables.user.Password;
            //    txtPassword2.Password = MyVariables.user.Password;

            //    cmbUserName.SelectedValue = MyVariables.user.UserID;
            //    cmbUserRole.SelectedValue = MyVariables.user.RoleID;
            //    cmbUserRole.IsHitTestVisible = false;


            //    chkAdministrator.IsChecked = MyVariables.user.IsADmin;
            //    chkAdministrator.IsEnabled = false;
            //}
        }
        public void SetUsersDataSet()
        {
            ErrorLabel.Content = "";
            try
            {
                Users = RunsStoredProc.RunStoredProc("Users_SEL", "", "", "", "", "", "", "", "", "", "", "", "", "", 0);
                DataRow usersrow = Users.Tables[0].NewRow();
                usersrow["UserID"] = "-1";
                usersrow["WinUserID"] = " --Select--";

                Users.Tables[0].Rows.Add(usersrow);
                Users.Tables[0].DefaultView.Sort = "UserName";
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;

            }
        }
        public void SetRolesDataSet()
        {
            ErrorLabel.Content = "";
            try
            {

                roles = RunsStoredProc.RunStoredProc("MiscValue_SEL", "@MiscTypeID", "1", "", "", "", "", "", "", "", "", "", "", "", 0);
                DataRow usersrow = roles.Tables[0].NewRow();
                usersrow["MiscTypeValueID"] = "-1";
                usersrow["Value"] = " --Select--";

                roles.Tables[0].Rows.Add(usersrow);
                roles.Tables[0].DefaultView.Sort = "Value";
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;

            }
        }
        public void SetUserComboBox()
        {
            ErrorLabel.Content = "";
            try
            {
                cmbUserID.ItemsSource = Users.Tables[0].DefaultView;
                cmbUserID.DisplayMemberPath = Users.Tables[0].Columns["WinUserID"].ToString();
                cmbUserID.SelectedValuePath = Users.Tables[0].Columns["UserID"].ToString();

            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;

            }

        }
        public void SetRolesComboBox()
        {
            ErrorLabel.Content = "";
            try
            {
                cmbUserRole.ItemsSource = roles.Tables[0].DefaultView;
                cmbUserRole.DisplayMemberPath = roles.Tables[0].Columns["Value"].ToString();
                cmbUserRole.SelectedValuePath = roles.Tables[0].Columns["MiscTypeValueID"].ToString();

            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;

            }

        }
        private void CmbUserID_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            { 
                int userID = Convert.ToInt32(cmbUserID.SelectedValue);
                AppUser user = new AppUser(Convert.ToInt32(cmbUserID.SelectedValue));
            
                txtUserName.Text = user.UserName;
                txtPassword.Password = user.Password;
                cmbUserRole.SelectedValue = user.RoleID;
                chkAdministrator.IsChecked = user.IsADmin;
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
            }
        }


        private void UsersSaveButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
                int userid;

                if (cmbUserID.Visibility == System.Windows.Visibility.Hidden)
                    userid = -1;
                else
                    userid = Convert.ToInt32(cmbUserID.SelectedValue);

                AppUser user = new AppUser();

                user.UserID = userid;
                if (cmbUserID.Visibility == System.Windows.Visibility.Hidden)
                    user.UserName = txtUserName.Text;

                if (chkAdministrator.IsChecked == true)
                    user.IsADmin = true;
                else
                {
                    user.IsADmin = false;
                }

                user.RoleID = Convert.ToInt32(cmbUserRole.SelectedValue);
                user.UserName = txtUserName.Text;
                user.WinUserID = txtUserID.Text;
                user.Password = txtPassword.Password;
                user.Email = txtUserName.Text;
                user.OfficeID = 14;
                user.Save();

                //SetUsersDataSet();
                //SetUserComboBox();

                ErrorLabel.Content = "Saved Successfully";
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Green);
            }
            catch (Exception ex)
            {

                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }

        private void UsersNewButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
                cmbUserID.Visibility = System.Windows.Visibility.Hidden;
                txtUserID.Visibility = System.Windows.Visibility.Visible;

                txtPassword.Visibility = System.Windows.Visibility.Visible;
                lblPassword.Visibility = System.Windows.Visibility.Visible;
                cmbUserRole.SelectedIndex = 0;
                txtUserName.Text = "";
                txtPassword.Password = "";


            }
            catch (Exception ex)
            {

                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }
      
    }
}
