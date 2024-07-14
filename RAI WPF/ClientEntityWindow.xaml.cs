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
using log4net;

namespace RAI_WPF
{
    /// <summary>
    /// Interaction logic for ClientEntityWindow.xaml
    /// </summary>
    public partial class ClientEntityWindow : Window
    {
        DataSet Clients;
        DataSet Entities;
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public ClientEntityWindow(DataSet clients, DataSet entities)
        {
            InitializeComponent();
            Clients = clients;
            Entities = entities;
            SetClientComboBoxForClientPopup();
            SetEntityComboBoxForClientPopup();

            //cmbEntity.Visibility = System.Windows.Visibility.Visible;
            //txtEntity.Visibility = System.Windows.Visibility.Hidden;
        }
        public void SetClientComboBoxForClientPopup()
        {
            try
            {
                cmbClientC.ItemsSource = Clients.Tables[0].DefaultView;
                cmbClientC.DisplayMemberPath = Clients.Tables[0].Columns["ClientName"].ToString();
                cmbClientC.SelectedValuePath = Clients.Tables[0].Columns["ClientID"].ToString();
               

            }
            catch (Exception ex)
            {
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }
        }
        public void SetEntityComboBoxForClientPopup()
        {
            try
            {
                cmbEntity.ItemsSource = Entities.Tables[0].DefaultView;
                cmbEntity.DisplayMemberPath = Entities.Tables[0].Columns["EntityName"].ToString();
                cmbEntity.SelectedValuePath = Entities.Tables[0].Columns["EntityID"].ToString();

            }
            catch (Exception ex)
            {
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }
        }
        private void CmbEntity_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ErrorLabel.Content = "";
            TableFundingClientEntity clientEntity = new TableFundingClientEntity(Convert.ToInt32(cmbClientC.SelectedValue), Convert.ToInt32(cmbEntity.SelectedValue));

            txtClientNumber.Text = clientEntity.ClientNumber;
            txtBank.Text = clientEntity.ClientEntityBank;
            txtABA.Text = clientEntity.Aba;
            txtAccount.Text = clientEntity.Account;
            
        }

        private void ClientEntityNewButton_Click(object sender, RoutedEventArgs e)
        {

            //ErrorLabel.Content = "";
            try
            {
                //cmbEntity.Visibility = System.Windows.Visibility.Hidden;
                //txtEntity.Visibility = System.Windows.Visibility.Visible;

                txtClientNumber.Text = "";
                txtBank.Text = "";
                txtAccount.Text = "";
                txtABA.Text = "";
                

            }
            catch (Exception ex)
            {
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }
        private void ClientEntitySaveButton_Click(object sender, RoutedEventArgs e)
        {

            //ErrorLabel.Content = "";
            try
            {
                //int entityID;

                //if (cmbEntity.Visibility == System.Windows.Visibility.Hidden)
                //    entityID = -1;
                //else
                //    entityID = Convert.ToInt32(cmbEntity.SelectedValue);

                TableFundingClientEntity clientEntity = new TableFundingClientEntity(Convert.ToInt32(cmbClientC.SelectedValue), Convert.ToInt32(cmbEntity.SelectedValue));
                clientEntity.ClientNumber = txtClientNumber.Text;
                clientEntity.ClientID = Convert.ToInt32(cmbClientC.SelectedValue);
                clientEntity.EntityID = Convert.ToInt32(cmbEntity.SelectedValue);
                clientEntity.ClientEntityBank = txtBank.Text;
                clientEntity.Aba = txtABA.Text;
                clientEntity.Account = txtAccount.Text;

                //if (cmbEntity.Visibility == System.Windows.Visibility.Hidden)
                //    clientEntity.EntityName = txtEntity.Text;

                ErrorLabel.Content = "Saved Successfully";
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Green);

                clientEntity.Save();

                
            }
            catch (Exception ex)
            {
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }
    }
}
