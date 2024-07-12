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
    /// Interaction logic for SubWindow.xaml
    /// </summary>
    public partial class ClientWindow : Window
    {
        DataSet clients;
        public ClientWindow()
        {
            InitializeComponent();
            SetClientDataSet();
            SetClientComboBoxForClientPopup();

        }
       
       private void ClientSaveButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
               
                int clientID;

                if (cmbClientC.Visibility == System.Windows.Visibility.Hidden)
                    clientID = -1;
                else
                    clientID = Convert.ToInt32(cmbClientC.SelectedValue);

                TableFundingClient client = new TableFundingClient(clientID);

                if (cmbClientC.Visibility == System.Windows.Visibility.Hidden)
                    client.ClientName = txtClientC.Text;

                bool isNumber;
                string numberStr = txtAdvanceRateClient.Text;
                numberStr = numberStr.Replace("%", "");
                double numberDbl = 0;

                isNumber = double.TryParse(numberStr, out numberDbl);
                if (isNumber)
                    client.AdvanceRate = numberDbl/100;
                //
                numberStr = txtMinimumInterest.Text;
                numberStr = numberStr.Replace("%", "");
                numberDbl = 0;

                isNumber = double.TryParse(numberStr, out numberDbl);
                if (isNumber)
                    client.MinimumInterest = numberDbl/100;
                else
                    client.MinimumInterest = 0;
                //
                //
                numberStr = txtPrimeRate.Text;
                numberStr = numberStr.Replace("%", "");
                numberDbl = 0;

                isNumber = double.TryParse(numberStr, out numberDbl);
                if (isNumber)
                    client.ClientPrimeRate = numberDbl / 100;
                else
                    client.ClientPrimeRate = 0;
                //
                //
                numberStr = txtPrimeRateSpread.Text;
                numberStr = numberStr.Replace("%", "");
                numberDbl = 0;

                isNumber = double.TryParse(numberStr, out numberDbl);
                if (isNumber)
                    client.ClientPrimeRateSpread = numberDbl / 100;
                else
                    client.ClientPrimeRateSpread = 0;
                //
                if (client.MinimumInterest > 0 && (client.ClientPrimeRateSpread > 0 || client.ClientPrimeRate > 0))
                {

                    ErrorLabel.Content = "Please fill in Minimum Interest OR Prime Rate Spread";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }
                numberStr = txtOriginationDiscount.Text;
                numberStr = numberStr.Replace("%", "");
                numberDbl = 0;

                isNumber = double.TryParse(numberStr, out numberDbl);
                if (isNumber)
                    client.OriginationDiscount = numberDbl/100;
                //
                numberStr = txtOriginationDiscount2.Text;
                numberStr = numberStr.Replace("%", "");
                numberDbl = 0;

                isNumber = double.TryParse(numberStr, out numberDbl);
                if (isNumber)
                    client.OriginationDiscount2 = numberDbl / 100;
                
                //
                numberStr = txtOriginationDiscountNumDays.Text;
                numberStr = numberStr.Replace(",", "");
                int numberInt = 0;

                isNumber = int.TryParse(numberStr, out numberInt);
                if (isNumber)
                    client.OriginationDiscountNumDays = numberInt;
                //
                numberStr = txtOriginationDiscountNumDays2.Text;
                numberStr = numberStr.Replace(",", "");
                numberInt = 0;

                isNumber = int.TryParse(numberStr, out numberInt);
                if (isNumber)
                    client.OriginationDiscountNumDays2 = numberInt;
                //


                if (chkInterestBasedOnAdvance.IsChecked == true)
                    client.InterestBasedOnAdvance = true;
                else
                    client.InterestBasedOnAdvance = false;

                if (chkOriginationBasedOnAdvance.IsChecked == true)
                    client.OriginationBasedOnAdvance = true;
                else
                    client.OriginationBasedOnAdvance = false;

                if (chkNoInterest.IsChecked == true)
                    client.NoInterest = true;
                else
                    client.NoInterest = false;

                client.Save();

                LoadData();
                ErrorLabel.Content = "Client Saved Successfully";
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Green);
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }
       private void CmbClientC_SelectionChanged(object sender, SelectionChangedEventArgs e)
       {
            ErrorLabel.Content = "";
            try
            { 
                LoadData();
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }
        }
        private void LoadData()
        {
            ErrorLabel.Content = "";
            try
            { 
                TableFundingClient client = new TableFundingClient(Convert.ToInt32(cmbClientC.SelectedValue));

                txtAdvanceRateClient.Text = FormatPcnt(client.AdvanceRate.ToString());
                txtMinimumInterest.Text = FormatPcnt(client.MinimumInterest.ToString());
                txtPrimeRate.Text = FormatPcnt(client.ClientPrimeRate.ToString());
                txtPrimeRateSpread.Text = FormatPcnt(client.ClientPrimeRateSpread.ToString());
                txtOriginationDiscount.Text = FormatPcnt(client.OriginationDiscount.ToString());
                txtOriginationDiscountNumDays.Text = FormatNumberCommasNoDec(client.OriginationDiscountNumDays.ToString());
                txtOriginationDiscount2.Text = FormatPcnt(client.OriginationDiscount2.ToString());
                txtOriginationDiscountNumDays2.Text = FormatNumberCommasNoDec(client.OriginationDiscountNumDays2.ToString());

                if (client.InterestBasedOnAdvance)
                    chkInterestBasedOnAdvance.IsChecked = true;

                if (client.OriginationBasedOnAdvance)
                    chkOriginationBasedOnAdvance.IsChecked = true;

                if (client.NoInterest)
                    chkNoInterest.IsChecked = true;

            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
            }

        }
        private void ClientNewButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            {
                cmbClientC.Visibility = System.Windows.Visibility.Hidden;
                txtClientC.Visibility = System.Windows.Visibility.Visible;

                txtAdvanceRateClient.Text = "";
                txtMinimumInterest.Text = "";

                txtPrimeRate.Text = "";
                txtPrimeRateSpread.Text = "";
                txtOriginationDiscount.Text = "";
                txtOriginationDiscountNumDays.Text = "";
                txtOriginationDiscount2.Text = "";
                txtOriginationDiscountNumDays2.Text = "";

                chkInterestBasedOnAdvance.IsChecked = false;
                chkNoInterest.IsChecked = false;

            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }
        public void SetClientDataSet()
        {
            ErrorLabel.Content = "";
            try
            {
                clients = RunsStoredProc.RunStoredProc("TableFundingClient_SEL", "", "", "", "", "", "", "", "", "", "", "", "", "", 0);
                DataRow clientsrow = clients.Tables[0].NewRow();
                clientsrow["ClientID"] = "-1";
                clientsrow["ClientName"] = " --Select--";

                clients.Tables[0].Rows.Add(clientsrow);
                clients.Tables[0].DefaultView.Sort = "ClientName";
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }
        }
        public void SetClientComboBoxForClientPopup()
        {
            ErrorLabel.Content = "";
            try
            {
                cmbClientC.ItemsSource = clients.Tables[0].DefaultView;
                cmbClientC.DisplayMemberPath = clients.Tables[0].Columns["ClientName"].ToString();
                cmbClientC.SelectedValuePath = clients.Tables[0].Columns["ClientID"].ToString();

            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }
        }
        public string FormatNumberCommas(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("#,##0.00");
        }
        public string FormatNumberCommasNoDec(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("#,##0");
        }
        public string FormatPcnt(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("0.00%");
        }
    }
}
