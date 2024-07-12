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
    /// Interaction logic for Investorsxaml.xaml
    /// </summary>
    public partial class InvestorWindow : Window
    {
        DataSet investors;
        public InvestorWindow()
        {
            InitializeComponent();
            SetInvestorDataSet();
            SetInvestorComboBoxForInvestorPopup();
            cmbInvestorI.Visibility = System.Windows.Visibility.Visible;
            txtInvestorI.Visibility = System.Windows.Visibility.Hidden;
        }
        private void InvestorSaveButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
                int investorID;

                if (cmbInvestorI.Visibility == System.Windows.Visibility.Hidden)
                    investorID = -1;
                else
                    investorID = Convert.ToInt32(cmbInvestorI.SelectedValue);

                TableFundingInvestor investor = new TableFundingInvestor(investorID);

                if (cmbInvestorI.Visibility == System.Windows.Visibility.Hidden)
                    investor.InvestorName = txtInvestorI.Text;

                investor.CustodianName = txtCustodian.Text;

                investor.Save();

                SetInvestorDataSet();
                SetInvestorComboBoxForInvestorPopup();

                ErrorLabel.Content = "Saved Successfully";
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Green);

            }
            catch (Exception ex)
            {

                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }

        private void InvestorNewButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
                cmbInvestorI.Visibility = System.Windows.Visibility.Hidden;
                txtInvestorI.Visibility = System.Windows.Visibility.Visible;

                txtCustodian.Text = "";


            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;

            }

        }
        public string FormatNumberCommas(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("#,##0");
        }
        public string FormatPcnt(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("0.0%");
        }
        public void SetInvestorDataSet()
        {
            ErrorLabel.Content = "";
            try
            {
                investors = RunsStoredProc.RunStoredProc("TableFundingInvestor_SEL", "", "", "", "", "", "", "", "", "", "", "", "", "", 0);
                DataRow investorRow = investors.Tables[0].NewRow();
                investorRow["InvestorID"] = "-1";
                investorRow["InvestorName"] = " --Select--";

                investors.Tables[0].Rows.Add(investorRow);
                investors.Tables[0].DefaultView.Sort = "InvestorName";
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }
        public void SetInvestorComboBoxForInvestorPopup()
        {
            ErrorLabel.Content = "";
            try
            {
                cmbInvestorI.ItemsSource = investors.Tables[0].DefaultView;
                cmbInvestorI.DisplayMemberPath = investors.Tables[0].Columns["InvestorName"].ToString();
                cmbInvestorI.SelectedValuePath = investors.Tables[0].Columns["InvestorID"].ToString();

            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }
        private void CmbInvestorI_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            { 
                TableFundingInvestor investor = new TableFundingInvestor(Convert.ToInt32(cmbInvestorI.SelectedValue));
                txtCustodian.Text = investor.CustodianName;
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }

    }
}
