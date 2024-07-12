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
    /// Interaction logic for AccountingWindow.xaml
    /// </summary>
    public partial class AccountingWindow : Window
    {
        int LoanID;
        TableFundingLoanFunding loanFunding;
        TableFundingFees loanFees;
        public AccountingWindow(int loanID)
        {
            try
            { 
                InitializeComponent();
               if (MyVariables.userRole.Equals("Accounting") || MyVariables.userRole.Equals( "Underwriter"))
                {

                }
                else
                {
                    AccountingSaveButton.Visibility = Visibility.Hidden;
                }

                LoanID = loanID;
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
                loanFunding = new TableFundingLoanFunding(LoanID);
            
                loanFunding.LoanID = LoanID;
                loanFees = new TableFundingFees(LoanID);
                loanFees.LoanID = LoanID;
                dtDateDepositedInEscrow.SelectedDate = loanFunding.DateDepositedInEscrow;

                dtBaileeLetterDate.SelectedDate = loanFunding.BaileeLetterDate;
                dtInvestorProceedsDate.SelectedDate = loanFunding.InvestorProceedsDate;
                dtClosingDate.SelectedDate = loanFunding.ClosingDate;
                txtInvestorProceeds.Text = FormatNumberCommas(loanFunding.InvestorProceeds.ToString());
           
                lblInterestFee.Content = FormatNumberCommas(loanFees.InterestFee.ToString());
                txtCustSvcInterestDiscount.Text = FormatNumberCommas(loanFees.CustSvcInterestDiscount.ToString());
                lblUnderwritingFee.Content = FormatNumberCommas(loanFees.UnderwritingFee.ToString());
                txtCustSvcUnderwritingDiscount.Text = FormatNumberCommas(loanFees.CustSvcUnderwritingDiscount.ToString());
                lblOriginationFee.Content = FormatNumberCommas(loanFees.OriginationFee.ToString());
                txtCustSvcOriginationDiscount.Text = FormatNumberCommas(loanFees.CustSvcOriginationDiscount.ToString());
                lblTotalFee.Content = FormatNumberCommas(loanFunding.TotalFees.ToString());
            }
            catch (Exception ex)
            {                
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }
        }
        private void AccountingSaveButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                
                if (!dtDateDepositedInEscrow.SelectedDate.HasValue)
                    loanFunding.DateDepositedInEscrow = null;
                else
                {                    
                    loanFunding.DateDepositedInEscrow = dtDateDepositedInEscrow.SelectedDate.Value;
                }

                if (!dtBaileeLetterDate.SelectedDate.HasValue)
                    loanFunding.BaileeLetterDate = null;
                else
                    loanFunding.BaileeLetterDate = dtBaileeLetterDate.SelectedDate.Value;

                if (!dtInvestorProceedsDate.SelectedDate.HasValue)
                    loanFunding.InvestorProceedsDate = null;
                else
                    loanFunding.InvestorProceedsDate = dtInvestorProceedsDate.SelectedDate.Value;

                if (!dtClosingDate.SelectedDate.HasValue)
                    loanFunding.ClosingDate = null;
                else
                    loanFunding.ClosingDate = dtClosingDate.SelectedDate.Value;


                bool isNumber;
                string proceeds = txtInvestorProceeds.Text;
                proceeds = proceeds.Replace(",", "");
                proceeds = proceeds.Replace("$", "");
                double proceedsDbl = 0;

                isNumber = double.TryParse(proceeds, out proceedsDbl);
                if (isNumber)
                    loanFunding.InvestorProceeds = proceedsDbl;
                else
                {
                    ErrorLabel.Content = "Investor Proceeds Format Error";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;

                }
                //
                proceeds = txtCustSvcInterestDiscount.Text;
                proceeds = proceeds.Replace(",", "");
                proceeds = proceeds.Replace("$", "");
                proceedsDbl = 0;

                isNumber = double.TryParse(proceeds, out proceedsDbl);
                if (isNumber)
                    loanFees.CustSvcInterestDiscount = proceedsDbl;
                else
                {
                    ErrorLabel.Content = "Cust Service Interest Discount Format Error";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;

                }

                //
                proceeds = txtCustSvcUnderwritingDiscount.Text;
                proceeds = proceeds.Replace(",", "");
                proceeds = proceeds.Replace("$", "");
                proceedsDbl = 0;

                isNumber = double.TryParse(proceeds, out proceedsDbl);
                if (isNumber)
                    loanFees.CustSvcUnderwritingDiscount = proceedsDbl;
                else
                {
                    ErrorLabel.Content = "Cust Service Underwriting Discount Format Error";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;

                }

                //
                proceeds = txtCustSvcOriginationDiscount.Text;
                proceeds = proceeds.Replace(",", "");
                proceeds = proceeds.Replace("$", "");
                proceedsDbl = 0;

                isNumber = double.TryParse(proceeds, out proceedsDbl);
                if (isNumber)
                    loanFees.CustSvcOriginationDiscount = proceedsDbl;
                else
                {
                    ErrorLabel.Content = "Cust Service Origination Discount Format Error";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;

                }


                TableFundingLoan loan = new TableFundingLoan(LoanID);
                loan.LoanUpdateUserID = MyVariables.user.UserID;
                loan.Save();

                loanFunding.Save();
                
                loanFees.Save();

                //Reload because of fee calculation
                loanFees = new TableFundingFees(LoanID);
                string escrowDate;
                string invProceedsDate;

                if (loanFunding.DateDepositedInEscrow == null)
                    escrowDate = " ";
                else
                    escrowDate = loanFunding.DateDepositedInEscrow.ToString();


                if (loanFunding.InvestorProceedsDate == null)
                    invProceedsDate = " ";
                else
                    invProceedsDate = loanFunding.InvestorProceedsDate.ToString();

                LogWriter.WriteLog(LoanID, "Loan Accounting Save", "Escrow Date - " + escrowDate + " - Proceeds Date - " + invProceedsDate + " - Interest - " + loanFees.InterestFee.ToString());

                LoadData();
                MyVariables.needsRefresh = true;

                ErrorLabel.Content = "Saved Successfully";
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Green);
                
                Mouse.OverrideCursor = null;
            }
            catch (Exception ex)
            {

                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                Mouse.OverrideCursor = null;

            }

        }

        public string FormatNumberCommas(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("$#,##0.00");
        }
        public string FormatPcnt(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("0.0%");
        }
        public static bool IsDateTime(string txtDate)
        {
            DateTime tempDate;
            return DateTime.TryParse(txtDate, out tempDate);
        }

    }
}
