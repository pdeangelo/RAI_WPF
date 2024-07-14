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
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Spire.Doc;
using Microsoft.WindowsAPICodePack.Dialogs;
using log4net;
namespace RAI_WPF
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class UnderwritingWindow : Window
    {
        int LoanID;
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public TableFundingLoan Loan;
        DataSet Clients;
        DataSet LoanType;
        DataSet LoanDwellingType;
        DataSet States;
        DataSet Entities;
        DataSet Investors;
        
        public UnderwritingWindow(TableFundingLoan loan, DataSet clients, DataSet loanType, DataSet entities, DataSet investors, DataSet loanDwellingType, DataSet states)
        {
            try
            { 
                InitializeComponent();

                if (MyVariables.userRole != "Underwriter")
                {
                    UnderwritingSaveButton.Visibility = Visibility.Hidden;
                }
                LoanID = loan.LoanID;
                Loan = loan;
                Clients = clients;
                LoanType = loanType;
                LoanDwellingType = loanDwellingType;
                States = states;
                Entities = entities;
                Investors = investors;
                LoadData();
            }
            catch (Exception ex)
            {                
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }

        }
        private void LoadData(bool resetErrorMessage = true)
        {

            if (resetErrorMessage)
                ErrorLabel.Content = "";
            try
            { 

                Loan.LoanUW = new TableFundingLoanUW(LoanID);
                Loan.LoanFunding = new TableFundingLoanFunding(LoanID);
                Loan.LoanFunding.LoanID = LoanID;
                
                dtFundingDate.SelectedDate = Loan.LoanFundingDate;

                dtDateDepositedInEscrow.SelectedDate = Loan.LoanFunding.DateDepositedInEscrow;

                cmbClient.ItemsSource = Clients.Tables[0].DefaultView;
                cmbClient.DisplayMemberPath = Clients.Tables[0].Columns["ClientName"].ToString();
                cmbClient.SelectedValuePath = Clients.Tables[0].Columns["ClientID"].ToString();

                cmbLoanType.ItemsSource = LoanType.Tables[0].DefaultView;
                cmbLoanType.DisplayMemberPath = LoanType.Tables[0].Columns["Value"].ToString();
                cmbLoanType.SelectedValuePath = LoanType.Tables[0].Columns["MiscTypeValueID"].ToString();

                cmbLoanDwellingType.ItemsSource = LoanDwellingType.Tables[0].DefaultView;
                cmbLoanDwellingType.DisplayMemberPath = LoanDwellingType.Tables[0].Columns["Value"].ToString();
                cmbLoanDwellingType.SelectedValuePath = LoanDwellingType.Tables[0].Columns["MiscTypeValueID"].ToString();

                cmbState.ItemsSource = States.Tables[0].DefaultView;
                cmbState.DisplayMemberPath = States.Tables[0].Columns["Value"].ToString();
                cmbState.SelectedValuePath = States.Tables[0].Columns["MiscTypeValueID"].ToString();

                cmbEntity.ItemsSource = Entities.Tables[0].DefaultView;
                cmbEntity.DisplayMemberPath = Entities.Tables[0].Columns["EntityName"].ToString();
                cmbEntity.SelectedValuePath = Entities.Tables[0].Columns["EntityID"].ToString();

                cmbInvestor.ItemsSource = Investors.Tables[0].DefaultView;
                cmbInvestor.DisplayMemberPath = Investors.Tables[0].Columns["InvestorName"].ToString();
                cmbInvestor.SelectedValuePath = Investors.Tables[0].Columns["InvestorID"].ToString();

                cmbClient.SelectedValue = Loan.LoanClientID;
                txtLoanNumber.Text = Loan.LoanNumber;
                txtCustomerName.Text = Loan.LoanMortgagee;
                txtBusinessName.Text = Loan.LoanMortgageeBusiness;
                cmbLoanType.SelectedValue = Loan.LoanType;
                cmbLoanDwellingType.SelectedValue = Loan.LoanDwellingType;

                cmbState.SelectedValue = Loan.State;
                cmbEntity.SelectedValue = Loan.LoanEntityID;
                cmbInvestor.SelectedValue = Loan.LoanInvestorID;
                txtPropertyAddress.Text = Loan.LoanPropertyAddress;
                txtAppraisalAmount.Text = FormatNumberCommasMoney(Loan.LoanUW.LoanUWAppraisalAmount.ToString());
                txtPostRepairAppraisalAmount.Text = FormatNumberCommasMoney(Loan.LoanUW.LoanUWPostRepairAppraisalAmount.ToString());
                txtInterestRate.Text = FormatPcnt(Loan.LoanInterestRate.ToString());
                txtMortgageAmount.Text = FormatNumberCommasMoney(Loan.LoanMortgageAmount.ToString());
                txtAdvanceRate.Text = FormatPcnt(Loan.LoanAdvanceRate.ToString());
                txtZillowSquareFootage.Text = FormatNumberCommasNoDec(Loan.LoanUW.LoanUWZillowSqFt.ToString());
                lblCompletedBy.Content = Loan.LoanUW.CompletedByName;

                if (Loan.LoanUW.LoanUWAppraisal)
                    chkAppraisal.IsChecked = true;

                if (Loan.LoanUW.LoanUWAllongePromissoryNote)
                    chkAllonge.IsChecked = true;

                if (Loan.LoanUW.LoanUW10031008LoanApplication)
                    chkLoanApplication.IsChecked = true;

                if (Loan.LoanUW.LoanUWCreditReport)
                    chkCreditReport.IsChecked = true;

                if (Loan.LoanUW.LoanUWFloodCertificate)
                    chkFloodCertificate.IsChecked = true;

                if (Loan.LoanUW.LoanUWAssignmentOfMortgage)
                    chkAssignmentOfMortgage.IsChecked = true;

                if (Loan.LoanUW.LoanUWBackgroundCheck)
                    chkBackgroundCheck.IsChecked = true;

                if (Loan.LoanUW.LoanUWCertofGoodStandingFormation)
                    chkCertificateOfGoodStanding.IsChecked = true;

                if (Loan.LoanUW.LoanUWClosingProtectionLetter)
                    chkClosingProtectionLetter.IsChecked = true;

                if (Loan.LoanUW.LoanUWCommitteeReview)
                    chkCommitteeReview.IsChecked = true;

                if (Loan.LoanUW.LoanUWHomeownersInsurance)
                    chkHomeownersInsurance.IsChecked = true;

                if (Loan.LoanUW.LoanUWLoanPackage)
                    chkLoanPackage.IsChecked = true;

                //if (Loan.LoanUW.LoanUWLoanSizerLoanSubmissionTape)
                //    chkLoanSizer.IsChecked = true;

                if (Loan.LoanUW.LoanUWTitleCommitment)
                    chkTitleCommitment.IsChecked = true;

                if (Loan.LoanUW.LoanUWClaytonReportApprovalEmail)
                    chkToorak.IsChecked = true;

                if (Loan.LoanUW.LoanUWHUD1SettlementStatement)
                    chkHud.IsChecked = true;

                if (Loan.LoanUW.LoanUWIsComplete)
                    chkUnderwritingComplete.IsChecked = true;
            }
            catch (Exception ex)
            {                
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }
        }
        private void UnderwritingSaveButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

                bool status = true;
                if (chkUnderwritingComplete.IsChecked == true)
                {
                    status = VerifyReqFieldsOnSave();

                    if (!status)
                    {
                        ErrorLabel.Content = "ERROR - Mandatory fields are blank - NOT SAVED";
                        ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                        Mouse.OverrideCursor = null;
                        return;
                    }

                    status = true;
                    status = VerifyOptFieldsOnSave();

                    if (!status)
                    {
                        ErrorLabel.Content = "WARNING - Optional fields are blank. - OTHER DATA WILL BE SAVED";
                        ErrorLabel.Foreground = new SolidColorBrush(Colors.Yellow);
                    }
                }

                if (txtLoanNumber.Text.Length > 0)
                {
                    TableFundingLoans checkLoanNumber = new TableFundingLoans(-1, -1, -1, false, txtLoanNumber.Text, false);
                    bool dupeLoan = false;

                    foreach (TableFundingLoan checkDupeLoan in checkLoanNumber)
                    {
                        if (checkDupeLoan.LoanID != Loan.LoanID)
                            dupeLoan = true;
                    }
                    if (dupeLoan)
                    {
                        ErrorLabel.Content = "Loan Number Already Exists";
                        ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                        Mouse.OverrideCursor = null;
                        return;
                    }
                }

                if (dtDateDepositedInEscrow.SelectedDate.HasValue && chkUnderwritingComplete.IsChecked == false)
                {
                    ErrorLabel.Content = "Underwriting must be complete to enter an escrow date";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;
                }

                if (!dtFundingDate.SelectedDate.HasValue)
                    Loan.LoanFundingDate = null;
                else
                    Loan.LoanFundingDate = dtFundingDate.SelectedDate.Value;

                if (!dtDateDepositedInEscrow.SelectedDate.HasValue)
                    Loan.LoanFunding.DateDepositedInEscrow = null;
                else
                    Loan.LoanFunding.DateDepositedInEscrow = dtDateDepositedInEscrow.SelectedDate.Value;

                Loan.LoanClientID = Convert.ToInt32(cmbClient.SelectedValue);
                Loan.LoanInvestorID = Convert.ToInt32(cmbInvestor.SelectedValue);
                Loan.LoanEntityID = Convert.ToInt32(cmbEntity.SelectedValue);
                Loan.LoanType = Convert.ToInt32(cmbLoanType.SelectedValue);
                Loan.LoanDwellingType = Convert.ToInt32(cmbLoanDwellingType.SelectedValue);
                Loan.State = Convert.ToInt32(cmbState.SelectedValue);

                Loan.LoanNumber = txtLoanNumber.Text;
                Loan.LoanMortgagee = txtCustomerName.Text;
                Loan.LoanMortgageeBusiness = txtBusinessName.Text;
                Loan.LoanPropertyAddress = txtPropertyAddress.Text;

                bool isNumber;
                string interestRate = txtInterestRate.Text;
                interestRate = interestRate.Replace("%", "");
                double interestRateDbl = 0;

                isNumber = double.TryParse(interestRate, out interestRateDbl);
                if (isNumber)
                    Loan.LoanInterestRate = interestRateDbl / 100;
                else
                {
                    ErrorLabel.Content = "Interest Rate Format Error";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;

                }

                string mortgageAmount = txtMortgageAmount.Text;
                mortgageAmount = mortgageAmount.Replace(",", "");
                mortgageAmount = mortgageAmount.Replace("$", "");
                double mortgageAmountDbl = 0;

                isNumber = double.TryParse(mortgageAmount, out mortgageAmountDbl);
                if (isNumber)
                    Loan.LoanMortgageAmount = mortgageAmountDbl;
                else
                {
                    ErrorLabel.Content = "Mortgage Amount Format Error";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;

                }

                string advanceRate = txtAdvanceRate.Text;
                advanceRate = advanceRate.Replace("%", "");
                double advanceRateDbl = 0;

                isNumber = double.TryParse(advanceRate, out advanceRateDbl);
                if (isNumber)
                    Loan.LoanAdvanceRate = advanceRateDbl / 100;
                else
                {
                    ErrorLabel.Content = "Advance Rate Format Error";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;

                }
                Loan.LoanUpdateUserID = MyVariables.user.UserID;
                Loan.Save();
                Loan.LoanFunding.Save();
                string escrowDate;

                if (Loan.LoanFunding.DateDepositedInEscrow == null)
                    escrowDate = " ";
                else
                    escrowDate = Loan.LoanFunding.DateDepositedInEscrow.ToString();

                String message = "Escrow Date - " + escrowDate + " - Complete Flag - " + chkUnderwritingComplete.IsChecked.ToString();
                LogWriter.WriteLog(Loan.LoanID,"Loan Underwriting Save", message);

                LoanID = Loan.LoanID;

                string appraisal = txtAppraisalAmount.Text;
                appraisal = appraisal.Replace(",", "");
                appraisal = appraisal.Replace("$", "");
                double appraisalDbl = 0;

                isNumber = double.TryParse(appraisal, out appraisalDbl);
                if (isNumber)
                    Loan.LoanUW.LoanUWAppraisalAmount = appraisalDbl;
                else
                {
                    ErrorLabel.Content = "Appraisal Amount Format Error";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;

                }
                string postRepairappraisal = txtPostRepairAppraisalAmount.Text;
                postRepairappraisal = postRepairappraisal.Replace(",", "");
                postRepairappraisal = postRepairappraisal.Replace("$", "");
                double postRepairappraisalDbl = 0;

                isNumber = double.TryParse(postRepairappraisal, out postRepairappraisalDbl);
                if (isNumber)
                    Loan.LoanUW.LoanUWPostRepairAppraisalAmount = postRepairappraisalDbl;
                else
                {
                    ErrorLabel.Content = "Post Repair Appraisal Amount Format Error";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;

                }

                string zillowSqFt = txtZillowSquareFootage.Text;
                zillowSqFt = zillowSqFt.Replace(",", "");
                double zillowSqFtDbl = 0;

                isNumber = double.TryParse(zillowSqFt, out zillowSqFtDbl);
                if (isNumber)
                    Loan.LoanUW.LoanUWZillowSqFt = zillowSqFtDbl;
                else
                {
                    ErrorLabel.Content = "Zillow Sq Ft Format Error";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    Mouse.OverrideCursor = null;
                    return;

                }

                if (chkAppraisal.IsChecked == true)
                    Loan.LoanUW.LoanUWAppraisal = true;
                else
                    Loan.LoanUW.LoanUWAppraisal = false;

                if (chkAllonge.IsChecked == true)
                    Loan.LoanUW.LoanUWAllongePromissoryNote = true;
                else
                    Loan.LoanUW.LoanUWAllongePromissoryNote = false;

                if (chkLoanApplication.IsChecked == true)
                    Loan.LoanUW.LoanUW10031008LoanApplication = true;
                else
                    Loan.LoanUW.LoanUW10031008LoanApplication = false;

                if (chkCreditReport.IsChecked == true)
                    Loan.LoanUW.LoanUWCreditReport = true;
                else
                    Loan.LoanUW.LoanUWCreditReport = false;

                if (chkFloodCertificate.IsChecked == true)
                    Loan.LoanUW.LoanUWFloodCertificate = true;
                else
                    Loan.LoanUW.LoanUWFloodCertificate = false;

                if (chkAssignmentOfMortgage.IsChecked == true)
                    Loan.LoanUW.LoanUWAssignmentOfMortgage = true;
                else
                    Loan.LoanUW.LoanUWAssignmentOfMortgage = false;

                if (chkBackgroundCheck.IsChecked == true)
                    Loan.LoanUW.LoanUWBackgroundCheck = true;
                else
                    Loan.LoanUW.LoanUWBackgroundCheck = false;

                if (chkCertificateOfGoodStanding.IsChecked == true)
                    Loan.LoanUW.LoanUWCertofGoodStandingFormation = true;
                else
                    Loan.LoanUW.LoanUWCertofGoodStandingFormation = false;

                if (chkClosingProtectionLetter.IsChecked == true)
                    Loan.LoanUW.LoanUWClosingProtectionLetter = true;
                else
                    Loan.LoanUW.LoanUWClosingProtectionLetter = false;

                if (chkHomeownersInsurance.IsChecked == true)
                    Loan.LoanUW.LoanUWHomeownersInsurance = true;
                else
                    Loan.LoanUW.LoanUWHomeownersInsurance = false;

                if (chkCommitteeReview.IsChecked == true)
                    Loan.LoanUW.LoanUWCommitteeReview = true;
                else
                    Loan.LoanUW.LoanUWCommitteeReview = false;

                if (chkLoanPackage.IsChecked == true)
                    Loan.LoanUW.LoanUWLoanPackage = true;
                else
                    Loan.LoanUW.LoanUWLoanPackage = false;

                //if (chkLoanSizer.IsChecked == true)
                //    Loan.LoanUW.LoanUWLoanSizerLoanSubmissionTape = true;
                //else
                //    Loan.LoanUW.LoanUWLoanSizerLoanSubmissionTape = false;

                if (chkTitleCommitment.IsChecked == true)
                    Loan.LoanUW.LoanUWTitleCommitment = true;
                else
                    Loan.LoanUW.LoanUWTitleCommitment = false;

                if (chkToorak.IsChecked == true)
                    Loan.LoanUW.LoanUWClaytonReportApprovalEmail = true;
                else
                    Loan.LoanUW.LoanUWClaytonReportApprovalEmail = false;


                if (chkHud.IsChecked == true)
                    Loan.LoanUW.LoanUWHUD1SettlementStatement = true;
                else
                    Loan.LoanUW.LoanUWHUD1SettlementStatement = false;

                if (chkUnderwritingComplete.IsChecked == true)
                    Loan.LoanUW.LoanUWIsComplete = true;
                else
                    Loan.LoanUW.LoanUWIsComplete = false;

                if (Loan.LoanUW.LoanUWIsComplete)
                    Loan.LoanUW.CompletedBy = MyVariables.user.UserID;

                Loan.LoanUW.LoanID = LoanID;
                Loan.LoanUW.Save();
                Loan = new TableFundingLoan(LoanID);

                TableFundingClient client = new TableFundingClient(Loan.LoanClientID);

                TableFundingFees fees = new TableFundingFees(Loan.LoanID);
                fees.LoanID = Loan.LoanID;
                if (client.MinimumInterest > (client.ClientPrimeRate + client.ClientPrimeRateSpread))
                    fees.MinimumInterest = client.MinimumInterest;
                else                    
                    fees.MinimumInterest = client.ClientPrimeRate + client.ClientPrimeRateSpread;

                fees.OriginationDiscount= client.OriginationDiscount;
                fees.OriginationDiscount2 = client.OriginationDiscount2;
                fees.OriginationDiscountNumDays = client.OriginationDiscountNumDays;
                fees.OriginationDiscountNumDays2 = client.OriginationDiscountNumDays2;
                fees.InterestBasedOnAdvance = client.InterestBasedOnAdvance;
                fees.OriginationBasedOnAdvance = client.OriginationBasedOnAdvance;
                fees.NoInterest = client.NoInterest;
                fees.ClientPrimeRate = client.ClientPrimeRate;
                fees.ClientPrimeRateSpread = client.ClientPrimeRateSpread;
                fees.Save();
                MyVariables.needsRefresh = true;
                if (ErrorLabel.Content.ToString() != "WARNING - Optional fields are blank. - OTHER DATA WILL BE SAVED")
                { 
                    ErrorLabel.Content = "Saved Successfully";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Green);
                }
                LoadData(false);
                Mouse.OverrideCursor = null;
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                Mouse.OverrideCursor = null;
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }

        }
        public bool VerifyReqFieldsOnSave()
        {
            bool status = true;

            bool isNumber;
            double number = 0;
            string numberText;

            if (!dtFundingDate.SelectedDate.HasValue)
                status = false;

            if (cmbClient.SelectedValue is null)
                status = false;

            if (txtLoanNumber.Text.Length == 0)
                status = false;

            if (txtCustomerName.Text.Length == 0)
                status = false;

            if (cmbLoanType.SelectedValue is null)
                status = false;

            if (cmbLoanDwellingType.SelectedValue is null)
                status = false;

            if (cmbState.SelectedValue is null)
                status = false;

            if (cmbEntity.SelectedValue is null)
                status = false;

            if (txtPropertyAddress.Text.Length == 0)
                status = false;


            numberText = txtInterestRate.Text;
            numberText = numberText.Replace("%", "");
            isNumber = double.TryParse(numberText, out number);

            if (number == 0)
                status = false;

            numberText = txtAdvanceRate.Text;
            numberText = numberText.Replace("%", "");
            isNumber = double.TryParse(numberText, out number);

            if (number == 0)
                status = false;


            numberText = txtMortgageAmount.Text;
            numberText = numberText.Replace(",", "");
            numberText = numberText.Replace("$", "");

            isNumber = double.TryParse(numberText, out number);

            if (number == 0)
                status = false;

            if (cmbInvestor.SelectedValue is null)
                status = false;

            return status;
        }
        public bool VerifyOptFieldsOnSave()
        {
            bool status = true;
            bool isNumber;
            double number = 0;
            string numberText;


            numberText = txtInterestRate.Text;
            numberText = numberText.Replace("%", "");
            numberText = numberText.Replace(",", "");
            isNumber = double.TryParse(numberText, out number);

            if (number == 0)
                status = false;


            if (txtZillowSquareFootage.Text.Length == 0)
                status = false;

            if (chkAppraisal.IsChecked == false)
                status = false;

            if (chkAllonge.IsChecked == false)
                status = false;

            if (chkLoanApplication.IsChecked == false)
                status = false;

            if (chkCreditReport.IsChecked == false)
                status = false;

            if (chkFloodCertificate.IsChecked == false)
                status = false;

            if (chkAssignmentOfMortgage.IsChecked == false)
                status = false;

            if (chkBackgroundCheck.IsChecked == false)
                status = false;

            if (chkCertificateOfGoodStanding.IsChecked == false)
                status = false;

            if (chkClosingProtectionLetter.IsChecked == false)
                status = false;

            if (chkHomeownersInsurance.IsChecked == false)
                status = false;

            if (chkCommitteeReview.IsChecked == false)
                status = false;

            if (chkLoanPackage.IsChecked == false)
                status = false;

            //if (chkLoanSizer.IsChecked == false)
            //    status = false;

            if (chkTitleCommitment.IsChecked == false)
                status = false;

            if (chkToorak.IsChecked == false)
                status = false;
            
            if (chkHud.IsChecked == false)
                status = false;

            return status;
        }
        public string FormatNumberCommas(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("#,##0.00");
        }
        public string FormatNumberCommasMoney(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("$#,##0.00");
        }

        public string FormatNumberCommasNoDec(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("#,##0");
        }
        public string FormatPcnt(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("0.000%");
        }

        private void CmbClient_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            { 
                int clientID = Convert.ToInt32(cmbClient.SelectedValue);
                TableFundingClient client = new TableFundingClient(clientID);
                string advanceRate = txtAdvanceRate.Text;
                advanceRate = advanceRate.Replace("%", "");
                double advanceRateDbl = 0;

                bool isNumber = double.TryParse(advanceRate, out advanceRateDbl);
                if (isNumber && advanceRateDbl == 0)
                    txtAdvanceRate.Text = FormatPcnt(client.AdvanceRate.ToString());

            }
            catch (Exception ex)
            {                
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }
        }

        private void UnderwritingPrintButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            { 
            Spire.Doc.Document document = new Spire.Doc.Document();
            //Add Section

            Spire.Doc.Section section1 = document.AddSection();

            //Add Paragraph

            Spire.Doc.Documents.Paragraph paragraph1 = section1.AddParagraph();
            Spire.Doc.Table tblInv = section1.AddTable(true);
            tblInv.ResetCells(34, 2);
            
            //row 1
            Spire.Doc.TableRow DataRow = tblInv.Rows[0];
            Spire.Doc.Documents.Paragraph r1p0 = DataRow.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r1tr0 = r1p0.AppendText("Closing/Funding Date");

            Spire.Doc.Documents.Paragraph r1p1 = DataRow.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r1tr1 = r1p1.AppendText(Convert.ToDateTime(Loan.LoanFundingDate).ToShortDateString());
            
            //row 2
            Spire.Doc.TableRow DataRow2 = tblInv.Rows[1];
            Spire.Doc.Documents.Paragraph r2p0 = DataRow2.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r2tr0 = r2p0.AppendText("Client");

            Spire.Doc.Documents.Paragraph r2p1 = DataRow2.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r2tr1 = r2p1.AppendText(Loan.ClientName);

            //row 3
            Spire.Doc.TableRow DataRow3 = tblInv.Rows[2];
            Spire.Doc.Documents.Paragraph r3p0 = DataRow3.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r3tr0 = r3p0.AppendText("Loan #");

            Spire.Doc.Documents.Paragraph r3p1 = DataRow3.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r3tr1 = r3p1.AppendText(Loan.LoanNumber);

            //row 4
            Spire.Doc.TableRow DataRow4 = tblInv.Rows[3];
            Spire.Doc.Documents.Paragraph r4p0 = DataRow4.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r4tr0 = r4p0.AppendText("Customer Name");

            Spire.Doc.Documents.Paragraph r4p1 = DataRow4.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r4tr1 = r4p1.AppendText(Loan.LoanMortgagee);

            //row 5
            Spire.Doc.TableRow DataRow5 = tblInv.Rows[4];
            Spire.Doc.Documents.Paragraph r5p0 = DataRow5.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r5tr0 = r5p0.AppendText("Business Name");

            Spire.Doc.Documents.Paragraph r5p1 = DataRow5.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r5tr1 = r5p1.AppendText(Loan.LoanMortgageeBusiness);

            //row 6
            Spire.Doc.TableRow DataRow6 = tblInv.Rows[5];
            Spire.Doc.Documents.Paragraph r6p0 = DataRow6.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r6tr0 = r6p0.AppendText("Loan Type");

            Spire.Doc.Documents.Paragraph r6p1 = DataRow6.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r6tr1 = r6p1.AppendText(Loan.LoanTypeName);

            //row 7
            Spire.Doc.TableRow DataRow7 = tblInv.Rows[6];
            Spire.Doc.Documents.Paragraph r7p0 = DataRow7.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r7tr0 = r7p0.AppendText("Entity");

            Spire.Doc.Documents.Paragraph r7p1 = DataRow7.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r7tr1 = r7p1.AppendText(Loan.EntityName);

            //row 8
            Spire.Doc.TableRow DataRow8 = tblInv.Rows[7];
            Spire.Doc.Documents.Paragraph r8p0 = DataRow8.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r8tr0 = r8p0.AppendText("Property Address");

            Spire.Doc.Documents.Paragraph r8p1 = DataRow8.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r8tr1 = r8p1.AppendText(Loan.LoanPropertyAddress);
            
            //row 9
            Spire.Doc.TableRow DataRow9 = tblInv.Rows[8];
            Spire.Doc.Documents.Paragraph r9p0 = DataRow9.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r9tr0 = r9p0.AppendText("Pre Closing");
            r9tr0.CharacterFormat.Bold = true;

            //row 10
            Spire.Doc.TableRow DataRow10 = tblInv.Rows[9];
            Spire.Doc.Documents.Paragraph r10p0 = DataRow10.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r10tr0 = r10p0.AppendText("1003/1008 (Loan Application)");

            Spire.Doc.Documents.Paragraph r10p1 = DataRow10.Cells[1].AddParagraph();
            if (chkLoanApplication.IsChecked == true)
                r10p1.AppendCheckBox("chkLoanApplication", true);
            else
                r10p1.AppendCheckBox("chkLoanApplication", false);

            //row 11
            Spire.Doc.TableRow DataRow11 = tblInv.Rows[10];
            Spire.Doc.Documents.Paragraph r11p0 = DataRow11.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r11tr0 = r11p0.AppendText("Allonge Promissory Note");

            Spire.Doc.Documents.Paragraph r11p1 = DataRow11.Cells[1].AddParagraph();
            if (chkAllonge.IsChecked == true)
                r11p1.AppendCheckBox("chkAllonge", true);
            else
                r11p1.AppendCheckBox("chkAllonge", false);

            //row 12
            Spire.Doc.TableRow DataRow12 = tblInv.Rows[11];
            Spire.Doc.Documents.Paragraph r12p0 = DataRow12.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r12tr0 = r12p0.AppendText("Appraisal");


            Spire.Doc.Documents.Paragraph r12p1 = DataRow12.Cells[1].AddParagraph();
            if (chkAppraisal.IsChecked == true)
                r12p1.AppendCheckBox("chkAppraisal", true);
            else
                r12p1.AppendCheckBox("chkAppraisal", false);


            //row 13
            Spire.Doc.TableRow DataRow13 = tblInv.Rows[12];
            Spire.Doc.Documents.Paragraph r13p0 = DataRow13.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r13tr0 = r13p0.AppendText("Appraisal Amount");


            Spire.Doc.Documents.Paragraph r13p1 = DataRow13.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r13tr1 = r13p1.AppendText(FormatNumberCommasMoney(Loan.LoanUW.LoanUWAppraisalAmount.ToString()));

            //row 14
            Spire.Doc.TableRow DataRow14 = tblInv.Rows[13];
            Spire.Doc.Documents.Paragraph r14p0 = DataRow14.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r14tr0 = r14p0.AppendText("Post Repair Appraisal Amount");


            Spire.Doc.Documents.Paragraph r14p1 = DataRow14.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r14tr1 = r14p1.AppendText(FormatNumberCommasMoney(Loan.LoanUW.LoanUWPostRepairAppraisalAmount.ToString()));

            //row 15
            Spire.Doc.TableRow DataRow15 = tblInv.Rows[14];
            Spire.Doc.Documents.Paragraph r15p0 = DataRow15.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r15tr0 = r15p0.AppendText("Assignment of Mortgage");


            Spire.Doc.Documents.Paragraph r15p1 = DataRow15.Cells[1].AddParagraph();
            if (chkAssignmentOfMortgage.IsChecked == true)
                r15p1.AppendCheckBox("chkAssignmentOfMortgage", true);
            else
                r15p1.AppendCheckBox("chkAssignmentOfMortgage", false);

            //row 16
            Spire.Doc.TableRow DataRow16 = tblInv.Rows[15];
            Spire.Doc.Documents.Paragraph r16p0 = DataRow16.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r16tr0 = r16p0.AppendText("Background Check");


            Spire.Doc.Documents.Paragraph r16p1 = DataRow16.Cells[1].AddParagraph();
            if (chkBackgroundCheck.IsChecked == true)
                r16p1.AppendCheckBox("chkBackgroundCheck", true);
            else
                r16p1.AppendCheckBox("chkBackgroundCheck", false);

            //row 17
            Spire.Doc.TableRow DataRow17 = tblInv.Rows[16];
            Spire.Doc.Documents.Paragraph r17p0 = DataRow17.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r17tr0 = r17p0.AppendText("Certificate of Good Standing/Formation");


            Spire.Doc.Documents.Paragraph r17p1 = DataRow17.Cells[1].AddParagraph();
            if (chkCertificateOfGoodStanding.IsChecked == true)
                r17p1.AppendCheckBox("chkCertificateOfGoodStanding", true);
            else
                r17p1.AppendCheckBox("chkCertificateOfGoodStanding", false);

            //row 18
            Spire.Doc.TableRow DataRow18 = tblInv.Rows[17];
            Spire.Doc.Documents.Paragraph r18p0 = DataRow18.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r18tr0 = r18p0.AppendText("Closing Protection Letter");


            Spire.Doc.Documents.Paragraph r18p1 = DataRow18.Cells[1].AddParagraph();
            if (chkClosingProtectionLetter.IsChecked == true)
                r18p1.AppendCheckBox("chkClosingProtectionLetter", true);
            else
                r18p1.AppendCheckBox("chkClosingProtectionLetter", false);

            //row 19
            Spire.Doc.TableRow DataRow19 = tblInv.Rows[18];
            Spire.Doc.Documents.Paragraph r19p0 = DataRow19.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r19tr0 = r19p0.AppendText("Committee Review");


            Spire.Doc.Documents.Paragraph r19p1 = DataRow19.Cells[1].AddParagraph();
            if (chkCommitteeReview.IsChecked == true)
                r19p1.AppendCheckBox("chkCommitteeReview", true);
            else
                r19p1.AppendCheckBox("chkCommitteeReview", false);

            //row 20
            Spire.Doc.TableRow DataRow20 = tblInv.Rows[19];
            Spire.Doc.Documents.Paragraph r20p0 = DataRow20.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r20tr0 = r20p0.AppendText("Credit Report");


            Spire.Doc.Documents.Paragraph r20p1 = DataRow20.Cells[1].AddParagraph();
            if (chkCreditReport.IsChecked == true)
                r20p1.AppendCheckBox("chkCreditReport", true);
            else
                r20p1.AppendCheckBox("chkCreditReport", false);

            //row 21
            Spire.Doc.TableRow DataRow21 = tblInv.Rows[20];
            Spire.Doc.Documents.Paragraph r21p0 = DataRow21.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r21tr0 = r21p0.AppendText("Flood Certificate");


            Spire.Doc.Documents.Paragraph r21p1 = DataRow21.Cells[1].AddParagraph();
            if (chkFloodCertificate.IsChecked == true)
                r21p1.AppendCheckBox("chkFloodCertificate", true);
            else
                r21p1.AppendCheckBox("chkFloodCertificate", false);

            //row 22
            Spire.Doc.TableRow DataRow22 = tblInv.Rows[21];
            Spire.Doc.Documents.Paragraph r22p0 = DataRow22.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r22tr0 = r22p0.AppendText("HUD-1 (Settlement Statement)");


            Spire.Doc.Documents.Paragraph r22p1 = DataRow22.Cells[1].AddParagraph();
            if (chkHud.IsChecked == true)
                r22p1.AppendCheckBox("chkHud", true);
            else
                r22p1.AppendCheckBox("chkHud", false);

            //row 23
            Spire.Doc.TableRow DataRow23 = tblInv.Rows[22];
            Spire.Doc.Documents.Paragraph r23p0 = DataRow23.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r23tr0 = r23p0.AppendText("Insurance (Homeowners)");


            Spire.Doc.Documents.Paragraph r23p1 = DataRow23.Cells[1].AddParagraph();
            if (chkHomeownersInsurance.IsChecked == true)
                r23p1.AppendCheckBox("chkHomeownersInsurance", true);
            else
                r23p1.AppendCheckBox("chkHomeownersInsurance", false);

            //row 24
            Spire.Doc.TableRow DataRow24 = tblInv.Rows[23];
            Spire.Doc.Documents.Paragraph r24p0 = DataRow24.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r24tr0 = r24p0.AppendText("Loan Package (Closing Package and Instructions)");


            Spire.Doc.Documents.Paragraph r24p1 = DataRow24.Cells[1].AddParagraph();
            if (chkLoanPackage.IsChecked == true)
                r24p1.AppendCheckBox("chkLoanPackage", true);
            else
                r24p1.AppendCheckBox("chkLoanPackage", false);

                //row 25
                //Spire.Doc.TableRow DataRow25 = tblInv.Rows[24];
                //Spire.Doc.Documents.Paragraph r25p0 = DataRow25.Cells[0].AddParagraph();
                //Spire.Doc.Fields.TextRange r25tr0 = r25p0.AppendText("Loan Sizer (Toorak)/Loan Submission Tape (SG/AF/V)");


            //    Spire.Doc.Documents.Paragraph r25p1 = DataRow25.Cells[1].AddParagraph();
            //if (chkLoanSizer.IsChecked == true)
            //    r25p1.AppendCheckBox("chkLoanSizer", true);
            //else
            //    r25p1.AppendCheckBox("chkLoanSizer", false);

            //row 26
            Spire.Doc.TableRow DataRow26 = tblInv.Rows[24];
            Spire.Doc.Documents.Paragraph r26p0 = DataRow26.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r26tr0 = r26p0.AppendText("Title Commitment");


            Spire.Doc.Documents.Paragraph r26p1 = DataRow26.Cells[1].AddParagraph();
            if (chkTitleCommitment.IsChecked == true)
                r26p1.AppendCheckBox("chkTitleCommitment", true);
            else
                r26p1.AppendCheckBox("chkTitleCommitment", false);

            //row 27
            Spire.Doc.TableRow DataRow27 = tblInv.Rows[25];
            Spire.Doc.Documents.Paragraph r27p0 = DataRow27.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r27tr0 = r27p0.AppendText("Loan Sizer/Approval");


            Spire.Doc.Documents.Paragraph r27p1 = DataRow27.Cells[1].AddParagraph();
            if (chkToorak.IsChecked == true)
                r27p1.AppendCheckBox("chkToorak", true);
            else
                r27p1.AppendCheckBox("chkToorak", false);

            //row 28
            Spire.Doc.TableRow DataRow28 = tblInv.Rows[26];
            Spire.Doc.Documents.Paragraph r28p0 = DataRow28.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r28tr0 = r28p0.AppendText("Interest Rate");



            Spire.Doc.Documents.Paragraph r28p1 = DataRow28.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r28tr1 = r28p1.AppendText(FormatPcnt(Loan.LoanInterestRate.ToString()));


            //row 29
            Spire.Doc.TableRow DataRow29 = tblInv.Rows[27];
            Spire.Doc.Documents.Paragraph r29p0 = DataRow29.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r29tr0 = r29p0.AppendText("Request Loan Amount");
            Spire.Doc.Documents.Paragraph r29p1 = DataRow29.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r29tr1 = r29p1.AppendText(FormatNumberCommasMoney(Loan.LoanMortgageAmount.ToString()));

            //row 30
            Spire.Doc.TableRow DataRow30 = tblInv.Rows[28];
            Spire.Doc.Documents.Paragraph r30p0 = DataRow30.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r30tr0 = r30p0.AppendText("Advance Rate");
            Spire.Doc.Documents.Paragraph r30p1 = DataRow30.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r30tr1 = r30p1.AppendText(FormatPcnt(Loan.LoanAdvanceRate.ToString()));

            //row 31
            Spire.Doc.TableRow DataRow31 = tblInv.Rows[29];
            Spire.Doc.Documents.Paragraph r31p0 = DataRow31.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r31tr0 = r31p0.AppendText("Zillow Square Footage");
            Spire.Doc.Documents.Paragraph r31p1 = DataRow31.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r31tr1 = r31p1.AppendText(FormatNumberCommasNoDec(Loan.LoanUW.LoanUWZillowSqFt.ToString()));

            //row 32
            Spire.Doc.TableRow DataRow32 = tblInv.Rows[30];
            Spire.Doc.Documents.Paragraph r32p0 = DataRow32.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r32tr0 = r32p0.AppendText("Funded By");
            Spire.Doc.Documents.Paragraph r32p1 = DataRow32.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r32tr1 = r32p1.AppendText(Loan.InvestorName);

            //row 33
            Spire.Doc.TableRow DataRow33 = tblInv.Rows[31];
            Spire.Doc.Documents.Paragraph r33p0 = DataRow33.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r33tr0 = r33p0.AppendText("Underwriting Complete?");


            Spire.Doc.Documents.Paragraph r33p1 = DataRow33.Cells[1].AddParagraph();
            if (chkUnderwritingComplete.IsChecked == true)
                r33p1.AppendCheckBox("chkUnderwritingComplete", true);
            else
                r33p1.AppendCheckBox("chkUnderwritingComplete", false);
            
            //row 34
            Spire.Doc.TableRow DataRow34 = tblInv.Rows[32];
            Spire.Doc.Documents.Paragraph r34p0 = DataRow34.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r34tr0 = r34p0.AppendText("Completed By");
            Spire.Doc.Documents.Paragraph r34p1 = DataRow34.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r34tr1 = r34p1.AppendText(Loan.LoanUW.CompletedByName);

            //row 35
            Spire.Doc.TableRow DataRow35 = tblInv.Rows[33];
            Spire.Doc.Documents.Paragraph r35p0 = DataRow35.Cells[0].AddParagraph();
            Spire.Doc.Fields.TextRange r35tr0 = r35p0.AppendText("Date Deposited in Escrow");
            Spire.Doc.Documents.Paragraph r35p1 = DataRow35.Cells[1].AddParagraph();
            Spire.Doc.Fields.TextRange r35tr1 = r35p1.AppendText(Convert.ToDateTime(Loan.LoanFunding.DateDepositedInEscrow).ToShortDateString());

            tblInv.TableFormat.Borders.Right.BorderType = Spire.Doc.Documents.BorderStyle.None;
            tblInv.TableFormat.Borders.Left.BorderType = Spire.Doc.Documents.BorderStyle.None;
            tblInv.TableFormat.Borders.Top.BorderType = Spire.Doc.Documents.BorderStyle.None;
            tblInv.TableFormat.Borders.Bottom.BorderType = Spire.Doc.Documents.BorderStyle.None;
            tblInv.TableFormat.Borders.Vertical.BorderType = Spire.Doc.Documents.BorderStyle.None;
            tblInv.TableFormat.Borders.Horizontal.BorderType = Spire.Doc.Documents.BorderStyle.None;

            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = "C:\\Users";
            dialog.IsFolderPicker = true;
                
            string fileName;

            fileName = "";

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                fileName = dialog.FileName + "\\" + "Underwriting_" + Loan.LoanNumber + ".docx";

                }

            //string fileName = "Underwriting_" + Loan.LoanNumber + ".docx";
            document.SaveToFile(fileName, FileFormat.Docx);

            System.Diagnostics.Process.Start(fileName);
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }
        }
    }
}
