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
using Microsoft.WindowsAPICodePack.Dialogs;
using Clearwater.DataAccess;
using log4net;
namespace RAI_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class Loans : Window
    {
        DataSet clients;
        DataSet loanType = RunsStoredProc.RunStoredProc("MiscValue_SEL", "@MiscTypeID", "3", "", "", "", "", "", "", "", "", "", "", "", 0);
        DataSet loanDwellingType = RunsStoredProc.RunStoredProc("MiscValue_SEL", "@MiscTypeID", "6", "", "", "", "", "", "", "", "", "", "", "", 0);
        DataSet states = RunsStoredProc.RunStoredProc("MiscValue_SEL", "@MiscTypeID", "7", "", "", "", "", "", "", "", "", "", "", "", 0);
        DataSet entities = RunsStoredProc.RunStoredProc("TableFundingEntity_SEL", "", "", "", "", "", "", "", "", "", "", "", "", "", 0);
        DataTable entityTable;
        DataSet investors;
        DataSet loanStatus = RunsStoredProc.RunStoredProc("MiscValue_SEL", "@MiscTypeID", "2", "", "", "", "", "", "", "", "", "", "", "", 0);
        //public bool needsRefresh;
        public static bool needsRefresh;
        public int test = 1;
        //protected static bool needsRefresh = false;
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);



        TableFundingLoans loans;

        public DataTable EntityList { get; private set; }
        public DataRow CurrentEntity { get; private set; }

        public Loans()
        {
            try
            { 
                MyVariables.needsRefresh = false;

                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                loans = new TableFundingLoans(false);
                InitializeComponent();
                SetClientDataSet();
                SetInvestorDataSet();
                //SetClientComboBoxForClientPopup();
                SetClientComboBoxForFilter();

                SetInvestorComboBoxForFilter();
                //SetUsersDataSet();
                //SetUserComboBox();x
                SetComboBoxes();

                LoansDataGrid.ItemsSource = loans;
                MyVariables.letterFilePath = ConfigurationManager.AppSettings["LetterFilePath"].ToString();
                MyVariables.letterPage1FilePath = ConfigurationManager.AppSettings["LetterPage1FilePath"].ToString();
                MyVariables.excelReportFilePath = ConfigurationManager.AppSettings["ExcelReportFilePath"].ToString();
                Mouse.OverrideCursor = null;

                if (!MyVariables.userIsAdmin)
                {
                    UsersButton.Visibility = Visibility.Hidden;
                    ClientButton.Visibility = Visibility.Hidden;
                    InvestorButton.Visibility = Visibility.Hidden;
                }
            }
            catch (Exception ex)
            {                
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }
        }
        public void refreshMainWindowGrid()
        {
            ErrorLabel.Content = "";
            try
            {
                int clientID = Convert.ToInt32(cmbClientFilter.SelectedValue);
                int investorID = Convert.ToInt32(cmbInvestorFilter.SelectedValue);
                int statusID = Convert.ToInt32(cmbLoanStatus.SelectedValue);
                bool showCompleted;

                if (chkShowCompleted.IsChecked == true)
                    showCompleted = true;
                else
                    showCompleted = false;

                if (clientID == 0)
                    clientID = -1;

                if (statusID == 0)
                    statusID = -1;

                if (investorID == 0)
                    investorID = -1;

                loans = new TableFundingLoans(clientID, investorID, statusID, false, "-1", showCompleted);
                //loans = new TableFundingLoans(false);
                //LoansDataGrid.ItemsSource = null;
                LoansDataGrid.ItemsSource = loans;
                //LoansDataGrid.Items.Refresh();
            }
            catch (Exception ex)
            {                
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }
        }

        private void MenuUnderwriting_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
                TableFundingLoan loan = (TableFundingLoan)LoansDataGrid.SelectedItem;

                UnderwritingWindow uWindow = new UnderwritingWindow(loan, clients, loanType, entities, investors, loanDwellingType, states);
                uWindow.Show();

            }

            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }

        }

        private void MenuAccounting_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
                TableFundingLoan loan = (TableFundingLoan)LoansDataGrid.SelectedItem;

                AccountingWindow aWindow = new AccountingWindow(loan.LoanID);
                aWindow.Show();
            }

            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }

        }

        private void ClientButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {


                ClientWindow clientWindow = new ClientWindow();
                clientWindow.Show();

            }

            catch (Exception ex)
            {
               ErrorLabel.Content = ex.Message;
               ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
               log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }

        }

        private void InvestorButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {

                //cmbInvestorI.Visibility = System.Windows.Visibility.Visible;
                //txtInvestorI.Visibility = System.Windows.Visibility.Hidden;

                //cmbInvestorI.SelectedIndex = 0;
                //txtInvestorI.Text = "";
                //txtCustodian.Text = "";
                //InvestorPopup.IsOpen = true;

                InvestorWindow investorWindow = new InvestorWindow();
                investorWindow.Show();
            }

            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }

        }


        public string FormatNumberCommas(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("#,##0");
        }
        public string FormatNumberCommas2dec(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("#,##0.00;(#,##0.00)");
        }

        public string FormatPcnt(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("0.000%");
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
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }
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
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }

        }

        public void SetInvestorComboBoxForFilter()
        {
            ErrorLabel.Content = "";
            try
            {
                cmbInvestorFilter.ItemsSource = investors.Tables[0].DefaultView;
                cmbInvestorFilter.DisplayMemberPath = investors.Tables[0].Columns["InvestorName"].ToString();
                cmbInvestorFilter.SelectedValuePath = investors.Tables[0].Columns["InvestorID"].ToString();

            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }
        }

        public void SetClientComboBoxForFilter()
        {
            ErrorLabel.Content = "";
            try
            {
                cmbClientFilter.ItemsSource = clients.Tables[0].DefaultView;
                cmbClientFilter.DisplayMemberPath = clients.Tables[0].Columns["ClientName"].ToString();
                cmbClientFilter.SelectedValuePath = clients.Tables[0].Columns["ClientID"].ToString();

            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

            }

        }
        public void SetComboBoxes()
        {
            ErrorLabel.Content = "";
            try
            {
                entityTable = entities.Tables[0];

                DataRow statusRow = loanStatus.Tables[0].NewRow();
                statusRow["MiscTypeValueID"] = "-1";
                statusRow["Value"] = " --Select--";

                loanStatus.Tables[0].Rows.Add(statusRow);
                loanStatus.Tables[0].DefaultView.Sort = "Value";

                cmbLoanStatus.ItemsSource = loanStatus.Tables[0].DefaultView;
                cmbLoanStatus.DisplayMemberPath = loanStatus.Tables[0].Columns["Value"].ToString();
                cmbLoanStatus.SelectedValuePath = loanStatus.Tables[0].Columns["MiscTypeValueID"].ToString();
                //cmbLoanStatus.SelectedIndex = 0;


                //EntityColumn.ItemsSource = entities.Tables[0].DefaultView;
                //EntityColumn.DisplayMemberPath = entities.Tables[0].Columns["EntityName"].ToString();
                //EntityColumn.SelectedValuePath = entities.Tables[0].Columns["EntityID"].ToString();
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }

        }
        private void CmbLoanStatus_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            ErrorLabel.Content = "";
            try
            {
                int clientID = Convert.ToInt32(cmbClientFilter.SelectedValue);
                int investorID = Convert.ToInt32(cmbInvestorFilter.SelectedValue);
                int statusID = Convert.ToInt32(cmbLoanStatus.SelectedValue);
                bool showCompleted;

                if (chkShowCompleted.IsChecked == true)
                    showCompleted = true;
                else
                    showCompleted = false;

                if (clientID == 0)
                    clientID = -1;

                if (statusID == 0)
                    statusID = -1;

                if (investorID == 0)
                    investorID = -1;

                loans = new TableFundingLoans(clientID, investorID, statusID, false, "-1", showCompleted);

                LoansDataGrid.ItemsSource = loans;
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

                Mouse.OverrideCursor = null;
            }

            Mouse.OverrideCursor = null;
        }

        private void CmbClientFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            ErrorLabel.Content = "";
            try
            { 
                int clientID = Convert.ToInt32(cmbClientFilter.SelectedValue);
                int investorID = Convert.ToInt32(cmbInvestorFilter.SelectedValue);
                int statusID = Convert.ToInt32(cmbLoanStatus.SelectedValue);
                bool showCompleted;

                if (chkShowCompleted.IsChecked == true)
                    showCompleted = true;
                else
                    showCompleted = false;

                if (clientID == 0)
                    clientID = -1;

                if (statusID == 0)
                    statusID = -1;

                if (investorID == 0)
                    investorID = -1;

                loans = new TableFundingLoans(clientID, investorID, statusID, false, "-1", showCompleted);

                LoansDataGrid.ItemsSource = loans;

            }
            catch (Exception ex)
            {                
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

                Mouse.OverrideCursor = null;
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }

            Mouse.OverrideCursor = null;
        }

        private void CmbInvestorFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            ErrorLabel.Content = "";
            try
            { 
                int clientID = Convert.ToInt32(cmbClientFilter.SelectedValue);
                int investorID = Convert.ToInt32(cmbInvestorFilter.SelectedValue);
                int statusID = Convert.ToInt32(cmbLoanStatus.SelectedValue);
                bool showCompleted;

                if (chkShowCompleted.IsChecked == true)
                    showCompleted = true;
                else
                    showCompleted = false;

                if (clientID == 0)
                    clientID = -1;

                if (statusID == 0)
                    statusID = -1;

                if (investorID == 0)
                    investorID = -1;

                loans = new TableFundingLoans(clientID, investorID, statusID, false, "-1", showCompleted);

                LoansDataGrid.ItemsSource = loans;
            }
            catch (Exception ex)
            {                
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);

                Mouse.OverrideCursor = null;
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }

            Mouse.OverrideCursor = null;
        }

        private void ChkShowCompleted_Checked(object sender, RoutedEventArgs e)
        {

            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            ErrorLabel.Content = "";
            showAll();

            Mouse.OverrideCursor = null;

        }
        private void showAll()
        {
            try
            {

                int clientID = Convert.ToInt32(cmbClientFilter.SelectedValue);
                int investorID = Convert.ToInt32(cmbInvestorFilter.SelectedValue);
                int statusID = Convert.ToInt32(cmbLoanStatus.SelectedValue);
                bool showCompleted;

                if (chkShowCompleted.IsChecked == true)
                    showCompleted = true;
                else
                    showCompleted = false;

                if (clientID == 0)
                    clientID = -1;

                if (statusID == 0)
                    statusID = -1;

                if (investorID == 0)
                    investorID = -1;

                TableFundingLoans loans = new TableFundingLoans(clientID, investorID, statusID, false, "-1", showCompleted);

                LoansDataGrid.ItemsSource = loans;
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }
        }

            private void UsersButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            {
                UserWindow uwindow = new UserWindow();
                uwindow.Show();
            }

            catch (Exception ex)
            {                
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }


        }
        private void ChangePasswordButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            {                
                ChangePasswordWindow uwindow = new ChangePasswordWindow();
                uwindow.Show();
            }

            catch (Exception ex)
            {

                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }


        }

        private void ClientEntities_Click(object sender, RoutedEventArgs e)
        {
            try
            { 
                ClientEntityWindow eWindow = new ClientEntityWindow(clients, entities);
                eWindow.Show();
            }
        
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }       
        }

        private void NewLoanButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            { 
                TableFundingLoan loan = new TableFundingLoan();
                UnderwritingWindow uWindow = new UnderwritingWindow(loan, clients, loanType, entities, investors, loanDwellingType, states);
                uWindow.Show();
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }       

        }
        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            { 
                refreshMainWindowGrid();
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }       
        }

        private void TrackingReportButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            {
                ErrorLabel.Content = "";
                DataSet report = RunsStoredProc.RunStoredProc("TableFunding_TrackingReport", "", "", "", "", "", "", "", "", "", "", "", "", "", 0);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                
                using (ExcelPackage pck = new ExcelPackage())
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Tracking Report");

                    ws.Cells["A4"].LoadFromDataTable(report.Tables[0], true);
                    ws.Cells["A2"].Value = "Tracking Report";
                    ws.Cells["A3"].Formula = "=Today()";

                    ws.Cells["A3"].Style.Numberformat.Format = "yyyy-mm-dd";

                    ws.Cells["A2"].Style.Font.Size = 18;
                    ws.Cells["A3"].Style.Font.Size = 18;

                    ws.Cells[4, 1, 4, 38].Style.Font.Size = 14;
                    ws.Cells[4, 1, 4, 38].Style.Fill.PatternType = ExcelFillStyle.Solid;

                    ws.Cells[4, 1, 4, 38].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
                    ws.Cells[4, 1, 4, 38].Style.Font.Color.SetColor(System.Drawing.Color.White);

                    string client = "";
                    string entity = "";
                    int clientStartRow = 0;
                    int clientEntityStartRow = 0;
                    string rowType = "";

                    int dataStartRow = 5;
                    int dataEndRow = 0;
                    for (int row = dataStartRow; row <= 100000; row++)
                    {
                        if (ws.Cells[row, 1].Value == null)
                        {
                            break;
                        }

                        if (row == dataStartRow)
                        {
                            client = ws.Cells[row, 1].Text;
                            entity = ws.Cells[row, 29].Text;
                            clientStartRow = row;
                            clientEntityStartRow = row;
                        }

                        rowType = ws.Cells[row, 39].Text;
                        if (rowType == "Grand Total")
                        {
                            ws.Cells[row, 2].Value = "";
                            ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + dataStartRow + ":E" + (row - 1) + ")";
                            ws.Cells[row, 6].Value = "";
                            ws.Cells[row, 7].Value = "";
                            ws.Cells[row, 8].Formula = "=SUBTOTAL(9,H" + dataStartRow + ":H" + (row - 1) + ")";
                            ws.Cells[row, 9].Formula = "=SUBTOTAL(9,I" + dataStartRow + ":I" + (row - 1) + ")";
                            ws.Cells[row, 10].Formula = "=SUBTOTAL(9,J" + dataStartRow + ":J" + (row - 1) + ")";
                            ws.Cells[row, 11].Value = "";
                            ws.Cells[row, 12].Value = "";
                            ws.Cells[row, 13].Value = "";
                            ws.Cells[row, 14].Value = "";
                            ws.Cells[row, 15].Value = "";
                            ws.Cells[row, 16].Value = "";
                            ws.Cells[row, 17].Value = "";
                            ws.Cells[row, 18].Formula = "=SUBTOTAL(9,R" + dataStartRow + ":R" + (row - 1) + ")";
                            ws.Cells[row, 19].Value = "";
                            ws.Cells[row, 20].Formula = "=SUBTOTAL(9,T" + dataStartRow + ":T" + (row - 1) + ")";
                            ws.Cells[row, 21].Formula = "=SUBTOTAL(9,U" + dataStartRow + ":U" + (row - 1) + ")";
                            ws.Cells[row, 22].Formula = "=SUBTOTAL(9,V" + dataStartRow + ":V" + (row - 1) + ")";
                            ws.Cells[row, 23].Formula = "=SUBTOTAL(9,W" + dataStartRow + ":W" + (row - 1) + ")";
                            ws.Cells[row, 24].Formula = "=SUBTOTAL(9,X" + dataStartRow + ":X" + (row - 1) + ")";
                            ws.Cells[row, 25].Formula = "=SUBTOTAL(9,Y" + dataStartRow + ":Y" + (row - 1) + ")";
                            ws.Cells[row, 26].Formula = "=SUBTOTAL(9,Z" + dataStartRow + ":z" + (row - 1) + ")";
                            ws.Cells[row, 27].Value = "";
                            ws.Cells[row, 28].Value = "";
                            ws.Cells[row, 29].Value = "";
                            ws.Cells[row, 30].Value = "";
                            ws.Cells[row, 31].Value = "";
                            ws.Cells[row, 32].Value = "";
                            ws.Cells[row, 33].Value = "";
                            ws.Cells[row, 34].Value = "";
                            ws.Cells[row, 35].Value = "";
                            ws.Cells[row, 36].Value = "";
                            ws.Cells[row, 37].Value = "";
                            ws.Cells[row, 38].Value = "";
                            ws.Cells[row, 38].Value = "";

                            ws.Cells[row, 1, row, 41].Style.Font.Bold = true;
                            ws.Cells[row, 1, row, 41].Style.Font.Size = 14;


                            ws.Cells[row, 1, row, 41].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            ws.Cells[row, 1, row, 41].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

                            ws.Cells[row, 1, row, 41].Style.Font.Color.SetColor(System.Drawing.Color.White);



                        }
                        if (rowType == "Client Total")
                        {
                            ws.Cells[row, 2].Value = "";
                            ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + clientStartRow + ":E" + (row - 1) + ")";
                            ws.Cells[row, 6].Value = "";
                            ws.Cells[row, 7].Value = "";
                            ws.Cells[row, 8].Formula = "=SUBTOTAL(9,H" + clientStartRow + ":H" + (row - 1) + ")";
                            ws.Cells[row, 9].Formula = "=SUBTOTAL(9,I" + clientStartRow + ":I" + (row - 1) + ")";
                            ws.Cells[row, 10].Formula = "=SUBTOTAL(9,J" + clientStartRow + ":J" + (row - 1) + ")";
                            ws.Cells[row, 11].Value = "";
                            ws.Cells[row, 12].Value = "";
                            ws.Cells[row, 13].Value = "";
                            ws.Cells[row, 14].Value = "";
                            ws.Cells[row, 15].Value = "";
                            ws.Cells[row, 16].Value = "";
                            ws.Cells[row, 17].Value = "";
                            ws.Cells[row, 18].Formula = "=SUBTOTAL(9,R" + clientStartRow + ":R" + (row - 1) + ")";
                            ws.Cells[row, 19].Value = "";
                            ws.Cells[row, 20].Formula = "=SUBTOTAL(9,T" + clientStartRow + ":T" + (row - 1) + ")";
                            ws.Cells[row, 21].Formula = "=SUBTOTAL(9,U" + clientStartRow + ":U" + (row - 1) + ")";
                            ws.Cells[row, 22].Formula = "=SUBTOTAL(9,V" + clientStartRow + ":V" + (row - 1) + ")";
                            ws.Cells[row, 23].Formula = "=SUBTOTAL(9,W" + clientStartRow + ":W" + (row - 1) + ")";
                            ws.Cells[row, 24].Formula = "=SUBTOTAL(9,X" + clientStartRow + ":X" + (row - 1) + ")";
                            ws.Cells[row, 25].Formula = "=SUBTOTAL(9,Y" + clientStartRow + ":Y" + (row - 1) + ")";
                            ws.Cells[row, 26].Formula = "=SUBTOTAL(9,Z" + clientStartRow + ":z" + (row - 1) + ")";
                            ws.Cells[row, 27].Value = "";
                            ws.Cells[row, 28].Value = "";
                            ws.Cells[row, 29].Value = "";
                            ws.Cells[row, 30].Value = "";
                            ws.Cells[row, 31].Value = "";
                            ws.Cells[row, 32].Value = "";
                            ws.Cells[row, 33].Value = "";
                            ws.Cells[row, 34].Value = "";
                            ws.Cells[row, 35].Value = "";
                            ws.Cells[row, 36].Value = "";
                            ws.Cells[row, 37].Value = "";
                            ws.Cells[row, 38].Value = "";
                            ws.Cells[row, 39].Value = "";

                            ws.Cells[row, 1, row, 41].Style.Font.Size = 13;
                            ws.Cells[row, 1, row, 41].Style.Font.Bold = true;

                            ws.Cells[row, 1, row, 41].Style.Fill.PatternType = ExcelFillStyle.Solid;


                            ws.Cells[row, 1, row, 41].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                        }

                        if (rowType == "Client Entity Total")
                        {
                            ws.Cells[row, 2].Value = "";
                            ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + clientEntityStartRow + ":E" + (row - 1) + ")";
                            ws.Cells[row, 6].Value = "";
                            ws.Cells[row, 7].Value = "";
                            ws.Cells[row, 8].Formula = "=SUBTOTAL(9,H" + clientEntityStartRow + ":H" + (row - 1) + ")";
                            ws.Cells[row, 9].Formula = "=SUBTOTAL(9,I" + clientEntityStartRow + ":I" + (row - 1) + ")";
                            ws.Cells[row, 10].Formula = "=SUBTOTAL(9,J" + clientEntityStartRow + ":J" + (row - 1) + ")";
                            ws.Cells[row, 11].Value = "";
                            ws.Cells[row, 12].Value = "";
                            ws.Cells[row, 13].Value = "";
                            ws.Cells[row, 14].Value = "";
                            ws.Cells[row, 15].Value = "";
                            ws.Cells[row, 16].Value = "";
                            ws.Cells[row, 17].Value = "";
                            ws.Cells[row, 18].Formula = "=SUBTOTAL(9,R" + clientEntityStartRow + ":R" + (row - 1) + ")";
                            ws.Cells[row, 19].Value = "";
                            ws.Cells[row, 20].Formula = "=SUBTOTAL(9,T" + clientEntityStartRow + ":T" + (row - 1) + ")";
                            ws.Cells[row, 21].Formula = "=SUBTOTAL(9,U" + clientEntityStartRow + ":U" + (row - 1) + ")";
                            ws.Cells[row, 22].Formula = "=SUBTOTAL(9,V" + clientEntityStartRow + ":V" + (row - 1) + ")";
                            ws.Cells[row, 23].Formula = "=SUBTOTAL(9,W" + clientEntityStartRow + ":W" + (row - 1) + ")";
                            ws.Cells[row, 24].Formula = "=SUBTOTAL(9,X" + clientEntityStartRow + ":X" + (row - 1) + ")";
                            ws.Cells[row, 25].Formula = "=SUBTOTAL(9,Y" + clientEntityStartRow + ":Y" + (row - 1) + ")";
                            ws.Cells[row, 26].Formula = "=SUBTOTAL(9,Z" + clientEntityStartRow + ":z" + (row - 1) + ")";
                            ws.Cells[row, 27].Value = "";
                            ws.Cells[row, 28].Value = "";
                            ws.Cells[row, 29].Value = "";
                            ws.Cells[row, 30].Value = "";
                            ws.Cells[row, 31].Value = "";
                            ws.Cells[row, 32].Value = "";
                            ws.Cells[row, 33].Value = "";
                            ws.Cells[row, 34].Value = "";
                            ws.Cells[row, 35].Value = "";
                            ws.Cells[row, 36].Value = "";
                            ws.Cells[row, 37].Value = "";
                            ws.Cells[row, 38].Value = "";
                            ws.Cells[row, 38].Value = "";

                            ws.Cells[row, 1, row, 41].Style.Font.Bold = true;
                            ws.Cells[row, 1, row, 41].Style.Font.Size = 12;
                            
                            ws.Cells[row,1,row, 41].Style.Fill.PatternType = ExcelFillStyle.Solid;

                            ws.Cells[row, 1, row, 41].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        if (rowType == "Detail")
                        {
                            

                        }
                        if (ws.Cells[row, 1].Text != client && rowType == "Detail")
                        {
                            entity = "";
                            client = ws.Cells[row, 1].Text;
                            clientStartRow = row;
                        }
                        if (ws.Cells[row, 29].Text != entity && rowType == "Detail")
                        {
                            entity = ws.Cells[row, 29].Text;
                            clientEntityStartRow = row;
                        }

                        dataEndRow = row;
                    }

                    ExcelRange dataRange = ws.Cells[dataStartRow - 1, 1, dataEndRow, 10];
                    dataRange.AutoFilter = true;
                    ws.Cells["B:B"].Style.Numberformat.Format = "#,##0";
                    ws.Cells["E:AH"].Style.Numberformat.Format = "#,##0.00";

                    ws.Cells["F:F"].Style.Numberformat.Format = "#0%";
                    ws.Cells["K:K"].Style.Numberformat.Format = "yyyy-mm-dd";
                    ws.Cells["Q:Q"].Style.Numberformat.Format = "yyyy-mm-dd";
                    ws.Cells["S:S"].Style.Numberformat.Format = "#,##0";
                    ws.Cells["AA:AA"].Style.Numberformat.Format = "#,##0";
                    ws.Cells["AI:AI"].Style.Numberformat.Format = "yyyy-mm-dd";
                    ws.Cells["L:P"].Style.Numberformat.Format = "#0.00%";
                    ws.Cells["O:O"].Style.Numberformat.Format = "#0.0000%";
                    ws.Cells["AB:AB"].Style.Numberformat.Format = "#0.000%";
                    ws.Cells["AG:AG"].Style.Numberformat.Format = "#0.00%";

                    ws.Cells["A:AZ"].AutoFitColumns();

                    ws.Column(20).Width = 20;
                    ws.Column(21).Width = 20;
                    ws.Column(22).Width = 20;
                    ws.Column(23).Width = 20;
                    ws.Column(24).Width = 20;
                    
                    CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                    dialog.InitialDirectory = "C:\\Users";
                    dialog.IsFolderPicker = true;

                    string fileName;

                    fileName = "";

                    if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                    {
                        fileName = dialog.FileName + "\\" + "TrackingReport.xlsx";

                    }
                    System.IO.FileInfo tableFunding = new System.IO.FileInfo(fileName);

                    tableFunding.Delete();

                    pck.SaveAs(tableFunding);
                    System.Diagnostics.Process.Start(fileName);
                }
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }
        }

        private void AgingReportButton_Click(object sender, RoutedEventArgs e)
        {
            AgingReportWindow uwindow = new AgingReportWindow();
            uwindow.Show();
            //ErrorLabel.Content = "";
            //try
            //{
            //    DataSet report = RunsStoredProc.RunStoredProc("TableFunding_AgingReport", "", "", "", "", "", "", "", "", "", "", "", "", "", 0);
            //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //    using (ExcelPackage pck = new ExcelPackage())
            //    {
            //        ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Aging Report");

            //        ws.Cells["A4"].LoadFromDataTable(report.Tables[0], true);
            //        ws.Cells["A2"].Value = "Master Receivable Report";
            //        ws.Cells["A3"].Formula = "=Today()";

            //        ws.Cells["A3"].Style.Numberformat.Format = "yyyy-mm-dd";

            //        ws.Cells["E:F"].Style.Numberformat.Format = "#,##0.00";

            //        ws.Cells["G:G"].Style.Numberformat.Format = "yyyy-mm-dd";
            //        ws.Cells["H:H"].Style.Numberformat.Format = "#,##0";
            //        ws.Cells["I:I"].Style.Numberformat.Format = "#,##0.00";
            //        ws.Cells["J:J"].Style.Numberformat.Format = "#0.00%";

            //        ws.Cells["A2"].Style.Font.Size = 18;
            //        ws.Cells["A3"].Style.Font.Size = 18;

            //        ws.Cells["A4"].Style.Font.Size = 14;
            //        ws.Cells["B4"].Style.Font.Size = 14;
            //        ws.Cells["C4"].Style.Font.Size = 14;
            //        ws.Cells["D4"].Style.Font.Size = 14;
            //        ws.Cells["E4"].Style.Font.Size = 14;
            //        ws.Cells["F4"].Style.Font.Size = 14;
            //        ws.Cells["G4"].Style.Font.Size = 14;
            //        ws.Cells["H4"].Style.Font.Size = 14;
            //        ws.Cells["I4"].Style.Font.Size = 14;
            //        ws.Cells["J4"].Style.Font.Size = 14;


            //        ws.Cells["A4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        ws.Cells["B4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        ws.Cells["C4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        ws.Cells["D4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        ws.Cells["E4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        ws.Cells["F4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        ws.Cells["G4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        ws.Cells["H4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        ws.Cells["I4"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        ws.Cells["J4"].Style.Fill.PatternType = ExcelFillStyle.Solid;

            //        ws.Cells["A4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //        ws.Cells["B4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //        ws.Cells["C4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //        ws.Cells["D4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //        ws.Cells["E4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //        ws.Cells["F4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //        ws.Cells["G4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //        ws.Cells["H4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //        ws.Cells["I4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //        ws.Cells["J4"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

            //        ws.Cells["A4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //        ws.Cells["B4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //        ws.Cells["C4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //        ws.Cells["D4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //        ws.Cells["E4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //        ws.Cells["F4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //        ws.Cells["G4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //        ws.Cells["H4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //        ws.Cells["I4"].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //        ws.Cells["J4"].Style.Font.Color.SetColor(System.Drawing.Color.White);


            //        ws.Cells["K4"].Value = "";
            //        ws.Cells["L4"].Value = "";
            //        ws.Cells["M4"].Value = "";

            //        string client = "";
            //        string entity = "";
            //        int clientStartRow = 0;
            //        int clientEntityStartRow = 0;
            //        string rowType = "";
            //        //double totalOpenMortgageBalance = 0;
            //        //double totalOpenAdvanceBalance = 0;
            //        int dataStartRow = 5;
            //        int dataEndRow = 0;
            //        for (int row = dataStartRow; row <= 1000; row++)
            //        {
            //            if (ws.Cells[row, 1].Value == null)
            //            {
            //                break;
            //            }

            //            if (row == dataStartRow)
            //            {
            //                client = ws.Cells[row, 1].Text;
            //                entity = ws.Cells[row, 2].Text;
            //                clientStartRow = row;
            //                clientEntityStartRow = row;
            //            }

            //            rowType = ws.Cells[row, 11].Text;
            //            if (rowType == "Grand Total")
            //            {
            //                ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + dataStartRow + ":E" + (row - 1) + ")";
            //                ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + dataStartRow + ":F" + (row - 1) + ")";
            //                ws.Cells[row, 7].Value = "";
            //                ws.Cells[row, 9].Value = "";
            //                ws.Cells[row, 10].Value = "";
            //                ws.Cells[row, 11].Value = "";
            //                ws.Cells[row, 12].Value = "";
            //                ws.Cells[row, 13].Value = "";
            //                ws.Cells[row, 1].Style.Font.Bold = true;
            //                ws.Cells[row, 1].Style.Font.Size = 14;
            //                ws.Cells[row, 5].Style.Font.Size = 14;
            //                ws.Cells[row, 5].Style.Font.Bold = true;
            //                ws.Cells[row, 6].Style.Font.Bold = true;
            //                ws.Cells[row, 6].Style.Font.Size = 14;


            //                ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;


            //                ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //                ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //                ws.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //                ws.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //                ws.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //                ws.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //                ws.Cells[row, 7].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //                ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //                ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);
            //                ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

            //                ws.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //                ws.Cells[row, 2].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //                ws.Cells[row, 3].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //                ws.Cells[row, 4].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //                ws.Cells[row, 5].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //                ws.Cells[row, 6].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //                ws.Cells[row, 7].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //                ws.Cells[row, 8].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //                ws.Cells[row, 9].Style.Font.Color.SetColor(System.Drawing.Color.White);
            //                ws.Cells[row, 10].Style.Font.Color.SetColor(System.Drawing.Color.White);



            //            }
            //            if (rowType == "Client Total")
            //            {
            //                ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + clientStartRow + ":E" + (row - 1) + ")";
            //                ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";
            //                ws.Cells[row, 7].Value = "";
            //                ws.Cells[row, 9].Value = "";
            //                ws.Cells[row, 10].Value = "";
            //                ws.Cells[row, 11].Value = "";
            //                ws.Cells[row, 12].Value = "";
            //                ws.Cells[row, 13].Value = ""; ws.Cells[row, 1].Style.Font.Bold = true;
            //                ws.Cells[row, 1].Style.Font.Size = 13;
            //                ws.Cells[row, 1].Style.Font.Bold = true;
            //                ws.Cells[row, 5].Style.Font.Bold = true;
            //                ws.Cells[row, 5].Style.Font.Size = 13;
            //                ws.Cells[row, 6].Style.Font.Bold = true;
            //                ws.Cells[row, 6].Style.Font.Size = 13;

            //                ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;


            //                ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //                ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //                ws.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //                ws.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //                ws.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //                ws.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //                ws.Cells[row, 7].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //                ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //                ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //                ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //            }

            //            if (rowType == "Client Entity Total")
            //            {
            //                ws.Cells[row, 5].Formula = "=SUBTOTAL(9,E" + clientEntityStartRow + ":E" + (row - 1) + ")";
            //                ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + clientEntityStartRow + ":F" + (row - 1) + ")";
            //                ws.Cells[row, 7].Value = "";
            //                ws.Cells[row, 9].Value = "";
            //                ws.Cells[row, 10].Value = "";
            //                ws.Cells[row, 11].Value = "";
            //                ws.Cells[row, 12].Value = "";
            //                ws.Cells[row, 13].Value = ""; ws.Cells[row, 1].Style.Font.Bold = true;
            //                ws.Cells[row, 1].Style.Font.Size = 12;
            //                ws.Cells[row, 1].Style.Font.Bold = true;
            //                ws.Cells[row, 2].Style.Font.Size = 12;
            //                ws.Cells[row, 2].Style.Font.Bold = true;
            //                ws.Cells[row, 5].Style.Font.Bold = true;
            //                ws.Cells[row, 5].Style.Font.Size = 12;
            //                ws.Cells[row, 6].Style.Font.Bold = true;
            //                ws.Cells[row, 6].Style.Font.Size = 12;

            //                ws.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 9].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //                ws.Cells[row, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;


            //                ws.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            //                ws.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            //                ws.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            //                ws.Cells[row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            //                ws.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            //                ws.Cells[row, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            //                ws.Cells[row, 7].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            //                ws.Cells[row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            //                ws.Cells[row, 9].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            //                ws.Cells[row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

            //            }

            //            if (rowType == "Detail")
            //            {
            //                ws.Cells[row, 8].Formula = "=A3-G" + row.ToString();
            //                ws.Cells[row, 11].Value = "";
            //                ws.Cells[row, 12].Value = "";
            //                ws.Cells[row, 13].Value = "";

            //            }
            //            if (ws.Cells[row, 1].Text != client && rowType == "Detail")
            //            {
            //                entity = "";
            //                client = ws.Cells[row, 1].Text;
            //                clientStartRow = row;
            //            }
            //            if (ws.Cells[row, 2].Text != entity && rowType == "Detail")
            //            {
            //                entity = ws.Cells[row, 2].Text;
            //                clientEntityStartRow = row;
            //            }

            //            dataEndRow = row;
            //        }

            //        ExcelRange dataRange = ws.Cells[dataStartRow - 1, 1, dataEndRow, 10];
            //        dataRange.AutoFilter = true;

            //        ws.Cells["A:J"].AutoFitColumns();
            //        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //        CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            //        dialog.InitialDirectory = "C:\\Users";
            //        dialog.IsFolderPicker = true;

            //        string fileName;

            //        fileName = "";

            //        if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            //        {
            //            fileName = dialog.FileName + "\\" + "AgingReport.xlsx";

            //        }

            //        System.IO.FileInfo tableFunding = new System.IO.FileInfo(fileName);


            //        tableFunding.Delete();

            //        pck.SaveAs(tableFunding);

            //        System.Diagnostics.Process.Start(fileName);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    ErrorLabel.Content = ex.Message;
            //    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
            //}
        }

        private void CollectionsReportButton_Click(object sender, RoutedEventArgs e)
        {
            CollectionsReport uwindow = new CollectionsReport();
            uwindow.Show();
         
        }

        private void SalesReportButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            SalesReport uwindow = new SalesReport();
            uwindow.Show();

            //try
            //{
            //    DataSet report = RunsStoredProc.RunStoredProc("TableFunding_SalesReport", "", "", "", "", "", "", "", "", "", "", "", "", "", 0);
            //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //    //tableFunding.Delete();
            //    using (ExcelPackage pck = new ExcelPackage())
            //    {
            //        ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sales Report");

            //        ws.Cells["A4"].LoadFromDataTable(report.Tables[0], true);
            //        ws.Cells["A2"].Value = "Sales Report";
            //        ws.Cells["A3"].Formula = "=Today()";

            //        ws.Cells["A3"].Style.Numberformat.Format = "yyyy-mm-dd";

            //        ws.Cells["F:G"].Style.Numberformat.Format = "#,##0.00";

            //        ws.Cells["H:H"].Style.Numberformat.Format = "yyyy-mm-dd";

            //        ws.Cells["A2"].Style.Font.Size = 18;
            //        ws.Cells["A3"].Style.Font.Size = 18;

            //        ws.Cells[4, 1, 4, 38].Style.Font.Size = 14;


            //        ws.Cells[4, 1, 4, 38].Style.Fill.PatternType = ExcelFillStyle.Solid;

            //        ws.Cells[4, 1, 4, 38].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

            //        ws.Cells[4, 1, 4, 38].Style.Font.Color.SetColor(System.Drawing.Color.White);


            //        string client = "";
            //        string investor = "";
            //        int clientStartRow = 0;
            //        int clientInvestorStartRow = 0;
            //        string rowType = "";
            //        //double totalOpenMortgageBalance = 0;
            //        //double totalOpenAdvanceBalance = 0;
            //        int dataStartRow = 5;
            //        int dataEndRow = 0;
            //        for (int row = dataStartRow; row <= 1000; row++)
            //        {
            //            if (ws.Cells[row, 1].Value == null)
            //            {
            //                break;
            //            }

            //            if (row == dataStartRow)
            //            {
            //                client = ws.Cells[row, 1].Text;
            //                investor = ws.Cells[row, 3].Text;
            //                clientStartRow = row;
            //                clientInvestorStartRow = row;
            //            }

            //            rowType = ws.Cells[row, 11].Text;
            //            if (rowType == "Grand Total")
            //            {
            //                ws.Cells[row, 4].Value = "";
            //                ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + dataStartRow + ":F" + (row - 1) + ")";
            //                ws.Cells[row, 7].Formula = "=SUBTOTAL(9,G" + dataStartRow + ":G" + (row - 1) + ")";

            //                ws.Cells[row, 8].Value = "";
            //                ws.Cells[row, 1, row, 8].Style.Font.Bold = true;
            //                ws.Cells[row, 1, row, 8].Style.Font.Size = 14;

            //                ws.Cells[row, 1, row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;

            //                ws.Cells[row, 1, row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkGray);

            //                ws.Cells[row, 1, row, 8].Style.Font.Color.SetColor(System.Drawing.Color.White);

            //            }
                      
            //            if (rowType == "Client Total")
            //            {
            //                ws.Cells[row, 4].Value = "";
            //                ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";
            //                ws.Cells[row, 7].Formula = "=SUBTOTAL(9,G" + clientStartRow + ":G" + (row - 1) + ")";

            //                ws.Cells[row, 8].Value = "";
            //                ws.Cells[row, 1, row, 8].Style.Font.Bold = true;
            //                ws.Cells[row, 1, row, 8].Style.Font.Size = 12;

            //                ws.Cells[row, 1, row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;

            //                ws.Cells[row, 1, row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            //            }

            //            if (rowType == "Client Investor Total")
            //            {
            //                ws.Cells[row, 4].Value = "";
            //                ws.Cells[row, 6].Formula = "=SUBTOTAL(9,F" + clientInvestorStartRow + ":F" + (row - 1) + ")";
            //                ws.Cells[row, 7].Formula = "=SUBTOTAL(9,G" + clientInvestorStartRow + ":G" + (row - 1) + ")";

            //                ws.Cells[row, 8].Value = "";
            //                ws.Cells[row, 1, row, 8].Style.Font.Bold = true;
            //                ws.Cells[row, 1, row, 8].Style.Font.Size = 13;

            //                ws.Cells[row, 1, row, 8].Style.Fill.PatternType = ExcelFillStyle.Solid;

            //                ws.Cells[row, 1, row, 8].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
            //            }
            //            if (rowType == "Detail")
            //            {
                           

            //            }
            //            if (ws.Cells[row, 1].Text != client && rowType == "Detail")
            //            {
            //                investor = "";
            //                client = ws.Cells[row, 1].Text;
            //                clientStartRow = row;
            //            }
            //            if (ws.Cells[row, 3].Text != investor && rowType == "Detail")
            //            {
            //                investor = ws.Cells[row, 3].Text;
            //                clientInvestorStartRow = row;
            //            }

            //            dataEndRow = row;
            //        }

            //        ExcelRange dataRange = ws.Cells[dataStartRow - 1, 1, dataEndRow, 10];
            //        dataRange.AutoFilter = true;

            //        ws.Cells["A:M"].AutoFitColumns();
            //        ws.Column(6).Width = 20;
            //        ws.Column(7).Width = 20;

            //        CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            //        dialog.InitialDirectory = "C:\\Users";
            //        dialog.IsFolderPicker = true;

            //        string fileName;

            //        fileName = "";

            //        if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            //        {
            //            fileName = dialog.FileName + "\\" + "SalesReport.xlsx";

            //        }



            //        //System.IO.FileInfo tableFunding = new System.IO.FileInfo(MyVariables.excelReportFilePath + "SalesReport.xlsx");

            //        System.IO.FileInfo tableFunding = new System.IO.FileInfo(fileName);


            //        tableFunding.Delete();

            //        pck.SaveAs(tableFunding);
            //        System.Diagnostics.Process.Start(fileName);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    ErrorLabel.Content = ex.Message;
            //    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
            //}
        }

        private void LoansDataGrid_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            { 
                TableFundingLoan loan = (TableFundingLoan)e.Row.Item;
                loan.LoanID = -1;
                loan.Save();
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }       

        }


        private void BaileeButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";

            try
            {
                TableFundingLoans loans = new TableFundingLoans();

                foreach (TableFundingLoan loan in LoansDataGrid.ItemsSource)
                {
                    if (loan.IsChecked)
                    {
                        loan.LoanUW = new TableFundingLoanUW(loan.LoanID);
                        loan.LoanFunding = new TableFundingLoanFunding(loan.LoanID);
                        loans.Add(loan);

                    }
                }
                if (loans.Count() == 0)
                {
                    ErrorLabel.Content = "Please select at least one loan";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }

                TableFundingClient client = new TableFundingClient(loans[0].LoanClientID);
                //TableFundingClientEntity entity = new TableFundingClientEntity(loans[0].LoanClientID, loans[0].LoanEntityID);
                TableFundingEntity entity = new TableFundingEntity(loans[0].LoanEntityID);

                Spire.Doc.Document document = new Spire.Doc.Document();
                //Add Section

                Spire.Doc.Section section1 = document.AddSection();

                section1.PageSetup.Orientation = Spire.Doc.Documents.PageOrientation.Portrait;

                //Add Paragraph

                Spire.Doc.Documents.Paragraph parExhibitA = section1.AddParagraph();


                parExhibitA.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                //Add Paragraph

                Spire.Doc.Documents.Paragraph parSchedofInv = section1.AddParagraph();
                parSchedofInv.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                Spire.Doc.Documents.Paragraph parDate = section1.AddParagraph();
                parDate.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                Spire.Doc.Table tblInv = section1.AddTable(true);
                Spire.Doc.Fields.TextRange trExhibitA = parExhibitA.AppendText("EXHIBIT A\v");
                trExhibitA.CharacterFormat.UnderlineStyle = Spire.Doc.Documents.UnderlineStyle.Single;


                parSchedofInv.AppendText("SCHEDULE OF ASSETS\v");

                parDate.AppendText("Date:" + DateTime.Now.ToShortDateString() + "\v");

                //Create Header and Data

                String[] Header = { "Loan ID", "Borrower", "Business Name", "Address" };

                //
                // Table Logic
                //
                //Add Cells
                tblInv.ResetCells(loans.Count() + 1, Header.Length);
                //Header Row
                Spire.Doc.TableRow FRow = tblInv.Rows[0];
                FRow.IsHeader = true;

                //Row Height
                FRow.Height = 23;

                //Header Format
                //FRow.RowFormat.BackColor = Color.AliceBlue;

                for (int i = 0; i < Header.Length; i++)
                {
                    //Cell Alignment
                    Spire.Doc.Documents.Paragraph p = FRow.Cells[i].AddParagraph();
                    FRow.Cells[i].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    //Data Format
                    Spire.Doc.Fields.TextRange TR = p.AppendText(Header[i]);
                    //TR.CharacterFormat.FontName = "Calibri";
                    TR.CharacterFormat.FontSize = 14;
                    //TR.CharacterFormat.TextColor = Color.Teal;                
                    TR.CharacterFormat.Bold = true;

                }

                //Data Row

                string loanNumberList = "";
                int r = 0;
                //for (int r = 0; r < data.Length; r++)
                foreach (TableFundingLoan loan in loans)
                {

                    Spire.Doc.TableRow DataRow = tblInv.Rows[r + 1];

                    //Row Height

                    DataRow.Height = 20;

                    //C Represents Column.

                    //for (int c = 0; c < data[r].Length; c++)                    
                    //{                    
                    //Cell Alignment
                    DataRow.Cells[0].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[1].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[2].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[3].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    //Fill Data in Rows

                    loanNumberList += loan.LoanNumber + " "; ;

                    Spire.Doc.Documents.Paragraph pLoanNumber = DataRow.Cells[0].AddParagraph();
                    Spire.Doc.Fields.TextRange TRLoanNumber = pLoanNumber.AppendText(loan.LoanNumber);

                    Spire.Doc.Documents.Paragraph pBorrower = DataRow.Cells[1].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBorrower = pBorrower.AppendText(loan.LoanMortgagee);

                    Spire.Doc.Documents.Paragraph pBusiness = DataRow.Cells[2].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBusiness = pBusiness.AppendText(loan.LoanMortgageeBusiness);

                    Spire.Doc.Documents.Paragraph pAddress = DataRow.Cells[3].AddParagraph();
                    Spire.Doc.Fields.TextRange TRAddress = pAddress.AppendText(loan.LoanPropertyAddress);

                    //Format Cells

                    pLoanNumber.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pBorrower.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pAddress.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                    //TRLoanNumber.CharacterFormat.FontName = "Calibri";
                    //TRBorrower.CharacterFormat.FontName = "Calibri";
                    //TRAddress.CharacterFormat.FontName = "Calibri";

                    TRLoanNumber.CharacterFormat.FontSize = 12;
                    TRBorrower.CharacterFormat.FontSize = 12;
                    TRAddress.CharacterFormat.FontSize = 12;

                    //TR2.CharacterFormat.TextColor = Color.Brown;

                    //}
                    r++;

                }
                //
                // End Table Logic
                //
                Spire.Doc.Documents.Paragraph parWireInfo = section1.AddParagraph();
                parWireInfo.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

                if (loans[0].LoanFunding.BaileeLetterDate == null)
                {
                    ErrorLabel.Content = "Bailee Letter Date Missing";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }
                if (loans[0].LoanFunding.ClosingDate == null)
                {
                    ErrorLabel.Content = "Closing Date Missing";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }
                DateTime baileeLetterDate = (DateTime)loans[0].LoanFunding.BaileeLetterDate;
                DateTime closingDate = (DateTime)loans[0].LoanFunding.ClosingDate;

                parWireInfo.AppendText("Purchase Advice Date:" + baileeLetterDate.ToShortDateString() + "\v\v");
                parWireInfo.AppendText("Closing Date:" + closingDate.ToShortDateString() + "\v\v");
                parWireInfo.AppendText("Wiring Instructions:\v");

                Spire.Doc.Documents.Paragraph parBankInfo = section1.AddParagraph();
                parBankInfo.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                parBankInfo.Format.LeftIndent = 80;

                parBankInfo.AppendText("Bank: " + entity.EntityBank + "\v");
                parBankInfo.AppendText("ABA: " + entity.Aba + "\v");
                parBankInfo.AppendText("Account: " + entity.Account + "\v");
                parBankInfo.AppendText("Account Name: " + entity.EntityName + "\v");
                parBankInfo.AppendText("Ref: " + client.ClientName + "\v");
                Spire.Doc.Break pageBreak = new Spire.Doc.Break(document, Spire.Doc.Documents.BreakType.PageBreak);

                string baileeFileName;
                baileeFileName = "Bailee_" + client.ClientName + "_" + entity.EntityName + ".docx";


                Spire.Doc.Document DocOne = new Spire.Doc.Document();
                OpenFileDialog fileDialog = new OpenFileDialog();
                string baileeletterFileName = "";

                if (fileDialog.ShowDialog() == true)
                {
                    baileeletterFileName = fileDialog.FileName;
                }

                DocOne.LoadFromFile(baileeletterFileName, FileFormat.Docx);
                // document.SaveToFile(MyVariables.letterFilePath + "result.docx", FileFormat.Docx);

                DateTime today = DateTime.Today;
                string todayStr;
                todayStr = string.Format("{0:MMMM dd, yyyy}", today);

                DocOne.Replace("TODAYSDATE", todayStr, false, true);

                foreach (Spire.Doc.Section sec in document.Sections)
                {
                    DocOne.Sections.Add(sec.Clone());
                }
                string fileName;
                string todayStrFile = string.Format("{0:yyyy_MM_dd}", today);

                fileName = "Bailee_" + client.ClientName + "_" + entity.EntityName + "_" + loanNumberList + ".docx";
                CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                dialog.InitialDirectory = "C:\\Users";
                dialog.IsFolderPicker = true;

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    fileName = dialog.FileName + "\\" + fileName;

                }

                DocOne.SaveToFile(fileName, FileFormat.Docx);
                
                System.Diagnostics.Process.Start(fileName);
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }

        }
        private void ReleaseButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";

            try
            {
                TableFundingLoans loans = new TableFundingLoans();

                foreach (TableFundingLoan loan in LoansDataGrid.ItemsSource)
                {
                    if (loan.IsChecked)
                    {
                        loan.LoanUW = new TableFundingLoanUW(loan.LoanID);
                        loan.LoanFunding = new TableFundingLoanFunding(loan.LoanID);
                        loans.Add(loan);
                        if (loan.LoanStatus != "4 - Completed")
                        {
                            ErrorLabel.Content = "Loan Not in Completed Status";
                            ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                            return;
                        }
                    }
                }
                if (loans.Count() == 0)
                {
                    ErrorLabel.Content = "Please select at least one loan";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }

                TableFundingClient client = new TableFundingClient(loans[0].LoanClientID);
                //TableFundingClientEntity entity = new TableFundingClientEntity(loans[0].LoanClientID, loans[0].LoanEntityID);
                TableFundingEntity entity = new TableFundingEntity(loans[0].LoanEntityID);

                if (loans[0].LoanFunding.BaileeLetterDate == null)
                {
                    ErrorLabel.Content = "Bailee Letter Date Missing";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }

                DateTime baileeLetterDate = (DateTime)loans[0].LoanFunding.BaileeLetterDate;

                string baileeLetterDateStr = string.Format("{0:MMMM dd, yyyy}", baileeLetterDate);
                Spire.Doc.Document document = new Spire.Doc.Document();
                //Add Section

                Spire.Doc.Section section1 = document.AddSection();
                section1.PageSetup.Orientation = Spire.Doc.Documents.PageOrientation.Portrait;

                //Add Paragraph

                Spire.Doc.Documents.Paragraph parSchedule1 = section1.AddParagraph();
                parSchedule1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                Spire.Doc.Fields.TextRange trSchedule1 = parSchedule1.AppendText("Schedule 1\v");
                trSchedule1.CharacterFormat.UnderlineStyle = Spire.Doc.Documents.UnderlineStyle.Single;

                Spire.Doc.Documents.Paragraph parMort = section1.AddParagraph();
                parMort.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                parMort.AppendText("Mortgage Loan Schedule\v");

                Spire.Doc.Table tblSch = section1.AddTable(true);
                //
                // Table Logic
                //

                String[] Header = { "Loan ID", "Borrower", "Address" };

                //Add Cells
                tblSch.ResetCells(loans.Count() + 1, Header.Length);
                //Header Row
                Spire.Doc.TableRow FRow2 = tblSch.Rows[0];
                FRow2.IsHeader = true;

                //Row Height
                FRow2.Height = 23;

                //Header Format
                //FRow.RowFormat.BackColor = Color.AliceBlue;

                for (int i = 0; i < Header.Length; i++)
                {
                    //Cell Alignment
                    Spire.Doc.Documents.Paragraph p = FRow2.Cells[i].AddParagraph();
                    FRow2.Cells[i].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    //Data Format
                    Spire.Doc.Fields.TextRange TR = p.AppendText(Header[i]);
                    //TR.CharacterFormat.FontName = "Calibri";
                    TR.CharacterFormat.FontSize = 14;
                    //TR.CharacterFormat.TextColor = Color.Teal;                
                    TR.CharacterFormat.Bold = true;

                }

                //Data Row
                string loanNumberList = "";
                int r = 0;
                //for (int r = 0; r < data.Length; r++)
                foreach (TableFundingLoan loan in loans)
                {

                    Spire.Doc.TableRow DataRow = tblSch.Rows[r + 1];

                    //Row Height

                    DataRow.Height = 20;

                    //C Represents Column.

                    //for (int c = 0; c < data[r].Length; c++)                    
                    //{                    
                    //Cell Alignment
                    DataRow.Cells[0].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[1].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[2].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    //Fill Data in Rows
                    loanNumberList += loan.LoanNumber + " ";

                    Spire.Doc.Documents.Paragraph pLoanNumber = DataRow.Cells[0].AddParagraph();
                    Spire.Doc.Fields.TextRange TRLoanNumber = pLoanNumber.AppendText(loan.LoanNumber);

                    Spire.Doc.Documents.Paragraph pBorrower = DataRow.Cells[1].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBorrower = pBorrower.AppendText(loan.LoanMortgagee);

                    Spire.Doc.Documents.Paragraph pAddress = DataRow.Cells[2].AddParagraph();
                    Spire.Doc.Fields.TextRange TRAddress = pAddress.AppendText(loan.LoanPropertyAddress);

                    //Format Cells

                    pLoanNumber.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pBorrower.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pAddress.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                    //TRLoanNumber.CharacterFormat.FontName = "Calibri";
                    //TRBorrower.CharacterFormat.FontName = "Calibri";
                    //TRAddress.CharacterFormat.FontName = "Calibri";

                    TRLoanNumber.CharacterFormat.FontSize = 12;
                    TRBorrower.CharacterFormat.FontSize = 12;
                    TRAddress.CharacterFormat.FontSize = 12;

                    //TR2.CharacterFormat.TextColor = Color.Brown;

                    //}
                    r++;
                }
                //
                // End Table Logic
                //

                //document.SaveToFile(MyVariables.letterFilePath + "result.docx", FileFormat.Docx);

                Spire.Doc.Document DocOne = new Spire.Doc.Document();

                OpenFileDialog fileDialog = new OpenFileDialog();
                string releaseFileName = "";
                //releaseFileName = "Release_" + client.ClientName + "_" + entity.EntityName + ".docx";

                if (fileDialog.ShowDialog() == true)
                {
                    releaseFileName = fileDialog.FileName;
                }

                DocOne.LoadFromFile(releaseFileName, FileFormat.Docx);

                DateTime today = DateTime.Today;
                string todayStr;
                todayStr = string.Format("{0:MMMM dd, yyyy}", today);

                DocOne.Replace("TODAYSDATE", todayStr, false, true);

                DocOne.Replace("BAILEELETTERDATE", baileeLetterDateStr, false, true);

                foreach (Spire.Doc.Section sec in document.Sections)
                {
                    DocOne.Sections.Add(sec.Clone());
                }
                string fileName;
                string todayStrFile = string.Format("{0:yyyy_MM_dd}", today);

                fileName = "Release_" + client.ClientName + "_" + entity.EntityName + "_" + loanNumberList + ".docx";


                CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                dialog.InitialDirectory = "C:\\Users";
                dialog.IsFolderPicker = true;

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    fileName = dialog.FileName + "\\" + fileName;

                }
                DocOne.SaveToFile(fileName, FileFormat.Docx);

                System.Diagnostics.Process.Start(fileName);
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }


        }
        private void AdvanceButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";

            try
            {
                TableFundingLoans loans = new TableFundingLoans();

                foreach (TableFundingLoan loan in LoansDataGrid.ItemsSource)
                {
                    if (loan.IsChecked)
                    {
                        //Get a refreshed loan because status might have changed
                        TableFundingLoan loanNew = new TableFundingLoan(loan.LoanID);
                        loan.LoanUW = new TableFundingLoanUW(loan.LoanID);
                        loan.LoanFunding = new TableFundingLoanFunding(loan.LoanID);
                        loans.Add(loan);
                        if (loanNew.LoanStatus != "3 - Open")
                        {
                            ErrorLabel.Content = "Loan Not in Open Status";
                            ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                            return;
                        }
                    }
                }
                if (loans.Count() == 0)
                {
                    ErrorLabel.Content = "Please select at least one loan";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }

                TableFundingClient client = new TableFundingClient(loans[0].LoanClientID);
                //TableFundingClientEntity entity = new TableFundingClientEntity(loans[0].LoanClientID, loans[0].LoanEntityID);
                TableFundingEntity entity = new TableFundingEntity(loans[0].LoanEntityID);

                Spire.Doc.Document document = new Spire.Doc.Document();
                //Add Section

                Spire.Doc.Section section1 = document.AddSection();

                section1.PageSetup.Orientation = Spire.Doc.Documents.PageOrientation.Landscape;
                //Add Paragraph

                Spire.Doc.Documents.Paragraph parSchedule1 = section1.AddParagraph();
                parSchedule1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                Spire.Doc.Fields.TextRange trSchedule1 = parSchedule1.AppendText(entity.EntityName);
                //trSchedule1.CharacterFormat.UnderlineStyle = Spire.Doc.Documents.UnderlineStyle.Single;

                Spire.Doc.Documents.Paragraph parMort = section1.AddParagraph();
                parMort.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                parMort.AppendText("Advance Report");

                Spire.Doc.Documents.Paragraph parClient = section1.AddParagraph();
                parClient.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                parClient.AppendText(client.ClientName);

                Spire.Doc.Documents.Paragraph parRemitReport = section1.AddParagraph();
                parRemitReport.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                parRemitReport.AppendText("Advance Date " + DateTime.Today.ToShortDateString() + "\v");

                Spire.Doc.Table tblSch = section1.AddTable(true);
                //
                // Table Logic
                //

                String[] Header = {"Business Name","Client Name", "Loan Number", "Collateral Description", "Mortgage Amount", "Reserve Amount", "Total Transfer" };

                //Add Cells
                // Number of rows is # of loans + header + footer + line for wire fee
                tblSch.ResetCells(loans.Count() + 3, Header.Length);
                int footerRow;
                int wireRow;

                footerRow = loans.Count() + 2;
                wireRow = loans.Count() + 1;
                //Header Row
                Spire.Doc.TableRow FRow2 = tblSch.Rows[0];
                FRow2.IsHeader = true;

                //Row Height
                FRow2.Height = 30;

                //Wire Fee Row
                Spire.Doc.TableRow FRowWire = tblSch.Rows[wireRow];

                FRowWire.Height = 30;
                Spire.Doc.Documents.Paragraph pWire = FRowWire.Cells[0].AddParagraph();
                Spire.Doc.Fields.TextRange TRWire = pWire.AppendText("Wire Transfer Fee If Applicable");
                pWire.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

                Spire.Doc.Documents.Paragraph pWireFee = FRowWire.Cells[6].AddParagraph();
                Spire.Doc.Fields.TextRange TRWireFee = pWireFee.AppendText("(25.00)");
                pWireFee.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

                //Merge first 4 cells
                tblSch.ApplyHorizontalMerge((wireRow), 0, 4);

                Spire.Doc.TableRow FRowFooter = tblSch.Rows[footerRow];

                //Merge first 3 cells
                tblSch.ApplyHorizontalMerge((footerRow), 0, 2);

                Spire.Doc.Documents.Paragraph pFooter = FRowFooter.Cells[0].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooter = pFooter.AppendText("Total Due " + client.ClientName);
                pFooter.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                
                //Row Height
                FRowFooter.Height = 23;

                //Header Format
                //FRow.RowFormat.BackColor = Color.AliceBlue;

                for (int i = 0; i < Header.Length; i++)
                {
                    //Cell Alignment
                    Spire.Doc.Documents.Paragraph p = FRow2.Cells[i].AddParagraph();
                    FRow2.Cells[i].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    //Data Format
                    Spire.Doc.Fields.TextRange TR = p.AppendText(Header[i]);
                    //TR.CharacterFormat.FontName = "Calibri";
                    TR.CharacterFormat.FontSize = 14;
                    //TR.CharacterFormat.TextColor = Color.Teal;                
                    TR.CharacterFormat.Bold = true;

                }

                //Data Row

                double totalMortgage = 0;
                double totalAdvance = -25;  // -25 for wire fee
                double totalReserve = 0;

                int r = 0;
                //for (int r = 0; r < data.Length; r++)
                foreach (TableFundingLoan loan in loans)
                {

                    TableFundingLoan loanNew = new TableFundingLoan(loan.LoanID);
                    Spire.Doc.TableRow DataRow = tblSch.Rows[r + 1];

                    //Row Height

                    DataRow.Height = 30;

                    //C Represents Column.

                    //for (int c = 0; c < data[r].Length; c++)                    
                    //{                    
                    //Cell Alignment
                    DataRow.Cells[0].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    DataRow.Cells[1].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[2].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[3].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[4].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[5].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[6].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;



                    //Fill Data in Rows

                    Spire.Doc.Documents.Paragraph pBusiness = DataRow.Cells[0].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBusiness = pBusiness.AppendText(loanNew.LoanMortgageeBusiness);

                    Spire.Doc.Documents.Paragraph pBorrower = DataRow.Cells[1].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBorrower = pBorrower.AppendText(loanNew.LoanMortgagee);

                    Spire.Doc.Documents.Paragraph pLoanNumber = DataRow.Cells[2].AddParagraph();
                    Spire.Doc.Fields.TextRange TRLoanNumber = pLoanNumber.AppendText(loanNew.LoanNumber);

                    Spire.Doc.Documents.Paragraph pAddress = DataRow.Cells[3].AddParagraph();
                    Spire.Doc.Fields.TextRange TRAddress = pAddress.AppendText(loanNew.LoanPropertyAddress);

                    Spire.Doc.Documents.Paragraph pMortgageAmount = DataRow.Cells[4].AddParagraph();
                    Spire.Doc.Fields.TextRange trMortgageAmount = pMortgageAmount.AppendText(FormatNumberCommas2dec(loanNew.LoanMortgageAmount.ToString()));

                    Spire.Doc.Documents.Paragraph pReserveAmount = DataRow.Cells[5].AddParagraph();
                    Spire.Doc.Fields.TextRange trReserveAmount = pReserveAmount.AppendText(FormatNumberCommas2dec(loanNew.LoanReserveAmount.ToString()));

                    Spire.Doc.Documents.Paragraph pAdvanceAmount = DataRow.Cells[6].AddParagraph();
                    Spire.Doc.Fields.TextRange trAdvanceAmount = pAdvanceAmount.AppendText(FormatNumberCommas2dec(loanNew.LoanAdvanceAmount.ToString()));

                    totalMortgage += loanNew.LoanMortgageAmount;
                    totalAdvance += loanNew.LoanAdvanceAmount;
                    totalReserve += loanNew.LoanReserveAmount;

                    //Format Cells

                    pLoanNumber.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pBorrower.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pBusiness.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pAddress.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pMortgageAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pReserveAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pAdvanceAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    //TRLoanNumber.CharacterFormat.FontName = "Calibri";
                    //TRBorrower.CharacterFormat.FontName = "Calibri";
                    //TRAddress.CharacterFormat.FontName = "Calibri";

                    TRLoanNumber.CharacterFormat.FontSize = 12;
                    TRBorrower.CharacterFormat.FontSize = 12;
                    TRAddress.CharacterFormat.FontSize = 12;

                    TRAddress.CharacterFormat.FontSize = 12;
                    trMortgageAmount.CharacterFormat.FontSize = 12;
                    trReserveAmount.CharacterFormat.FontSize = 12;
                    trAdvanceAmount.CharacterFormat.FontSize = 12;
                    //TR2.CharacterFormat.TextColor = Color.Brown;

                    //}
                    r++;
                }

                Spire.Doc.TableRow FRowFooterMortgage = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterMortgage = FRowFooterMortgage.Cells[4].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterMortgage = pFooterMortgage.AppendText(FormatNumberCommas2dec(totalMortgage.ToString()));
                pFooterMortgage.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterMortgage.CharacterFormat.Bold = true;

                Spire.Doc.TableRow FRowFooterReserve = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterReserve = FRowFooterReserve.Cells[5].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterReserve = pFooterReserve.AppendText(FormatNumberCommas2dec(totalReserve.ToString()));
                pFooterReserve.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterReserve.CharacterFormat.Bold = true;

                Spire.Doc.TableRow FRowFooterAdvance = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterAdvance = FRowFooterAdvance.Cells[6].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterAdvance = pFooterAdvance.AppendText(FormatNumberCommas2dec(totalAdvance.ToString()));
                pFooterAdvance.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterAdvance.CharacterFormat.Bold = true;

                //
                // End Table Logic
                //

                Spire.Doc.Documents.Paragraph parEscrow = section1.AddParagraph();
                parEscrow.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

                parEscrow.AppendText("Wire Sent to LaRocca Escrow Account");

                //document.SaveToFile("C:/Temp/result.docx", FileFormat.Docx);

                //Spire.Doc.Document DocOne = new Spire.Doc.Document();

                //DocOne.LoadFromFile("C:/Temp/Release Text.docx", FileFormat.Docx);

                //DocOne.Replace("TODAYSDATE", todayStr, false, true);

                //DocOne.Replace("BAILEELETTERDATE", baileeLetterDateStr, false, true);

                //foreach (Spire.Doc.Section sec in document.Sections)
                //{
                //    DocOne.Sections.Add(sec.Clone());
                //}
                DateTime today = DateTime.Today;
                string todayStr;
                todayStr = string.Format("{0:MMMM dd, yyyy}", today);

                string fileName;
                string todayStrFile = string.Format("{0:yyyy_MM_dd}", today);

                fileName = "Advance_" + client.ClientName + "_" + entity.EntityName + "_" + todayStrFile + ".docx";
                CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                dialog.InitialDirectory = "C:\\Users";
                dialog.IsFolderPicker = true;

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    fileName = dialog.FileName + "\\" + fileName;

                }

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
        private void RemittanceButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";

            try
            {
                TableFundingLoans loans = new TableFundingLoans();

                foreach (TableFundingLoan loan in LoansDataGrid.ItemsSource)
                {
                    if (loan.IsChecked)
                    {
                        loan.LoanUW = new TableFundingLoanUW(loan.LoanID);
                        loan.LoanFunding = new TableFundingLoanFunding(loan.LoanID);
                        loan.LoanFees = new TableFundingFees(loan.LoanID);
                        if (loan.LoanStatus != "4 - Completed")
                        {
                            ErrorLabel.Content = "Loan Not in Completed Status";
                            ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                            return;
                        }
                        loans.Add(loan);

                    }
                }
                if (loans.Count() == 0)
                {
                    ErrorLabel.Content = "Please select at least one loan";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }

                TableFundingClient client = new TableFundingClient(loans[0].LoanClientID);
                //TableFundingClientEntity entity = new TableFundingClientEntity(loans[0].LoanClientID, loans[0].LoanEntityID);
                TableFundingEntity entity = new TableFundingEntity(loans[0].LoanEntityID);

                if (loans[0].LoanFunding.BaileeLetterDate == null)
                {
                    ErrorLabel.Content = "Bailee Letter Date Missing";
                    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    return;
                }
                DateTime baileeLetterDate = (DateTime)loans[0].LoanFunding.BaileeLetterDate;

                string baileeLetterDateStr = string.Format("{0:MMMM dd, yyyy}", baileeLetterDate);
                Spire.Doc.Document document = new Spire.Doc.Document();
                //Add Section

                Spire.Doc.Section section1 = document.AddSection();

                section1.PageSetup.Orientation = Spire.Doc.Documents.PageOrientation.Landscape;
                //Add Paragraph

                Spire.Doc.Documents.Paragraph parSchedule1 = section1.AddParagraph();
                parSchedule1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                Spire.Doc.Fields.TextRange trSchedule1 = parSchedule1.AppendText(entity.EntityName);
                //trSchedule1.CharacterFormat.UnderlineStyle = Spire.Doc.Documents.UnderlineStyle.Single;

                Spire.Doc.Documents.Paragraph parMort = section1.AddParagraph();
                parMort.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                parMort.AppendText("Remittance Report");

                Spire.Doc.Documents.Paragraph parClient = section1.AddParagraph();
                parClient.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                parClient.AppendText(client.ClientName);

                Spire.Doc.Documents.Paragraph parRemitReport = section1.AddParagraph();
                parRemitReport.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                parRemitReport.AppendText("Remittance Date " + DateTime.Today.ToShortDateString() + "\v");

                Spire.Doc.Table tblSch = section1.AddTable(true);
                //
                // Table Logic
                //

                String[] Header = { "Business Name", "Client Name", "Loan Number", "Interest Percentage", "Mortgage Amount", "Proceeds", "Loan Amount", "Interest", "Origination Fee", "Underwriting/ Administrative Fee", "Total Transfer" };

                //Add Cells
                // Number of rows is # of loans + header + footer + line for wire fee
                tblSch.ResetCells(loans.Count() + 3, Header.Length);
                int footerRow;
                int wireRow;

                footerRow = loans.Count() + 2;
                wireRow = loans.Count() + 1;
                //Header Row
                Spire.Doc.TableRow FRow2 = tblSch.Rows[0];
                FRow2.IsHeader = true;

                //Row Height
                FRow2.Height = 23;

                //Wire Fee Row
                Spire.Doc.TableRow FRowWire = tblSch.Rows[wireRow];

                Spire.Doc.Documents.Paragraph pWire = FRowWire.Cells[0].AddParagraph();
                Spire.Doc.Fields.TextRange TRWire = pWire.AppendText("Wire Transfer Fee If Applicable");
                pWire.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                TRWire.CharacterFormat.FontSize = 12;

                Spire.Doc.Documents.Paragraph pWireFee = FRowWire.Cells[10].AddParagraph();
                Spire.Doc.Fields.TextRange TRWireFee = pWireFee.AppendText("(25.00)");
                pWireFee.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRWireFee.CharacterFormat.FontSize = 12;

                //Merge first 4 cells
                tblSch.ApplyHorizontalMerge((wireRow), 0, 8);

                Spire.Doc.TableRow FRowFooter = tblSch.Rows[footerRow];

                //Merge first 3 cells
                tblSch.ApplyHorizontalMerge((footerRow), 0, 2);

                Spire.Doc.Documents.Paragraph pFooter = FRowFooter.Cells[0].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooter = pFooter.AppendText("Total Transfer Due " + client.ClientName);
                pFooter.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                TRFooter.CharacterFormat.FontSize = 12;

                //Row Height
                FRowFooter.Height = 23;

                //Header Format
                //FRow.RowFormat.BackColor = Color.AliceBlue;

                for (int i = 0; i < Header.Length; i++)
                {
                    //Cell Alignment
                    Spire.Doc.Documents.Paragraph p = FRow2.Cells[i].AddParagraph();
                    FRow2.Cells[i].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    //Data Format
                    Spire.Doc.Fields.TextRange TR = p.AppendText(Header[i]);
                    //TR.CharacterFormat.FontName = "Calibri";
                    TR.CharacterFormat.FontSize = 12;
                    //TR.CharacterFormat.TextColor = Color.Teal;                
                    TR.CharacterFormat.Bold = true;

                }

                //Data Row

                double totalMortgage = 0;
                double totalProceeds = 0;
                double totalLoanAmount = 0;
                double totalInterest = 0;
                double totalOriginationFee = 0;
                double totalUnderwritingFee = 0;
                double totalTransfer = -25; // -25 for wire fee

                int r = 0;
                //for (int r = 0; r < data.Length; r++)
                foreach (TableFundingLoan loan in loans)
                {

                    Spire.Doc.TableRow DataRow = tblSch.Rows[r + 1];

                    //Row Height

                    DataRow.Height = 20;

                    //C Represents Column.

                    //for (int c = 0; c < data[r].Length; c++)                    
                    //{                    
                    //Cell Alignment
                    DataRow.Cells[0].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    DataRow.Cells[1].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[2].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[3].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[4].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[5].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[6].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[7].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[8].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[9].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[10].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;



                    //Fill Data in Rows
                    Spire.Doc.Documents.Paragraph pBusiness = DataRow.Cells[0].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBusiness = pBusiness.AppendText(loan.LoanMortgageeBusiness);

                    Spire.Doc.Documents.Paragraph pBorrower = DataRow.Cells[1].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBorrower = pBorrower.AppendText(loan.LoanMortgagee);

                    Spire.Doc.Documents.Paragraph pLoanNumber = DataRow.Cells[2].AddParagraph();
                    Spire.Doc.Fields.TextRange TRLoanNumber = pLoanNumber.AppendText(loan.LoanNumber);

                    Spire.Doc.Documents.Paragraph pAddress = DataRow.Cells[3].AddParagraph();
                    Spire.Doc.Fields.TextRange TRAddress = pAddress.AppendText(FormatPcnt(loan.LoanMinimumInterest.ToString()));

                    Spire.Doc.Documents.Paragraph pMortgageAmount = DataRow.Cells[4].AddParagraph();
                    Spire.Doc.Fields.TextRange trMortgageAmount = pMortgageAmount.AppendText(FormatNumberCommas2dec(loan.LoanMortgageAmount.ToString()));

                    Spire.Doc.Documents.Paragraph pProceeds = DataRow.Cells[5].AddParagraph();
                    Spire.Doc.Fields.TextRange trProceeds = pProceeds.AppendText(FormatNumberCommas2dec(loan.LoanFunding.InvestorProceeds.ToString()));

                    Spire.Doc.Documents.Paragraph pLoanAmount = DataRow.Cells[6].AddParagraph();
                    Spire.Doc.Fields.TextRange trLoanAmount = pLoanAmount.AppendText(FormatNumberCommas2dec(loan.LoanAdvanceAmount.ToString()));

                    double interestFee = loan.LoanFunding.InterestFee + loan.LoanFees.CustSvcInterestDiscount;
                    Spire.Doc.Documents.Paragraph pInterest = DataRow.Cells[7].AddParagraph();
                    Spire.Doc.Fields.TextRange trInterest = pInterest.AppendText(FormatNumberCommas2dec(interestFee.ToString()));

                    Spire.Doc.Documents.Paragraph pOrig = DataRow.Cells[8].AddParagraph();
                    double originationFee = loan.LoanFunding.OriginationFee + loan.LoanFees.CustSvcOriginationDiscount;
                    Spire.Doc.Fields.TextRange trOrig = pOrig.AppendText(FormatNumberCommas2dec(originationFee.ToString()));

                    double underwritingFee = loan.LoanFunding.UnderwritingFee + loan.LoanFees.CustSvcUnderwritingDiscount;
                    Spire.Doc.Documents.Paragraph pUW = DataRow.Cells[9].AddParagraph();
                    Spire.Doc.Fields.TextRange trUW = pUW.AppendText(FormatNumberCommas2dec(underwritingFee.ToString()));

                    Spire.Doc.Documents.Paragraph pTotalTransfer = DataRow.Cells[10].AddParagraph();
                    Spire.Doc.Fields.TextRange trTotalTransfer = pTotalTransfer.AppendText(FormatNumberCommas2dec(loan.LoanFunding.TotalTransfer.ToString()));

                    totalMortgage += loan.LoanMortgageAmount;
                    totalProceeds += loan.LoanFunding.InvestorProceeds;
                    totalLoanAmount += loan.LoanAdvanceAmount;
                    totalInterest += loan.LoanFunding.InterestFee + loan.LoanFees.CustSvcInterestDiscount;
                    totalOriginationFee += loan.LoanFunding.OriginationFee + loan.LoanFees.CustSvcOriginationDiscount;
                    totalUnderwritingFee += loan.LoanFunding.UnderwritingFee + loan.LoanFees.CustSvcUnderwritingDiscount;
                    totalTransfer += loan.LoanFunding.TotalTransfer;

                    //Format Cells

                    pLoanNumber.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pBorrower.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pAddress.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pMortgageAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pProceeds.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pLoanAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pInterest.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pOrig.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pUW.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pTotalTransfer.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    //TRLoanNumber.CharacterFormat.FontName = "Calibri";
                    //TRBorrower.CharacterFormat.FontName = "Calibri";
                    //TRAddress.CharacterFormat.FontName = "Calibri";

                    TRLoanNumber.CharacterFormat.FontSize = 12;
                    TRBorrower.CharacterFormat.FontSize = 12;
                    TRAddress.CharacterFormat.FontSize = 12;

                    TRAddress.CharacterFormat.FontSize = 12;
                    trMortgageAmount.CharacterFormat.FontSize = 12;
                    trProceeds.CharacterFormat.FontSize = 12;
                    trLoanAmount.CharacterFormat.FontSize = 12;
                    trInterest.CharacterFormat.FontSize = 12;
                    trOrig.CharacterFormat.FontSize = 12;
                    trUW.CharacterFormat.FontSize = 12;
                    trTotalTransfer.CharacterFormat.FontSize = 12;
                    //TR2.CharacterFormat.TextColor = Color.Brown;

                    //}
                    r++;
                }

                Spire.Doc.TableRow FRowFooterMortgage = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterMortgage = FRowFooterMortgage.Cells[4].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterMortgage = pFooterMortgage.AppendText(FormatNumberCommas2dec(totalMortgage.ToString()));
                pFooterMortgage.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterMortgage.CharacterFormat.Bold = true;
                TRFooterMortgage.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterProceeds = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterProceeds = FRowFooterProceeds.Cells[5].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterProceeds = pFooterProceeds.AppendText(FormatNumberCommas2dec(totalProceeds.ToString()));
                pFooterProceeds.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterProceeds.CharacterFormat.Bold = true;
                TRFooterProceeds.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterLoanAmount = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterLoanAmount = FRowFooterLoanAmount.Cells[6].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterLoanAmount = pFooterLoanAmount.AppendText(FormatNumberCommas2dec(totalLoanAmount.ToString()));
                pFooterLoanAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterLoanAmount.CharacterFormat.Bold = true;
                TRFooterLoanAmount.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterInterest = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterInterest = FRowFooterInterest.Cells[7].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterInterest = pFooterInterest.AppendText(FormatNumberCommas2dec(totalInterest.ToString()));
                pFooterInterest.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterInterest.CharacterFormat.Bold = true;
                TRFooterInterest.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterOrig = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterOrig = FRowFooterOrig.Cells[8].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterOrig = pFooterOrig.AppendText(FormatNumberCommas2dec(totalOriginationFee.ToString()));
                pFooterOrig.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterOrig.CharacterFormat.Bold = true;
                TRFooterOrig.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterUW = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterUW = FRowFooterUW.Cells[9].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterUW = pFooterUW.AppendText(FormatNumberCommas2dec(totalUnderwritingFee.ToString()));
                pFooterUW.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterUW.CharacterFormat.Bold = true;
                TRFooterUW.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterTotalTransfer = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterTotalTransfer = FRowFooterTotalTransfer.Cells[10].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterTotalTransfer = pFooterTotalTransfer.AppendText(FormatNumberCommas2dec(totalTransfer.ToString()));
                pFooterTotalTransfer.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterTotalTransfer.CharacterFormat.Bold = true;
                TRFooterTotalTransfer.CharacterFormat.FontSize = 12;

                //
                // End Table Logic
                //

                Spire.Doc.Documents.Paragraph parEscrow = section1.AddParagraph();
                parEscrow.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

                parEscrow.AppendText("Wire Sent to " + client.ClientName);

                //document.SaveToFile("C:/Temp/result.docx", FileFormat.Docx);

                //Spire.Doc.Document DocOne = new Spire.Doc.Document();

                //DocOne.LoadFromFile("C:/Temp/Release Text.docx", FileFormat.Docx);

                //DocOne.Replace("TODAYSDATE", todayStr, false, true);

                //DocOne.Replace("BAILEELETTERDATE", baileeLetterDateStr, false, true);

                //foreach (Spire.Doc.Section sec in document.Sections)
                //{
                //    DocOne.Sections.Add(sec.Clone());
                //}
                DateTime today = DateTime.Today;
                string todayStr;
                todayStr = string.Format("{0:MMMM dd, yyyy}", today);

                string fileName;
                string todayStrFile = string.Format("{0:yyyy_MM_dd}", today);

                fileName = "Remittance_" + client.ClientName + "_" + entity.EntityName + "_" + todayStrFile + ".docx";
                CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                dialog.InitialDirectory = "C:\\Users";
                dialog.IsFolderPicker = true;
                
                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    fileName = dialog.FileName + "\\" + fileName;
                    
                }
                

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

        private void LoansDataGrid_GotFocus(object sender, RoutedEventArgs e)
        {
            
            try
            { 
                if (MyVariables.needsRefresh)
                {
                    Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
                    bool showCompleted = false;
                    
                    if (chkShowCompleted.IsChecked == true)
                        showCompleted = true;

                    refreshMainWindowGrid();
                    if (showCompleted)
                        showAll();

                    Mouse.OverrideCursor = null;
                    MyVariables.needsRefresh = false;
                }
            }
            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
            }
        }

        private void MenuCancel_Click(object sender, RoutedEventArgs e)
        {
            try

            {
                TableFundingLoan loan = (TableFundingLoan)LoansDataGrid.SelectedItem;

                ChangeLoanStatusWindow win = new ChangeLoanStatusWindow(loan);
                win.Show();

                ////Read only users cannot save
                ////
                //if (MyVariables.userRole.Contains("Read"))
                //{
                //    ErrorLabel.Content = "Permission Denied";
                //    ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                //    return;
                //}
                //TableFundingLoan loan = (TableFundingLoan)LoansDataGrid.SelectedItem;
                //int cancelledStatusID = 0;
                //foreach (DataRow row in loanStatus.Tables[0].Rows)
                //{
                //    if (row[3].ToString().Contains("Cancel"))
                //        cancelledStatusID = Convert.ToInt32(row[2].ToString());
                //}
                //loan.LoanStatusID = cancelledStatusID;
                //loan.LoanUpdateUserID = MyVariables.user.UserID;
                
                //loan.Save();
                //refreshMainWindowGrid();
            }

            catch (Exception ex)
            {
                ErrorLabel.Content = ex.Message;
                ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);

            }


        }

        private void DashboardButton_Click(object sender, RoutedEventArgs e)
        {
            ErrorLabel.Content = "";
            try
            {
                System.Diagnostics.Process.Start("https://app.powerbi.com/view?r=eyJrIjoiZmU4MDdjZjUtOTFmZC00ZDU2LTgxYWMtMTk1NTBhZDQ1MjMxIiwidCI6IjcyOWQxY2U4LTUxN2EtNGJjYi1iY2JlLTgzYThmNzJkYzMzMSIsImMiOjF9");
                //Dashboard dwindow = new Dashboard();
                //dwindow.Show();
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

