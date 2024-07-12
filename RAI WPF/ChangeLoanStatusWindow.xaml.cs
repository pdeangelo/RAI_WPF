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
    public partial class ChangeLoanStatusWindow : Window
    {
        TableFundingLoan Loan = new TableFundingLoan();
        public ChangeLoanStatusWindow(TableFundingLoan loan)
        {
            InitializeComponent();
            DataSet loanStatus = RunsStoredProc.RunStoredProc("MiscValue_SEL", "@MiscTypeID", "2", "", "", "", "", "", "", "", "", "", "", "", 0);

            cmbLoanStatus.ItemsSource = loanStatus.Tables[0].DefaultView;
            cmbLoanStatus.DisplayMemberPath = loanStatus.Tables[0].Columns["Value"].ToString();
            cmbLoanStatus.SelectedValuePath = loanStatus.Tables[0].Columns["MiscTypeValueID"].ToString();

            cmbLoanStatus.SelectedValue = loan.LoanStatusID;
            Loan = loan;
        }
       
        private void StatusSaveButton_Click(object sender, RoutedEventArgs e)
        {

            ErrorLabel.Content = "";
            try
            {
                Loan.LoanStatusID = Convert.ToInt32(cmbLoanStatus.SelectedValue);
                Loan.Save();
                MyVariables.needsRefresh = true;

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
