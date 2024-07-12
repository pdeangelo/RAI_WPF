using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


using Clearwater.DataAccess;



    /// ---------------------------------------------
    /// Project		: CashFlow.TableFunding
    /// Class		: Table funding Loan Collection
    /// ---------------------------------------------
    /// <summary>
    /// Represent collection of the security type
    /// </summary>
    /// <history>
    /// 	[rrb] 		Created
    /// </history>
    /// --------------------------------------------------
    public class TableFundingLoans: List<TableFundingLoan>
      {

        #region Constants

        #endregion
        #region Private Members
        #endregion

        #region Member Variables
      
        /// <summary>

        #endregion
        #region Constructor

        /// <summary>
        /// Constructor. Initialize the class member variables
        /// </summary>
        public TableFundingLoans()
        {
            //intialize the member variables
            Initialize();
         
        }
        public TableFundingLoans(bool getUnderlyingData = true, bool showCompleted = false)
        {
            //intialize the member variables
            Initialize();
            LoadData(getUnderlyingData, showCompleted);
        }
        public TableFundingLoans(int clientID, int investorID, int statusID, bool getUnderlyingData, string loanNumber, bool showCompleted = false)
        {
            //intialize the member variables
            Initialize();
            LoadData(clientID, investorID, statusID, getUnderlyingData, showCompleted, loanNumber);
        }

        ///// <summary>
        ///// Constructor. Initialize the class member variables
        ///// </summary>
        //public TableFundingClients(int issuerId, int year)
        //{
        //    //intialize the member variables
        //    Initialize();
        //    LoadData(issuerId, year);
        //}

        #endregion

        #region Public  Methods


        #endregion

        #region Private Methods

        /// <summary>
        /// Initialize the member variables
        /// </summary>
        private void Initialize()
        {

        }
      
        private void LoadData(bool getUnderlyingData, bool showCompleted)
        {
            try
            {
                ParameterCollection parameters = new ParameterCollection();
                parameters.Add(new Parameter(TableFundingLoan.Param_ShowCompleted, showCompleted));

                DBHelper.ExecuteToList<TableFundingLoan>(TableFundingLoan.SP_TRANSACTION_SELECT, parameters, this);

                if (getUnderlyingData)
                { 
                    foreach (TableFundingLoan loan in this)
                    {
                        TableFundingLoanUW loanUW = new TableFundingLoanUW(loan.LoanID);
                        loan.LoanUW = loanUW;
                        TableFundingLoanFunding loanFunding = new TableFundingLoanFunding(loan.LoanID);
                        loan.LoanFunding = loanFunding;
                        TableFundingFees loanFees = new TableFundingFees(loan.LoanID);
                        loan.LoanFees = loanFees;
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private void LoadData(int clientID, int investorID, int statusID, bool getUnderlyingData, bool showCompleted, string loanNumber)
        {
            try
            {
                ParameterCollection parameters = new ParameterCollection();
                parameters.Add(new Parameter(TableFundingLoan.Param_ClientID, clientID));
                parameters.Add(new Parameter(TableFundingLoan.Param_InvestorID, investorID));
                parameters.Add(new Parameter(TableFundingLoan.Param_StatusID, statusID));
                parameters.Add(new Parameter(TableFundingLoan.Param_ShowCompleted, showCompleted));
                parameters.Add(new Parameter(TableFundingLoan.Param_LoanNumber, loanNumber));

                DBHelper.ExecuteToList<TableFundingLoan>(TableFundingLoan.SP_TRANSACTION_SELECT, parameters, this);

                if (getUnderlyingData)
                { 
                    foreach (TableFundingLoan loan in this)
                    {
                        TableFundingLoanUW loanUW = new TableFundingLoanUW(loan.LoanID);
                        loan.LoanUW = loanUW;
                        TableFundingLoanFunding loanFunding = new TableFundingLoanFunding(loan.LoanID);
                        loan.LoanFunding = loanFunding;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        //private void LoadData(int issuerID, int year)
        //{
        //    try
        //    {
        //        ParameterCollection parameters = new ParameterCollection();
        //        DBHelper.ExecuteToList<ESGQuestion>(SP_ESGQuestion_SELECT, parameters, this);
        //        //Get the answer for the selected report date
        //        foreach (ESGQuestion question in this)
        //        {
        //            question.ESGAnswer = new ESGAnswer(question.ESGQuestionID, issuerID, year);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }

        //}

        #endregion

        #region Internal Methods

        #endregion

    }
