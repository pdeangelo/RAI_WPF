using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Clearwater.DataAccess;
using System.Data.SqlClient;
using log4net;
public class TableFundingLoan
    {
        #region Constants
        /// <summary>
        /// Define all the procedure name here
        /// </summary>
        internal const String SP_TRANSACTION_SELECT = "TableFundingLoans_Sel";
        internal const String SP_TRANSACTION_SAVE = "TableFundingLoans_AddEdit";
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// Define all the procedure parameter fields as a constant here
        /// </summary>
        internal const String Param_LoanID = "LoanID";
        internal const String Param_ClientID = "LoanClientID";
        internal const String Param_ClientName = "ClientName";
        internal const String Param_EntityID = "LoanEntityID";
        internal const String Param_EntityName = "EntityName";
        internal const String Param_InvestorID = "LoanInvestorID";
        internal const String Param_InvestorName = "InvestorName";
        internal const String Param_StatusID = "LoanStatusID";
        internal const String Param_Status = "LoanStatus";
        internal const String Param_LoanNumber = "LoanNumber";
        internal const String Param_LoanFundingDate = "LoanFundingDate";
        internal const String Param_LoanMortgagee = "LoanMortgagee";
        internal const String Param_LoanPropertyAddress = "LoanPropertyAddress";
        internal const String Param_LoanInterestRate = "LoanInterestRate";
        internal const String Param_LoanMortgageAmount = "LoanMortgageAmount";
        internal const String Param_LoanAdvanceAmount = "LoanAdvanceAmount";
        internal const String Param_LoanAdvanceRate = "LoanAdvanceRate";
        internal const String Param_LoanReserveAmount = "LoanReserveAmount";
        internal const String Param_LoanType = "LoanType";
        internal const String Param_LoanTypeName = "LoanTypeName";
        internal const String Param_LoanDwellingType = "LoanDwellingType";
        internal const String Param_LoanDwellingTypeName = "LoanDwellingTypeName";
        internal const String Param_ShowCompleted = "ShowCompleted";
        internal const String Param_LoanEnteredDate = "LoanEnteredDate";
        internal const String Param_LoanUpdateDate = "LoanUpdateDate";
        internal const String Param_LoanUpdateUserID = "LoanUpdateUserID";
        internal const String Param_LoanMortgageeBusiness = "LoanMortgageeBusiness";
        internal const String Param_LoanMinimumInterest = "LoanMinimumInterest";
        internal const String Param_InvestorProceedsDate = "InvestorProceedsDate";
        internal const String Param_DateDepositedInEscrow = "DateDepositedInEscrow";
        internal const String Param_State = "State";
        internal const String Param_StateName = "StateName";

    #endregion

    #region Member Variables

    private int _loanID;
        private int _loanClientID;
        private int _loanEntityID;
        private int _loanType;
        private string _loanTypeName;
        private int _loanDwellingType;
        private string _loanDwellingTypeName;
        private string _entityName;
        private string _investorName;
        private string _clientName;
        private int _loanInvestorID;
        private int _loanStatusID;
        private int _loanUpdateUserID;
        private string _loanStatus;
        private int _state;
        private string _stateName;
        private string _loanNumber;
        private DateTime? _loanFundingDate;
        private DateTime? _loanEnteredDate;
        private DateTime? _loanUpdateDate;
        private DateTime? _investorProceedsDate;
        private DateTime? _dateDepositedInEscrow;

        private string _loanMortgagee;
        private string _loanMortgageeBusiness;

        private double _loanMinimumInterest;
        private double _loanInterestRate;
        private string _loanPropertyAddress;
        
        private double _loanMortgageAmount;
        private double _loanAdvanceRate;
        private double _loanAdvanceAmount;
        private double _loanReserveAmount;

        private TableFundingLoanUW _loanUW;
        private TableFundingLoanFunding _loanFunding;
        private bool _isChecked;
        private TableFundingFees _loanFees;

        #endregion

        #region Constructor
        public TableFundingLoan()
        {
            Initialize();
        }
        public TableFundingLoan(int loanID)
        {
            //intialize the member variables
            Initialize();
            LoadData(loanID);
        }

        #endregion

        #region Public Property

        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public int LoanID
        {
            get { return _loanID; }
            set { _loanID = value; }
        }

        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public int LoanClientID
        {
            get { return _loanClientID; }
            set { _loanClientID = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public int LoanType
        {
            get { return _loanType; }
            set { _loanType = value; }
        }

        public int State
        {
            get { return _state; }
            set { _state = value; }
        }
        public string StateName
        {
            get { return _stateName; }
            set { _stateName = value; }
        }
        public int LoanDwellingType
        {
            get { return _loanDwellingType; }
            set { _loanDwellingType = value; }
        }

        /// <summary>
        /// Gets/sets the Transaction Date of the transaction
        /// </summary>
        public int LoanEntityID
        {
            get { return _loanEntityID; }
            set { _loanEntityID = value; }
        }


        /// <summary>
        /// Get/Set the Loan Number value 
        /// </summary>
        public int LoanInvestorID
        {
            get { return _loanInvestorID; }
            set { _loanInvestorID = value; }
        }

        /// <summary>
        /// Get/Set the Loan Number value 
        /// </summary>
        public int LoanUpdateUserID
        {
            get { return _loanUpdateUserID; }
            set { _loanUpdateUserID = value; }
        }

        /// <summary>
        /// Get/Set the Dealer Number
        /// </summary>
        public int LoanStatusID
        {
            get { return _loanStatusID; }
            set { _loanStatusID = value; }
        }


        /// <summary>
        /// Get/Set the Loan Type 
        /// </summary>
        public string LoanTypeName
        {
            get { return _loanTypeName; }
            set { _loanTypeName = value; }
        }
        public string LoanDwellingTypeName
        {
            get { return _loanDwellingTypeName; }
            set { _loanDwellingTypeName = value; }
        }
        /// <summary>
        /// Get/Set the Mortgagee Amount 
        /// </summary>
        public string LoanNumber
        {
            get { return _loanNumber; }
            set { _loanNumber = value; }
        }

        /// <summary>
        /// Gets/sets the Entity Amount
        /// </summary>    
        public DateTime? LoanFundingDate
        {
            get { return _loanFundingDate; }
            set { _loanFundingDate = value; }
        }
        public DateTime? LoanEnteredDate
        {
            get { return _loanEnteredDate; }
            set { _loanEnteredDate = value; }
        }
        public DateTime? LoanUpdateDate
        {
            get { return _loanUpdateDate; }
            set { _loanUpdateDate = value; }
        }
        public DateTime? InvestorProceedsDate
        {
            get { return _investorProceedsDate; }
            set { _investorProceedsDate = value; }
        }
        public DateTime? DateDepositedInEscrow
        {
            get { return _dateDepositedInEscrow; }
            set { _dateDepositedInEscrow = value; }
        }

        /// <summary>
        /// Gets/sets the Property Address
        /// </summary>    
        public string LoanMortgagee
        {
            get { return _loanMortgagee; }
            set { _loanMortgagee = value; }
        }
        public string LoanMortgageeBusiness
        {
            get { return _loanMortgageeBusiness; }
            set { _loanMortgageeBusiness = value; }
        }

        /// <summary>
        /// Gets/sets the Property Address
        /// </summary>    
        public string LoanPropertyAddress
        {
            get { return _loanPropertyAddress; }
            set { _loanPropertyAddress = value; }
        }

        /// <summary>
        /// Get/Set the Mortgage Amount
        /// </summary>
        public double LoanMortgageAmount
        {
            get { return _loanMortgageAmount; }
            set { _loanMortgageAmount = value; }
        }
        public double LoanAdvanceAmount
        {
            get { return _loanAdvanceAmount; }
            set { _loanAdvanceAmount = value; }
        }
        public double LoanReserveAmount
        {
            get { return _loanReserveAmount; }
            set { _loanReserveAmount = value; }
        }


        /// <summary>
        /// Get/Set the Mortgage Amount
        /// </summary>
        public double LoanAdvanceRate
        {
            get { return _loanAdvanceRate; }
            set { _loanAdvanceRate = value; }
        }



        /// <summary>
        /// Gets/sets the interest amount of the transaction
        /// </summary>                    
        public double LoanInterestRate
        {
            get { return _loanInterestRate; }
            set { _loanInterestRate = value; }
        }
        public double LoanMinimumInterest
        {
            get { return _loanMinimumInterest; }
            set { _loanMinimumInterest = value; }
        }

        public string EntityName
        {
            get { return _entityName; }
            set { _entityName = value; }
        }

        public string InvestorName
        {
            get { return _investorName; }
            set { _investorName = value; }
        }

        public string ClientName
        {
            get { return _clientName; }
            set { _clientName = value; }
        }
        public string LoanStatus
        {
            get { return _loanStatus; }
            set { _loanStatus = value; }
        }
        public TableFundingLoanUW LoanUW
        {
            get { return _loanUW; }
            set { _loanUW = value; }
        }
        public TableFundingLoanFunding LoanFunding
        {
            get { return _loanFunding; }
            set { _loanFunding = value; }
        }
        public TableFundingFees LoanFees
        {
            get { return _loanFees; }
            set { _loanFees = value; }
        }
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool IsChecked
        {
            get { return _isChecked; }
            set { _isChecked = value; }
        }
        #endregion
        #region Internal Functions
        private void Initialize()
        {
            _loanMortgageAmount = 0;
            _loanInterestRate = 0;
            _loanAdvanceRate = 0;
            _loanUW = new TableFundingLoanUW();
        }
        private void LoadData(int loanID)
        {
            SqlDataReader dr = null;

            try
            {

                //Fill the procedure parameters
                ParameterCollection parmeters = new ParameterCollection();
                parmeters.Add(new Parameter(Param_LoanID, loanID));

                dr = (SqlDataReader)DBHelper.ExecuteReader(SP_TRANSACTION_SELECT, parmeters);

                if (dr != null)
                {
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
        
                            _loanID = (int)dr[Param_LoanID];
                            _loanClientID = (int)dr[Param_ClientID];
                            _clientName = dr[Param_ClientName].ToString();
                            _loanEntityID = (int)dr[Param_EntityID];

                            _loanUpdateUserID = (int)dr[Param_LoanUpdateUserID];
                            _entityName = dr[Param_EntityName].ToString();
                            _loanTypeName = dr[Param_LoanTypeName].ToString();
                            _loanDwellingTypeName = dr[Param_LoanDwellingTypeName].ToString();
                            _loanInvestorID = (int)dr[Param_InvestorID];
                            _investorName = dr[Param_InvestorName].ToString();
                            _loanStatusID = (int)dr[Param_StatusID];
                            _loanStatus = dr[Param_Status].ToString();
                            _loanNumber = dr[Param_LoanNumber].ToString();

                            _state = (int)dr[Param_State];
                            _stateName = dr[Param_StateName].ToString();
                            if (dr[Param_LoanFundingDate].ToString().Length == 0)
                                _loanFundingDate = null;
                            else
                                _loanFundingDate = Convert.ToDateTime(dr[Param_LoanFundingDate]);

                            if (dr[Param_LoanEnteredDate].ToString().Length == 0)
                                _loanEnteredDate = null;
                            else
                                _loanEnteredDate = Convert.ToDateTime(dr[Param_LoanEnteredDate]);

                            if (dr[Param_LoanUpdateDate].ToString().Length == 0)
                                _loanUpdateDate = null;
                            else
                                _loanUpdateDate = Convert.ToDateTime(dr[Param_LoanUpdateDate]);

                            _loanMortgagee = dr[Param_LoanMortgagee].ToString();
                            _loanMortgageeBusiness = dr[Param_LoanMortgageeBusiness].ToString();
                            _loanPropertyAddress = dr[Param_LoanPropertyAddress].ToString();
                            _loanInterestRate = (double)dr[Param_LoanInterestRate];
                            _loanMinimumInterest = (double)dr[Param_LoanMinimumInterest];
                            _loanMortgageAmount = (double)dr[Param_LoanMortgageAmount];
                            _loanAdvanceAmount = (double)dr[Param_LoanAdvanceAmount];
                            _loanAdvanceRate = (double)dr[Param_LoanAdvanceRate];
                            _loanReserveAmount = (double)dr[Param_LoanReserveAmount];
                            _loanType = (int)dr[Param_LoanType];
                            _loanDwellingType = (int)dr[Param_LoanDwellingType];

                        TableFundingLoanUW loanUW = new TableFundingLoanUW(_loanID);
                            _loanUW = loanUW;
                           

                        }//End of dr.Read

                    }//End of dr.HasRows

                }//End of dr!= Null
            }
            catch (Exception ex)
            {
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
                throw ex;
            }
            finally
            {
                //Close the connection
                DBHelper.DisposeOff(dr);
            }
        }
        #endregion
        #region Public Methods
        public ReturnResult Save()
        {

            ReturnResult result = new ReturnResult();
            StringBuilder errMsg = new StringBuilder();
            DBHelper db = null;

            try
            {
                //Instantiate the db helper object
                db = new DBHelper();

                //Initiate the database transaction
                db.BeginTransaction();

                ParameterCollection parameters = new ParameterCollection();
                parameters.Add(new Parameter(Param_LoanID, _loanID));
                parameters.Add(new Parameter(Param_ClientID, _loanClientID));
                parameters.Add(new Parameter(Param_EntityID, _loanEntityID));
                parameters.Add(new Parameter(Param_InvestorID, _loanInvestorID));
                parameters.Add(new Parameter(Param_StatusID, _loanStatusID));
                parameters.Add(new Parameter(Param_LoanNumber, _loanNumber));
                parameters.Add(new Parameter(Param_LoanFundingDate, _loanFundingDate));
                parameters.Add(new Parameter(Param_LoanMortgagee, _loanMortgagee));
                parameters.Add(new Parameter(Param_LoanPropertyAddress, _loanPropertyAddress));
                parameters.Add(new Parameter(Param_LoanInterestRate, _loanInterestRate));
                parameters.Add(new Parameter(Param_LoanMortgageAmount, _loanMortgageAmount));
                parameters.Add(new Parameter(Param_LoanAdvanceRate, _loanAdvanceRate));
                parameters.Add(new Parameter(Param_LoanType, _loanType));
                parameters.Add(new Parameter(Param_LoanUpdateUserID, _loanUpdateUserID));
                parameters.Add(new Parameter(Param_LoanMortgageeBusiness, _loanMortgageeBusiness));
                parameters.Add(new Parameter(Param_LoanDwellingType, _loanDwellingType));
                parameters.Add(new Parameter(Param_State, _state));

                //Set the Exception Message  
                errMsg.Append("Error occured while saving the Loan for  -:");
                errMsg.Append("Loan Name:" + LoanID.ToString());
                result.Status = true;

                //Save Data
                LoanID = db.Update(SP_TRANSACTION_SAVE, parameters);

                if (LoanID == AppConstants.FAILURE)
                {
                    result.Message = errMsg.ToString();
                    result.Status = false;
                    db.RollBack();  //Roll back the transaction
                    return result;

                }

                //Check if the database updated has executed success or not
                if (result.Status == false)
                    db.RollBack();  //Roll back the transaction
                else
                    db.Commit();    //Commits the transactions

            }
            catch (Exception ex)
            {
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
                errMsg.Append(ex.Message);
                result.Message = errMsg.ToString();
                result.Status = false;

                db.RollBack();  //Roll back the transaction
            }

            return result;

        }

        #endregion
    }

