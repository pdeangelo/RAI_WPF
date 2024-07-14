using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Clearwater.DataAccess;
using System.Data;
using System.Data.SqlClient;
using log4net;
public class TableFundingLoanFunding
    {
        #region Constants
        /// <summary>
        /// Define all the procedure name here
        /// </summary>
        internal const String SP_TRANSACTION_SELECT = "TableFundingLoanFunding_Sel";
        internal const String SP_TRANSACTION_SAVE = "TableFundingLoanFunding_AddEdit";
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// Define all the procedure parameter fields as a constant here
        /// </summary>
        internal const String Param_FundingID = "FundingID";
        internal const String Param_LoanID = "LoanID";
        internal const String Param_DateDepositedInEscrow = "DateDepositedInEscrow";
        internal const String Param_InvestorProceedsDate = "InvestorProceedsDate";
        internal const String Param_InvestorProceeds = "InvestorProceeds";
        internal const String Param_BaileeLetterDate = "BaileeLetterDate";
        internal const String Param_ClosingDate = "ClosingDate";
        internal const String Param_UnderwritingFee = "UnderwritingFee";

        internal const String Param_InterestFee = "InterestFee";
        internal const String Param_OriginationFee = "OriginationFee";
        internal const String Param_TotalFees = "TotalFees";
        internal const String Param_TotalTransfer = "TotalTransfer";

        #endregion

        #region Member Variables

        private int _fundingID;
        private int _loanID;
        private DateTime? _dateDepositedInEscrow;
        private DateTime? _investorProceedsDate;
        private DateTime? _closingDate;
        private double _investorProceeds;
        private DateTime? _baileeLetterDate;
        private double _interestFee;
        private double _underwritingFee;
        private double _originationFee;
        private double _totalFees;
        private double _totalTransfer;
        #endregion

        #region Constructor
        public TableFundingLoanFunding()
        {
        }

        public TableFundingLoanFunding(int loanID)
        {
            try
            {
                //initialize the private members
                Initialze();

                //Load the memeber variable for the given security
                LoadData(loanID);
            }
            catch (Exception ex)
            {
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
                throw ex;
            }


        }
        #endregion

        #region Public Property

        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public int FundingID
        {
            get { return _fundingID; }
            set { _fundingID = value; }
        }
        public int LoanID
        {
            get { return _loanID; }
            set { _loanID = value; }
        }
        public DateTime? DateDepositedInEscrow
        {
            get { return _dateDepositedInEscrow; }
            set { _dateDepositedInEscrow = value; }
        }
        public DateTime? InvestorProceedsDate
        {
            get { return _investorProceedsDate; }
            set { _investorProceedsDate = value; }
        }
        public DateTime? ClosingDate
        {
            get { return _closingDate; }
            set { _closingDate = value; }
        }
        public double InvestorProceeds
        {
            get { return _investorProceeds; }
            set { _investorProceeds = value; }
        }
        
        public DateTime? BaileeLetterDate
        {
            get { return _baileeLetterDate; }
            set { _baileeLetterDate = value; }
        }
      
        public double InterestFee
        {
            get { return _interestFee; }
            set { _interestFee = value; }
        }
        public double OriginationFee
        {
            get { return _originationFee; }
            set { _originationFee = value; }
        }
        public double UnderwritingFee
        {
            get { return _underwritingFee; }
            set { _underwritingFee = value; }
        }

        public double TotalFees
        {
            get { return _totalFees; }
            set { _totalFees = value; }
        }
        public double TotalTransfer
        {
            get { return _totalTransfer; }
            set { _totalTransfer = value; }
        }
        #endregion
        #region Internal Functions
        private void Initialze()
        {

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

                            _fundingID = (int) dr[Param_FundingID];
                            _loanID = (int)dr[Param_LoanID];

                            if (dr[Param_DateDepositedInEscrow].ToString().Length == 0)
                                _dateDepositedInEscrow = null;
                            else
                                _dateDepositedInEscrow = Convert.ToDateTime(dr[Param_DateDepositedInEscrow]);

                            if (dr[Param_InvestorProceedsDate].ToString().Length == 0)
                                _investorProceedsDate = null;
                            else
                                _investorProceedsDate = Convert.ToDateTime(dr[Param_InvestorProceedsDate]);

                            _investorProceeds = Convert.ToDouble(dr[Param_InvestorProceeds]);
                            if (dr[Param_BaileeLetterDate].ToString().Length == 0)
                                _baileeLetterDate = null;
                            else
                                _baileeLetterDate = Convert.ToDateTime(dr[Param_BaileeLetterDate]);
                            if (dr[Param_ClosingDate].ToString().Length == 0)
                                _closingDate = null;
                            else
                                _closingDate = Convert.ToDateTime(dr[Param_ClosingDate]);

                            _interestFee = Convert.ToDouble(dr[Param_InterestFee]);
                            

                            _originationFee = Convert.ToDouble(dr[Param_OriginationFee]);

                            _underwritingFee = Convert.ToDouble(dr[Param_UnderwritingFee]);
                           

                            _totalFees = Convert.ToDouble(dr[Param_TotalFees]);
                            _totalTransfer = Convert.ToDouble(dr[Param_TotalTransfer]);
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
                parameters.Add(new Parameter(Param_FundingID, _fundingID));
                parameters.Add(new Parameter(Param_LoanID, _loanID));
                parameters.Add(new Parameter(Param_DateDepositedInEscrow, _dateDepositedInEscrow));
                parameters.Add(new Parameter(Param_BaileeLetterDate, _baileeLetterDate));
                parameters.Add(new Parameter(Param_InvestorProceedsDate, _investorProceedsDate));
                parameters.Add(new Parameter(Param_InvestorProceeds, _investorProceeds));
                parameters.Add(new Parameter(Param_ClosingDate, _closingDate));



                //Set the Exception Message  
                errMsg.Append("Error occured while saving the Loan Funding for  -:");
                errMsg.Append("Loan Funding ID:" + FundingID.ToString());
                result.Status = true;

                //Save Data
                FundingID = db.Update(SP_TRANSACTION_SAVE, parameters);

                if (FundingID == AppConstants.FAILURE)
                {
                    result.Message = errMsg.ToString();
                    result.Status = false;

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
