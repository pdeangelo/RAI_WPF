using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Clearwater.DataAccess;
using System.Data.SqlClient;
using log4net;

public class TableFundingFees
    {
        #region Constants
        /// <summary>
        /// Define all the procedure name here
        /// </summary>
        internal const String SP_TRANSACTION_SELECT = "TableFundingFees_Sel";
        internal const String SP_TRANSACTION_SAVE = "TableFundingFees_AddEdit";
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

    /// <summary>
    /// Define all the procedure parameter fields as a constant here
    /// </summary>
    internal const String Param_LoanID = "LoanID";
        internal const String Param_MinimumInterest = "MinimumInterest";
        internal const String Param_ClientPrimeRate = "ClientPrimeRate";
        internal const String Param_ClientPrimeRateSpread = "ClientPrimeRateSpread";
        internal const String Param_OriginationDiscount = "OriginationDiscount";
        internal const String Param_OriginationDiscount2 = "OriginationDiscount2";
        internal const String Param_OriginationDiscountNumDays = "OriginationDiscountNumDays";
        internal const String Param_OriginationDiscountNumDays2 = "OriginationDiscountNumDays2";
        internal const String Param_InterestBasedOnAdvance = "InterestBasedOnAdvance";
        internal const String Param_NoInterest = "NoInterest";
        internal const String Param_OriginationBasedOnAdvance = "OriginationBasedOnAdvance";
        internal const String Param_InterestFee = "InterestFee";
        internal const String Param_OriginationFee = "OriginationFee";
        internal const String Param_UnderwritingFee = "UnderwritingFee";
        internal const String Param_CustSvcUnderwritingDiscount = "CustSvcUnderwritingDiscount";
        internal const String Param_CustSvcOriginationDiscount = "CustSvcOriginationDiscount";
        internal const String Param_CustSvcInterestDiscount = "CustSvcInterestDiscount";
        #endregion

        #region Member Variables

        private int _loanID;
        private double _minimumInterest;
        private double _clientPrimeRate;
        private double _clientPrimeRateSpread;
        private double _originationDiscount;
        private double _originationDiscount2;
        private int _originationDiscountNumDays;
        private int _originationDiscountNumDays2;
        private bool _interestBasedOnAdvance;
        private bool _originationBasedOnAdvance;
        private bool _noInterest;
        private double _originationFee;
        private double _interestFee;
        private double _underwritingFee;

        private double _custSvcInterestDiscount;
        private double _custSvcUnderwritingDiscount;
        private double _custSvcOriginationDiscount;
        #endregion

        #region Constructor
        public TableFundingFees()
        {
            Initialize();
        }
        public TableFundingFees(int loanID)
        {
            //intialize the member variables
            Initialize();
            if (loanID > -1)
                LoadData(loanID);
        }
        #endregion

        #region Internal Functions
        private void Initialize()
        {
            _loanID = 0;
            _minimumInterest = 0;
            _clientPrimeRate = 0;
            _clientPrimeRateSpread = 0;
            _originationDiscount = 0;
            _originationDiscount2 = 0;
            _originationDiscountNumDays = 0;
            _originationDiscountNumDays2 = 0;
            _interestBasedOnAdvance = false;
            _originationBasedOnAdvance = false;
            _noInterest = false;
            _originationFee = 0; ;
            _interestFee = 0;
            _underwritingFee = 0;
            _custSvcInterestDiscount = 0;
            _custSvcOriginationDiscount = 0;
            _custSvcUnderwritingDiscount = 0;

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
                            _minimumInterest = (double)dr[Param_MinimumInterest];
                            _clientPrimeRate = (double)dr[Param_ClientPrimeRate];
                            _clientPrimeRateSpread = (double)dr[Param_ClientPrimeRateSpread];
                            _originationDiscount = (double)dr[Param_OriginationDiscount];
                            _originationDiscount2 = (double)dr[Param_OriginationDiscount2];
                            _originationDiscountNumDays = (int)dr[Param_OriginationDiscountNumDays];
                            _originationDiscountNumDays2 = (int)dr[Param_OriginationDiscountNumDays2];
                            _interestBasedOnAdvance = (bool)dr[Param_InterestBasedOnAdvance];
                            _originationBasedOnAdvance = (bool)dr[Param_OriginationBasedOnAdvance];
                            _noInterest = (bool)dr[Param_NoInterest];
                            _interestFee = (double)dr[Param_InterestFee];
                            _originationFee = (double)dr[Param_OriginationFee];
                            _underwritingFee = (double)dr[Param_UnderwritingFee];
                            _custSvcOriginationDiscount = Convert.ToDouble(dr[Param_CustSvcOriginationDiscount]);
                            _custSvcUnderwritingDiscount = Convert.ToDouble(dr[Param_CustSvcUnderwritingDiscount]);
                            _custSvcInterestDiscount = Convert.ToDouble(dr[Param_CustSvcInterestDiscount]);

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
        #region Public Property

        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public int LoanID
        {
            get { return _loanID; }
            set { _loanID = value; }
        }

        public double MinimumInterest
        {
            get { return _minimumInterest; }
            set { _minimumInterest = value; }
    }
        public double ClientPrimeRate
        {
            get { return _clientPrimeRate; }
            set { _clientPrimeRate = value; }
        }
        public double ClientPrimeRateSpread
    {
            get { return _clientPrimeRateSpread; }
            set { _clientPrimeRateSpread = value; }
        }
        public double OriginationDiscount
        {
            get { return _originationDiscount; }
            set { _originationDiscount = value; }
        }
        public double OriginationDiscount2
        {
            get { return _originationDiscount2; }
            set { _originationDiscount2 = value; }
        }
        public int OriginationDiscountNumDays
        {
            get { return _originationDiscountNumDays; }
            set { _originationDiscountNumDays = value; }
        }
        public int OriginationDiscountNumDays2
        {
            get { return _originationDiscountNumDays2; }
            set { _originationDiscountNumDays2 = value; }
        }

        public bool InterestBasedOnAdvance
        {
            get { return _interestBasedOnAdvance; }
            set { _interestBasedOnAdvance = value; }
        }
        public bool OriginationBasedOnAdvance
        {
            get { return _originationBasedOnAdvance; }
            set { _originationBasedOnAdvance = value; }
        }
        public bool NoInterest
        {
            get { return _noInterest; }
            set { _noInterest = value; }
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

        public double CustSvcInterestDiscount
        {
            get { return _custSvcInterestDiscount; }
            set { _custSvcInterestDiscount = value; }
        }
        public double CustSvcOriginationDiscount
        {
            get { return _custSvcOriginationDiscount; }
            set { _custSvcOriginationDiscount = value; }
        }
        public double CustSvcUnderwritingDiscount
        {
            get { return _custSvcUnderwritingDiscount; }
            set { _custSvcUnderwritingDiscount = value; }
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
                parameters.Add(new Parameter(Param_MinimumInterest, _minimumInterest));
                parameters.Add(new Parameter(Param_OriginationDiscount, _originationDiscount));
                parameters.Add(new Parameter(Param_OriginationDiscount2, _originationDiscount2));
                parameters.Add(new Parameter(Param_OriginationDiscountNumDays, _originationDiscountNumDays));
                parameters.Add(new Parameter(Param_OriginationDiscountNumDays2, _originationDiscountNumDays2));
                parameters.Add(new Parameter(Param_InterestBasedOnAdvance, _interestBasedOnAdvance));
                parameters.Add(new Parameter(Param_OriginationBasedOnAdvance, _originationBasedOnAdvance));
                parameters.Add(new Parameter(Param_NoInterest, _noInterest));
                parameters.Add(new Parameter(Param_InterestFee, _interestFee));
                parameters.Add(new Parameter(Param_OriginationFee, _originationFee));
                parameters.Add(new Parameter(Param_UnderwritingFee, _underwritingFee));
                parameters.Add(new Parameter(Param_CustSvcUnderwritingDiscount, _custSvcUnderwritingDiscount));
                parameters.Add(new Parameter(Param_CustSvcInterestDiscount, _custSvcInterestDiscount));
                parameters.Add(new Parameter(Param_CustSvcOriginationDiscount, _custSvcOriginationDiscount));
                parameters.Add(new Parameter(Param_ClientPrimeRate, _clientPrimeRate));
                parameters.Add(new Parameter(Param_ClientPrimeRateSpread, _clientPrimeRateSpread));

                //Set the Exception Message  
                errMsg.Append("Error occured while saving the Fees for  -:");
                errMsg.Append("Loan:" + LoanID.ToString());
                result.Status = true;

                //Save Data
                LoanID = db.Update(SP_TRANSACTION_SAVE, parameters);

                if (LoanID == AppConstants.FAILURE)
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
