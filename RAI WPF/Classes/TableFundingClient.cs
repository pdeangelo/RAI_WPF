using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Clearwater.DataAccess;
using System.Data.SqlClient;


    public class TableFundingClient
    {
        #region Constants
        /// <summary>
        /// Define all the procedure name here
        /// </summary>
        internal const String SP_TRANSACTION_SELECT = "TableFundingClient_Sel";
        internal const String SP_TRANSACTION_SAVE = "TableFundingClient_AddEdit";

        /// <summary>
        /// Define all the procedure parameter fields as a constant here
        /// </summary>
        internal const String Param_ClientID = "ClientID";
        internal const String Param_ClientName = "ClientName";
        internal const String Param_AdvanceRate = "AdvanceRate";
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
        #endregion

        #region Member Variables

        private int _clientID;
        private String _clientName;
        private double _advanceRate;
        private double _minimumInterest;
        private double _clientPrimeRate;
        private double _clientPrimeRateSpread;
        private double _originationDiscount;
        private double _originationDiscount2;
        private int _originationDiscountNumDays;
        private int _originationDiscountNumDays2;
        //private TableFundingClientEntities _clientEntities;
        private bool _interestBasedOnAdvance;
        private bool _originationBasedOnAdvance;
        private bool _noInterest;
        #endregion

        #region Constructor
        public TableFundingClient()
        {
            Initialize();
        }
        public TableFundingClient(int clientID)
        {
            //intialize the member variables
            Initialize();
            if (clientID > -1)
                LoadData(clientID);
        }
        #endregion

        #region Internal Functions
        private void Initialize()
        {
            _clientID = 0;
            _clientName = "";
            _advanceRate = 0;
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
            //_clientEntities = new TableFundingClientEntities();
        }
        private void LoadData(int clientID)
        {
            SqlDataReader dr = null;

            try
            {

                //Fill the procedure parameters
                ParameterCollection parmeters = new ParameterCollection();
                parmeters.Add(new Parameter(Param_ClientID, clientID));

                dr = (SqlDataReader)DBHelper.ExecuteReader(SP_TRANSACTION_SELECT, parmeters);

                if (dr != null)
                {
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            _clientID = (int)dr[Param_ClientID];
                            _clientName = dr[Param_ClientName].ToString();
                            _advanceRate = (double)dr[Param_AdvanceRate];
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
                           // _clientEntities = new TableFundingClientEntities(clientID);
                           

                        }//End of dr.Read

                    }//End of dr.HasRows

                }//End of dr!= Null
            }
            catch (Exception ex)
            {
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
        public int ClientID
        {
            get { return _clientID; }
            set { _clientID = value; }
        }

        /// <summary>
        /// Get/Set the Loan Number value 
        /// </summary>
        public String ClientName
        {
            get { return _clientName; }
            set { _clientName = value; }
        }
        
        public double AdvanceRate
        {
            get { return _advanceRate; }
            set { _advanceRate = value; }
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

        //public TableFundingClientEntities ClientEntities
        //{
        //    get { return _clientEntities; }
        //    set { _clientEntities = value; }
        //}
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

                parameters.Add(new Parameter(Param_ClientID, _clientID));
                parameters.Add(new Parameter(Param_ClientName, _clientName));
                parameters.Add(new Parameter(Param_AdvanceRate, _advanceRate));
                parameters.Add(new Parameter(Param_MinimumInterest, _minimumInterest));
                parameters.Add(new Parameter(Param_OriginationDiscount, _originationDiscount));
                parameters.Add(new Parameter(Param_OriginationDiscount2, _originationDiscount2));
                parameters.Add(new Parameter(Param_OriginationDiscountNumDays, _originationDiscountNumDays));
                parameters.Add(new Parameter(Param_OriginationDiscountNumDays2, _originationDiscountNumDays2));
                parameters.Add(new Parameter(Param_InterestBasedOnAdvance, _interestBasedOnAdvance));
                parameters.Add(new Parameter(Param_NoInterest, _noInterest));
                parameters.Add(new Parameter(Param_OriginationBasedOnAdvance, _originationBasedOnAdvance));
                parameters.Add(new Parameter(Param_ClientPrimeRate, _clientPrimeRate));
                parameters.Add(new Parameter(Param_ClientPrimeRateSpread, _clientPrimeRateSpread));

                //Set the Exception Message  
                errMsg.Append("Error occured while saving the Client for  -:");
                errMsg.Append("Client Name:" + ClientName.ToString());
                result.Status = true;

                //Save Data
                ClientID = db.Update(SP_TRANSACTION_SAVE, parameters);

                if (ClientID == AppConstants.FAILURE)
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
                errMsg.Append(ex.Message);
                result.Message = errMsg.ToString();
                result.Status = false;

                db.RollBack();  //Roll back the transaction
            }

            return result;

        }

        #endregion
    }
