using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Clearwater.DataAccess;
using System.Data.SqlClient;


    public class TableFundingInvestor
    {
        #region Constants
        /// <summary>
        /// Define all the procedure name here
        /// </summary>
        internal const String SP_TRANSACTION_SELECT = "TableFundingInvestor_Sel";
        internal const String SP_TRANSACTION_SAVE = "TableFundingInvestor_AddEdit";

        /// <summary>
        /// Define all the procedure parameter fields as a constant here
        /// </summary>
        internal const String Param_InvestorID = "InvestorID";
        internal const String Param_InvestorName = "InvestorName";
        internal const String Param_CustodianName = "CustodianName";
        #endregion

        #region Member Variables

        private int _investorID;
        private String _InvestorName;
        private String _custodianName;
        #endregion

        #region Constructor
        public TableFundingInvestor()
        {
            Initialize();
        }
        public TableFundingInvestor(int investorID)
        {
            //intialize the member variables
            
            Initialize();
            if (investorID > -1)
                LoadData(investorID);
        }
        #endregion

        #region Internal Functions
        private void Initialize()
        {
            _investorID = 0;
            _InvestorName = "";
            _custodianName = "";
           
        }
        private void LoadData(int investorID)
        {
            SqlDataReader dr = null;

            try
            {

                //Fill the procedure parameters
                ParameterCollection parmeters = new ParameterCollection();
                parmeters.Add(new Parameter(Param_InvestorID, investorID));

                dr = (SqlDataReader)DBHelper.ExecuteReader(SP_TRANSACTION_SELECT, parmeters);

                if (dr != null)
                {
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            _investorID = (int)dr[Param_InvestorID];
                            _InvestorName = dr[Param_InvestorName].ToString();
                            _custodianName = dr[Param_CustodianName].ToString();
                           

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
        public int InvestorID
        {
            get { return _investorID; }
            set { _investorID = value; }
        }
        /// <summary>
        /// Get/Set the Loan Number value 
        /// </summary>
        public String InvestorName
        {
            get { return _InvestorName; }
            set { _InvestorName = value; }
        }

        /// <summary>
        /// Get/Set the Loan Number value 
        /// </summary>
        public String CustodianName
        {
            get { return _custodianName; }
            set { _custodianName = value; }
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

                parameters.Add(new Parameter(Param_InvestorID, _investorID));
                parameters.Add(new Parameter(Param_InvestorName, _InvestorName));
                parameters.Add(new Parameter(Param_CustodianName, _custodianName));
                //Set the Exception Message  
                errMsg.Append("Error occured while saving the Investor for  -:");
                errMsg.Append("Investor Name:" + InvestorName.ToString());
                result.Status = true;

                //Save Data
                InvestorID = db.Update(SP_TRANSACTION_SAVE, parameters);

                if (InvestorID == AppConstants.FAILURE)
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

