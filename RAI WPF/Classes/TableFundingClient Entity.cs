using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Clearwater.DataAccess;
using System.Data;
using System.Data.SqlClient;


    public class TableFundingClientEntity
    {
        #region Constants
        /// <summary>
        /// Define all the procedure name here
        /// </summary>
        internal const String SP_TRANSACTION_SELECT = "TableFundingClientEntity_Sel";
        internal const String SP_TRANSACTION_SAVE = "TableFundingClientEntity_AddEdit";

        /// <summary>
        /// Define all the procedure parameter fields as a constant here
        /// </summary>
        internal const String Param_ClientEntityID = "ClientEntityID";
        internal const String Param_ClientID = "ClientID";
        internal const String Param_EntityID = "EntityID";
        internal const String Param_ClientName = "ClientName";
        internal const String Param_EntityName = "EntityName";
        internal const String Param_ClientNumber = "ClientNumber";
        internal const String Param_ClientEntityBank = "ClientEntityBank";
        internal const String Param_ABA = "ABA";
        internal const String Param_Account = "Account";

        #endregion

        #region Member Variables

        private int _clientEntityID;
        private int _clientID;

        private int _entityID;
        private String _clientName;
        private String _entityName;
        private String _clientNumber;
        private String _clientEntityBank;
        private String _aba;
        private String _account;
        #endregion

        #region Constructor
        public TableFundingClientEntity()
        {
            Initialize();
        }
        public TableFundingClientEntity(int clientID, int entityID)
        {
            //intialize the member variables
            Initialize();
            if (entityID > -1)
                LoadData(clientID, entityID);
        }
        #endregion
        #region Internal Functions
        private void Initialize()
        {
            _clientID = 0;
            _clientName = "";
            _clientNumber = "";
            _clientEntityBank = "";
            _aba = "";
            _account = "";
        }
        private void LoadData(int clientID, int entityID)
        {
            SqlDataReader dr = null;

            try
            {

                //Fill the procedure parameters
                ParameterCollection parmeters = new ParameterCollection();
                parmeters.Add(new Parameter(Param_ClientID, clientID));
                parmeters.Add(new Parameter(Param_EntityID, entityID));

                dr = (SqlDataReader)DBHelper.ExecuteReader(SP_TRANSACTION_SELECT, parmeters);

                if (dr != null)
                {
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            _clientID = (int)dr[Param_ClientID];
                            _entityID = (int)dr[Param_EntityID];
                            _clientName = dr[Param_ClientName].ToString();
                            _clientNumber = dr[Param_ClientNumber].ToString();
                            _clientEntityBank = dr[Param_ClientEntityBank].ToString();
                            _aba = dr[Param_ABA].ToString();
                            _account = dr[Param_Account].ToString();
                            _entityName = dr[Param_EntityName].ToString();


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
        public int ClientEntityID
        {
            get { return _clientEntityID; }
            set { _clientEntityID = value; }
        }
        public int ClientID
        {
            get { return _clientID; }
            set { _clientID = value; }
        }
        public int EntityID
        {
            get { return _entityID; }
            set { _entityID = value; }
        }

        /// <summary>
        /// Get/Set the Loan Number value 
        /// </summary>
        public String ClientName
        {
            get { return _clientName; }
            set { _clientName = value; }
        }
        public String EntityName
        {
            get { return _entityName; }
            set { _entityName = value; }
        }

        public String ClientNumber
        {
            get { return _clientNumber; }
            set { _clientNumber = value; }
        }
        public String ClientEntityBank
        {
            get { return _clientEntityBank; }
            set { _clientEntityBank = value; }
        }
        public String Aba
        {
            get { return _aba; }
            set { _aba = value; }
        }
        public String Account
        {
            get { return _account; }
            set { _account = value; }
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
                parameters.Add(new Parameter(Param_ClientEntityID, _clientEntityID));
                parameters.Add(new Parameter(Param_ClientID, _clientID));
                parameters.Add(new Parameter(Param_EntityID, _entityID));
                parameters.Add(new Parameter(Param_ClientNumber, _clientNumber));
                parameters.Add(new Parameter(Param_ClientEntityBank, _clientEntityBank));
                parameters.Add(new Parameter(Param_ABA, _aba));
                parameters.Add(new Parameter(Param_Account, _account));

                //Set the Exception Message  
                errMsg.Append("Error occured while saving the Client Entity for  -:");
                errMsg.Append("Client Entity:" + ClientEntityID.ToString());
                result.Status = true;

                //Save Data
                ClientEntityID = db.Update(SP_TRANSACTION_SAVE, parameters);

                if (ClientEntityID == AppConstants.FAILURE)
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

