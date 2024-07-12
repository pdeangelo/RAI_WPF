using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Clearwater.DataAccess;
using System.Data;
using System.Data.SqlClient;

    public class TableFundingEntity
    {
        #region Constants
        /// <summary>
        /// Define all the procedure name here
        /// </summary>
        internal const String SP_TRANSACTION_SELECT = "TableFundingEntity_Sel";
        internal const String SP_TRANSACTION_SAVE = "TableFundingEntity_AddEdit";

        /// <summary>
        /// Define all the procedure parameter fields as a constant here
        /// </summary>
        internal const String Param_EntityID = "EntityID";
        internal const String Param_EntityName = "EntityName";
        internal const String Param_EntityBank = "EntityBank";
        internal const String Param_ABA = "ABA";
        internal const String Param_Account = "Account";

        #endregion

        #region Member Variables

        private int _entityID;
        private String _entityName;
        private String _entityBank;
        private String _aba;
        private String _account;
        #endregion

        #region Constructor
        public TableFundingEntity()
        {
            Initialize();
        }
        public TableFundingEntity(int entityID)
        {
            //intialize the member variables
            Initialize();
            if (entityID > -1)
                LoadData(entityID);
        }
        #endregion
        #region Internal Functions
        private void Initialize()
        {
            _entityBank = "";
            _aba = "";
            _account = "";
        }
        private void LoadData(int entityID)
        {
            SqlDataReader dr = null;

            try
            {

                //Fill the procedure parameters
                ParameterCollection parmeters = new ParameterCollection();
                parmeters.Add(new Parameter(Param_EntityID, entityID));

                dr = (SqlDataReader)DBHelper.ExecuteReader(SP_TRANSACTION_SELECT, parmeters);

                if (dr != null)
                {
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            _entityID = (int)dr[Param_EntityID];
                            _entityBank = dr[Param_EntityBank].ToString();
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
     
        public int EntityID
        {
            get { return _entityID; }
            set { _entityID = value; }
        }

     
        public String EntityName
        {
            get { return _entityName; }
            set { _entityName = value; }
        }
        public String EntityBank
        {
            get { return _entityBank; }
            set { _entityBank = value; }
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
                parameters.Add(new Parameter(Param_EntityID, _entityID));
                parameters.Add(new Parameter(Param_EntityBank, _entityBank));
                parameters.Add(new Parameter(Param_ABA, _aba));
                parameters.Add(new Parameter(Param_Account, _account));

                //Set the Exception Message  
                errMsg.Append("Error occured while saving the Entity for  -:");
                errMsg.Append("Entity:" + EntityID.ToString());
                result.Status = true;

                //Save Data
                EntityID = db.Update(SP_TRANSACTION_SAVE, parameters);

                if (EntityID == AppConstants.FAILURE)
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
