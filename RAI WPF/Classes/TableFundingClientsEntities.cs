using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using log4net;

using Clearwater.DataAccess;


    /// ---------------------------------------------
    /// Project		: CashFlow.TableFunding
    /// Class		: Table funding Client Collection
    /// ---------------------------------------------
    /// <summary>
    /// Represent collection of the security type
    /// </summary>
    /// <history>
    /// 	[rrb] 		Created
    /// </history>
    /// --------------------------------------------------
    public class TableFundingClientEntities: List<TableFundingClientEntity>
      {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
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
        public TableFundingClientEntities()
            {
                //intialize the member variables
                Initialize();
                //LoadData();
            }
            public TableFundingClientEntities(int clientID)
            {
                //intialize the member variables
                Initialize();
                LoadData(clientID);
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
      
        private void LoadData()
        {
            try
            {
                ParameterCollection parameters = new ParameterCollection();

                DBHelper.ExecuteToList<TableFundingClientEntity>(TableFundingClientEntity.SP_TRANSACTION_SELECT, parameters, this);

            }
            catch (Exception ex)
            {
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
                throw ex;
            }

        }
        private void LoadData(int clientID)
        {
            try
            {
                ParameterCollection parameters = new ParameterCollection();
                parameters.Add(new Parameter(TableFundingClientEntity.Param_ClientID, clientID));
                DBHelper.ExecuteToList<TableFundingClientEntity>(TableFundingClientEntity.SP_TRANSACTION_SELECT, parameters, this);

            }
            catch (Exception ex)
            {
                log.Error("Error in logon - Exception Message:" + ex.Message + " Exception: " + ex);
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
