using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Clearwater.DataAccess;
using System.Data;
using System.Data.SqlClient;
using log4net;
public class TableFundingLoanUW
    {
        #region Constants
        /// <summary>
        /// Define all the procedure name here
        /// </summary>
        internal const String SP_TRANSACTION_SELECT = "TableFundingLoansUW_SEL";
        internal const String SP_TRANSACTION_SAVE = "TableFundingLoansUW_AddEdit";
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


    /// <summary>
    /// Define all the procedure parameter fields as a constant here
    /// </summary>
    internal const String Param_LoanUWID = "LoanUWID";
        internal const String Param_LoanID = "LoanID";
        internal const String Param_LoanUW10031008LoanApplication = "LoanUW10031008LoanApplication";
        internal const String Param_LoanUWAllongePromissoryNote = "LoanUWAllongePromissoryNote";
        internal const String Param_LoanUWAppraisal = "LoanUWAppraisal";
        internal const String Param_LoanUWAppraisalAmount = "LoanUWAppraisalAmount";
        internal const String Param_LoanUWPostRepairAppraisalAmount = "LoanUWPostRepairAppraisalAmount";
        internal const String Param_LoanUWAssignmentOfMortgage = "LoanUWAssignmentOfMortgage";
        internal const String Param_LoanUWBackgroundCheck = "LoanUWBackgroundCheck";
        internal const String Param_LoanUWCertofGoodStandingFormation = "LoanUWCertofGoodStandingFormation";
        internal const String Param_LoanUWClosingProtectionLetter = "LoanUWClosingProtectionLetter";
        internal const String Param_LoanUWCommitteeReview = "LoanUWCommitteeReview";
        internal const String Param_LoanUWCreditReport = "LoanUWCreditReport";
        internal const String Param_LoanUWFloodCertificate = "LoanUWFloodCertificate";
        internal const String Param_LoanUWHUD1SettlementStatement = "LoanUWHUD1SettlementStatement";
        internal const String Param_LoanUWHomeownersInsurance = "LoanUWHomeownersInsurance";
        internal const String Param_LoanUWLoanPackage = "LoanUWLoanPackage";
        internal const String Param_LoanUWLoanSizerLoanSubmissionTape = "LoanUWLoanSizerLoanSubmissionTape";
        internal const String Param_LoanUWTitleCommitment = "LoanUWTitleCommitment";
        internal const String Param_LoanUWClaytonReportApprovalEmail = "LoanUWClaytonReportApprovalEmail";
        internal const String Param_LoanUWZillowSqFt = "LoanUWZillowSqFt";
        internal const String Param_LoanUWIsComplete = "LoanUWIsComplete";
        internal const String Param_CompletedBy = "CompletedBy";
        internal const String Param_CompletedByName = "CompletedByName";

        #endregion

        #region Member Variables
        private int _loanUWID;
        private int _loanID;
        private int _completedBy;
        private string _completedByName;
        private bool _loanUW10031008LoanApplication;
        private bool _loanUWAllongePromissoryNote;
        private bool _loanUWAppraisal;
        private double _loanUWAppraisalAmount;
        private double _loanUWPostRepairAppraisalAmount;
        private bool _loanUWAssignmentOfMortgage;
        private bool _loanUWBackgroundCheck;
        private bool _loanUWCertofGoodStandingFormation;
        private bool _loanUWClosingProtectionLetter;
        private bool _loanUWCommitteeReview;
        private bool _loanUWCreditReport;
        private bool _loanUWFloodCertificate;
        private bool _loanUWHUD1SettlementStatement;
        private bool _loanUWHomeownersInsurance;
        private bool _loanUWLoanPackage;
        private bool _loanUWLoanSizerLoanSubmissionTape;
        private bool _loanUWTitleCommitment;
        private bool _loanUWClaytonReportApprovalEmail;
        private double _loanUWZillowSqFt;
        private bool _loanUWIsComplete;
        

        #endregion

        #region Constructor
        /// <summary>
        /// Constructor. 
        /// </summary>
        public TableFundingLoanUW()
        {
        }

        public TableFundingLoanUW(int loanID)
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
        public int LoanUWID
        {
            get { return _loanUWID; }
            set { _loanUWID = value; }
        }
        public int LoanID
        {
            get { return _loanID; }
            set { _loanID = value; }
        }
        public int CompletedBy
        {
            get { return _completedBy; }
            set { _completedBy = value; }
        }
        public string CompletedByName
        {
            get { return _completedByName; }
            set { _completedByName = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUW10031008LoanApplication
        {
            get { return _loanUW10031008LoanApplication; }
            set { _loanUW10031008LoanApplication = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWAllongePromissoryNote
        {
            get { return _loanUWAllongePromissoryNote; }
            set { _loanUWAllongePromissoryNote = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWAppraisal
        {
            get { return _loanUWAppraisal; }
            set { _loanUWAppraisal = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public double LoanUWAppraisalAmount
        {
            get { return _loanUWAppraisalAmount; }
            set { _loanUWAppraisalAmount = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public double LoanUWPostRepairAppraisalAmount
        {
            get { return _loanUWPostRepairAppraisalAmount; }
            set { _loanUWPostRepairAppraisalAmount = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWAssignmentOfMortgage
        {
            get { return _loanUWAssignmentOfMortgage; }
            set { _loanUWAssignmentOfMortgage = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWBackgroundCheck
        {
            get { return _loanUWBackgroundCheck; }
            set { _loanUWBackgroundCheck = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWCertofGoodStandingFormation
        {
            get { return _loanUWCertofGoodStandingFormation; }
            set { _loanUWCertofGoodStandingFormation = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWClosingProtectionLetter
        {
            get { return _loanUWClosingProtectionLetter; }
            set { _loanUWClosingProtectionLetter = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWCommitteeReview
        {
            get { return _loanUWCommitteeReview; }
            set { _loanUWCommitteeReview = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWCreditReport
        {
            get { return _loanUWCreditReport; }
            set { _loanUWCreditReport = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWFloodCertificate
        {
            get { return _loanUWFloodCertificate; }
            set { _loanUWFloodCertificate = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWHUD1SettlementStatement
        {
            get { return _loanUWHUD1SettlementStatement; }
            set { _loanUWHUD1SettlementStatement = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWHomeownersInsurance
        {
            get { return _loanUWHomeownersInsurance; }
            set { _loanUWHomeownersInsurance = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWLoanPackage
        {
            get { return _loanUWLoanPackage; }
            set { _loanUWLoanPackage = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWLoanSizerLoanSubmissionTape
        {
            get { return _loanUWLoanSizerLoanSubmissionTape; }
            set { _loanUWLoanSizerLoanSubmissionTape = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWTitleCommitment
        {
            get { return _loanUWTitleCommitment; }
            set { _loanUWTitleCommitment = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWClaytonReportApprovalEmail
        {
            get { return _loanUWClaytonReportApprovalEmail; }
            set { _loanUWClaytonReportApprovalEmail = value; }
        }
        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public double LoanUWZillowSqFt
        {
            get { return _loanUWZillowSqFt; }
            set { _loanUWZillowSqFt = value; }
        }

        /// <summary>
        /// Get the Security Type associated with the transaction
        /// </summary>
        public bool LoanUWIsComplete
        {
            get { return _loanUWIsComplete; }
            set { _loanUWIsComplete = value; }
        } /// <summary>
       
        #endregion
        #region Internal Functions
        private void Initialze()
        {
            _loanUWZillowSqFt = 0;
            _loanUWAppraisalAmount = 0;
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
                            _loanUWID = (int)dr[Param_LoanUWID];
                            _loanID = (int)dr[Param_LoanID];
                            _loanUW10031008LoanApplication = (bool)dr[Param_LoanUW10031008LoanApplication];
                            _loanUWAllongePromissoryNote = (bool)dr[Param_LoanUWAllongePromissoryNote];
                            _loanUWAppraisal = (bool)dr[Param_LoanUWAppraisal];
                            _loanUWAppraisalAmount = (double)dr[Param_LoanUWAppraisalAmount];
                            _loanUWPostRepairAppraisalAmount = (double)dr[Param_LoanUWPostRepairAppraisalAmount];
                            _loanUWAssignmentOfMortgage = (bool)dr[Param_LoanUWAssignmentOfMortgage];
                            _loanUWBackgroundCheck = (bool)dr[Param_LoanUWBackgroundCheck];
                            _loanUWCertofGoodStandingFormation = (bool)dr[Param_LoanUWCertofGoodStandingFormation];
                            _loanUWClosingProtectionLetter = (bool)dr[Param_LoanUWClosingProtectionLetter];
                            _loanUWCommitteeReview = (bool)dr[Param_LoanUWCommitteeReview];
                            _loanUWCreditReport = (bool)dr[Param_LoanUWCreditReport];
                            _loanUWFloodCertificate = (bool)dr[Param_LoanUWFloodCertificate];
                            _loanUWHUD1SettlementStatement = (bool)dr[Param_LoanUWHUD1SettlementStatement];
                            _loanUWHomeownersInsurance = (bool)dr[Param_LoanUWHomeownersInsurance];
                            _loanUWLoanPackage = (bool)dr[Param_LoanUWLoanPackage];
                            _loanUWLoanSizerLoanSubmissionTape = (bool)dr[Param_LoanUWLoanSizerLoanSubmissionTape];
                            _loanUWTitleCommitment = (bool)dr[Param_LoanUWTitleCommitment];
                            _loanUWClaytonReportApprovalEmail = (bool)dr[Param_LoanUWClaytonReportApprovalEmail];
                            _loanUWZillowSqFt = (double)dr[Param_LoanUWZillowSqFt];
                            _loanUWIsComplete = (bool)dr[Param_LoanUWIsComplete];
                            _completedBy = (int)dr[Param_CompletedBy];
                            _completedByName = dr[Param_CompletedByName].ToString();
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

                parameters.Add(new Parameter(Param_LoanUWID, _loanUWID));
                parameters.Add(new Parameter(Param_LoanID, _loanID));
                parameters.Add(new Parameter(Param_LoanUW10031008LoanApplication, _loanUW10031008LoanApplication));
                parameters.Add(new Parameter(Param_LoanUWAllongePromissoryNote, _loanUWAllongePromissoryNote));
                parameters.Add(new Parameter(Param_LoanUWAppraisal, _loanUWAppraisal));
                parameters.Add(new Parameter(Param_LoanUWAppraisalAmount, _loanUWAppraisalAmount));
                parameters.Add(new Parameter(Param_LoanUWPostRepairAppraisalAmount, _loanUWPostRepairAppraisalAmount));
                parameters.Add(new Parameter(Param_LoanUWAssignmentOfMortgage, _loanUWAssignmentOfMortgage));
                parameters.Add(new Parameter(Param_LoanUWBackgroundCheck, _loanUWBackgroundCheck));
                parameters.Add(new Parameter(Param_LoanUWCertofGoodStandingFormation, _loanUWCertofGoodStandingFormation));
                parameters.Add(new Parameter(Param_LoanUWClosingProtectionLetter, _loanUWClosingProtectionLetter));
                parameters.Add(new Parameter(Param_LoanUWCommitteeReview, _loanUWCommitteeReview));
                parameters.Add(new Parameter(Param_LoanUWCreditReport, _loanUWCreditReport));
                parameters.Add(new Parameter(Param_LoanUWFloodCertificate, _loanUWFloodCertificate));
                parameters.Add(new Parameter(Param_LoanUWHUD1SettlementStatement, _loanUWHUD1SettlementStatement));
                parameters.Add(new Parameter(Param_LoanUWHomeownersInsurance, _loanUWHomeownersInsurance));
                parameters.Add(new Parameter(Param_LoanUWLoanPackage, _loanUWLoanPackage));
                parameters.Add(new Parameter(Param_LoanUWLoanSizerLoanSubmissionTape, _loanUWLoanSizerLoanSubmissionTape));
                parameters.Add(new Parameter(Param_LoanUWTitleCommitment, _loanUWTitleCommitment));
                parameters.Add(new Parameter(Param_LoanUWClaytonReportApprovalEmail, _loanUWClaytonReportApprovalEmail));
                parameters.Add(new Parameter(Param_LoanUWZillowSqFt, _loanUWZillowSqFt));
                parameters.Add(new Parameter(Param_LoanUWIsComplete, _loanUWIsComplete));
                parameters.Add(new Parameter(Param_CompletedBy, _completedBy));

                //Set the Exception Message  
                errMsg.Append("Error occured while saving the Loan UW for  -:");
                errMsg.Append("Loan UW ID:" + LoanUWID.ToString());
                result.Status = true;

                //Save Data
                LoanUWID = db.Update(SP_TRANSACTION_SAVE, parameters);

                if (LoanUWID == AppConstants.FAILURE)
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
