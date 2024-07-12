using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient; 

//using CashFlow.Common;
using Clearwater.DataAccess ;


  /// ---------------------------------------------
  /// Project	: CashFlow.BusinessEntity
  /// Class		: AppUser
  /// ---------------------------------------------
  /// <summary>
  /// Represent a single instance for the user
  /// </summary>
  /// <history>
  /// 	[rrb] 		Created
  /// </history>
  /// --------------------------------------------------
  public class AppUser
  {
    #region Constants

    /// <summary>
    /// Define all the procedure name here
    /// </summary>
    const string SP_User_SEL = "Users_SEL";    
    const string SP_UserByWinsUserID_SEL = "UsersByWinUserID_SEL";
    const string SP_User_AddEdit = "Users_AddEdit";
    const string SP_User_CheckUserExist  = "Users_CheckIfUserExist";

    /// <summary>
    /// Define all the procedure parameter fields as a constant here
    /// </summary>
    const string Param_UserID = "@UserID";
    const string Param_UserName = "@UserName";
    const string Param_WinUserID = "@WinUserID";
    const string Param_RoleID = "@RoleID";
    const string Param_RoleName = "@RoleName";
    const string Param_OfficeID = "@OfficeID";
    const string Param_Email = "@Email";
    const string Param_IsADmin = "@IsAdmin";
    const string Param_Status = "@Status";
    const string Param_FinanceGroup = "@FinanceGroup";
    const string Param_Client = "@Client";
    const string Param_Password = "@Password";

    #endregion

    #region Members Variables
        private int _userID = 0;
        private string _userName = "";
        private string _winUserID = "";
        private int _roleID = 0;
        private string _roleName = "";
        private int _officeID = 0;
        private string _email = "";
        private bool _isADmin = false;
        private int _status = 0;
        private int _financeGroup = 0;
        private int _client = 0;
        //private UserPermission _userPermission;
        private bool _allSecurityAccess = false;
        private bool _allPipelineAccess = false;
        private bool _isAnalyst = false;
        private bool _isFinance = false;
        private bool _isManager = false;
        private bool _fundLevelAccess = false;
        private string _password = "";
    #endregion

    #region Constructor
    /// <summary>
    /// Default Constructor. Creates the 
    /// </summary>
    public AppUser()
    {
      Initialze();
    }

    /// <summary>
    /// Constructor. 
    /// </summary>
    public AppUser(int UserID)
    {
      //initialize the private members
      Initialze();

      //if user id is -1 then don't load the data
      if (UserID == AppConstants.NewRecord)
        return ;

      //Load the memeber variable for the given user
      LoadData(UserID, -1);
    }

    /// <summary>
    /// Constructor. 
    /// </summary>
    public AppUser(int UserID, int roleID)
    {
        //initialize the private members
        Initialze();

        //if user id is -1 then don't load the data
        if (UserID == AppConstants.NewRecord)
            return;

        //Load the memeber variable for the given user
        LoadData(UserID, roleID);
    }

    /// <summary>
    /// Constructor. 
    /// </summary>
    public AppUser(string email, bool emailCon)
    {
        //initialize the private members
        Initialze();

        //Load the memeber variable for the given user
        LoadDataEmail(email);
    }
    /// <summary>
    /// Constructor. Initialize the object by loading the data for the given windows user id.
    /// </summary>
    /// <param name="WindUserID">String</param>
    public AppUser(string WindUserID)
    {
        //initialize the private members
        Initialze();

        //Load the memeber variable for the given user
        LoadData(WindUserID);
    }
    
    #endregion

    #region Public Property
    /// <summary>
    /// Gets the User id of the User
    /// </summary>
    public int UserID
    {
      get { return _userID; }
      set { _userID = value; }
    }
    /// <summary>
    /// Gets/sets the User name
    /// </summary>
    public string UserName
    {
      get { return _userName; }
      set { _userName = value; }
    }
    /// <summary>
    /// Gets/sets the Windows User ID
    /// </summary>
    public string WinUserID
    {
      get { return _winUserID; }
      set { _winUserID = value; }
    }

    /// <summary>
    /// Gets/sets the Role ID of the User
    /// </summary>
    public int RoleID
    {
      get { return _roleID; }
      set { _roleID = value; }
    }

    /// <summary>
    /// Gets/sets the CLient of the User
    /// </summary>
    public int Client
    {
        get { return _client; }
        set { _client = value; }
    }

    /// <summary>
    /// Gets the role name associated with the user.
    /// </summary>
    public string RoleName
    {
        get { return _roleName; } 
    }    

    /// <summary>
    /// Gets/sets the Office ID of the user
    /// </summary>
    public int OfficeID
    {
      get { return _officeID; }
      set { _officeID = value; }
    }
    /// <summary>
    /// Gets/sets the Email address of the user
    /// </summary>
    public string Email
    {
      get { return _email; }
      set { _email = value; }
    }

    /// <summary>
    /// Gets/sets the user type value.
    /// </summary>
    public bool IsADmin
    {
      get { return _isADmin; }
      set { _isADmin = value; }
    }

    /// <summary>
    /// Gets/sets the user type value.
    /// </summary>
    public bool AllSecurityAccess
    {
        get { return _allSecurityAccess; }
        set { _allSecurityAccess = value; }
    }

    /// <summary>
    /// Gets/sets the user type value.
    /// </summary>
    public bool AllPipelineAccess
    {
        get { return _allPipelineAccess; }
        set { _allPipelineAccess = value; }
    }

    /// <summary>
    /// Gets/sets the user type value.
    /// </summary>
    public bool IsAnalyst
    {
        get { return _isAnalyst; }
        set { _isAnalyst = value; }
    }

    /// <summary>
    /// Gets/sets the user type value.
    /// </summary>
    public bool IsFinance
    {
        get { return _isFinance; }
        set { _isFinance = value; }
    }

    /// <summary>
    /// Gets/sets the user type FundLevelAccess.
    /// </summary>
    public bool FundLevelAccess
    {
        get { return _fundLevelAccess; }
        set { _fundLevelAccess = value; }
    }
    
      /// <summary>
    /// Gets/sets the user type value.
    /// </summary>
    public bool IsManager
    {
        get { return _isManager; }
        set { _isManager = value; }
    }
    /// <summary>
    /// Gets/sets the status of user.
    /// </summary>
    public int Status
    {
      get { return _status; }
      set { _status = value; }
    }

    /// <summary>
    /// Gets/sets the Finance group.
    /// </summary>
    public int FinanceGroup
    {
        get { return _financeGroup; }
        set { _financeGroup = value; }
    }

    /// <summary>
    /// Gets bool value indicating whether the object data is valid or not
    /// </summary>
    public bool IsValidUser
    {
        get
        {
            if (_userID == AppConstants.NewRecord)
                return false;
            else
                return true;

        }
    }
    /// <summary>
    /// Gets/sets the password of the user
    /// </summary>
    public string Password
    {
        get { return _password; }
        set { _password = value; }
    }

    /// <summary>
    /// Get the object UserPermission which has colection of user accessible web pages and roles
    /// </summary>
    //public UserPermission UserPermissions
    //{
    //    get { return _userPermission; }
    //    set { _userPermission = value; }
    //}

    #endregion

    #region Public Method
        /// <summary>
        /// Insert/Update User Details.
        /// </summary>
        /// <returns></returns>
        public ReturnResult  Save()
        {
            Int64 returnID = 0;

            ReturnResult result = new ReturnResult();

            //Check if the user record already exists
            result  = AppUser.CheckUserExist(_userID, _userName, _winUserID, _email  );
            if (result.Status == false)
                return result;
            
            //START SAVING THE RECORD

            DBHelper db = new DBHelper();
            ParameterCollection parameters = new ParameterCollection();
            parameters.Add(new Parameter(Param_UserID, _userID));
            parameters.Add(new Parameter(Param_UserName, _userName));
            parameters.Add(new Parameter(Param_WinUserID, _winUserID));
            parameters.Add(new Parameter(Param_RoleID, _roleID));
            parameters.Add(new Parameter(Param_OfficeID, _officeID));
            parameters.Add(new Parameter(Param_Email, _email));
            parameters.Add(new Parameter(Param_IsADmin, _isADmin));
            parameters.Add(new Parameter(Param_Status, _status));
            parameters.Add(new Parameter(Param_FinanceGroup, _financeGroup));


            parameters.Add(new Parameter("AllSecurityAccess", _allSecurityAccess));
            parameters.Add(new Parameter("AllPipelineAccess", _allPipelineAccess));
            parameters.Add(new Parameter("IsAnalyst", _isAnalyst));
            parameters.Add(new Parameter("IsFinance", _isFinance));
            parameters.Add(new Parameter("IsManager", _isManager));
            parameters.Add(new Parameter("FundLevelAccess", _fundLevelAccess));
            parameters.Add(new Parameter(Param_Client, _client));
            parameters.Add(new Parameter(Param_Password, _password));
            returnID = db.Update(SP_User_AddEdit, parameters);
            db.Dispose();

            if (returnID == AppConstants.FAILURE)
            {
                result.Status = false;
                result.Message = "Error occurred while saving the user record.";
            }

            else
            {
                result.Status = true ;
                result.Message = "User record saved successfully.";
                _userID = Convert.ToInt16(returnID);
            }

            return result;
        }

        /// <summary>
        /// Populate user permission data on demand
        /// </summary>
        /// <param name="userID"></param>
        //public void LoadUserPermission()
        //{
        //    //Filled UserPermission object for given userID
        //    _userPermission = new UserPermission(_userID);
        //}
 
    #endregion

    #region Private Methods
        private void Initialze()
        {
          _userID = AppConstants.NewRecord;
          _userName = " ";
          _winUserID = " ";
          _roleID = 0;
          _roleName = "";
          _officeID = 0;
          _email = " ";
          _isADmin = false;
          _status = 0;
          _financeGroup = 0;
          _client = 0;
          _allPipelineAccess = false;
          _allSecurityAccess = false;
          _isAnalyst = false;
          _isFinance = false;
          _isManager = false;
          //_userPermission = null;
          _fundLevelAccess = false;
          _password = "";
        }

        /// <summary>
        ///User details based on User ID.
        /// </summary>
        /// <return> </return>
        private void LoadData(int UserID, int roleId)
        {
            SqlDataReader dr = null;
            //Fill the procedure parameters
            try
            {

                ParameterCollection parmeters = new ParameterCollection();
                parmeters.Add(new Parameter(Param_UserID, UserID));
                parmeters.Add(new Parameter(Param_RoleID, roleId));

                dr = (SqlDataReader)DBHelper.ExecuteReader(SP_User_SEL, parmeters);

                //Set member variable
                SetData(dr);
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

        /// <summary>
        ///User details based on Win User ID.
        /// </summary>
        /// <return> </return>
        private void LoadData(string WinUserID)
        {
            SqlDataReader dr = null;

            try
            {
                //Fill the procedure parameters
                ParameterCollection parmeters = new ParameterCollection();
                parmeters.Add(new Parameter(Param_WinUserID, WinUserID));

                dr = (SqlDataReader)DBHelper.ExecuteReader(SP_UserByWinsUserID_SEL, parmeters);

                //Set member variable
                SetData(dr);
            }
            catch (Exception ex)
            {
                //Close the connection
                DBHelper.DisposeOff(dr);
                throw ex;
            }
            finally
            {
                //Close the connection
                DBHelper.DisposeOff(dr);
            }
        }
        /// <summary>
        ///User details based on Win User ID.
        /// </summary>
        /// <return> </return>
        private void LoadDataEmail(string email)
        {
            SqlDataReader dr = null;

            try
            {
                //Fill the procedure parameters
                ParameterCollection parmeters = new ParameterCollection();
                parmeters.Add(new Parameter(Param_UserID, -1));
                parmeters.Add(new Parameter(Param_RoleID, -1));
                parmeters.Add(new Parameter(Param_Status, -1));
                parmeters.Add(new Parameter(Param_Email, email));

                dr = (SqlDataReader)DBHelper.ExecuteReader(SP_User_SEL, parmeters);

                //Set member variable
                SetData(dr);
            }
            catch (Exception ex)
            {
                //Close the connection
                DBHelper.DisposeOff(dr);
                throw ex;
            }
            finally
            {
                //Close the connection
                DBHelper.DisposeOff(dr);
            }
        }

        /// <summary>
        /// Load the object by setting the respective value for the data reader.
        /// </summary>
        /// <param name="dr">SqlDataReader</param>
        private void SetData(SqlDataReader dr)
        {
            if (dr != null)
            {
                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        _userID = (int)dr["UserID"];
                        _userName = dr["UserName"].ToString();
                        _winUserID = dr["WinUserID"].ToString();
                        _roleID = (int)dr["RoleID"];
                        _roleName = dr["RoleName"].ToString(); 
                        _officeID = (int)dr["OfficeID"];
                        _email = dr["Email"].ToString();
                        _isADmin = (bool)dr["IsADmin"];
                        _status = (int)dr["Status"];
                        _financeGroup = (int)dr["FinanceGroup"];
                        _client = (int)dr["Client"];
                        _allSecurityAccess = (bool)dr["AllSecurityAccess"];
                        _allPipelineAccess = (bool)dr["AllPipelineAccess"];
                        _isAnalyst = (bool)dr["IsAnalyst"];
                        _isFinance = (bool)dr["IsFinance"];
                        _isManager = (bool)dr["IsManager"];
                        _fundLevelAccess = (bool)dr["FundLevelAccess"];
                        _password = dr["Password"].ToString();
                    }//End of dr.Read

                    ////Filled UserPermission object for given userID
                    //_userPermission = new UserPermission(_userID);

                }//End of dr.HasRows
            }//End of dr!= Null
        }
    #endregion

    #region Public Static Methods
    /// <summary>
    /// Static. Return the dataset object containing the all the Users Stored in the application.
    /// </summary>
    /// <returns></returns>
        public static DataSet GetUsers(int status)
        {
          DataSet ds = new DataSet();
          ParameterCollection parameters = new ParameterCollection();
          parameters.Add(new Parameter(Param_UserID, -1));
          parameters.Add(new Parameter(Param_RoleID, -1));
          parameters.Add(new Parameter(Param_Status, status)); 
          ds = DBHelper.ExecuteDataset(SP_User_SEL,   parameters);   
          return ds;           
        }

      /// <summary>
      /// Get users By Role
      /// </summary>
      /// <returns></returns>
        public static DataSet GetUsers(int roleID, int status)
        {
            DataSet ds = new DataSet();
            ParameterCollection parameters = new ParameterCollection();
            parameters.Add(new Parameter(Param_UserID, -1));
            parameters.Add(new Parameter(Param_RoleID, roleID));
            parameters.Add(new Parameter(Param_Status, status));
            ds = DBHelper.ExecuteDataset(SP_User_SEL, parameters);
            return ds;
        }

        /// <summary>
        /// Get users By Role
        /// </summary>
        /// <returns></returns>
        public static ReturnResult  CheckUserExist(int userID, String userName, String winUserName, String email )
        {
            long value;
            ReturnResult result = new ReturnResult(); 
            ParameterCollection parameters = new ParameterCollection();
            parameters.Add(new Parameter(Param_UserID, userID));
            parameters.Add(new Parameter(Param_UserName  , userName  ));
            parameters.Add(new Parameter(Param_WinUserID, winUserName));
            parameters.Add(new Parameter(Param_Email , email));
            value = DBHelper.ExecuteScaler(SP_User_CheckUserExist, parameters);

            switch (value)
            {
                case 1:
                    result.Status = false;
                    result.Message = "User name already exist.";  
                    break;
                case 2:
                    result.Status = false;
                    result.Message = "Windows User name already assinged to another user.";  
                    break;
                case 3:
                    result.Status = false;
                    result.Message = "Email address already assinged to another user.";  
                    break;
                default :
                    result.Status = true;
                    result.Message = "";
                    break;

            }

            return result ;
        }
      
  #endregion 

  }
