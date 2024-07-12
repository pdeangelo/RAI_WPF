using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


    /// ---------------------------------------------
    /// Project		: CashFlow.BusinessEntity.WebSecurity.
    /// Class		: UserWebPages
    /// ---------------------------------------------
    /// <summary>
    /// Represent a single instance for the UserWebPages
    /// </summary>
    /// <history>
    /// 	[rrb] 		Created
    /// </history>
    /// --------------------------------------------------
    public class UserPermission
    {
        #region Constants
        #endregion


        #region Member Variables

        //private UserRolesCollection _roles;
        //private UserWebPageCollection _accessiblePages;

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor. Initialize the class member variables
        /// </summary>
        public UserPermission()
        {
            //Initialize the member variables
            Initialize();
        }

        /// <summary>
        /// Constructor. Initialize the class member variables
        /// </summary>
        public UserPermission(int userID)
        {
            //Initialize the member variables
            Initialize();

            //_roles = new UserRolesCollection(userID);
            //_accessiblePages = new UserWebPageCollection(userID);
        }

        #endregion

        #region Public Property

        /// <summary>
        /// Get the Collection of user assigned roles
        /// </summary>
        //public UserRolesCollection Roles
        //{
        //    get { return _roles; }
        //    set { _roles = value; }
        //}

        /// <summary>
        /// Get the Collection of user assigned web pages
        /// </summary>
        //public UserWebPageCollection AccessiblePages
        //{
        //    get { return _accessiblePages; }
        //    set { _accessiblePages = value; }
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
            //_roles = new UserRolesCollection();
            //_accessiblePages = new UserWebPageCollection();
        }
        #endregion
    }
