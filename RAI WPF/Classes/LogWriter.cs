using System;
using System.Configuration;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Clearwater.DataAccess;
using System.Data.SqlClient;



public static class LogWriter
{
    public static void WriteLog(int loanID, string action, string logText)
    {
        DBHelper db = null;

        try
        {
            //Instantiate the db helper object
            db = new DBHelper();


            ParameterCollection parameters = new ParameterCollection();

            parameters.Add(new Parameter("LoanID", loanID));
            parameters.Add(new Parameter("UserID", MyVariables.user.UserID));
            parameters.Add(new Parameter("Action", action));
            parameters.Add(new Parameter("Message", logText));
            
            //Save Data
            db.Update("LogTable_Write", parameters);

          
        }
        catch (Exception ex)
        {
           
        }


    }
}





