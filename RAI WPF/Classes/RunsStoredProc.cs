using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Clearwater.DataAccess;
using System.Data;

class RunsStoredProc
    {
    public static DataSet RunStoredProc(string procName, string param1Name, string param1, string param2Name, string param2, string param3Name, string param3, string param4Name, string param4, string param5Name, string param5, string param6Name, string param6, string param7Name, int param7)
    {
        DataSet ds = new DataSet();
        ParameterCollection parameters = new ParameterCollection();
        if (param1Name.Length > 0)
            parameters.Add(new Parameter(param1Name, param1));
        if (param2Name.Length > 0)
            parameters.Add(new Parameter(param2Name, param2));
        if (param3Name.Length > 0)
            parameters.Add(new Parameter(param3Name, param3));
        if (param4Name.Length > 0)
            parameters.Add(new Parameter(param4Name, param4));
        if (param5Name.Length > 0)
            parameters.Add(new Parameter(param5Name, param5));
        if (param6Name.Length > 0)
            parameters.Add(new Parameter(param6Name, param6));
        if (param7Name.Length > 0)
            parameters.Add(new Parameter(param7Name, param7));

        ds = DBHelper.ExecuteDataset(procName, parameters);
        return ds;
    }
}
