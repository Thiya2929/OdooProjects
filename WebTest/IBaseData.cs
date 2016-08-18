using System;
using System.Collections.Generic;
using System.Web;
using System.Data.SqlClient;

namespace DAT_HHD
{
    public interface IBaseData
    {
        void GetConnection();
        void CloseConnection();
        string convertto(string str);
        string convertfrom(string str);
        SqlTransaction begintrans();
        void commit();
        string ExceptionLog(string error);


    }
}
