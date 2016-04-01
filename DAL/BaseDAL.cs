using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace DAL
{
    public class BaseDAL
    {
        public static String DbFile { get; set; }

        public static OleDbConnection getConn()
        {
            string connstr = "Provider=Microsoft.ACE.OLEDB.12.0;Persist Security Info=False;Data Source=" + DbFile;

            OleDbConnection tempconn = new OleDbConnection(connstr);

            try
            {
                tempconn.Open();
            }
            catch (Exception)
            {
                tempconn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + DbFile;
            }

            return tempconn;
        }
    }
}
