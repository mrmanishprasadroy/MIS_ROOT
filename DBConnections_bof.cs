using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Oracle.DataAccess.Client;

namespace BOFInterface
{
    public class DBConnections
    {
        public string Sql_con_string = System.Configuration.ConfigurationSettings.AppSettings["SqlDB_Con"];
        public string BOF_con_string = System.Configuration.ConfigurationSettings.AppSettings["BOFL2_Con"];
        bool statusMIS;
        public DataTable DBSelectQueryMIS_Table(string QueryMIS)
        {
            DataTable TempDataTableMIS = new DataTable();
            TempDataTableMIS.Clear();
            try
            {
                SqlConnection DbconMIS = new SqlConnection(Sql_con_string);
                SqlDataAdapter TempAdapterMIS = new SqlDataAdapter(QueryMIS, DbconMIS);
                DbconMIS.Open();
                TempAdapterMIS.Fill(TempDataTableMIS);
                DbconMIS.Close();

            }
            catch
            {
                // Pdi_Msg.alert("CoilId Updated Successfully")
            }

            return TempDataTableMIS;
        }


        public bool DBInsertUpdateDeleteMIS(string QueryMIS)
        {
            try
            {
                SqlConnection DbconMIS = new SqlConnection(Sql_con_string);
                DbconMIS.Open();
                SqlCommand cmdMIS = new SqlCommand(QueryMIS, DbconMIS);
                cmdMIS.CommandType = CommandType.Text;
                cmdMIS.ExecuteNonQuery();
                DbconMIS.Close();
                statusMIS = true;
            }
            catch
            {
                statusMIS = false;
            }
            return statusMIS;
        }
        public DataTable DBSelectQueryStatusScreen_Table(string QueryBhushan_StatusScreen)
        {
            DataTable TempTbl_StatusScreen = new DataTable();
            TempTbl_StatusScreen.Clear();
            try
            {
                OracleConnection DbconStatusScreen = new OracleConnection(BOF_con_string);
                OracleDataAdapter TempAdpStatusScreen = new OracleDataAdapter(QueryBhushan_StatusScreen, DbconStatusScreen);
                DbconStatusScreen.Open();
                TempAdpStatusScreen.Fill(TempTbl_StatusScreen);
                DbconStatusScreen.Close();

            }
            catch
            {
                // Pdi_Msg.alert("CoilId Updated Successfully")
            }

            return TempTbl_StatusScreen;
        }

    }
}
