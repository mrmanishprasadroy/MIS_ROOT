using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.UI;

using System.Xml.Linq;
using System.Data.SqlClient;
using Oracle.DataAccess.Client;
/// <summary>
/// Summary description for DBConnections
/// </summary>
public class DBConnections
{
    //string con_stringMIS = System.Configuration.ConfigurationSettings.AppSettings["ConnectionStringMIS"];
    //string con_stringSMSII = System.Configuration.ConfigurationSettings.AppSettings["ConnectionStringSMSII"];
    //public static SqlConnection con_stringMIS = new SqlConnection("server=10.10.25.114;database=MIS;uid=sa;password=bhushan123#;");
    //public static OracleConnection con_stringSMSII = new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.15.20.85)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=SMCDB)));User Id=bhushan;Password=bhushan;");
    //public static string DbconSql = "server=10.10.25.114\\SQLEXPRESS;database=MIS;uid=sa;password=bhushan123#;";
    public static string DbconSql = "server=10.10.58.171\\SQLEXPRESS;database=MIS;uid=sa;password=bhushan123#;";
    public static string DbconSqlLIMS = "server=10.15.20.91\\SQLEXPRESS;database=Bhushan;uid=sa;password=Metsys123#";
    public static string con_stringSMSII = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.15.20.85)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=SMCDB)));User Id=bhushan;Password=bhushan;";

    public static string con_stringCASTERI = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.15.10.22)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=CCSDB1)));User Id=L2CCS;Password=L2CCS;";
    public static string con_stringCASTERII = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.16.10.22)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=CCSDB)));User Id=L2CCS;Password=L2CCS;";
    public static string con_stringCASTERIII = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.17.10.22)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=CCSDB)));User Id=L2CCS;Password=L2CCS;";

    //public static string con_stringSTATUS = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST = SMC-DEV)(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = SMCDB)));User Id=BHUSHAN;Password=BHUSHAN;";
    public static string con_stringSTATUS = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST = 10.15.20.85)(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = SMCDB)));User Id=BHUSHAN;Password=BHUSHAN;";

    public string StrQueryMIS;
    public string StrQuerySMSII;
    public string StrQueryCASTERII;
    public string StrDateFrom;
    bool statusMIS;

    //public static string StrConarcARunning;
    //public static string StrConarcBRunning;

    public string connectionMIS(string con_stringMIS)
    {

        //Get connection string from Web.Config
        con_stringMIS = System.Configuration.ConfigurationSettings.AppSettings["ConnectionStringMIS"];

        return con_stringMIS;
    }

    public string connectionSMSII(string con_stringSMSII)
    {

        //Get connection string from Web.Config
        con_stringSMSII = System.Configuration.ConfigurationSettings.AppSettings["ConnectionStringSMSII"];

        return con_stringSMSII;
    }
    public DataTable DBSelectQueryLMIS_Table(string QueryMIS)
    {

        DataTable TempDataTableMIS = new DataTable();
        TempDataTableMIS.Clear();
        try
        {
            SqlConnection DbconMIS = new SqlConnection(DbconSqlLIMS);
            SqlDataAdapter TempAdapterMIS = new SqlDataAdapter(QueryMIS, DbconMIS);
            DbconMIS.Open();
            TempAdapterMIS.Fill(TempDataTableMIS);
            DbconMIS.Close();
        }
        catch (Exception ex)
        {
            //txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
        }
        return TempDataTableMIS;

    }
    public DataSet DBSelectQueryMIS(string QueryMIS)
    {
        DataSet TempDatasetMIS = new DataSet();
        TempDatasetMIS.Tables.Clear();

        SqlConnection DbconMIS = new SqlConnection();
        SqlDataAdapter TempAdapterMIS = new SqlDataAdapter(QueryMIS, DbconMIS);
        DbconMIS.Open();
        TempAdapterMIS.Fill(TempDatasetMIS);
        DbconMIS.Close();

        return TempDatasetMIS;
    }
    public DataTable DBSelectQueryMIS_Table(string QueryMIS)
    {

        DataTable TempDataTableMIS = new DataTable();
        TempDataTableMIS.Clear();
        try
        {
            SqlConnection DbconMIS = new SqlConnection(DbconSql);
            SqlDataAdapter TempAdapterMIS = new SqlDataAdapter(QueryMIS, DbconMIS);
            DbconMIS.Open();
            TempAdapterMIS.Fill(TempDataTableMIS);
            DbconMIS.Close();
        }
        catch (Exception ex)
        {
            //txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
        }
        return TempDataTableMIS;

    }
    public DataSet DBSelectQuerySMSII(string QuerySMSII)
    {

        DataSet TempDatasetSMSII = new DataSet();
        TempDatasetSMSII.Tables.Clear();
        try
        {
            OracleConnection DbconSMSII = new OracleConnection(con_stringSMSII);
            OracleDataAdapter TempAdapterSMSII = new OracleDataAdapter(QuerySMSII, DbconSMSII);
            DbconSMSII.Open();
            TempAdapterSMSII.Fill(TempDatasetSMSII);
            DbconSMSII.Close();
        }
        catch (Exception ex)
        {
            // txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
        }
        return TempDatasetSMSII;

    }
    public DataTable DBSelectQuerySMSII_Table(string QuerySMSII)
    {

        DataTable TempDataTableSMSII = new DataTable();
        TempDataTableSMSII.Clear();
        try
        {
            OracleConnection DbconSMSII = new OracleConnection(con_stringSMSII);
            OracleDataAdapter TempAdapterSMSII = new OracleDataAdapter(QuerySMSII, DbconSMSII);
            DbconSMSII.Open();
            TempAdapterSMSII.Fill(TempDataTableSMSII);
            DbconSMSII.Close();
        }
        catch (Exception ex)
        {
            //txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
        }
        return TempDataTableSMSII;
    }

    public DataTable DBSelectQueryCASTERII_Table(string QueryCASTERII)
    {

        DataTable TempDataTableCASTERII = new DataTable();
        OracleConnection DbconCASTERII = new OracleConnection(con_stringCASTERII);
        OracleDataAdapter TempAdapterSMSII = new OracleDataAdapter(QueryCASTERII, DbconCASTERII);
        TempDataTableCASTERII.Clear();
        try
        {
            DbconCASTERII.Open();
            TempAdapterSMSII.Fill(TempDataTableCASTERII);
            DbconCASTERII.Close();
        }
        catch (Exception ex)
        {
            //  txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
        }
        finally
        {
            if (DbconCASTERII.State == ConnectionState.Open)
            {
                DbconCASTERII.Close();
            }
        }
        return TempDataTableCASTERII;
    }
    public DataTable DBSelectQueryCASTERIII_Table(string QueryCASTERIII)
    {

        DataTable TempDataTableCASTERIII = new DataTable();
        OracleConnection DbconCASTERIII = new OracleConnection(con_stringCASTERIII);
        OracleDataAdapter TempAdapterSMSIII = new OracleDataAdapter(QueryCASTERIII, DbconCASTERIII);
        TempDataTableCASTERIII.Clear();
        try
        {
            DbconCASTERIII.Open();
            TempAdapterSMSIII.Fill(TempDataTableCASTERIII);
            DbconCASTERIII.Close();
        }
        catch (Exception ex)
        {
            //  txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
        }
        finally
        {
            if (DbconCASTERIII.State == ConnectionState.Open)
            {
                DbconCASTERIII.Close();
            }
        }
        return TempDataTableCASTERIII;
    }
    public DataTable DBSelectQueryCASTERI_Table(string QueryCASTERI)
    {
        DataTable TempDataTableCASTERI = new DataTable();
        TempDataTableCASTERI.Clear();
        try
        {
            OracleConnection DbconCASTERI = new OracleConnection(con_stringCASTERI);
            OracleDataAdapter TempAdapterSMSI = new OracleDataAdapter(QueryCASTERI, DbconCASTERI);
            DbconCASTERI.Open();
            TempAdapterSMSI.Fill(TempDataTableCASTERI);
            DbconCASTERI.Close();
        }
        catch
        {
        }
        return TempDataTableCASTERI;
    }

    public bool DBInsertUpdateDeleteMIS(string QueryMIS)
    {
        try
        {

            SqlConnection DbconMIS = new SqlConnection(DbconSql);
            DbconMIS.Open();
            SqlCommand cmdMIS = new SqlCommand(QueryMIS, DbconMIS);
            cmdMIS.CommandType = CommandType.Text;
            cmdMIS.ExecuteNonQuery();
            DbconMIS.Close();
            statusMIS = true;
        }
        catch (Exception ex)
        {
            //   txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
        }
        return statusMIS;
    }

    #region "STATUS SCREEN"

    public DataTable DBSelectQueryStatusScreen_Table(string QueryBhushan_StatusScreen)
    {
        DataTable TempTbl_StatusScreen = new DataTable();
        try
        {
            TempTbl_StatusScreen.Clear();

            OracleConnection DbconStatusScreen = new OracleConnection(con_stringSTATUS);
            OracleDataAdapter TempAdpStatusScreen = new OracleDataAdapter(QueryBhushan_StatusScreen, DbconStatusScreen);
            DbconStatusScreen.Open();
            TempAdpStatusScreen.Fill(TempTbl_StatusScreen);
            DbconStatusScreen.Close();

        }
        catch
        {
        }
        return TempTbl_StatusScreen;
    }


    #endregion



}
