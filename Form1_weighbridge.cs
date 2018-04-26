using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Globalization;

namespace TrpDataFetchingMIS
{
    public partial class Form1 : Form
    {
        public string con_string = System.Configuration.ConfigurationSettings.AppSettings["SqlDBMIS_Con"];
        public string Trp_con_string = System.Configuration.ConfigurationSettings.AppSettings["AccessTrp_Con"];
        public string Act_con_string;
        // string con_string = "Data Source=10.10.58.171\\SQLEXPRESS;Initial Catalog=MIS;User ID=sa;Password=bhushan123#";
        string SerialNoLog = "";
        OleDbConnection con = new OleDbConnection();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataTable dt3 = new DataTable();
        string Query = "";
        public Form1()
        {
            InitializeComponent();
            // +Manish Insert log at start of the form 
            string Message = "Form 1 componenet initialized";
            LogManager.CreateLogFile(Message.ToString());
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                FnCheckConnectionMIS();
                FnCheckConnectionTRP();
                // +Manish Insert log at start of the form 
                string Message = "Form 1_Load object";
                LogManager.CreateLogFile(Message.ToString());
            }
            catch
            {
            }
        }
        public void GetDataTrpBrigde()
        {
            try
            {
                ds = new DataSet();
                ds.Clear();
                con = new OleDbConnection();
                con.ConnectionString = Trp_con_string;
                
                con.Open();
                // +Manish test Code for checking 
                ConnectionState mdb_state = con.State;
                if (mdb_state == ConnectionState.Closed || mdb_state == ConnectionState.Broken)
                {
                    LogManager.CreateLogFile(" could not Connected to Wagon.mdb on this Location:- " + Trp_con_string.ToString());
                }

                string str = "SELECT Weighment.[SerialNo], Weighment.[RakeNo], Weighment.[Material], Weighment.[Source], Weighment.[Destiantion]," +
                                      "Weighment.[TareWt], Weighment.[TDate], Weighment.[TTime], Weighment.[GrossWt], Weighment.[GDate], Weighment.[GTime]," +
                                      "Weighment.[NetWt], Weighment.[Shift] FROM Weighment where  ( Weighment.[RakeNo]) like '%" + DateTime.Now.ToString("MMddyyyy") + "%' and Weighment.[GrossWt]<>0";


                OleDbDataAdapter adp = new OleDbDataAdapter(str, con);
                adp.Fill(ds);
                dt = ds.Tables[0];
                if (dt.Rows.Count == 0)
                {
                    Act_con_string = System.Configuration.ConfigurationSettings.AppSettings["AccessTrp_Alt"];
                    LogManager.CreateLogFile("using other location to search the database<C:\\RIMSoft18-Torpedo\\RailInMotion\\wagon.mdb");
                }
                else
                {
                    Act_con_string = System.Configuration.ConfigurationSettings.AppSettings["AccessTrp_Con"];
                }

                try
                {
                    ds.Clear();
                    con = new OleDbConnection();
                    con.ConnectionString = Act_con_string;
                    con.Open();
                    // +Manish test Code for checking 
                    ConnectionState mdb1_state = con.State;
                    if (mdb1_state == ConnectionState.Closed || mdb1_state == ConnectionState.Broken)
                    {
                        LogManager.CreateLogFile(" could not Connected to Wagon.mdb");
                    }


                    // string str = "SELECT Weighment.RakeNo FROM Weighment  WHERE (((Weighment.RakeNo) Like '%05142014%'))";
                    //string str="SELECT Weighment.[SerialNo], Weighment.[RakeNo], Weighment.[Material], Weighment.[Source], Weighment.[Destiantion],"+
                    //        "Weighment.[TareWt], Weighment.[TDate], Weighment.[TTime], Weighment.[GrossWt], Weighment.[GDate], Weighment.[GTime],"+
                    //        "Weighment.[NetWt], Weighment.[Shift] FROM Weighment where  ( Weighment.[RakeNo]) like '%"+DateTime.Now.ToString("MMddyyyy")+"%'"+
                    //        "and Weighment.[TTime]= (select max(Weighment.[TTime]) from Weighment where (Weighment.[RakeNo]) like '%"+DateTime.Now.ToString("MMddyyyy")+"%')";
                    string str1 = "SELECT Weighment.[SerialNo], Weighment.[RakeNo], Weighment.[Material], Weighment.[Source], Weighment.[Destiantion]," +
                                          "Weighment.[TareWt], Weighment.[TDate], Weighment.[TTime], Weighment.[GrossWt], Weighment.[GDate], Weighment.[GTime]," +
                                          "Weighment.[NetWt], Weighment.[Shift] FROM Weighment where  ( Weighment.[RakeNo]) like '%" + DateTime.Now.ToString("MMddyyyy") + "%' and Weighment.[GrossWt]<>0";


                    OleDbDataAdapter adp1 = new OleDbDataAdapter(str, con);
                    adp1.Fill(ds);
                    dt = ds.Tables[0];
                    double hotMetalwt = 0;

                    if (dt.Rows.Count > 0)
                    {
                        // +Manish Insert Row Count satement to log file
                        string CountRows = "GetDataTrpBrigde() ->Select Statment Returns total: ";
                        CountRows = CountRows.ToString() + dt.Rows.Count.ToString();
                        LogManager.CreateLogFile(CountRows);
                        foreach (DataRow dr in dt.Rows)
                        {
                            // dr["GTime"].ToString()
                            if (dr["GTime"].ToString() != "")
                            {
                                string InsertedDate = DateTime.Now.ToString("yyyy-MM-dd") + " " + (dr["GTime"].ToString()).Substring(10, (dr["GTime"].ToString().Length - 10));
                                Query = "SELECT *FROM Trp_WeightBridge where Serial_No='" + dr["SerialNo"].ToString() + "'";
                                dt2 = DBSelectQueryLevel(Query);
                                if (dt2.Rows.Count == 0)
                                {
                                    if ((dr["GrossWt"].ToString() == "" || dr["GrossWt"].ToString() == "0") || (dr["TareWt"].ToString() == "" || dr["TareWt"].ToString() == "0"))
                                    {
                                        hotMetalwt = 0;
                                    }
                                    else
                                    {
                                        hotMetalwt = (Convert.ToDouble(dr["GrossWt"].ToString()) - Convert.ToDouble(dr["TareWt"].ToString()));
                                    }
                                    string StrUpdateMisMasterSap = "INSERT INTO Trp_WeightBridge(Serial_No,Rake_No,Material,Source,Destiantion,Tare_wt,TDate,TTime,Gross_Wt,GDate,GTime,HotMetal_Wt,RevDateTime) VALUES('" + dr["SerialNo"].ToString() + "','" + dr["RakeNo"].ToString() + "','" + dr["Material"].ToString() + "','" + dr["Source"].ToString() + "','" + dr["Destiantion"].ToString() + "'," + dr["TareWt"].ToString() + ",'" + Convert.ToDateTime(dr["TDate"].ToString(), CultureInfo.CurrentCulture).ToString("yyyy-MM-dd") + "','" + (dr["TTime"]).ToString() + "'," + dr["GrossWt"].ToString() + ",'" + Convert.ToDateTime(dr["GDate"].ToString(), CultureInfo.CurrentCulture).ToString("yyyy-MM-dd") + "','" + dr["GTime"].ToString() + "'," + hotMetalwt + ",'" + InsertedDate + "')";
                                    if (DB_MIS_MasterSapInsert(StrUpdateMisMasterSap) == true)
                                    {
                                        SerialNoLog = dr["SerialNo"].ToString();
                                        string InsertLogData = "Inserted SerialNo-" + dr["SerialNo"].ToString() + "|Source-" + dr["Source"].ToString() + "|Destiantion-" + dr["Destiantion"].ToString() + "|TareWt-" + dr["TareWt"].ToString() + "|TTime-" + Convert.ToDateTime(dr["TTime"].ToString()).ToShortTimeString() + "TDate-" + Convert.ToDateTime(dr["TDate"].ToString(), CultureInfo.CurrentCulture).ToString("yyyy-MM-dd");
                                        LogManager.CreateLogFile(InsertLogData);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    string ErrorLogData = "SerialNo-" + SerialNoLog + " " + ex.Message;
                    LogManager.CreateLogFile(ErrorLogData);
                }
                finally
                {
                    con.Close();
                    con.Dispose();
                }

            }
            catch (Exception e)
            {

                LogManager.CreateLogFile("Not Able to search Path in  App Config File ->" + e.Message);
            }
 }

        public void UpdDataTrpBrigde()
        {
            string str = "";
            DataSet dsTrp = new DataSet();
            DataTable dtTrp = new DataTable();
            try
            {
                   str = "SELECT Serial_No FROM Trp_WeightBridge where Tare_wt=0";
                    dt = DBSelectQueryLevel(str);
               
                
                if (dt.Rows.Count > 0)
                {
                  
                    foreach (DataRow dr in dt.Rows)
                    {
                        try
                        {
                            con = new OleDbConnection();
                            //con.ConnectionString = Trp_con_string;
                            con.ConnectionString = Act_con_string;
                            con.Open();
                            // +Manish test Code for checking 
                            ConnectionState mdb_state = con.State;
                            if (mdb_state == ConnectionState.Closed || mdb_state == ConnectionState.Broken)
                            {
                                LogManager.CreateLogFile(" could not Connected to Wagon.mdb" + Act_con_string.ToString());
                            }

                            str = "SELECT Weighment.[SerialNo], Weighment.[RakeNo], Weighment.[Material], Weighment.[Source], Weighment.[Destiantion]," +
                            "Weighment.[TareWt], Weighment.[TDate], Weighment.[TTime], Weighment.[GrossWt], Weighment.[GDate], Weighment.[GTime]," +
                            "Weighment.[NetWt], Weighment.[Shift] FROM Weighment where Weighment.[SerialNo]='" + dr["Serial_No"].ToString() + "' and Weighment.[TareWt]<>0";
                            OleDbDataAdapter adp = new OleDbDataAdapter(str, con);
                            adp.Fill(dsTrp);
                            dtTrp = dsTrp.Tables[0];
                            double hotMetalwtUpd = 0;
                            if (dtTrp.Rows.Count > 0)
                            {
                                // +Manish Insert Row Count satement to log file
                                string CountRows = "UpdDataTrpBrigde() ->Select Statment Returns total: ";
                                CountRows = CountRows.ToString() + dtTrp.Rows.Count.ToString();
                                LogManager.CreateLogFile(CountRows);
                                for (int i = 0; i < dtTrp.Rows.Count; i++)
                                {
                                    try
                                    {
                                        if ((dtTrp.Rows[i]["GrossWt"].ToString() == "" || dtTrp.Rows[i]["GrossWt"].ToString() == "0") || (dtTrp.Rows[i]["TareWt"].ToString() == "" || dtTrp.Rows[i]["TareWt"].ToString() == "0"))
                                        {
                                            hotMetalwtUpd = 0;
                                        }
                                        else
                                        {

                                            hotMetalwtUpd = (Convert.ToDouble(dtTrp.Rows[i]["GrossWt"].ToString()) - Convert.ToDouble(dtTrp.Rows[i]["TareWt"].ToString()));
                                        }
                                    }

                                    catch (Exception ex)
                                    {
                                        hotMetalwtUpd = 0;

                                    }
                                    try
                                    {
                                        str = "UPDATE Trp_WeightBridge SET [Material] = '" + dtTrp.Rows[i]["Material"].ToString() + "',[Source] ='" + dtTrp.Rows[i]["Source"].ToString() +
                                            "',[Destiantion] ='" + dtTrp.Rows[i]["Destiantion"].ToString() + "' ,[Tare_wt] ='" + dtTrp.Rows[i]["TareWt"].ToString() +
                                            "',[TDate] ='" + Convert.ToDateTime(dtTrp.Rows[i]["TDate"].ToString(), CultureInfo.CurrentCulture).ToString("yyyy-MM-dd") + "' ,[TTime] ='" + dtTrp.Rows[i]["TTime"].ToString() + "',[Gross_Wt] ='" +
                                             dtTrp.Rows[i]["GrossWt"].ToString() + "',[GDate] ='" + Convert.ToDateTime(dtTrp.Rows[i]["GDate"].ToString(), CultureInfo.CurrentCulture).ToString("yyyy-MM-dd") + "' ,[GTime] ='" + dtTrp.Rows[i]["GTime"].ToString() +
                                            "',[Net_Wt] ='" + dtTrp.Rows[i]["NetWt"].ToString() + "'  ,[Shift] ='" + dtTrp.Rows[i]["Shift"].ToString() + "',[HotMetal_Wt] ='" + hotMetalwtUpd +
                                            "' WHERE Serial_No='" + dtTrp.Rows[i]["SerialNo"].ToString() + "'";

                                       // SerialNoLog = dtTrp.Rows[i]["SerialNo"].ToString();
                                        if (DB_MIS_MasterSapInsert(str) == true)
                                        {
                                            string UpdLogData = "Updated SerialNo-" + dtTrp.Rows[i]["SerialNo"].ToString() +"|Source-" + dtTrp.Rows[i]["Source"].ToString() +"|Destiantion-" + dtTrp.Rows[i]["Destiantion"].ToString() + "|TareWt-" + dtTrp.Rows[i]["TareWt"].ToString() + "|TTime-" + Convert.ToDateTime(dtTrp.Rows[i]["TTime"].ToString()).ToShortTimeString() + "|TDate-" + Convert.ToDateTime(dtTrp.Rows[i]["TDate"].ToString(), CultureInfo.CurrentCulture).ToString("yyyy-MM-dd");
                                            LogManager.CreateLogFile(UpdLogData);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogManager.CreateLogFile(ex.Message);
                                    }
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            string ErrorLogData = "SerialNo-" + SerialNoLog + " " + ex.Message;
                            LogManager.CreateLogFile(ErrorLogData);
                        }
                        finally
                        {
                            con.Close();
                            con.Dispose();
                        }

                    }

                }
            }
            catch (Exception ex)
            {
                LogManager.CreateLogFile(ex.Message);
            }
            finally
            {
            }
        }        
        public bool DB_MIS_MasterSapInsert(string Query) 
        {
            bool sts = false;
            try
            {
                SqlConnection Dbcon = new SqlConnection(con_string);
                Dbcon.Open();
                SqlCommand cmd = new SqlCommand(Query, Dbcon);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
                Dbcon.Close();
                sts = true;
            }
            catch
            {
                sts = false;
            }

            return sts;
        }
        public DataTable DBSelectQueryLevel(string Query)
        {
            DataTable TempDataTable = new DataTable();
            TempDataTable.Clear();
            try
            {
                SqlConnection Dbcon = new SqlConnection(con_string);
                SqlDataAdapter TempAdapter = new SqlDataAdapter(Query, Dbcon);
                Dbcon.Open();
                TempAdapter.Fill(TempDataTable);
                Dbcon.Close();

            }
            catch
            {
                // Pdi_Msg.alert("CoilId Updated Successfully")
            }

            return TempDataTable;
        }

        private void timer_TrpDataFetch_Tick(object sender, EventArgs e)
        {
  
            GetDataTrpBrigde();
            UpdDataTrpBrigde();
        }
        public void FnCheckConnectionMIS()
        {
            try
            {

                SqlConnection DbconSql = new SqlConnection(con_string);
                DbconSql.Open();
                lblArrowMIS.ForeColor = System.Drawing.Color.Green;
                lbl_MIS_Status_txt.Text = "Connected";
                lbl_MIS_Status_txt.ForeColor = System.Drawing.Color.Green;
                lblMIS.BackColor = System.Drawing.Color.Green;
                lblMIS.ForeColor = System.Drawing.Color.White;
                DbconSql.Close();


            }
            catch
            {
                lblArrowMIS.ForeColor = System.Drawing.Color.Red;
                lbl_MIS_Status_txt.Text = "Not-Connected";
                lbl_MIS_Status_txt.ForeColor = System.Drawing.Color.Red;
                lblMIS.BackColor = System.Drawing.Color.Red;
                lblMIS.ForeColor = System.Drawing.Color.White;
            }

        }
        public void FnCheckConnectionTRP()
        {
            try
            {
                OleDbConnection AccessConOrc = new OleDbConnection(Trp_con_string);
                AccessConOrc.Open();
                lblArrowTRP.ForeColor = System.Drawing.Color.Green;
                lbl_TRP_Status_txt.Text = "Connected";
                lbl_TRP_Status_txt.ForeColor = System.Drawing.Color.Green;
                lblTRP.BackColor = System.Drawing.Color.Green;
                lblTRP.ForeColor = System.Drawing.Color.White;
                AccessConOrc.Close();
            }
            catch
            {
                lblArrowTRP.ForeColor = System.Drawing.Color.Red;
                lbl_TRP_Status_txt.Text = "Not-Connected";
                lbl_TRP_Status_txt.ForeColor = System.Drawing.Color.Red;
                lblTRP.BackColor = System.Drawing.Color.Red;
                lblTRP.ForeColor = System.Drawing.Color.White;
            }

        }
        private void timer_DbConnStatus_Tick(object sender, EventArgs e)
        {
            try
            {
                FnCheckConnectionMIS();
                FnCheckConnectionTRP();
                LogManager.DeleteLogFile();
            }
            catch
            {
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
            Hide();
        }
    }
}
