using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Oracle.DataAccess.Client;
using System.Net.Sockets;
using System.IO;

namespace BOFInterface
{
    public partial class Form1 : Form
    {
        public string Sql_con_string = System.Configuration.ConfigurationSettings.AppSettings["SqlDB_Con"];
        public string BOF_con_string = System.Configuration.ConfigurationSettings.AppSettings["BOFL2_Con"];
        public string XmlFilePath = System.Configuration.ConfigurationSettings.AppSettings["Xmlpath"];
        DBConnections clsObj = new DBConnections();
        DataTable dt = new DataTable();
        string Error_Line = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                FnCheckConnectionMIS();
                FnCheckConnectionBOF1();
                FnCheckConnectionMIS2();
                FnCheckConnectionCLAB();
            }
            catch
            {
            }
        }
        public void FnCheckConnectionCLAB()
        {
            try
            {
                if (System.IO.File.Exists(XmlFilePath))
                {
                    lblArrowLAB.ForeColor = System.Drawing.Color.Green;
                    lbl_LAB_Status_txt.Text = "Connected";
                    lbl_LAB_Status_txt.ForeColor = System.Drawing.Color.Green;
                    lblLAB.BackColor = System.Drawing.Color.Green;
                    lblLAB.ForeColor = System.Drawing.Color.White;
                }
                else
                {
                    lblArrowLAB.ForeColor = System.Drawing.Color.Red;
                    lbl_LAB_Status_txt.Text = "Not-Connected";
                    lbl_LAB_Status_txt.ForeColor = System.Drawing.Color.Red;
                    lblLAB.BackColor = System.Drawing.Color.Red;
                    lblLAB.ForeColor = System.Drawing.Color.White;
                }

            }
            catch
            {
                lblArrowLAB.ForeColor = System.Drawing.Color.Red;
                lbl_LAB_Status_txt.Text = "Not-Connected";
                lbl_LAB_Status_txt.ForeColor = System.Drawing.Color.Red;
                lblLAB.BackColor = System.Drawing.Color.Red;
                lblLAB.ForeColor = System.Drawing.Color.White;
            }

        }
        public void FnCheckConnectionMIS2()
        {
            try
            {

                SqlConnection DbconSql = new SqlConnection(Sql_con_string);
                DbconSql.Open();
                lblArrowMIS2.ForeColor = System.Drawing.Color.Green;
                lbl_MIS_Status_txt2.Text = "Connected";
                lbl_MIS_Status_txt2.ForeColor = System.Drawing.Color.Green;
                lblMIS2.BackColor = System.Drawing.Color.Green;
                lblMIS2.ForeColor = System.Drawing.Color.White;
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
        public void GetDataContainerLab()
        {
            //Author:-Mohan Singh Created on 18-07-2014
            //Purpose:-Read Data From Xml File(Results.xml) from Path 'D:\Mohan_Backup\XMLLABTOMISCOPY' of Bof Interface System 
            //Output:-Data is Stored In MIS table 'SMS_ContainerLabChem' and Created Log in 'D:\Mohan_Backup\XMLLABTOMISLOG'
            DataTable dt2 = new DataTable();
            string resultxml_path = XmlFilePath;
            DataSet ResultContainer = new DataSet();

            string elem_C = "", elem_Si = "", elem_Mn = "", elem_P = "", elem_S = "", elem_Cr = "", elem_Mo = "", elem_Ni = "", elem_Al = "";
            string elem_Co = "", elem_Cu = "", elem_Nb = "", elem_Ti = "", elem_W = "", elem_Pb = "", elem_Sn = "", elem_As = "", elem_Bi = "";
            string elem_Ca = "", elem_Sb = "", elem_B = "", elem_Zn = "", elem_N = "", elem_Fe = "", elem_Al_s = "", elem_V = "";
            string HeatId = "", TreatId = "", sample_No = "", Taken_Time = "";
            string Query = "", Measument_ID = "", PlantNo = "", Plant = "";
            string resultID_C = "", resultID_Si = "", resultID_Mn = "", resultID_P = "", resultID_S = "", resultID_Cr = "", resultID_Mo = "", resultID_Ni = "", resultID_Al = "";
            string resultID_Co = "", resultID_Cu = "", resultID_Nb = "", resultID_Ti = "", resultID_W = "", resultID_Pb = "", resultID_Sn = "", resultID_As = "", resultID_Bi = "";
            string resultID_Ca = "", resultID_Sb = "", resultID_B = "", resultID_Zn = "", resultID_N = "", resultID_Fe = "", resultID_Al_s = "", resultID_V = "";

            try
            {
                Error_Line = "0";
                //ResultContainer.ReadXml(@"E:\Results.xml");
                ResultContainer.ReadXml(resultxml_path.ToString());
                DataTable dtMeasurment = new DataTable();
                if (ResultContainer.Tables["Measurement"] != null)
                {
                    Error_Line = "1";
                    try
                    {
                        dtMeasurment = ResultContainer.Tables["Measurement"];
                        foreach (DataRow dr in dtMeasurment.Rows)
                        {
                            if (dr["Type"].ToString().Trim() == "Statistics")
                            {
                                Measument_ID = dr["Measurement_Id"].ToString();
                                Taken_Time = dr["Time"].ToString();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.CreateLogFile(DateTime.Now.ToString() + "Error in Line " + Error_Line + " " + ex.Message);
                    }
                }

                if (ResultContainer.Tables["SampleID"] != null)
                {
                    Error_Line = "2";
                    DataTable SampleIDs = new DataTable();
                    SampleIDs = ResultContainer.Tables["SampleID"];
                    try
                    {

                        if (SampleIDs.Rows.Count > 0)
                        {
                            sample_No = SampleIDs.Rows[0]["Value"].ToString();
                            HeatId = SampleIDs.Rows[1]["Value"].ToString();
                            TreatId = SampleIDs.Rows[2]["Value"].ToString();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.CreateLogFile(DateTime.Now.ToString() + "Error in Line " + Error_Line + " " + ex.Message);
                    }

                }
                if (ResultContainer.Tables["ResultValues"] != null)
                {
                    Error_Line = "3";
                    DataTable ResultValues = new DataTable();
                    ResultValues = ResultContainer.Tables["ResultValues"];
                    //Collect Results_Id Dynamically according to Measurement_Id of Statistics
                    #region Dynamically Results_Id according to Measurement_Id of Statistics
                    try
                    {
                        foreach (DataRow dr1 in ResultValues.Rows)
                        {
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "C"))
                            {
                                resultID_C = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Si"))
                            {
                                resultID_Si = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Mn"))
                            {
                                resultID_Mn = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "P"))
                            {
                                resultID_P = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "S"))
                            {
                                resultID_S = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Cr"))
                            {
                                resultID_Cr = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Mo"))
                            {
                                resultID_Mo = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Ni"))
                            {
                                resultID_Ni = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Al"))
                            {
                                resultID_Al = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Co"))
                            {
                                resultID_Co = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Cu"))
                            {
                                resultID_Cu = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Nb"))
                            {
                                resultID_Nb = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Ti"))
                            {
                                resultID_Ti = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "W"))
                            {
                                resultID_W = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Pb"))
                            {
                                resultID_Pb = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Sn"))
                            {
                                resultID_Sn = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "As"))
                            {
                                resultID_As = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Bi"))
                            {
                                resultID_Bi = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Ca"))
                            {
                                resultID_Ca = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Sb"))
                            {
                                resultID_Sb = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "B"))
                            {
                                resultID_B = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Zn"))
                            {
                                resultID_Zn = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "N"))
                            {
                                resultID_N = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Fe"))
                            {
                                resultID_Fe = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "Al_Sol"))
                            {
                                resultID_Al_s = dr1["ResultValues_Id"].ToString();
                            }
                            if ((dr1["Measurement_Id"].ToString().Trim() == Measument_ID) && (dr1["Type"].ToString().Trim() == "Element" && dr1["IDName"].ToString().Trim() == "V"))
                            {
                                resultID_V = dr1["ResultValues_Id"].ToString();
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.CreateLogFile(DateTime.Now.ToString() + "Error in Line " + Error_Line + " " + ex.Message);
                    }
                    #endregion
                }
                if (ResultContainer.Tables["ResultValue"] != null)
                {
                    Error_Line = "4";
                    DataTable ResultValue = new DataTable();
                    ResultValue = ResultContainer.Tables["ResultValue"];
                    if (ResultValue.Rows.Count > 0)
                    {
                        try
                        {
                            #region Dynmically Fetching The Elememts values
                            for (int i = 0; i < ResultValue.Rows.Count; i++)
                            {
                                // for element C 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_C)
                                {

                                    elem_C = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Si
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Si)
                                {

                                    elem_Si = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Mn 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Mn)
                                {
                                    elem_Mn = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element P 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_P)
                                {
                                    elem_P = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element S 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_S)
                                {
                                    elem_S = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Cr 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Cr)
                                {
                                    elem_Cr = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Mo
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Mo)
                                {
                                    elem_Mo = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Ni 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Ni)
                                {
                                    elem_Ni = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Al 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Al)
                                {
                                    elem_Al = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Co 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Co)
                                {
                                    elem_Co = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Cu 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Cu)
                                {
                                    elem_Cu = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Nb 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Nb)
                                {
                                    elem_Nb = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Ti 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Ti)
                                {
                                    elem_Ti = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element V 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_V)
                                {
                                    elem_V = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element W 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_W)
                                {

                                    elem_W = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Pb 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Pb)
                                {
                                    elem_Pb = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Sn 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Sn)
                                {
                                    elem_Sn = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element As 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_As)
                                {
                                    elem_As = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Bi 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Bi)
                                {
                                    elem_Bi = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Ca 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Ca)
                                {
                                    elem_Ca = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Sb 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Sb)
                                {
                                    elem_Sb = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element B result_id =229
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_B)
                                {
                                    elem_B = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Zn 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Zn)
                                {
                                    elem_Zn = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element N 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_N)
                                {

                                    elem_N = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Fe 
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Fe)
                                {

                                    elem_Fe = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }

                                // for element Al_sol
                                if (ResultValue.Rows[i]["Type"].ToString() == "Mean-Conc" && ResultValue.Rows[i]["ResultValues_Id"].ToString() == resultID_Al_s)
                                {
                                    elem_Al_s = ResultValue.Rows[i]["ResultValue_Text"].ToString();
                                }
                            #endregion
                            }
                        }
                        catch (Exception ex)
                        {
                            // txtError.Text = ex.Message + DateTime.Now.ToString();
                            LogManager.CreateLogFile(DateTime.Now.ToString() + "Error in Line " + Error_Line + " " + ex.Message);
                        }
                        finally
                        {
                        }
                    }

                }


                Query = "SELECT *FROM SMS_ContainerLabChem where HeatID='" + HeatId + "' and SampleNo='" + sample_No + "' and TreatID='" + TreatId + "'";
                //dt = DBSelectQueryLevel(Query);
                dt = clsObj.DBSelectQueryMIS_Table(Query);
                if (dt.Rows.Count == 0)
                {
                    string PlantStauts = GetPlantStatus(HeatId, TreatId);
                    string[] strPlant = PlantStauts.Split(' ');
                    Plant = strPlant[0].ToString();
                    PlantNo = strPlant[1].ToString();

                    string strQry = "INSERT INTO SMS_ContainerLabChem([HeatID],[SampleNo],[TreatID],[C],[Si],[Mn],[P],[S],[Cr],[Mo],[Ni],[Al],[Co],[Cu],[Nb]," +
                    "[Ti],[V],[W],[Pb],[Sn],[As],[Bi],[Ca],[Sb],[B],[Zn],[N],[Fe],[Al_Sol],[TimeOfAnalysis] ,[Flag],[Plant],[PlantNo],[Grade]) VALUES('" + HeatId + "','" + sample_No + "','" +
                    TreatId + "','" + elem_C + "','" + elem_Si + "','" + elem_Mn + "','" + elem_P + "','" + elem_S + "','" + elem_Cr + "','" + elem_Mo + "','" +
                   elem_Ni + "','" + elem_Al + "','" + elem_Co + "','" + elem_Cu + "','" + elem_Nb + "','" + elem_Ti + "','" + elem_V + "','" + elem_W + "','" +
                   elem_Pb + "','" + elem_Sn + "','" + elem_As + "','" + elem_Bi + "','" + elem_Ca + "','" + elem_Sb + "','" + elem_B + "','" + elem_Zn + "','" +
                   elem_N + "','" + elem_Fe + "','" + elem_Al_s + "','" + Taken_Time + "','0','" + Plant + "','" + PlantNo + "','0')";
                    //clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery)
                    if (clsObj.DBInsertUpdateDeleteMIS(strQry) == true)
                    {
                        string InsertLogData = "HeatID:" + HeatId + "|SampleNo:" + sample_No + "|TreatID:" + TreatId + "|TimeofAna:" + Taken_Time + "|C:" + elem_C + "|Si:" + elem_Si +
                    "|Mn:" + elem_Mn + "|P:" + elem_P + "|S:" + elem_S + "|Cr:" + elem_Cr + "|Mo:" + elem_Mo + "|Ni:" + elem_Ni + "|Al:" + elem_Al +
                    "|Co:" + elem_Co + "|Cu:" + elem_Cu + "|Nb:" + elem_Nb + "|Ti:" + elem_Ti + "|V:" + elem_V + "|W:" + elem_W + "|Pb:" + elem_Pb + "|Sn:" + elem_Sn +
                    "|As:" + elem_As + "|Bi:" + elem_Bi + "|Ca:" + elem_Ca + "|Sb:" + elem_Sb + "|B:" + elem_B + "|Zn:" + elem_Zn + "|N:" + elem_N + "|Fe:" + elem_Fe +
                    "|Alsol:" + elem_Al_s + "|Plant:" + Plant + "|PlantNo:" + PlantNo + "|Grade:0";
                        LogManager.CreateLogFile(InsertLogData);
                        TcpClient tcpclnt = new TcpClient();
                        try
                        {
                            //Sending Telegram to LIMS server from this Client.
                            Error_Line = "5";
                            tcpclnt.Connect("10.15.20.91", 8492);
                            //tcpclnt.Connect("10.10.25.67", 8492);
                            Stream stm = tcpclnt.GetStream();
                            ASCIIEncoding asen = new ASCIIEncoding();
                            byte[] ba = asen.GetBytes(InsertLogData);
                            stm.Write(ba, 0, ba.Length);
                            Query = "Update SMS_ContainerLabChem SET Flag ='1' where HeatID='" + HeatId + "' and SampleNo='" + sample_No + "' and TreatID='" + TreatId + "'";
                            clsObj.DBInsertUpdateDeleteMIS(Query);


                        }
                        catch (Exception ex)
                        {
                            LogManager.CreateLogFile(DateTime.Now.ToString() + "Error in Line " + Error_Line + "" + ex.Message);
                        }
                        finally
                        {
                            tcpclnt.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.CreateLogFile(DateTime.Now.ToString() + "Error in Line " + Error_Line + " " + ex.Message);

            }

        }

        public string GetPlantStatus(string HeatID, string TreatID)
        {
            string PlantNo = "", Plant = "";
            if (HeatID.Substring(2, 1) == "C" && TreatID.Substring(0, 1) == "B")
            {
                Plant = "5";
                PlantNo = "1";
            }
            else if (HeatID.Substring(2, 1) == "D" && TreatID.Substring(0, 1) == "B")
            {
                Plant = "5";
                PlantNo = "2";
            }
            else if (HeatID.Substring(2, 1) == "C" && TreatID.Substring(0, 1) == "A")
            {
                Plant = "6";
                PlantNo = "1";
            }
            else if (HeatID.Substring(2, 1) == "D" && TreatID.Substring(0, 1) == "A")
            {
                Plant = "6";
                PlantNo = "2";
            }
            else if (TreatID.Substring(0, 1) == "L")
            {
                PlantNo = "3";
                Plant = "3";
            }

            return Plant + " " + PlantNo;
        }
        public void FnCheckConnectionMIS()
        {
            try
            {

                SqlConnection DbconSql = new SqlConnection(Sql_con_string);
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
        public void FnCheckConnectionBOF1()
        {
            try
            {
                OracleConnection OrcConOrc = new OracleConnection(BOF_con_string);
                OrcConOrc.Open();
                lblArrowBOF1.ForeColor = System.Drawing.Color.Green;
                lbl_BOF1_Status_txt.Text = "Connected";
                lbl_BOF1_Status_txt.ForeColor = System.Drawing.Color.Green;
                lblBOF1.BackColor = System.Drawing.Color.Green;
                lblBOF1.ForeColor = System.Drawing.Color.White;
                OrcConOrc.Close();
            }
            catch
            {
                lblArrowBOF1.ForeColor = System.Drawing.Color.Red;
                lbl_BOF1_Status_txt.Text = "Not-Connected";
                lbl_BOF1_Status_txt.ForeColor = System.Drawing.Color.Red;
                lblBOF1.BackColor = System.Drawing.Color.Red;
                lblBOF1.ForeColor = System.Drawing.Color.White;
            }

        }
        private void timercheckCon_Tick(object sender, EventArgs e)
        {
            try
            {
                FnCheckConnectionMIS();
                FnCheckConnectionBOF1();
                FnCheckConnectionMIS2();
                FnCheckConnectionCLAB();
            }
            catch
            {
            }
        }

        private void timerSynchBOF1_data_Tick(object sender, EventArgs e)
        {
            try
            {
                FnBOF1Data();
                FnBOF_ScrpCnsmp();
            }
            catch
            {
            }
        }

        private void FnBOF1Data()
        {
            string StrPLANT = "";
            string StrMIS_PLANTNO = "";
            string StrBOF_PLANTNO = "";
            string StrLF_PLANTNO = "";
            string StrDate = System.DateTime.Now.ToString("yyyy-MM-dd");
            string StrPrevDate = System.DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            StrPLANT = "BOF"; StrMIS_PLANTNO = "5"; StrBOF_PLANTNO = "1";
            FnSMS_SMCDB_BOF(StrPLANT, StrMIS_PLANTNO, StrBOF_PLANTNO);
            FnSMS_SMCDB_UpdateBOF(StrPLANT, StrMIS_PLANTNO);
            FnOP_BOF_STATUS(StrBOF_PLANTNO);
            FnSMS_SMCDB_HeatTracking(StrPLANT, StrBOF_PLANTNO, StrMIS_PLANTNO, StrDate);
            Fn_HeatBlow_Arc(StrPLANT, StrBOF_PLANTNO, StrMIS_PLANTNO, StrDate);
            StrPLANT = "BOF"; StrMIS_PLANTNO = "6"; StrBOF_PLANTNO = "2";
            FnSMS_SMCDB_BOF(StrPLANT, StrMIS_PLANTNO, StrBOF_PLANTNO);
            FnSMS_SMCDB_UpdateBOF(StrPLANT, StrMIS_PLANTNO);
            FnOP_BOF_STATUS(StrBOF_PLANTNO);
            FnSMS_SMCDB_HeatTracking(StrPLANT, StrBOF_PLANTNO, StrMIS_PLANTNO, StrDate);
            Fn_HeatBlow_Arc(StrPLANT, StrBOF_PLANTNO, StrMIS_PLANTNO, StrDate);
            StrPLANT = "LF"; StrLF_PLANTNO = "3"; StrMIS_PLANTNO = "11";
            FnSMS_SMCDB_HeatTracking_LF(StrPLANT, StrLF_PLANTNO, StrMIS_PLANTNO, StrDate);
            FnSMS_SMCDB_HeatTracking_LF_Update(StrPLANT, StrLF_PLANTNO, StrMIS_PLANTNO, StrPrevDate);
            Fn_HeatBlow_Arc_LF(StrPLANT, StrLF_PLANTNO, StrMIS_PLANTNO, StrDate);

            //string StrUpdateQuery = "update  smsHeatTracker set actEnd='24' where actStart > actEnd and DateStamp<'" + StrDate + "'";
            //bool StrUStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
            FnSMC_LF();

        }
        private void Fn_HeatBlow_Arc_LF(string StrPLANT, string StrBOF_PLANTNO, string StrMIS_PLANTNO, string StrDate)
        {
            StrMIS_PLANTNO = "";
            string StrQueryOrclCONARC = "Select Phpr.Plant || Phpr.Plantno as Plant,Phpr.Heatid_Cust As Heat_Id,Phpr.Treatid As Treatid,replace(Phdm.Starttime,',','.') As Arcing_Start,replace(Phdm.Endtime,',','.') As Arcing_End,decode(replace(SUBSTR(Phdm.Starttime,11,6) ,':','.'),0,null,replace(SUBSTR(Phdm.Starttime,11,6) ,':','.')) AS ArcStart,decode(replace(SUBSTR(Phdm.Endtime,11,6) ,':','.'),0,null,replace(SUBSTR(Phdm.Endtime,11,6) ,':','.')) AS ArcEnd,Phdm.Elec_Cons As Electrical_Cons,Phd.Carno as Car_id,Phpr.Heatid As Int_Heatid,SUBSTR(Phdm.Starttime,1,10) as ArcdteStamp From Pd_Heat_Plant_Ref Phpr, Pdl_Heat_Data_Melt Phdm,Pdl_Heat_Data phd Where Phpr.Heatid = Phdm.Heatid And Phpr.Plant = Phdm.Plant And Phpr.Treatid = Phdm.Treatid And Phdm.Heatid = Phd.Heatid And Phpr.Plant = 'LF' And Phpr.Expirationdate = 'VALID' And SUBSTR(Phdm.Starttime,1,10)='" + StrDate + "'";

            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclCONARC);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                string StrHEATID = "";
                object StrBADesc = "";
                object StrPlant = "";
                object ArcStat = DBNull.Value;
                object ArcStrt = DBNull.Value;
                object ArcEnnd = DBNull.Value;
                object ArcStart = DBNull.Value;
                object ArcEnd = DBNull.Value;
                object ArcDteStamp = DBNull.Value;
                object carno = DBNull.Value;

                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_Conarc.Rows[i]["HEAT_ID"].ToString();

                    if (dt_Orcl_Conarc.Rows[i]["ARCING_START"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ARCING_START"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        ArcStrt = dt_Orcl_Conarc.Rows[i]["ARCING_START"].ToString();
                    }
                    else
                    { ArcStrt = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["ARCING_END"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ARCING_END"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        ArcEnnd = dt_Orcl_Conarc.Rows[i]["ARCING_END"].ToString();
                    }
                    else
                    { ArcEnnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["ARCSTART"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ARCSTART"].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        ArcStart = dt_Orcl_Conarc.Rows[i]["ARCSTART"].ToString().Replace(",", ".");
                    }
                    else
                    { ArcStart = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["ARCEND"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ARCEND"].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        ArcEnd = dt_Orcl_Conarc.Rows[i]["ARCEND"].ToString().Replace(",", ".");
                    }
                    else
                    { ArcEnd = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["ArcdteStamp"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ArcdteStamp"].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        ArcDteStamp = dt_Orcl_Conarc.Rows[i]["ArcdteStamp"].ToString();
                    }
                    else
                    { ArcDteStamp = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != null && dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrPlant = dt_Orcl_Conarc.Rows[i]["PLANT"].ToString();
                    }
                    else
                    { StrPlant = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["CAR_ID"].ToString() != null && dt_Orcl_Conarc.Rows[i]["CAR_ID"].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        carno = dt_Orcl_Conarc.Rows[i]["CAR_ID"].ToString();
                    }
                    else
                    { carno = DBNull.Value; }
                    if (StrPlant.ToString() == "LF3" && carno.ToString() == "1")
                    {
                        StrMIS_PLANTNO = "11";
                    }
                    else if (StrPlant.ToString() == "LF3" && carno.ToString() == "2")
                    {
                        StrMIS_PLANTNO = "12";
                    }
                    else
                    { }

                    DataTable dt_MIS_CONARC_Arc = new DataTable();
                    string StrQueryCONARC_ta = "select HeatNo from smsHeatBlowArc where HeatNo='" + StrHEATID + "' and Plant='" + StrPlant + "'  and baStart='" + ArcStart + "' and baEnd='" + ArcEnd + "' and blow_arc='Arcing Time' and plantNo='" + StrMIS_PLANTNO + "' and convert(varchar(10),DteStamp,120)=convert(varchar(10),'" + ArcDteStamp + "',120)";
                    dt_MIS_CONARC_Arc = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC_ta);
                    if (dt_MIS_CONARC_Arc.Rows.Count <= 0)
                    {
                        if (ArcStart != DBNull.Value && ArcEnd != DBNull.Value)
                        {
                            string StrInsertQuery = "insert smsHeatBlowArc(plantNo,heatNo,baStart,baEnd,blow_arc,DteStamp,Plant,baStat,batStrt,batEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + ArcStart + "','" + ArcEnd + "','Arcing Time','" + ArcDteStamp + "','" + StrPlant + "','2','" + ArcStrt + "','" + ArcEnnd + "')";
                            bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                        }

                    }

                }
            }
        }

        private void Fn_HeatBlow_Arc(string StrPLANT, string StrBOF_PLANTNO, string StrMIS_PLANTNO, string StrDate)
        {
            StrMIS_PLANTNO = "";
            //string StrQueryOrclCONARC = "select VHD.PLANT || VHD.PLANTNO as PLANT,VHD.HEATID_CUST as HEATID,VHRTB.BLOWSTART,VHRTB.BLOWEND,decode(replace(SUBSTR(VHRTB.BLOWSTART,11,6) ,':','.'),0,null,replace(SUBSTR(VHRTB.BLOWSTART,11,6) ,':','.')) AS baStart,decode(replace(SUBSTR(VHRTB.BLOWEND,11,6) ,':','.'),0,null,replace(SUBSTR(VHRTB.BLOWEND,11,6) ,':','.')) AS baEnd,VHRTB.TAPSTART,VHRTB.TAPEND,decode(replace(SUBSTR(VHRTB.TAPSTART,11,6) ,':','.'),0,null,replace(SUBSTR(VHRTB.TAPSTART,11,6) ,':','.')) AS taStart,decode(replace(SUBSTR(VHRTB.TAPEND,11,6) ,':','.'),0,null,replace(SUBSTR(VHRTB.TAPEND,11,6) ,':','.')) AS taEnd,SUBSTR(VHRTB.BLOWSTART,1,10) AS baDateStamp,SUBSTR(VHRTB.TAPSTART,1,10) AS taDateStamp,(select carno from Pdl_Heat_Data Pdh where Pdh.HEATID=VHD.HEATID) as carno from V_HEAT_DATA VHD inner join V_HEAT_REPORT_TIMES_BOF VHRTB on VHD.HEATID=VHRTB.HEAT_ID inner join V_HEAT_REPORT_MAIN_BOF VHRMB  on VHD.HEATID=VHRMB.HEAT_ID and VHRTB.HEAT_ID=VHRMB.HEAT_ID where VHD.TREATSTART_ACT in(select TREATSTART_ACT from V_HEAT_DATA where SUBSTR(TREATSTART_ACT,1,10)='" + StrDate + "') order by TREATSTART_ACT asc";
            string StrQueryOrclCONARC = "select VHD.PLANT || VHD.PLANTNO as PLANT,VHD.HEATID_CUST as HEATID,VHRTB.BLOWSTART,VHRTB.BLOWEND,decode(replace(SUBSTR(VHRTB.BLOWSTART,11,6) ,':','.'),0,null,replace(SUBSTR(VHRTB.BLOWSTART,11,6) ,':','.')) AS baStart,decode(replace(SUBSTR(VHRTB.BLOWEND,11,6) ,':','.'),0,null,replace(SUBSTR(VHRTB.BLOWEND,11,6) ,':','.')) AS baEnd,VHRTB.TAPSTART,VHRTB.TAPEND,decode(replace(SUBSTR(VHRTB.TAPSTART,11,6) ,':','.'),0,null,replace(SUBSTR(VHRTB.TAPSTART,11,6) ,':','.')) AS taStart,decode(replace(SUBSTR(VHRTB.TAPEND,11,6) ,':','.'),0,null,replace(SUBSTR(VHRTB.TAPEND,11,6) ,':','.')) AS taEnd,SUBSTR(VHRTB.BLOWSTART,1,10) AS baDateStamp,SUBSTR(VHRTB.TAPSTART,1,10) AS taDateStamp from V_HEAT_DATA VHD inner join V_HEAT_REPORT_TIMES_BOF VHRTB on VHD.HEATID=VHRTB.HEAT_ID inner join V_HEAT_REPORT_MAIN_BOF VHRMB  on VHD.HEATID=VHRMB.HEAT_ID and VHRTB.HEAT_ID=VHRMB.HEAT_ID where VHD.TREATSTART_ACT in(select TREATSTART_ACT from V_HEAT_DATA where SUBSTR(TREATSTART_ACT,1,10)='" + StrDate + "' and PLANT = 'BOF') order by TREATSTART_ACT asc";
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclCONARC);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                string StrHEATID = "";
                object StrBADesc = "";
                object StrPlant = "";
                object baStart = DBNull.Value;
                object baEnd = DBNull.Value;
                object baDteStamp = DBNull.Value;
                object taDteStamp = DBNull.Value;
                object baStat = DBNull.Value;
                object baStrt = DBNull.Value;
                object baEnnd = DBNull.Value;
                object taStart = DBNull.Value;
                object taEnd = DBNull.Value;
                object taStrt = DBNull.Value;
                object taEnnd = DBNull.Value;
                //object carno = DBNull.Value;

                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_Conarc.Rows[i]["HEATID"].ToString();

                    if (dt_Orcl_Conarc.Rows[i]["BASTART"].ToString() != null && dt_Orcl_Conarc.Rows[i]["BASTART"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baStart = dt_Orcl_Conarc.Rows[i]["BASTART"].ToString();
                    }
                    else
                    { baStart = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["BAEND"].ToString() != null && dt_Orcl_Conarc.Rows[i]["BAEND"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baEnd = dt_Orcl_Conarc.Rows[i]["BAEND"].ToString();
                    }
                    else
                    { baEnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["BLOWSTART"].ToString() != null && dt_Orcl_Conarc.Rows[i]["BLOWSTART"].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baStrt = dt_Orcl_Conarc.Rows[i]["BLOWSTART"].ToString().Replace(",", ".");
                    }
                    else
                    { baStrt = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["BLOWEND"].ToString() != null && dt_Orcl_Conarc.Rows[i]["BLOWEND"].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baEnnd = dt_Orcl_Conarc.Rows[i]["BLOWEND"].ToString().Replace(",", ".");
                    }
                    else
                    { baEnnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["TASTART"].ToString() != null && dt_Orcl_Conarc.Rows[i]["TASTART"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        taStart = dt_Orcl_Conarc.Rows[i]["TASTART"].ToString();
                    }
                    else
                    { taStart = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["TAEND"].ToString() != null && dt_Orcl_Conarc.Rows[i]["TAEND"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        taEnd = dt_Orcl_Conarc.Rows[i]["TAEND"].ToString();
                    }
                    else
                    { taEnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["TAPSTART"].ToString() != null && dt_Orcl_Conarc.Rows[i]["TAPSTART"].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        taStrt = dt_Orcl_Conarc.Rows[i]["TAPSTART"].ToString().Replace(",", ".");
                    }
                    else
                    { taStrt = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["TAPEND"].ToString() != null && dt_Orcl_Conarc.Rows[i]["TAPEND"].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        taEnnd = dt_Orcl_Conarc.Rows[i]["TAPEND"].ToString().Replace(",", ".");
                    }
                    else
                    { taEnnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["BADATESTAMP"].ToString() != null && dt_Orcl_Conarc.Rows[i]["BADATESTAMP"].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baDteStamp = dt_Orcl_Conarc.Rows[i]["BADATESTAMP"].ToString();
                    }
                    else
                    { baDteStamp = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["TADATESTAMP"].ToString() != null && dt_Orcl_Conarc.Rows[i]["TADATESTAMP"].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        taDteStamp = dt_Orcl_Conarc.Rows[i]["TADATESTAMP"].ToString();
                    }
                    else
                    { taDteStamp = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != null && dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrPlant = dt_Orcl_Conarc.Rows[i]["PLANT"].ToString();
                    }
                    else
                    { StrPlant = DBNull.Value; }
                    //if (dt_Orcl_Conarc.Rows[i]["carno"].ToString() != null && dt_Orcl_Conarc.Rows[i]["carno"].ToString() != string.Empty)
                    //{
                    //    //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                    //    carno = dt_Orcl_Conarc.Rows[i]["carno"].ToString();
                    //}
                    //else
                    //{ carno = DBNull.Value; }
                    if (StrPlant.ToString() == "BOF1")
                    {
                        StrMIS_PLANTNO = "5";
                    }
                    else if (StrPlant.ToString() == "BOF2")
                    {
                        StrMIS_PLANTNO = "6";
                    }
                    //else if (StrPlant.ToString() == "LF3" && carno.ToString() == "1")
                    //{
                    //    StrMIS_PLANTNO = "11";
                    //}
                    //else if (StrPlant.ToString() == "LF3" && carno.ToString() == "2")
                    //{
                    //    StrMIS_PLANTNO = "12";
                    //}
                    else
                    { }

                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select HeatNo from smsHeatBlowArc where HeatNo='" + StrHEATID + "' and Plant='" + StrPlant + "'  and baStart='" + baStart + "' and baEnd='" + baEnd + "' and blow_arc='Blowing Time' and plantNo='" + StrMIS_PLANTNO + "' and convert(varchar(10),DteStamp,120)=convert(varchar(10),'" + baDteStamp + "',120)";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count <= 0)
                    {
                        if (baStart != DBNull.Value && baEnd != DBNull.Value)
                        {
                            string StrInsertQuery = "insert smsHeatBlowArc(plantNo,heatNo,baStart,baEnd,blow_arc,DteStamp,Plant,baStat,batStrt,batEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + baStart + "','" + baEnd + "','Blowing Time','" + baDteStamp + "','" + StrPlant + "','1','" + baStrt + "','" + baEnnd + "')";
                            bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                        }
                    }
                    DataTable dt_MIS_CONARC_ta = new DataTable();
                    string StrQueryCONARC_ta = "select HeatNo from smsHeatBlowArc where HeatNo='" + StrHEATID + "' and Plant='" + StrPlant + "'  and baStart='" + taStart + "' and baEnd='" + taEnd + "' and blow_arc='Taping Time' and plantNo='" + StrMIS_PLANTNO + "' and convert(varchar(10),DteStamp,120)=convert(varchar(10),'" + taDteStamp + "',120)";
                    dt_MIS_CONARC_ta = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC_ta);
                    if (dt_MIS_CONARC_ta.Rows.Count <= 0)
                    {
                        if (taStart != DBNull.Value && taEnd != DBNull.Value)
                        {
                            string StrInsertQuery = "insert smsHeatBlowArc(plantNo,heatNo,baStart,baEnd,blow_arc,DteStamp,Plant,baStat,batStrt,batEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + taStart + "','" + taEnd + "','Taping Time','" + taDteStamp + "','" + StrPlant + "','3','" + taStrt + "','" + taEnnd + "')";
                            bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                        }

                    }

                }
            }
        }

        private void FnSMC_LF()
        {

            string StrHEATID = "";
            string StrTREATID = "";
            string StrStatus = "";
            string StrStartDate = "";
            //string StrQueryOrclLF = "SELECT p.PLANT  || p.PLANTNO AS UNIT,  p.HEATID_CUST HEAT_ID,  p.TREATID_CUST TREATID,  grade.STEELGRADECODE, status.HEATSTATUS, status.HEATSTATUSDESC STATUS_DESC,  SUBSTR(DECODE(p.ACTTREATSTART, NULL, p.PLANTREATSTART, p.ACTTREATSTART), 0, 19) START_T,  SUBSTR(DECODE(p.ACTTREATEND, NULL, p.ACTTREATEND), 0, 19) END_T,  de.LADLENO,  de.FIRST_POWER_ON,  de.LAST_POWER_OFF,  de.POWER_ON_DUR FROM PP_HEAT_PLANT p,  PP_HEAT ph,  (SELECT g.STEELGRADECODE, g.STEELGRADECODEDESC FROM PP_GRADE g  ) grade,  (SELECT UNIQUE s.ORD_STAT_DESC,    s.ORD_STAT_NO,    s.ORD_STAT_CODE,    o.PRODORDERID,    o.STEELGRADECODE   FROM GC_PRD_STAT s,     PP_ORDER o   WHERE o.ORD_STAT_NO = s.ORD_STAT_NO   ) ord,   (SELECT st.HEATSTATORDER,     st.HEATSTATUS,     st.HEATSTATUSDESC   FROM GC_HEAT_STAT st   ) status,  PDL_HEAT_DATA de WHERE ph.HEATID_CUST(+)     = p.HEATID_CUST AND grade.STEELGRADECODE(+) = ord.STEELGRADECODE AND ord.PRODORDERID(+)      = ph.PRODORDERID AND de.HEATID(+)            = p.HEATID AND de.TREATID(+)           = p.TREATID AND status.HEATSTATORDER(+) = ph.HEATSTATORDER AND (p.PLANT                = 'LF' AND p.EXPIRATIONDATE        = 'VALID' AND p.METAL_TYPE            = 'Steel' AND ph.METAL_TYPE           = 'Steel' AND p.ACTTREATSTART        IS NOT NULL AND p.ACTTREATEND          IS NULL)";
            //string StrQueryOrclLF = "SELECT p.PLANT  || p.PLANTNO AS UNIT,  p.HEATID_CUST HEAT_ID,  p.TREATID_CUST TREATID,  grade.STEELGRADECODE,  status.HEATSTATUS,  status.HEATSTATUSDESC STATUS_DESC,  SUBSTR(DECODE(p.ACTTREATSTART, NULL, p.PLANTREATSTART, p.ACTTREATSTART), 0, 19) START_T,  SUBSTR(DECODE(p.ACTTREATEND, NULL, p.ACTTREATEND, p.ACTTREATEND), 0, 19) END_T,  de.LADLENO,  de.FIRST_POWER_ON,  de.LAST_POWER_OFF,  de.POWER_ON_DUR FROM PP_HEAT_PLANT p,   PP_HEAT ph,   (SELECT g.STEELGRADECODE, g.STEELGRADECODEDESC FROM PP_GRADE g   ) grade,   (SELECT UNIQUE s.ORD_STAT_DESC,     s.ORD_STAT_NO,     s.ORD_STAT_CODE,     o.PRODORDERID,     o.STEELGRADECODE   FROM GC_PRD_STAT s,     PP_ORDER o   WHERE o.ORD_STAT_NO = s.ORD_STAT_NO   ) ord,   (SELECT st.HEATSTATORDER,     st.HEATSTATUS,     st.HEATSTATUSDESC   FROM GC_HEAT_STAT st   ) status,   PDL_HEAT_DATA de WHERE ph.HEATID_CUST(+)     = p.HEATID_CUST AND grade.STEELGRADECODE(+) = ord.STEELGRADECODE AND ord.PRODORDERID(+)      = ph.PRODORDERID AND de.HEATID(+)            = p.HEATID AND de.TREATID(+)           = p.TREATID AND status.HEATSTATORDER(+) = ph.HEATSTATORDER AND (p.EXPIRATIONDATE        = 'VALID' AND p.METAL_TYPE            = 'Steel' AND ph.METAL_TYPE           = 'Steel' AND SUBSTR(p.ACTTREATSTART,1,10)='" + System.DateTime.Now.Date.ToString("yyyy-MM-dd") + "') ORDER BY p.acttreatstart DESC";
            //string StrQueryOrclLF = "select hr.PLANT||3 AS UNIT,hr.HEAT_ID_CUST AS HEATID,hr.TREAT_ID_CUST AS TREATID,hr.GRADE_ACT as GRADE,'' AS STATUS,'' AS STATUS_DESC,hr.TREAT_START AS TEARTSTART,hr.TREAT_END AS TREATEND,'' AS LADLENO,hrt.POWER_ON AS FIRSTPOWERON,hrt.POWER_OFF AS FIRSTPOWEROFF,hrt.POWER_ON_DUR AS POWERDUR from V_HEAT_REPORT_LF hr  inner join V_HEAT_REPORT_LF_TIMEDUR hrt on hr.HEAT_ID=hrt.HEAT_ID and SUBSTR(hr.TREAT_END,1,10)='" + System.DateTime.Now.Date.ToString("yyyy-MM-dd") + "'";
            string StrQueryOrclLF = "SELECT Vhd.Plant ||Vhd.Plantno || decode(Pdh.carno,1,'CAR1',2,'CAR2') AS Unit,  Vhd.Heatid_Cust  AS Heatid,  Vhd.Treatid_Cust       AS Treatid,  Vhd.Steelgradecode_Act AS Grade,  Vhd.Statusno           AS Status,  (SELECT A.Heatstatus FROM Gc_Heat_Stat A WHERE A.Heatstatusno = Vhd.Statusno  )                  AS Status_Desc,  Vhd.Treatstart_Act AS Treatstart,  Vhd.Treatend_Act   As Treatend,  Vhd.Ladleno        As Ladleno,  Pdh.First_Power_On ,  Pdh.Last_Power_Off,  Pdh.Power_On_Dur From V_Heat_Data Vhd,Pdl_Heat_Data Pdh Where Vhd.Heatid = Pdh.Heatid And Vhd.Treatid = Pdh.Treatid AND Vhd.Plant = Pdh.Plant AND Vhd.Plant = 'LF'AND Vhd.Plantno = 3 AND SUBSTR(Vhd.Treatstart_Act,1,10)='" + System.DateTime.Now.Date.ToString("yyyy-MM-dd") + "' ORDER BY Vhd.Treatstart_Act DESC";// BETWEEN '2014-05-29 22:00:00' AND '2014-05-30 22:00:00' ORDER BY Vhd.Treatstart_Act DESC";

            DataTable dt_Orcl_LF = new DataTable();
            dt_Orcl_LF = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclLF);
            if (dt_Orcl_LF.Rows.Count > 0)
            {
                string StrUpdateDefault = ";UPDATE OP_LF_STATUS set END_T=null where END_T='1900-01-01 00:00:00.000';UPDATE OP_LF_STATUS set FIRST_POWER_ON=null where FIRST_POWER_ON='1900-01-01 00:00:00.000';UPDATE OP_LF_STATUS set LAST_POWER_OFF=null where LAST_POWER_OFF='1900-01-01 00:00:00.000'";
                for (int i = 0; i < dt_Orcl_LF.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_LF.Rows[i][1].ToString();
                    StrTREATID = dt_Orcl_LF.Rows[i][2].ToString();
                    StrStatus = dt_Orcl_LF.Rows[i][4].ToString();
                    StrStartDate = dt_Orcl_LF.Rows[i][5].ToString();
                    DataTable dt_MIS_LF = new DataTable();
                    string StrQueryMISLF = "select SLNo,HEAT_ID,TREAT_ID,STATUS,START_T from OP_LF_STATUS WHERE HEAT_ID='" + StrHEATID + "' AND TREAT_ID='" + StrTREATID + "' AND STATUS='" + StrStatus + "' and convert(varchar(10),START_T,120)=convert(varchar(10),'" + StrStartDate + "',120)";
                    dt_MIS_LF = clsObj.DBSelectQueryMIS_Table(StrQueryMISLF);
                    if (dt_MIS_LF.Rows.Count == 0)
                    {
                        string StrInsertQuery = "insert into OP_LF_STATUS(UNIT,HEAT_ID,TREAT_ID,STEL_GRADE_CODE,STATUS,STATUS_DESC,START_T,END_T,LADLENO,FIRST_POWER_ON,LAST_POWER_OFF,POWER_ON_DUR)values('" + dt_Orcl_LF.Rows[i][0].ToString() + "','" + dt_Orcl_LF.Rows[i][1].ToString() + "','" + dt_Orcl_LF.Rows[i][2].ToString() + "','" + dt_Orcl_LF.Rows[i][3].ToString() + "','" + dt_Orcl_LF.Rows[i][4].ToString() + "','" + dt_Orcl_LF.Rows[i][5].ToString().Replace(",", ".") + "','" + dt_Orcl_LF.Rows[i][6].ToString().Replace(",", ".") + "','" + dt_Orcl_LF.Rows[i][7].ToString().Replace(",", ".") + "','" + dt_Orcl_LF.Rows[i][8].ToString() + "','" + dt_Orcl_LF.Rows[i][9].ToString().Replace(",", ".") + "','" + dt_Orcl_LF.Rows[i][10].ToString().Replace(",", ".") + "','" + dt_Orcl_LF.Rows[i][11].ToString() + "')";
                        StrInsertQuery = StrInsertQuery + StrUpdateDefault;
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                    else
                    {
                        string StrUpdateQuery = "Update OP_LF_STATUS SET UNIT='" + dt_Orcl_LF.Rows[i][0].ToString() + "',HEAT_ID='" + dt_Orcl_LF.Rows[i][1].ToString() + "',TREAT_ID='" + dt_Orcl_LF.Rows[i][2].ToString() + "',STEL_GRADE_CODE='" + dt_Orcl_LF.Rows[i][3].ToString() + "',STATUS='" + dt_Orcl_LF.Rows[i][4].ToString() + "',STATUS_DESC='" + dt_Orcl_LF.Rows[i][5].ToString() + "',START_T='" + dt_Orcl_LF.Rows[i][6].ToString().Replace(",", ".") + "',END_T='" + dt_Orcl_LF.Rows[i][7].ToString().Replace(",", ".") + "',LADLENO='" + dt_Orcl_LF.Rows[i][8].ToString() + "',FIRST_POWER_ON='" + dt_Orcl_LF.Rows[i][9].ToString().Replace(",", ".") + "',LAST_POWER_OFF='" + dt_Orcl_LF.Rows[i][10].ToString().Replace(",", ".") + "',POWER_ON_DUR='" + dt_Orcl_LF.Rows[i][11].ToString() + "' where SLNo='" + dt_MIS_LF.Rows[0][0].ToString() + "'";
                        StrUpdateQuery = StrUpdateQuery + StrUpdateDefault;
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                }
            }

        }
        private void FnSMS_SMCDB_HeatTracking_LF(string StrPLANT, string StrLF_PLANTNO, string StrMIS_PLANTNO, string StrDate)
        {
            //System.DateTime.Now.ToString("yyyy-MM-dd")
           // string StrQueryOrclLF = "SELECT  hp.HEATID_CUST_PLAN AS HEAT_ID,hp.PLANT ||hp.PLANTNO || decode(Pdh.carno,1,'CAR1',2,'CAR2') AS PLANT,hp.HEATID AS TREAT_ID,hd.LADLETYPE || hd.LADLENO AS LADLE_ID,decode(replace(SUBSTR(hp.TREATSTART_PLAN,11,6) ,':','.'),0,null,replace(SUBSTR(hp.TREATSTART_PLAN,11,6) ,':','.')) AS PLAN_START_TIME, decode(replace(SUBSTR(hp.TREATEND_PLAN,11,6) ,':','.'),0,null,replace(SUBSTR(hp.TREATEND_PLAN,11,6) ,':','.'))   AS PLAN_END_TIME,  replace(SUBSTR(TREATSTART_ACT,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.'),0,null,  replace(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hd.TREATSTART_ACT,1,10) as Act_Start_Dt,hp.HEATID,hd.STEELGRADECODE_ACT as Grade,replace(TREATSTART_ACT ,',','.') AS actStrt,replace(TREATEND_ACT ,',','.') AS actEnnd,hd.FINALTEMP  FROM PP_HEAT_PLANT hp inner join PD_HEAT_DATA hd on hp.HEATID=hd.HEATID inner join Pdl_Heat_Data Pdh on Pdh.PLANT=hp.PLANT AND Pdh.HEATID=hp.HEATID AND Pdh.ELEC_CONS_TOTAL IS NOT NULL where hd.PLANT='" + StrPLANT + "'  and SUBSTR(hd.TREATSTART_ACT,1,10)='" + StrDate + "'   and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrLF_PLANTNO + "' ORDER BY hp.HEATID DESC";
            string StrQueryOrclLF = "select * from (SELECT Pdh.heatid as intheatid, phpr.heatid_cust AS HEAT_ID,phpr.plant||phpr.plantno||DECODE(Pdh.carno,1,'CAR1',2,'CAR2') AS PLANT,hd.treatid AS TREAT_ID,hd.LADLETYPE||hd.LADLENO AS LADLE_ID,DECODE(REPLACE(SUBSTR(hd.TREATSTART_PLAN,11,6) ,':','.'),0,NULL,REPLACE(SUBSTR(hd.TREATSTART_PLAN,11,6) ,':','.')) AS PLAN_START_TIME, DECODE(REPLACE(SUBSTR(hd.TREATEND_PLAN,11,6) ,':','.'),0,NULL,REPLACE(SUBSTR(hd.TREATEND_PLAN,11,6) ,':','.')) AS PLAN_END_TIME, REPLACE(SUBSTR(TREATSTART_ACT,11,6) ,':','.') AS ACT_START_DATE, DECODE(REPLACE(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.'),0,NULL, REPLACE(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.') ) AS ACT_END_DATE ,hd.ladleno,SUBSTR(hd.TREATSTART_ACT,1,10) AS Act_Start_Dt,hd.HEATID,hd.STEELGRADECODE_ACT as Grade,REPLACE(TREATSTART_ACT ,',','.') AS actStrt,REPLACE(TREATEND_ACT ,',','.')   AS actEnnd,hd.FINALTEMP,pdh.elec_cons_total as Total_Energy  FROM pd_heat_plant_ref phpr, PD_HEAT_DATA hd, Pdl_Heat_Data Pdh where Pdh.heatid = hd.heatid AND Pdh.treatid = hd.treatid AND Pdh.plant = hd.plant AND hd.heatid = phpr.heatid AND hd.treatid = phpr.treatid AND hd.plant = phpr.plant  AND phpr.EXPIRATIONDATE = 'VALID' AND phpr.plant ='" + StrPLANT + "' And phpr.plantno = '" + StrLF_PLANTNO + "' And SUBSTR(hd.TREATSTART_ACT,1,10) <= '" + StrDate + "'order by TREATEND_ACT desc) where rownum < = 7"; 
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclLF);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                string StrHEATID = "";
                string StrTREATID = "";
                object StrActStart = DBNull.Value;
                object StrActEnd = DBNull.Value;
                object StrPlndStart = DBNull.Value;
                object StrPlndEnd = DBNull.Value;
                object StrLadleNo = DBNull.Value;
                object StrActStartDte = DBNull.Value;
                object StrGrade = DBNull.Value;
                object StrActStrt = DBNull.Value;
                object StrActEnnd = DBNull.Value;
                object Plant = DBNull.Value;
                object StrLastTemp = DBNull.Value;
                object Total_Energy = DBNull.Value;
                object InternalHeatId = DBNull.Value;
                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_Conarc.Rows[i][0].ToString();
                    StrTREATID = dt_Orcl_Conarc.Rows[i][2].ToString();

                    if (dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != null && dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        Plant = dt_Orcl_Conarc.Rows[i]["PLANT"].ToString();
                    }
                    else
                    { Plant = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][6].ToString() != null && dt_Orcl_Conarc.Rows[i][6].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActStart = dt_Orcl_Conarc.Rows[i][6].ToString();
                    }
                    else
                    { StrActStart = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i][7].ToString() != null && dt_Orcl_Conarc.Rows[i][7].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActEnd = dt_Orcl_Conarc.Rows[i][7].ToString();
                    }
                    else
                    { StrActEnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][4].ToString() != null && dt_Orcl_Conarc.Rows[i][4].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrPlndStart = dt_Orcl_Conarc.Rows[i][4].ToString();
                    }
                    else
                    { StrPlndStart = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][5].ToString() != null && dt_Orcl_Conarc.Rows[i][5].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrPlndEnd = dt_Orcl_Conarc.Rows[i][5].ToString();
                    }
                    else
                    { StrPlndEnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][8].ToString() != null && dt_Orcl_Conarc.Rows[i][8].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrLadleNo = dt_Orcl_Conarc.Rows[i][8].ToString();
                    }
                    else
                    { StrLadleNo = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActStartDte = dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString();
                    }
                    else
                    { StrActStartDte = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["Grade"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Grade"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrGrade = dt_Orcl_Conarc.Rows[i]["Grade"].ToString();
                    }
                    else
                    { StrGrade = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["actStrt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["actStrt"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActStrt = dt_Orcl_Conarc.Rows[i]["actStrt"].ToString();
                    }
                    else
                    { StrActStrt = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["actEnnd"].ToString() != null && dt_Orcl_Conarc.Rows[i]["actEnnd"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActEnnd = dt_Orcl_Conarc.Rows[i]["actEnnd"].ToString();
                    }
                    else
                    { StrActEnnd = DBNull.Value; }
                    // +Manish@13092014 We are selecting actual Temperature So Required internal heatid to pass in table pd_sample for measured Temp.

                    if (dt_Orcl_Conarc.Rows[i]["intheatid"].ToString() != null && dt_Orcl_Conarc.Rows[i]["intheatid"].ToString() != string.Empty)
                    {
                        
                        InternalHeatId = dt_Orcl_Conarc.Rows[i]["intheatid"].ToString();

                        DataTable Temerature = new DataTable();

                        string FetchTemperature = " SELECT Se.Measvalue AS TEMP FROM pd_sample sa , Pd_Sample_Entry se Where Sa.Sample_Counter = Se.Sample_Counter And Sa.Plant = 'LF' And Sa.Plantno = '3' And Sa.Origin = '1' And Sa.Meastypeno = '5' and Se.Measvalue < 2000 And Sa.Heatid = '" + InternalHeatId  + "' order by Sa.Sampleno desc ";

                        Temerature = clsObj.DBSelectQueryStatusScreen_Table(FetchTemperature);

                        if (Temerature.Rows.Count < 0)
                        {
                            StrLastTemp = dt_Orcl_Conarc.Rows[0]["TEMP"].ToString();
                        }

                        else
                       { StrLastTemp = DBNull.Value; }
                        
                    }
                    else
                    { InternalHeatId = DBNull.Value; }
                    //if (dt_Orcl_Conarc.Rows[i]["FINALTEMP"].ToString() != null && dt_Orcl_Conarc.Rows[i]["FINALTEMP"].ToString() != string.Empty)
                    //{
                    //    //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                    //    StrLastTemp = dt_Orcl_Conarc.Rows[i]["FINALTEMP"].ToString();
                    //}
                    //else
                    //{ StrLastTemp = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["Total_Energy"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Total_Energy"].ToString() != string.Empty)
                    {
                        Total_Energy = dt_Orcl_Conarc.Rows[i]["Total_Energy"].ToString();
                    }
                    else
                    { Total_Energy = DBNull.Value; }
                    if (Plant.ToString() == "LF3CAR1")
                    {
                        StrMIS_PLANTNO = "11";
                    }
                    else
                    {
                        StrMIS_PLANTNO = "12";
                    }
                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select heatNo from smsHeatTracker where heatNo='" + StrHEATID + "' and plantNo='" + StrMIS_PLANTNO + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120)";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count > 0)
                    {
                        //string StrUpdateQuery = "Update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',plndStart='" + StrPlndStart + "',plndEnd='" + StrPlndEnd + "',LadleNo='" + StrLadleNo + "',actEnds='" + StrActEnd + "',Grade='" + StrGrade + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "',LastTemp='" + StrLastTemp + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120)";
                        string StrUpdateQuery = "Update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',plndStart='" + StrPlndStart + "',plndEnd='" + StrPlndEnd + "',LadleNo='" + StrLadleNo + "',actEnds='" + StrActEnd + "',Grade='" + StrGrade + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "',LastTemp='" + StrLastTemp + "',Tot_Energy='" + Total_Energy + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120)";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        //string StrInsertQuery = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,plndStart,plndEnd,DateStamp,LadleNo,actEnds,Grade,actStrt,actEnnd,LastTemp)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + StrActStart + "','" + StrActEnd + "','','" + StrPlndStart + "','" + StrPlndEnd + "','" + StrActStartDte + "','" + StrLadleNo + "','" + StrActEnd + "','" + StrGrade + "','" + StrActStrt + "','" + StrActEnnd + "','" + StrLastTemp + "')";
                        string StrInsertQuery = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,plndStart,plndEnd,DateStamp,LadleNo,actEnds,Grade,actStrt,actEnnd,LastTemp,Tot_Energy)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + StrActStart + "','" + StrActEnd + "','','" + StrPlndStart + "','" + StrPlndEnd + "','" + StrActStartDte + "','" + StrLadleNo + "','" + StrActEnd + "','" + StrGrade + "','" + StrActStrt + "','" + StrActEnnd + "','" + StrLastTemp + "','" + Total_Energy + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void FnSMS_SMCDB_HeatTracking_LF_Update(string StrPLANT, string StrLF_PLANTNO, string StrMIS_PLANTNO, string StrDate)
        {
            string StrUptQuery = "UPDATE t   SET t.actEnnd = c.actStrrt,t.actEnd=c.actEnd,t.actEnds=c.actEnd FROM smsHeatTracker t INNER JOIN (SELECT a.plantNo,a.heatNo,a.DateStamp,a.actStrt as actStart,b.actStrt,dateadd(minute,-1,b.actStrt) as actStrrt ,replace(convert(char(5), dateadd(minute,-1,b.actStrt), 108),':','.') as actEnd FROM ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt) , * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) b right JOIN ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt)+1, * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) a ON b.Id = a.Id) c ON  c.heatNo=t.heatNo and c.plantNo=t.plantNo and c.DateStamp=t.DateStamp and c.actStart=t.actStrt";
            bool StrUpdtStatus = clsObj.DBInsertUpdateDeleteMIS(StrUptQuery);
            return;

            DataTable dt_MIS_LF_Chk = new DataTable();
            string StrQueryCheck = "select heatNo from smsHeatTracker where actEnnd is NULL and plantNo='" + StrMIS_PLANTNO + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrDate + "',120)";
            dt_MIS_LF_Chk = clsObj.DBSelectQueryMIS_Table(StrQueryCheck);
            if (dt_MIS_LF_Chk.Rows.Count < 0)
            {
                return;
            }
            //System.DateTime.Now.ToString("yyyy-MM-dd")
            //string StrQueryOrclCONARC = "SELECT  hp.HEATID_CUST AS HEAT_ID,hp.PLANT ||hp.PLANTNO    AS PLANT,hp.TREATID_CUST AS TREAT_ID,hp.PLANLADLETYPE ||hp.PLANLADLENO AS LADLE_ID, decode(replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.')) AS PLAN_START_TIME,decode(replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'))   AS PLAN_END_TIME, replace(SUBSTR(ACTTREATSTART,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hp.ACTTREATSTART,1,10) as Act_Start_Dt,hp.HEATID FROM PP_HEAT_PLANT hp left outer join PD_HEATDATA hd on hp.HEATID=hd.HEATID   where hd.PLANT='" + StrPLANT + "' and  hp.EXPIRATIONDATE='VALID' and SUBSTR(hp.ACTTREATSTART,1,10)='2014-04-15' and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrPLANTNO + "' ORDER BY hp.HEATID DESC ";
            //string StrQueryOrclCONARC = "SELECT  hp.HEATID_CUST AS HEAT_ID,hp.PLANT ||hp.PLANTNO    AS PLANT,hp.TREATID_CUST AS TREAT_ID,hp.PLANLADLETYPE ||hp.PLANLADLENO AS LADLE_ID, decode(replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.')) AS PLAN_START_TIME,decode(replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'))   AS PLAN_END_TIME, replace(SUBSTR(ACTTREATSTART,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hp.ACTTREATSTART,1,10) as Act_Start_Dt,hp.HEATID FROM PP_HEAT_PLANT hp left outer join PD_HEATDATA hd on hp.HEATID=hd.HEATID   where hd.PLANT='" + StrPLANT + "' and  hp.EXPIRATIONDATE='VALID' and ACTTREATSTART>=(select max(ACTTREATSTART) from PP_HEAT_PLANT where ACTTREATSTART<(select max(ACTTREATSTART) from PP_HEAT_PLANT where PLANT='" + StrPLANT + "'and PLANTNO='" + StrPLANTNO + "' and SUBSTR(ACTTREATSTART,1,10)='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "') and PLANT='" + StrPLANT + "'and PLANTNO='" + StrPLANTNO + "') and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrPLANTNO + "' ORDER BY hp.HEATID DESC ";
            //string StrQueryOrclCONARC = "SELECT  hp.HEATID_CUST AS HEAT_ID,hp.PLANT ||hp.PLANTNO    AS PLANT,hp.TREATID_CUST AS TREAT_ID,hp.PLANLADLETYPE ||hp.PLANLADLENO AS LADLE_ID, decode(replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.')) AS PLAN_START_TIME,decode(replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'))   AS PLAN_END_TIME, replace(SUBSTR(ACTTREATSTART,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hp.ACTTREATSTART,1,10) as Act_Start_Dt,hp.HEATID FROM PP_HEAT_PLANT hp left outer join PD_HEATDATA hd on hp.HEATID=hd.HEATID   where hd.PLANT='" + StrPLANT + "' and  hp.EXPIRATIONDATE='VALID' and SUBSTR(hp.ACTTREATSTART,1,10)='" + StrDate + "' and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrLF_PLANTNO + "' ORDER BY hp.HEATID DESC ";
            //string StrQueryOrclLF = "SELECT  hp.HEATID_CUST_PLAN AS HEAT_ID,hp.PLANT ||hp.PLANTNO    AS PLANT,hp.HEATID AS TREAT_ID,hd.LADLETYPE || hd.LADLENO AS LADLE_ID,decode(replace(SUBSTR(hp.TREATSTART_PLAN,11,6) ,':','.'),0,null,replace(SUBSTR(hp.TREATSTART_PLAN,11,6) ,':','.')) AS PLAN_START_TIME, decode(replace(SUBSTR(hp.TREATEND_PLAN,11,6) ,':','.'),0,null,replace(SUBSTR(hp.TREATEND_PLAN,11,6) ,':','.'))   AS PLAN_END_TIME,  replace(SUBSTR(TREATSTART_ACT,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.'),0,null,  replace(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hd.TREATSTART_ACT,1,10) as Act_Start_Dt,hp.HEATID,hd.STEELGRADECODE_ACT as Grade,replace(TREATSTART_ACT ,',','.') AS actStrt,replace(TREATEND_ACT ,',','.') AS actEnnd FROM PP_HEAT_PLANT hp inner join PD_HEAT_DATA hd on hp.HEATID=hd.HEATID   where hd.PLANT='" + StrPLANT + "'  and SUBSTR(hd.TREATSTART_ACT,1,10)='" + StrDate + "'   and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrLF_PLANTNO + "' ORDER BY hp.HEATID DESC";
            string StrQueryOrclLF = "SELECT  hp.HEATID_CUST_PLAN AS HEAT_ID,hp.PLANT ||hp.PLANTNO || decode(Pdh.carno,1,'CAR1',2,'CAR2') AS PLANT,hp.HEATID AS TREAT_ID,hd.LADLETYPE || hd.LADLENO AS LADLE_ID,decode(replace(SUBSTR(hp.TREATSTART_PLAN,11,6) ,':','.'),0,null,replace(SUBSTR(hp.TREATSTART_PLAN,11,6) ,':','.')) AS PLAN_START_TIME, decode(replace(SUBSTR(hp.TREATEND_PLAN,11,6) ,':','.'),0,null,replace(SUBSTR(hp.TREATEND_PLAN,11,6) ,':','.'))   AS PLAN_END_TIME,  replace(SUBSTR(TREATSTART_ACT,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.'),0,null,  replace(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hd.TREATSTART_ACT,1,10) as Act_Start_Dt,hp.HEATID,hd.STEELGRADECODE_ACT as Grade,replace(TREATSTART_ACT ,',','.') AS actStrt,replace(TREATEND_ACT ,',','.') AS actEnnd,hd.FINALTEMP  FROM PP_HEAT_PLANT hp inner join PD_HEAT_DATA hd on hp.HEATID=hd.HEATID inner join Pdl_Heat_Data Pdh on Pdh.PLANT=hp.PLANT AND Pdh.HEATID=hp.HEATID AND Pdh.ELEC_CONS_TOTAL IS NOT NULL where hd.PLANT='" + StrPLANT + "'  and SUBSTR(hd.TREATSTART_ACT,1,10)='" + StrDate + "'   and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrLF_PLANTNO + "' ORDER BY hp.HEATID DESC";
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclLF);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                string StrHEATID = "";
                string StrTREATID = "";
                object StrActStart = DBNull.Value;
                object StrActEnd = DBNull.Value;
                object StrPlndStart = DBNull.Value;
                object StrPlndEnd = DBNull.Value;
                object StrLadleNo = DBNull.Value;
                object StrActStartDte = DBNull.Value;
                object StrGrade = DBNull.Value;
                object StrActStrt = DBNull.Value;
                object StrActEnnd = DBNull.Value;
                object Plant = DBNull.Value;
                object StrLastTemp = DBNull.Value;
                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_Conarc.Rows[i][0].ToString();
                    StrTREATID = dt_Orcl_Conarc.Rows[i][2].ToString();

                    if (dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != null && dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        Plant = dt_Orcl_Conarc.Rows[i]["PLANT"].ToString();
                    }
                    else
                    { Plant = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][6].ToString() != null && dt_Orcl_Conarc.Rows[i][6].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActStart = dt_Orcl_Conarc.Rows[i][6].ToString();
                    }
                    else
                    { StrActStart = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i][7].ToString() != null && dt_Orcl_Conarc.Rows[i][7].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActEnd = dt_Orcl_Conarc.Rows[i][7].ToString();
                    }
                    else
                    { StrActEnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][4].ToString() != null && dt_Orcl_Conarc.Rows[i][4].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrPlndStart = dt_Orcl_Conarc.Rows[i][4].ToString();
                    }
                    else
                    { StrPlndStart = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][5].ToString() != null && dt_Orcl_Conarc.Rows[i][5].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrPlndEnd = dt_Orcl_Conarc.Rows[i][5].ToString();
                    }
                    else
                    { StrPlndEnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][8].ToString() != null && dt_Orcl_Conarc.Rows[i][8].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrLadleNo = dt_Orcl_Conarc.Rows[i][8].ToString();
                    }
                    else
                    { StrLadleNo = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActStartDte = dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString();
                    }
                    else
                    { StrActStartDte = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["Grade"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Grade"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrGrade = dt_Orcl_Conarc.Rows[i]["Grade"].ToString();
                    }
                    else
                    { StrGrade = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["actStrt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["actStrt"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActStrt = dt_Orcl_Conarc.Rows[i]["actStrt"].ToString();
                    }
                    else
                    { StrActStrt = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["actEnnd"].ToString() != null && dt_Orcl_Conarc.Rows[i]["actEnnd"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActEnnd = dt_Orcl_Conarc.Rows[i]["actEnnd"].ToString();
                    }
                    else
                    { StrActEnnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["FINALTEMP"].ToString() != null && dt_Orcl_Conarc.Rows[i]["FINALTEMP"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrLastTemp = dt_Orcl_Conarc.Rows[i]["FINALTEMP"].ToString();
                    }
                    else
                    { StrLastTemp = DBNull.Value; }
                    if (Plant.ToString() == "LF3CAR1")
                    {
                        StrMIS_PLANTNO = "11";
                    }
                    else
                    {
                        StrMIS_PLANTNO = "12";
                    }
                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select heatNo from smsHeatTracker where heatNo='" + StrHEATID + "' and plantNo='" + StrMIS_PLANTNO + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120)";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',plndStart='" + StrPlndStart + "',plndEnd='" + StrPlndEnd + "',LadleNo='" + StrLadleNo + "',actEnds='" + StrActEnd + "',Grade='" + StrGrade + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "',LastTemp='" + StrLastTemp + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120)";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,plndStart,plndEnd,DateStamp,LadleNo,actEnds,Grade,actStrt,actEnnd,LastTemp)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + StrActStart + "','" + StrActEnd + "','','" + StrPlndStart + "','" + StrPlndEnd + "','" + StrActStartDte + "','" + StrLadleNo + "','" + StrActEnd + "','" + StrGrade + "','" + StrActStrt + "','" + StrActEnnd + "','" + StrLastTemp + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void FnSMS_SMCDB_HeatTracking(string StrPLANT, string StrBOF_PLANTNO, string StrMIS_PLANTNO, string StrDate)
        {
            //System.DateTime.Now.ToString("yyyy-MM-dd")
            //Old Query @01_08_2014 fr Revert Back
           // string StrQueryOrclLF = "SELECT  hp.HEATID_CUST_PLAN AS HEAT_ID,hp.PLANT ||hp.PLANTNO AS PLANT,hp.HEATID AS TREAT_ID,hd.LADLETYPE || hd.LADLENO AS LADLE_ID,decode(replace(SUBSTR(hp.TREATSTART_PLAN,11,6) ,':','.'),0,null,replace(SUBSTR(hp.TREATSTART_PLAN,11,6) ,':','.')) AS PLAN_START_TIME, decode(replace(SUBSTR(hp.TREATEND_PLAN,11,6) ,':','.'),0,null,replace(SUBSTR(hp.TREATEND_PLAN,11,6) ,':','.'))   AS PLAN_END_TIME,  replace(SUBSTR(TREATSTART_ACT,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.'),0,null,  replace(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hd.TREATSTART_ACT,1,10) as Act_Start_Dt,hp.HEATID,hd.STEELGRADECODE_ACT as Grade,replace(TREATSTART_ACT ,',','.') AS actStrt,replace(TREATEND_ACT ,',','.') AS actEnnd FROM PP_HEAT_PLANT hp inner join PD_HEAT_DATA hd on hp.HEATID=hd.HEATID where hd.PLANT='" + StrPLANT + "'  and SUBSTR(hd.TREATSTART_ACT,1,10)='" + StrDate + "'   and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrBOF_PLANTNO + "' ORDER BY hp.HEATID DESC";
            string StrQueryOrclLF = "SELECT * FROM (SELECT  hp.HEATID_CUST_PLAN AS HEAT_ID,hp.PLANT ||hp.PLANTNO AS PLANT,hp.HEATID AS TREAT_ID,hd.LADLETYPE || hd.LADLENO AS LADLE_ID,decode(replace(SUBSTR(hp.TREATSTART_PLAN,11,6) ,':','.'),0,null,replace(SUBSTR(hp.TREATSTART_PLAN,11,6) ,':','.')) AS PLAN_START_TIME, decode(replace(SUBSTR(hp.TREATEND_PLAN,11,6) ,':','.'),0,null,replace(SUBSTR(hp.TREATEND_PLAN,11,6) ,':','.'))   AS PLAN_END_TIME,  replace(SUBSTR(TREATSTART_ACT,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.'),0,null,  replace(SUBSTR(hd.TREATEND_ACT,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hd.TREATSTART_ACT,1,10) as Act_Start_Dt,hp.HEATID,hd.STEELGRADECODE_ACT as Grade,replace(TREATSTART_ACT ,',','.') AS actStrt,replace(TREATEND_ACT ,',','.') AS actEnnd FROM PP_HEAT_PLANT hp inner join PD_HEAT_DATA hd on hp.HEATID=hd.HEATID where hd.PLANT='" + StrPLANT + "'  and SUBSTR(hd.TREATSTART_ACT,1,10) <= '" + StrDate + "'   and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrBOF_PLANTNO + "' ORDER BY hp.HEATID DESC) where rownum <= 5";
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclLF);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                string StrHEATID = "";
                string StrTREATID = "";
                object StrActStart = DBNull.Value;
                object StrActEnd = DBNull.Value;
                object StrPlndStart = DBNull.Value;
                object StrPlndEnd = DBNull.Value;
                object StrLadleNo = DBNull.Value;
                object StrActStartDte = DBNull.Value;
                object StrGrade = DBNull.Value;
                object StrActStrt = DBNull.Value;
                object StrActEnnd = DBNull.Value;
                object Plant = DBNull.Value;
                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_Conarc.Rows[i][0].ToString();
                    StrTREATID = dt_Orcl_Conarc.Rows[i][2].ToString();

                    if (dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != null && dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        Plant = dt_Orcl_Conarc.Rows[i]["PLANT"].ToString();
                    }
                    else
                    { Plant = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][6].ToString() != null && dt_Orcl_Conarc.Rows[i][6].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActStart = dt_Orcl_Conarc.Rows[i][6].ToString();
                    }
                    else
                    { StrActStart = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i][7].ToString() != null && dt_Orcl_Conarc.Rows[i][7].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActEnd = dt_Orcl_Conarc.Rows[i][7].ToString();
                    }
                    else
                    { StrActEnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][4].ToString() != null && dt_Orcl_Conarc.Rows[i][4].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrPlndStart = dt_Orcl_Conarc.Rows[i][4].ToString();
                    }
                    else
                    { StrPlndStart = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][5].ToString() != null && dt_Orcl_Conarc.Rows[i][5].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrPlndEnd = dt_Orcl_Conarc.Rows[i][5].ToString();
                    }
                    else
                    { StrPlndEnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i][8].ToString() != null && dt_Orcl_Conarc.Rows[i][8].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrLadleNo = dt_Orcl_Conarc.Rows[i][8].ToString();
                    }
                    else
                    { StrLadleNo = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActStartDte = dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString();
                    }
                    else
                    { StrActStartDte = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["Grade"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Grade"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrGrade = dt_Orcl_Conarc.Rows[i]["Grade"].ToString();
                    }
                    else
                    { StrGrade = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["actStrt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["actStrt"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActStrt = dt_Orcl_Conarc.Rows[i]["actStrt"].ToString();
                    }
                    else
                    { StrActStrt = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["actEnnd"].ToString() != null && dt_Orcl_Conarc.Rows[i]["actEnnd"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActEnnd = dt_Orcl_Conarc.Rows[i]["actEnnd"].ToString();
                    }
                    else
                    { StrActEnnd = DBNull.Value; }

                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select heatNo from smsHeatTracker where heatNo='" + StrHEATID + "' and plantNo='" + StrMIS_PLANTNO + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120)";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',plndStart='" + StrPlndStart + "',plndEnd='" + StrPlndEnd + "',LadleNo='" + StrLadleNo + "',actEnds='" + StrActEnd + "',Grade='" + StrGrade + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120)";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,plndStart,plndEnd,DateStamp,LadleNo,actEnds,Grade,actStrt,actEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + StrActStart + "','" + StrActEnd + "','','" + StrPlndStart + "','" + StrPlndEnd + "','" + StrActStartDte + "','" + StrLadleNo + "','" + StrActEnd + "','" + StrGrade + "','" + StrActStrt + "','" + StrActEnnd + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void FnOP_BOF_STATUS(string StrBOF_PLANTNO)
        {
            //string StrQueryOrclBOF = "select  HEATID_CUST,replace(SUBSTR(TREATSTART_ACT,11,6) ,':','.') as TREATSTART_ACT,replace(SUBSTR(TREATEND_ACT,11,6) ,':','.') as TREATEND_ACT,replace(SUBSTR(TREATSTART_PLAN,11,6) ,':','.') as TREATSTART_PLAN,replace(SUBSTR(TREATEND_PLAN,11,6) ,':','.') as  TREATEND_PLAN from V_HEAT_DATA where HEATID_CUST=(select HEATID_CUST  from V_HEAT_DATA where PLANT='BOF' and TREATSTART_ACT=(select max(TREATSTART_ACT) from V_HEAT_DATA))and PLANT='BOF'";
            string StrQueryOrclBOF = "select VHD.PLANT || VHD.PLANTNO as UNIT,VHD.HEATID_CUST as HEATID,VHD.STEELGRADECODE_ACT as Grade ,VHD.STATUSNO,VHD.TREATSTART_ACT as TREATSTART,VHD.TREATEND_ACT as TREATEND,VHD.LADLENO,VHRTB.BLOWSTART,VHRTB.BLOWEND,VHRTB.BLOWINGDUR,VHRMB.OXY_TOTAL as OXYTOTAL,VHRTB.TAPSTART,VHRTB.TAPEND from V_HEAT_DATA VHD inner join V_HEAT_REPORT_TIMES_BOF VHRTB on VHD.HEATID=VHRTB.HEAT_ID inner join V_HEAT_REPORT_MAIN_BOF VHRMB  on VHD.HEATID=VHRMB.HEAT_ID and VHRTB.HEAT_ID=VHRMB.HEAT_ID where VHD.TREATSTART_ACT in (select TREATSTART_ACT from V_HEAT_DATA where SUBSTR(TREATSTART_ACT,1,10)='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "' and PLANT='BOF' and PLANTNO='" + StrBOF_PLANTNO + "') order by TREATSTART_ACT asc";
            DataTable dt_Orcl_BOF = new DataTable();
            dt_Orcl_BOF = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclBOF);
            if (dt_Orcl_BOF.Rows.Count > 0)
            {
                string StrHEATID = "";
                object StrTreatStart = DBNull.Value;
                object StrTreatEnd = DBNull.Value;
                object StrBlowStart = DBNull.Value;
                object StrBlowEnd = DBNull.Value;
                object StrTapStart = DBNull.Value;
                object StrTapEnd = DBNull.Value;
                for (int i = 0; i < dt_Orcl_BOF.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_BOF.Rows[i]["HEATID"].ToString();

                    if (dt_Orcl_BOF.Rows[i]["TREATSTART"].ToString() != null && dt_Orcl_BOF.Rows[i]["TREATSTART"].ToString() != string.Empty)
                    {
                        StrTreatStart = dt_Orcl_BOF.Rows[i]["TREATSTART"].ToString().Replace(",", ".");
                    }
                    else
                    { StrTreatStart = DBNull.Value; }

                    if (dt_Orcl_BOF.Rows[i]["TREATEND"].ToString() != null && dt_Orcl_BOF.Rows[i]["TREATEND"].ToString() != string.Empty)
                    {
                        StrTreatEnd = dt_Orcl_BOF.Rows[i]["TREATEND"].ToString().Replace(",", ".");
                    }
                    else
                    { StrTreatEnd = DBNull.Value; }
                    if (dt_Orcl_BOF.Rows[i]["BLOWSTART"].ToString() != null && dt_Orcl_BOF.Rows[i]["BLOWSTART"].ToString() != string.Empty)
                    {
                        StrBlowStart = dt_Orcl_BOF.Rows[i]["BLOWSTART"].ToString().Replace(",", ".");
                    }
                    else
                    { StrBlowStart = DBNull.Value; }

                    if (dt_Orcl_BOF.Rows[i]["BLOWEND"].ToString() != null && dt_Orcl_BOF.Rows[i]["BLOWEND"].ToString() != string.Empty)
                    {
                        StrBlowEnd = dt_Orcl_BOF.Rows[i]["BLOWEND"].ToString().Replace(",", ".");
                    }
                    else
                    { StrBlowEnd = DBNull.Value; }
                    if (dt_Orcl_BOF.Rows[i]["TAPSTART"].ToString() != null && dt_Orcl_BOF.Rows[i]["TAPSTART"].ToString() != string.Empty)
                    {
                        StrTapStart = dt_Orcl_BOF.Rows[i]["TAPSTART"].ToString().Replace(",", ".");
                    }
                    else
                    { StrTapStart = DBNull.Value; }
                    if (dt_Orcl_BOF.Rows[i]["TAPEND"].ToString() != null && dt_Orcl_BOF.Rows[i]["TAPEND"].ToString() != string.Empty)
                    {
                        StrTapEnd = dt_Orcl_BOF.Rows[i]["TAPEND"].ToString().Replace(",", ".");
                    }
                    else
                    { StrTapEnd = DBNull.Value; }
                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select HEAT_ID from OP_BOF_STATUS where HEAT_ID='" + StrHEATID + "'";// and DateStamp=Convert(varchar(10),'" + System.DateTime.Now.ToString() + "',120)";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update OP_BOF_STATUS set STEL_GRADE_CODE='" + dt_Orcl_BOF.Rows[i]["Grade"].ToString() + "',STATUS_DESC='" + dt_Orcl_BOF.Rows[i]["STATUSNO"].ToString() + "',HEAT_START_TIME='" + StrTreatStart + "',HEAT_END_TIME='" + StrTreatEnd + "',BLOW_START_TIME='" + StrBlowStart + "',BLOW_END_TIME='" + StrBlowEnd + "',BLOWDUR='" + dt_Orcl_BOF.Rows[i]["BLOWINGDUR"].ToString() + "',LADLENO='" + dt_Orcl_BOF.Rows[i]["LADLENO"].ToString() + "',OXY_CONSM='" + dt_Orcl_BOF.Rows[i]["OXYTOTAL"].ToString() + "',TAP_START_TIME='" + StrTapStart + "',TAP_END_TIME='" + StrTapEnd + "' where UNIT='" + dt_Orcl_BOF.Rows[i]["UNIT"].ToString() + "' and HEAT_ID='" + dt_Orcl_BOF.Rows[i]["HEATID"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert OP_BOF_STATUS(UNIT,HEAT_ID,STEL_GRADE_CODE,STATUS_DESC,HEAT_START_TIME,HEAT_END_TIME,BLOW_START_TIME,BLOW_END_TIME,BLOWDUR,LADLENO,OXY_CONSM,TAP_START_TIME,TAP_END_TIME)values('" + dt_Orcl_BOF.Rows[i]["UNIT"].ToString() + "','" + dt_Orcl_BOF.Rows[i]["HEATID"].ToString() + "','" + dt_Orcl_BOF.Rows[i]["Grade"].ToString() + "','" + dt_Orcl_BOF.Rows[i]["STATUSNO"].ToString() + "','" + StrTreatStart + "','" + StrTreatEnd + "','" + StrBlowStart + "','" + StrBlowEnd + "','" + dt_Orcl_BOF.Rows[i]["BLOWINGDUR"].ToString() + "','" + dt_Orcl_BOF.Rows[i]["LADLENO"].ToString() + "','" + dt_Orcl_BOF.Rows[i]["OXYTOTAL"].ToString() + "','" + StrTapStart + "','" + StrTapEnd + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }

        private void FnSMS_SMCDB_BOF(string StrPLANT, string StrMIS_PLANTNO, string StrBOF_PLANTNO)
        {

            string StrQueryOrclBOF = "select  HEATID_CUST,replace(SUBSTR(TREATSTART_ACT,11,6) ,':','.') as TREATSTART_ACT,replace(SUBSTR(TREATEND_ACT,11,6) ,':','.') as TREATEND_ACT,replace(SUBSTR(TREATSTART_PLAN,11,6) ,':','.') as TREATSTART_PLAN,replace(SUBSTR(TREATEND_PLAN,11,6) ,':','.') as  TREATEND_PLAN,SUBSTR(TREATSTART_ACT,1,10) as DATESTAMP,STEELGRADECODE_ACT as Grade,replace(TREATSTART_ACT ,',','.') AS actStrt,replace(TREATEND_ACT ,',','.') AS actEnnd from V_HEAT_DATA where HEATID_CUST in (select HEATID_CUST  from V_HEAT_DATA where PLANT='BOF' and TREATSTART_ACT in  (select TREATSTART_ACT from V_HEAT_DATA where SUBSTR(TREATSTART_ACT,1,10)='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "'))and PLANT='BOF' and PLANTNO='" + StrBOF_PLANTNO + "' order by TREATSTART_ACT asc";
            //string StrQueryOrclBOF = "select  HEATID_CUST,replace(SUBSTR(TREATSTART_ACT,11,6) ,':','.') as TREATSTART_ACT,replace(SUBSTR(TREATEND_ACT,11,6) ,':','.') as TREATEND_ACT,replace(SUBSTR(TREATSTART_PLAN,11,6) ,':','.') as TREATSTART_PLAN,replace(SUBSTR(TREATEND_PLAN,11,6) ,':','.') as  TREATEND_PLAN,SUBSTR(TREATSTART_ACT,1,10) as DATESTAMP,STEELGRADECODE_ACT as Grade,replace(TREATSTART_ACT ,',','.') AS actStrt,replace(TREATEND_ACT ,',','.') AS actEnnd from V_HEAT_DATA where HEATID_CUST in (select HEATID_CUST  from V_HEAT_DATA where PLANT='BOF' and TREATSTART_ACT in  (select TREATSTART_ACT from V_HEAT_DATA where SUBSTR(TREATSTART_ACT,1,10) between '2014-06-01' and '2014-07-07'))and PLANT='BOF' and PLANTNO='" + StrBOF_PLANTNO + "' order by TREATSTART_ACT asc";
            DataTable dt_Orcl_BOF = new DataTable();
            dt_Orcl_BOF = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclBOF);
            if (dt_Orcl_BOF.Rows.Count > 0)
            {
                string StrHEATID = "";
                object StrActStart = DBNull.Value;
                object StrActEnd = DBNull.Value;
                object StrPlndStart = DBNull.Value;
                object StrPlndEnd = DBNull.Value;
                object StrDateStamp = DBNull.Value;
                object StrGrade = DBNull.Value;
                object StrActStrt = DBNull.Value;
                object StrActEnnd = DBNull.Value;
                for (int i = 0; i < dt_Orcl_BOF.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_BOF.Rows[i][0].ToString();

                    if (dt_Orcl_BOF.Rows[i][1].ToString() != null && dt_Orcl_BOF.Rows[i][1].ToString() != string.Empty)
                    {
                        StrActStart = dt_Orcl_BOF.Rows[i][1].ToString();
                    }
                    else
                    { StrActStart = DBNull.Value; }

                    if (dt_Orcl_BOF.Rows[i][2].ToString() != null && dt_Orcl_BOF.Rows[i][2].ToString() != string.Empty)
                    {
                        StrActEnd = dt_Orcl_BOF.Rows[i][2].ToString();
                    }
                    else
                    { StrActEnd = DBNull.Value; }
                    if (dt_Orcl_BOF.Rows[i][3].ToString() != null && dt_Orcl_BOF.Rows[i][3].ToString() != string.Empty)
                    {
                        StrPlndStart = dt_Orcl_BOF.Rows[i][3].ToString();
                    }
                    else
                    { StrPlndStart = DBNull.Value; }
                    if (dt_Orcl_BOF.Rows[i][4].ToString() != null && dt_Orcl_BOF.Rows[i][4].ToString() != string.Empty)
                    {
                        StrPlndEnd = dt_Orcl_BOF.Rows[i][4].ToString();
                    }
                    else
                    { StrPlndEnd = DBNull.Value; }

                    if (dt_Orcl_BOF.Rows[i][5].ToString() != null && dt_Orcl_BOF.Rows[i][5].ToString() != string.Empty)
                    {
                        StrDateStamp = dt_Orcl_BOF.Rows[i][5].ToString();
                    }
                    else
                    { StrDateStamp = DBNull.Value; }
                    if (dt_Orcl_BOF.Rows[i]["Grade"].ToString() != null && dt_Orcl_BOF.Rows[i]["Grade"].ToString() != string.Empty)
                    {
                        StrGrade = dt_Orcl_BOF.Rows[i]["Grade"].ToString();
                    }
                    else
                    { StrGrade = DBNull.Value; }
                    if (dt_Orcl_BOF.Rows[i]["actStrt"].ToString() != null && dt_Orcl_BOF.Rows[i]["actStrt"].ToString() != string.Empty)
                    {
                        StrActStrt = dt_Orcl_BOF.Rows[i]["actStrt"].ToString();
                    }
                    else
                    { StrActStrt = DBNull.Value; }

                    if (dt_Orcl_BOF.Rows[i]["actEnnd"].ToString() != null && dt_Orcl_BOF.Rows[i]["actEnnd"].ToString() != string.Empty)
                    {
                        StrActEnnd = dt_Orcl_BOF.Rows[i]["actEnnd"].ToString();
                    }
                    else
                    { StrActEnnd = DBNull.Value; }
                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select heatNo from smsHeatTracker where heatNo='" + StrHEATID + "' and plantNo='" + StrMIS_PLANTNO + "'";// and DateStamp=Convert(varchar(10),'" + System.DateTime.Now.ToString() + "',120)";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count > 0)
                    {
                        if (StrActStart != DBNull.Value)
                        {
                            string StrUpdateQuery = "Update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',plndStart='" + StrPlndStart + "',plndEnd='" + StrPlndEnd + "',Grade='" + StrGrade + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "'";// and DateStamp=Convert(varchar(10),'" + System.DateTime.Now.ToString() + "',120)";
                            bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                        }
                    }
                    else
                    {
                        if (StrActStart != DBNull.Value)
                        {
                            string StrInsertQuery = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,plndStart,plndEnd,DateStamp,Grade,actStrt,actEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + StrActStart + "','" + StrActEnd + "','','" + StrPlndStart + "','" + StrPlndEnd + "','" + StrDateStamp + "','" + StrGrade + "','" + StrActStrt + "','" + StrActEnnd + "')";
                            bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                        }
                    }
                }
            }
        }

        private void FnSMS_SMCDB_UpdateBOF(string StrPLANT, string StrMIS_PLANTNO)
        {

            string StrUptQuery = "UPDATE t   SET t.actEnnd = c.actStrrt,t.actEnd=c.actEnd,t.actEnds=c.actEnd FROM smsHeatTracker t INNER JOIN (SELECT a.plantNo,a.heatNo,a.DateStamp,a.actStrt as actStart,b.actStrt,dateadd(minute,-1,b.actStrt) as actStrrt ,replace(convert(char(5), dateadd(minute,-1,b.actStrt), 108),':','.') as actEnd FROM ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt) , * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) b right JOIN ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt)+1, * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) a ON b.Id = a.Id) c ON  c.heatNo=t.heatNo and c.plantNo=t.plantNo and c.DateStamp=t.DateStamp and c.actStart=t.actStrt";
            bool StrUpdtStatus = clsObj.DBInsertUpdateDeleteMIS(StrUptQuery);
            return;


            DataTable dt_MIS_TimeUpdt = new DataTable();
            //string StrQueryEndTmeUpdate = "select heatNo from smsHeatTracker where convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + System.DateTime.Now.ToString("yyyy-MM-dd") + "',120) and (actEnd=0 or actEnd is null)  and plantNo='" + StrMIS_PLANTNO + "'";
            string StrQueryEndTmeUpdate = "select heatNo from smsHeatTracker where (actEnd=0 or actEnd is null)  and plantNo='" + StrMIS_PLANTNO + "'";
            dt_MIS_TimeUpdt = clsObj.DBSelectQueryMIS_Table(StrQueryEndTmeUpdate);
            if (dt_MIS_TimeUpdt.Rows.Count <= 0)
            {
                return;
            }
            for (int j = 0; j < dt_MIS_TimeUpdt.Rows.Count; j++)
            {
                string StrQueryOrclCONARC = "select  HEATID_CUST,replace(SUBSTR(TREATSTART_ACT,11,6) ,':','.') as TREATSTART_ACT,replace(SUBSTR(TREATEND_ACT,11,6) ,':','.') as TREATEND_ACT,replace(SUBSTR(TREATSTART_PLAN,11,6) ,':','.') as TREATSTART_PLAN,replace(SUBSTR(TREATEND_PLAN,11,6) ,':','.') as  TREATEND_PLAN,SUBSTR(TREATSTART_ACT,1,10) as DATESTAMP,STEELGRADECODE_ACT as Grade,replace(TREATSTART_ACT ,',','.') AS actStrt,replace(TREATEND_ACT ,',','.') AS actEnnd from V_HEAT_DATA where HEATID_CUST='" + dt_MIS_TimeUpdt.Rows[j][0].ToString() + "' and PLANT='BOF'";
                DataTable dt_Orcl_BOF = new DataTable();
                dt_Orcl_BOF = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclCONARC);
                if (dt_Orcl_BOF.Rows.Count > 0)
                {
                    string StrHEATID = "";
                    object StrActStart = DBNull.Value;
                    object StrActEnd = DBNull.Value;
                    object StrPlndStart = DBNull.Value;
                    object StrPlndEnd = DBNull.Value;
                    object StrGrade = DBNull.Value;
                    object StrActStrt = DBNull.Value;
                    object StrActEnnd = DBNull.Value;

                    for (int i = 0; i < dt_Orcl_BOF.Rows.Count; i++)
                    {


                        StrHEATID = dt_Orcl_BOF.Rows[i][0].ToString();

                        if (dt_Orcl_BOF.Rows[i][1].ToString() != null && dt_Orcl_BOF.Rows[i][1].ToString() != string.Empty)
                        {
                            StrActStart = dt_Orcl_BOF.Rows[i][1].ToString();
                        }
                        else
                        { StrActStart = DBNull.Value; }

                        if (dt_Orcl_BOF.Rows[i][2].ToString() != null && dt_Orcl_BOF.Rows[i][2].ToString() != string.Empty)
                        {
                            StrActEnd = dt_Orcl_BOF.Rows[i][2].ToString();
                        }
                        else
                        { StrActEnd = DBNull.Value; }
                        if (dt_Orcl_BOF.Rows[i][3].ToString() != null && dt_Orcl_BOF.Rows[i][3].ToString() != string.Empty)
                        {
                            StrPlndStart = dt_Orcl_BOF.Rows[i][3].ToString();
                        }
                        else
                        { StrPlndStart = DBNull.Value; }
                        if (dt_Orcl_BOF.Rows[i][4].ToString() != null && dt_Orcl_BOF.Rows[i][4].ToString() != string.Empty)
                        {
                            StrPlndEnd = dt_Orcl_BOF.Rows[i][4].ToString();
                        }
                        else
                        { StrPlndEnd = DBNull.Value; }
                        if (dt_Orcl_BOF.Rows[i]["Grade"].ToString() != null && dt_Orcl_BOF.Rows[i]["Grade"].ToString() != string.Empty)
                        {
                            StrGrade = dt_Orcl_BOF.Rows[i]["Grade"].ToString();
                        }
                        else
                        { StrGrade = DBNull.Value; }
                        if (dt_Orcl_BOF.Rows[i]["actStrt"].ToString() != null && dt_Orcl_BOF.Rows[i]["actStrt"].ToString() != string.Empty)
                        {
                            StrActStrt = dt_Orcl_BOF.Rows[i]["actStrt"].ToString();
                        }
                        else
                        { StrActStrt = DBNull.Value; }

                        if (dt_Orcl_BOF.Rows[i]["actEnnd"].ToString() != null && dt_Orcl_BOF.Rows[i]["actEnnd"].ToString() != string.Empty)
                        {
                            StrActEnnd = dt_Orcl_BOF.Rows[i]["actEnnd"].ToString();
                        }
                        else
                        { StrActEnnd = DBNull.Value; }
                        if (StrActStart != DBNull.Value)
                        {
                            string StrUpdateQuery = "Update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',plndStart='" + StrPlndStart + "',plndEnd='" + StrPlndEnd + "',Grade='" + StrGrade + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "'  where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "'";// and DateStamp=Convert(varchar(10),'" + System.DateTime.Now.ToString() + "',120)";
                            bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                        }
                    }
                }
            }
        }

        private void timer_GetDataLab_Tick(object sender, EventArgs e)
        {
            try
            {
                GetDataContainerLab();
            }
            catch
            {
            }
        }
        private void FnBOF_ScrpCnsmp()
        {
            DataTable dt_ScrpBOFCnsmp = new DataTable();
            DataTable dt_scrapMIs = new DataTable();
            string StrHeatId = "", StrMatCode = "", selectHeatMis = "";

            string StrQueryOrclBOF = "Select * from (Select v_mat.heat_id,v_mat.mat_code ,v_mat.mat_amount_act,mat.description ,v_mat.mat_source,v_rpt_bof.heat_id_cust,v_rpt_bof.heat_time from V_HEAT_REPORT_BOF_MAT_ADDI v_mat INNER JOIN v_daily_report_bof v_rpt_bof ON v_mat.heat_id= v_rpt_bof.heat_id INNER JOIN gt_mat mat ON v_mat.mat_code = mat.mat_code where v_mat.mat_source='Chute'and v_mat.recipename='Scrap' order by v_rpt_bof.heat_time desc ) where rownum<=5";
            //[0]           [1]                [2]               [3]                [4]                   [5]                   [6]                                                                                                                                                                                                                                                                                                     
            dt_ScrpBOFCnsmp = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclBOF);
            if (dt_ScrpBOFCnsmp.Rows.Count > 0)
            {
                for (int i = 0; i < dt_ScrpBOFCnsmp.Rows.Count; i++)
                {
                    StrHeatId = dt_ScrpBOFCnsmp.Rows[i][5].ToString();
                    StrMatCode = dt_ScrpBOFCnsmp.Rows[i][1].ToString();
                    selectHeatMis = "select * from Scrp_Mtrl_Consum where StrHeatId = '" + StrHeatId + "' and StrMatCode = '" + StrMatCode + "'";
                    dt_scrapMIs = clsObj.DBSelectQueryMIS_Table(selectHeatMis);
                    if (dt_ScrpBOFCnsmp.Rows[i][6].ToString() != "")
                    {
                        string Str_Date = dt_ScrpBOFCnsmp.Rows[i][6].ToString();

                        string[] DateTime_Only = Str_Date.Split(new char[] { ',' });
                        if (dt_scrapMIs.Rows.Count > 0)
                        {
                            string StrUpdateQuery = "UPDATE Scrp_Mtrl_Consum SET S_Date='" + DateTime_Only[0].ToString() + "',Mtrl_Descrp='" + dt_ScrpBOFCnsmp.Rows[i][3].ToString() + "',Qty='" + dt_ScrpBOFCnsmp.Rows[i][2].ToString() + "' where HeatId='" + StrHeatId + "' and Mat_Code='" + StrMatCode + "'";
                            bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);

                        }
                        else
                        {
                            string StrInsertQuery = "INSERT INTO Scrp_Mtrl_Consum (S_Date,HeatId,Mat_Code,Mtrl_Descrp,Qty) VALUES ('" + DateTime_Only[0].ToString() + "','" + StrHeatId + "','" + StrMatCode + "','" + dt_ScrpBOFCnsmp.Rows[i][3].ToString() + "','" + dt_ScrpBOFCnsmp.Rows[i][2].ToString() + "')";
                            bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                        }
                    }
                }
            }
        }





    }
}
