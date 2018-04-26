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
using System.Data.OleDb;
using MySql.Data.MySqlClient;



namespace InterfaceHMI
{
    public partial class Form1 : Form
    {
        //string StrConnectionSQL = System.Configuration.ConfigurationSettings.AppSettings["ConnectionStringMIS"];
        //string StrConnectionORACLE = System.Configuration.ConfigurationSettings.AppSettings["ConnectionStringSMSII"];
        //string StrConnectionstringCokeOven = System.Configuration.ConfigurationSettings.AppSettings["ConnectionStringCokeOven"];

        public static MySqlConnection DbconMySql = new MySqlConnection("Server=10.10.21.42;Database=lwtcokedb;uid=bslit;pwd=itbsl;");
        public static SqlConnection DbconSql = new SqlConnection("server=10.10.58.171\\SQLEXPRESS;database=MIS;uid=sa;password=bhushan123#;");
        //public static SqlConnection DbconSql = new SqlConnection("server=10.1.22.36;database=MIS;uid=sa;password=bhushan123#;");

        public static OracleConnection OrcConOrc = new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.15.20.85)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=SMCDB)));User Id=bhushan;Password=bhushan;");
        //public static OracleConnection OrcConOrcCasterI = new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.15.10.22)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=CCSDB_NEW)));User Id=L2CCS;Password=L2CCS;");
        public static OracleConnection OrcConOrcCasterI = new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.15.10.22)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=CCSDB1)));User Id=L2CCS;Password=L2CCS;");
        //(DESCRIPTION =    (ADDRESS_LIST =      (ADDRESS = (PROTOCOL = TCP)(HOST = 10.15.10.22)(PORT = 1521))    )    (CONNECT_DATA =      (SERVICE_NAME = CCSDB_NEW)    )  )
        public static OracleConnection OrcConOrcCasterII = new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.16.10.22)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=CCSDB)));User Id=L2CCS;Password=L2CCS;");
        //(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST = 10.16.10.22)(PORT = 1521) ))(CONNECT_DATA =(SERVICE_NAME = CCSDB)))
        public static OracleConnection OrcConOrcCasterIII = new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=10.17.10.22)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=CCSDB)));User Id=L2CCS;Password=L2CCS;");

        //public static string con_stringSTATUS = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST = SMC-DEV)(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = SMCDB)));User Id=BHUSHAN;Password=BHUSHAN;";
        public static string con_stringSTATUS = "Data Source=(DESCRIPTION =(ADDRESS_LIST =(ADDRESS =(PROTOCOL = TCP)(HOST = 10.15.20.85)(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = SMCDB)));User Id=BHUSHAN;Password=BHUSHAN;";

        string Query1 = "";
        string Query2 = "";
        string Query3 = "";
        DataSet ds1 = new DataSet();
        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();
        DBConnections clsObj = new DBConnections();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                FnCheckConnectionSMSII();
                FnCheckConnectionCASTERI();
                FnCheckConnectionCASTERII();
            }
            catch (Exception ex)
            {
                txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
            }
        }
        public void FnCheckConnectionSMSII()
        {
            try
            {

                DbconSql.Open();
                OrcConOrc.Open();
                lblINTERFACEtoMISSERVER.ForeColor = System.Drawing.Color.Green;
                lblINTERFACEtoSMSDB.ForeColor = System.Drawing.Color.Green;
                lblStatus1.Text = "Connected";
                lblStatus2.Text = "Connected";
                lblStatus1.ForeColor = System.Drawing.Color.Green;
                lblStatus2.ForeColor = System.Drawing.Color.Green;
                label4.BackColor = System.Drawing.Color.Green;
                label5.BackColor = System.Drawing.Color.Green;
                label4.ForeColor = System.Drawing.Color.White;
                label5.ForeColor = System.Drawing.Color.White;
                DbconSql.Close();
                OrcConOrc.Close();
            }
            catch
            {

            }

        }
        public void FnCheckConnectionCASTERI()
        {
            //uncheck//
            try
            {
                OrcConOrcCasterI.Open();
                lblInterfacetoMISDBI.ForeColor = System.Drawing.Color.Green;
                lblINTERFACETOCASTERI.ForeColor = System.Drawing.Color.Green;
                lblStatus3.Text = "Connected";
                lblStatus4.Text = "Connected";
                lblStatus3.ForeColor = System.Drawing.Color.Green;
                lblStatus4.ForeColor = System.Drawing.Color.Green;
                lblMISSERVERI.BackColor = System.Drawing.Color.Green;
                lblCasterI.BackColor = System.Drawing.Color.Green;
                lblMISSERVERI.ForeColor = System.Drawing.Color.White;
                lblCasterI.ForeColor = System.Drawing.Color.White;
                OrcConOrcCasterI.Close();
            }
            catch (Exception ex)
            {
                txtError.Text = ex.ToString();
            }
            //uncheck//
        }
        public void FnCheckConnectionCASTERII()
        {
            //uncheck//
            try
            {
                OrcConOrcCasterII.Open();
                lblInterfacetoMISDBII.ForeColor = System.Drawing.Color.Green;
                lblINTERFACETOCASTERII.ForeColor = System.Drawing.Color.Green;
                lblStatus5.Text = "Connected";
                lblStatus6.Text = "Connected";
                lblStatus5.ForeColor = System.Drawing.Color.Green;
                lblStatus6.ForeColor = System.Drawing.Color.Green;
                lblMISSERVERII.BackColor = System.Drawing.Color.Green;
                lblCasterII.BackColor = System.Drawing.Color.Green;
                lblMISSERVERII.ForeColor = System.Drawing.Color.White;
                lblCasterII.ForeColor = System.Drawing.Color.White;
                OrcConOrcCasterII.Close();
            }
            catch (Exception ex)
            {
                txtError.Text = ex.ToString();
            }
            //uncheck//
        }
        private void timer1_Tick(object sender, EventArgs e)
        {

            FnCheckConnectionSMSII();
            FnCheckConnectionCASTERI();
            FnCheckConnectionCASTERII();
            string StrConarcA = "";
            string StrConarcB = "";
            Query1 = "select heatid_cust from pp_heat where heatid=(select  heatid  from pd_heatdata where plant='CON' and plantno=1  and treatstart_plan=(Select max(treatstart_plan) from pd_heatdata where plant='CON' and plantno=1  ))";
            Query2 = "select heatid_cust from pp_heat where heatid=(select  heatid  from pd_heatdata where plant='CON' and plantno=2 and treatstart_plan=(Select max(treatstart_plan) from pd_heatdata where plant='CON' and plantno=2  ))";

            clsObj = new DBConnections();

            dt1 = new DataTable();
            dt1 = clsObj.DBSelectQuerySMSII_Table(Query1);
            if (dt1.Rows.Count > 0)
            {
                StrConarcA = dt1.Rows[0][0].ToString();
                if (lblConAShell1Anncd.Text == "--")
                {
                    lblConAShell1Anncd.Text = StrConarcA;

                }
                else if (lblConAShell1Anncd.Text == StrConarcA)
                {


                }
                else
                {

                    lblConAShell1Finsd.Text = lblConAShell1Anncd.Text;
                    lblConAShell1Anncd.Text = StrConarcA;

                    OrcConOrc.Open();
                    DataTable dt1SMSII = new DataTable();
                    string StrQueryString = "SELECT distinct pp_heat_plant.heatid_cust,pp_heat_plant.plant, pd_report.announcetime,pd_heatdata.treatstart_act,pd_report.tappingstarttime,pd_report.tappingendtime, pd_report.taptotapduration, pd_report.l1_dri_total_cons, pd_report.l1_lime_total_cons,pd_report.l1_dolo_total_cons, pd_report.l1_coal_total_cons,pdb_heat_data.o2_blow_dur, pdb_heat_data.total_o2_cons,pd_report.total_elec_egy, pde_heat_data.power_on_dur   FROM pp_heat_plant,pd_report,pd_heatdata,       pdb_heat_data, pde_heat_data,       pdf_heat_phase_data, pd_anl  WHERE   (pp_heat_plant.plant = pd_report.plant)        AND (pd_heatdata.plant = pd_report.plant)        AND (pd_heatdata.heatid = pd_report.heatid)        AND (pd_heatdata.treatid = pd_report.treatid)        AND (pd_heatdata.heatid = pdb_heat_data.heatid)        AND (pdb_heat_data.heatid = pde_heat_data.heatid)        AND (pp_heat_plant.heatid = pd_report.heatid)        AND (pde_heat_data.heatid = pdf_heat_phase_data.heatid)        AND (pdf_heat_phase_data.heatid = pd_anl.heatid)    AND ( pp_heat_plant.heatid_cust='" + StrConarcA + "')    AND (pp_heat_plant.plant='CON')";
                    OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                    adp.Fill(dt1SMSII);
                    DateTime dtAnnounceTime = System.DateTime.Now;
                    DateTime dtStartTime = System.DateTime.Now;

                    OrcConOrc.Close();
                    if (dt1SMSII.Rows.Count > 0)
                    {
                        String StrAnnounceTime = dt1SMSII.Rows[0][2].ToString().Replace(',', '.');
                        if (StrAnnounceTime != "")
                        {
                            dtAnnounceTime = Convert.ToDateTime(StrAnnounceTime);
                        }
                    }
                    string Str_AH_Insert = "Insert into lgc_heattimetable (HeatNo,C_datetime,Area) values('" + lblConAShell1Finsd.Text + "','" + System.DateTime.Now.ToString() + "','CONARC-A')";
                    Str_AH_Insert = Str_AH_Insert + " ; Insert into SCRAP_YARD (HeatNo,EDate)values('" + StrConarcA + "','" + dtAnnounceTime.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                    Str_AH_Insert = Str_AH_Insert + " ; Insert into r_material_consmn (HeatNo,EDate)values('" + StrConarcA + "','" + dtAnnounceTime.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                    Str_AH_Insert = Str_AH_Insert + " ; Insert into mnl_HMwtHMD_ConarcProcessData(HeatNo)values('" + StrConarcA + "')";
                    bool Status = clsObj.DBInsertUpdateDeleteMIS(Str_AH_Insert);

                    //FnInsertSMSIIdb1(lblConAShell1Finsd.Text);
                    //FnInsertSMSIIdb2(lblConAShell1Finsd.Text);
                    //FnInsertSMSIIdb3(lblConAShell1Finsd.Text);
                    //FnInsertSMSIIdb4(lblConAShell1Finsd.Text);
                    //FnInsertSMSIIdb5(lblConAShell1Finsd.Text);
                    //FnInsertSMSIIdb6(lblConAShell1Finsd.Text);
                    //FnUpdateFormulaFields(lblConAShell1Finsd.Text);
                    //FnUpdateFormulaFieldsDivideByZERO(lblConAShell1Finsd.Text);


                }

            }


            dt2 = new DataTable();
            dt2 = clsObj.DBSelectQuerySMSII_Table(Query2);
            if (dt2.Rows.Count > 0)
            {
                StrConarcB = dt2.Rows[0][0].ToString();
                if (lblConBShell2Anncd.Text == "--")
                {
                    lblConBShell2Anncd.Text = StrConarcB;
                }
                else if (lblConBShell2Anncd.Text == StrConarcB)
                {

                }
                else
                {

                    lblConBShell2Finsd.Text = lblConBShell2Anncd.Text;
                    lblConBShell2Anncd.Text = StrConarcB;

                    OrcConOrc.Open();
                    DataTable dt1SMSII = new DataTable();
                    string StrQueryString = "SELECT distinct pp_heat_plant.heatid_cust,pp_heat_plant.plant, pd_report.announcetime,pd_heatdata.treatstart_act,pd_report.tappingstarttime,pd_report.tappingendtime, pd_report.taptotapduration, pd_report.l1_dri_total_cons, pd_report.l1_lime_total_cons,pd_report.l1_dolo_total_cons, pd_report.l1_coal_total_cons,pdb_heat_data.o2_blow_dur, pdb_heat_data.total_o2_cons,pd_report.total_elec_egy, pde_heat_data.power_on_dur   FROM pp_heat_plant,pd_report,pd_heatdata,       pdb_heat_data, pde_heat_data,       pdf_heat_phase_data, pd_anl  WHERE   (pp_heat_plant.plant = pd_report.plant)        AND (pd_heatdata.plant = pd_report.plant)        AND (pd_heatdata.heatid = pd_report.heatid)        AND (pd_heatdata.treatid = pd_report.treatid)        AND (pd_heatdata.heatid = pdb_heat_data.heatid)        AND (pdb_heat_data.heatid = pde_heat_data.heatid)        AND (pp_heat_plant.heatid = pd_report.heatid)        AND (pde_heat_data.heatid = pdf_heat_phase_data.heatid)        AND (pdf_heat_phase_data.heatid = pd_anl.heatid)    AND ( pp_heat_plant.heatid_cust='" + StrConarcB + "')    AND (pp_heat_plant.plant='CON')";
                    OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                    adp.Fill(dt1SMSII);
                    DateTime dtAnnounceTime = System.DateTime.Now;
                    DateTime dtStartTime = System.DateTime.Now;

                    OrcConOrc.Close();
                    if (dt1SMSII.Rows.Count > 0)
                    {
                        String StrAnnounceTime = dt1SMSII.Rows[0][2].ToString().Replace(',', '.');
                        if (StrAnnounceTime != "")
                        {
                            dtAnnounceTime = Convert.ToDateTime(StrAnnounceTime);
                        }
                    }
                    string Str_AH_Insert = "Insert into lgc_heattimetable (HeatNo,C_datetime,Area) values('" + lblConBShell2Finsd.Text + "','" + System.DateTime.Now.ToString() + "','CONARC-B')";
                    Str_AH_Insert = Str_AH_Insert + " ; Insert into SCRAP_YARD (HeatNo,EDate)values('" + StrConarcB + "','" + dtAnnounceTime.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                    Str_AH_Insert = Str_AH_Insert + " ; Insert into r_material_consmn (HeatNo,EDate)values('" + StrConarcB + "','" + dtAnnounceTime.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                    Str_AH_Insert = Str_AH_Insert + " ; Insert into mnl_HMwtHMD_ConarcProcessData(HeatNo)values('" + StrConarcB + "')";
                    bool Status = clsObj.DBInsertUpdateDeleteMIS(Str_AH_Insert);

                    //FnInsertSMSIIdb1(lblConBShell2Finsd.Text);
                    //FnInsertSMSIIdb2(lblConBShell2Finsd.Text);
                    //FnInsertSMSIIdb3(lblConBShell2Finsd.Text);
                    //FnInsertSMSIIdb4(lblConBShell2Finsd.Text);
                    //FnInsertSMSIIdb5(lblConBShell2Finsd.Text);
                    //FnInsertSMSIIdb6(lblConBShell2Finsd.Text);
                    //FnUpdateFormulaFields(lblConBShell2Finsd.Text);
                    //FnUpdateFormulaFieldsDivideByZERO(lblConBShell2Finsd.Text);
                }
            }

            Query3 = "Select CONARC_A,CONARC_B from ANNOUNCED_HEAT";
            dt1 = new DataTable();

            dt1 = clsObj.DBSelectQueryMIS_Table(Query3);
            if (dt1.Rows.Count > 0)
            {
                string Str_AH_Update = "Update  ANNOUNCED_HEAT set CONARC_A='" + StrConarcA + "',CONARC_B='" + StrConarcB + "'";
                bool Status = clsObj.DBInsertUpdateDeleteMIS(Str_AH_Update);
            }
            else
            {
                string Str_AH_Insert = "Insert into ANNOUNCED_HEAT (CONARC_A,CONARC_B) values('" + StrConarcA + "','" + StrConarcB + "')";
                bool Status = clsObj.DBInsertUpdateDeleteMIS(Str_AH_Insert);

            }
            FnUpdateTotalAl();
            FnUpdateInjection();
            FnUpdateChemistry();
            FnUpdateSCRAPYARD();
            //FnInsertScrapMtrl();

        }

        public void FnUpdateSCRAPYARD()
        {

            string StrQueryString = "select HeatNo from ConarcOpGuide where B_ShreddScrap is null and B_CRCABundle is null and B_PI is null and B_SCPrec_Skull is null and  B_CASTIRON is null and B_HMS is null and B_RE_fc is null and B_PlantRecycle is null order by SLNo desc";
            DataTable dtSCRAPYARD = new DataTable();

            dtSCRAPYARD = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtSCRAPYARD.Rows.Count > 0)
            {
                for (int i = 0; i < dtSCRAPYARD.Rows.Count; i++)
                {
                    string StrHeatNo = dtSCRAPYARD.Rows[i][0].ToString();
                    if (StrHeatNo != "")
                    {
                        FnUpdateSCRAPYARD(StrHeatNo);
                    }
                }

            }


        }
        public string FnUpdateSCRAPYARD(string StrHeatNo)
        {
            string Status = "false";
            try
            {

                DbconSql.Open();
                DataTable dt_SC_SMSII = new DataTable();
                string StrQueryString = "Select SHREDDSCRAP,CRCABUNDLE,PI,SKULL,CASTIRON,HMS,RE_FC,PLANTRECYCLE,ShiftInchConarc  from SCRAP_YARD where HeatNo='" + StrHeatNo + "'";
                SqlDataAdapter adp = new SqlDataAdapter(StrQueryString, DbconSql);
                adp.Fill(dt_SC_SMSII);
                DbconSql.Close();
                if (dt_SC_SMSII.Rows.Count > 0)
                {
                    String StrSHREDDSCRAP = dt_SC_SMSII.Rows[0][0].ToString();
                    if (StrSHREDDSCRAP == "") { StrSHREDDSCRAP = "0"; }
                    String StrCRCABUNDLE = dt_SC_SMSII.Rows[0][1].ToString();
                    if (StrCRCABUNDLE == "") { StrCRCABUNDLE = "0"; }
                    String StrPI = dt_SC_SMSII.Rows[0][2].ToString();
                    if (StrPI == "") { StrPI = "0"; }
                    String StrSKULL = dt_SC_SMSII.Rows[0][3].ToString();
                    if (StrSKULL == "") { StrSKULL = "0"; }
                    String StrCASTIRON = dt_SC_SMSII.Rows[0][4].ToString();
                    if (StrCASTIRON == "") { StrCASTIRON = "0"; }
                    String StrHMS = dt_SC_SMSII.Rows[0][5].ToString();
                    if (StrHMS == "") { StrHMS = "0"; }
                    String StrRE_FC = dt_SC_SMSII.Rows[0][6].ToString();
                    if (StrRE_FC == "") { StrRE_FC = "0"; }
                    String StrPLANTRECYCLE = dt_SC_SMSII.Rows[0][7].ToString();
                    if (StrPLANTRECYCLE == "") { StrPLANTRECYCLE = "0"; }
                    String StrShiftInchConarc = dt_SC_SMSII.Rows[0][8].ToString();


                    //SqlConnection SqlCon = new SqlConnection(StrConnectionSQL);
                    DbconSql.Open();



                    SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set ShiftIncharge='" + StrShiftInchConarc + "',B_ShreddScrap='" + StrSHREDDSCRAP + "',B_CRCABundle='" + StrCRCABUNDLE + "',B_PI='" + StrPI + "',B_SCPrec_Skull='" + StrSKULL + "',B_CASTIRON='" + StrCASTIRON + "',B_HMS='" + StrHMS + "',B_RE_fc='" + StrRE_FC + "',B_PlantRecycle='" + StrPLANTRECYCLE + "'  where HeatNo='" + StrHeatNo + "'", DbconSql);
                    Cmd.ExecuteNonQuery();
                    DbconSql.Close();
                }

            }
            catch
            {

            }
            return Status = "true";
        }

        //public void FnInsertScrapMtrl()
        //{
            
        //    string Status = "false";
        //    string StrScrpMtrlMstrQuery = "Select MAT_CODE,DESCR from GT_MAT ";
        //    DataTable dt_mtrl = new DataTable();
        //    dt_mtrl = clsObj.DBSelectQueryMIS_Table(StrScrpMtrlMstrQuery);
        //    if (dt_mtrl.Rows.Count > 0)
        //    {
        //        for (int i = 0; i < dt_mtrl.Rows.Count; i++)
        //        {
        //            string MAT_CODE = dt_mtrl.Rows[i][0].ToString();
        //            string MRTL_DESCRP = dt_mtrl.Rows[i][1].ToString();
        //        }

        //    }
        //    OrcConOrc.Open();
        //    string StrOrcConInsertQuery = "INSERT INTO Scrap_Mtrl_Mstr (Mat_Code,Descr,Mode) VALUES ('" + MAT_CODE + "','" + MRTL_DESCRP + "','1')";
        //    OracleCommand cmdscrp = new OracleCommand(StrOrcConInsertQuery, OrcConOrc);
        //    try
        //    {
        //        cmdscrp.ExecuteNonQuery();
        //        Status = "true";
        //    }
        //    catch (Exception ex)
        //    {
        //        Status = "false";
        //        txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
        //    }
        //    OrcConOrc.Close();
        //    if (Status == "true")
        //    {
        //        DbconSql.Open();
        //        SqlCommand cmdscrpmtrl = new SqlCommand("Update Scrap_Mtrl_Mstr set SAP_CODE = '0' where Mat_Code='" + MAT_CODE + "' and Mode= '" + MRTL_DESCRP + "' ", DbconSql);
        //        cmdscrpmtrl.ExecuteNonQuery();
        //        DbconSql.Close();
        //    }
        //}
        public void FnUpdateChemistry()
        {

            string StrQueryString = "select HeatNo from ConarcopGuide where CHEM_C IS NULL and CHEM_Mn is null and CHEM_N2ppm is null and CHEM_P is null and CHEM_S is null and CHEM_Si is null order by SLNo desc";
            DataTable dtChem = new DataTable();

            dtChem = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtChem.Rows.Count > 0)
            {
                for (int i = 0; i < dtChem.Rows.Count; i++)
                {
                    string StrHeatNo = dtChem.Rows[0][0].ToString();
                    if (StrHeatNo != "")
                    {
                        FnInsertSMSIIdb3(StrHeatNo);
                    }
                }

            }


        }
        public void FnUpdateInjection()
        {

            string StrQueryString = "select HeatNo from ConarcopGuide where INJ_O2fromTL IS NULL or INJ_O2fromEBTlance_doorlance is null order by SLNo desc";
            DataTable dtTotalINJ = new DataTable();

            dtTotalINJ = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtTotalINJ.Rows.Count > 0)
            {
                for (int i = 0; i < dtTotalINJ.Rows.Count; i++)
                {
                    string StrHeatNo = dtTotalINJ.Rows[i][0].ToString();
                    if (StrHeatNo != "")
                    {
                        FnInsertSMSIIdb5(StrHeatNo);
                    }
                }

            }


        }
        public void FnUpdateTotalAl()
        {
            try
            {
                string StrQueryString = "select HeatNo from ConarcopGuide where AlWireinKg IS NULL   order by SLNo desc";
                DataTable dtTotalAl = new DataTable();

                dtTotalAl = clsObj.DBSelectQueryMIS_Table(StrQueryString);
                if (dtTotalAl.Rows.Count > 0)
                {
                    for (int i = 0; i < dtTotalAl.Rows.Count; i++)
                    {


                        string StrHeatNo = dtTotalAl.Rows[i][0].ToString();
                        if (StrHeatNo != "")
                        {
                            OrcConOrc.Open();
                            DataTable dt1SMSII = new DataTable();
                            string StrQueryString1 = "SELECT pp_heat.heatid_cust, vr_lf_heat_report.plant,vr_lf_heat_report.treatid_cust,vr_lf_heat_report.arrival,vr_lf_heat_report.departure FROM vr_lf_heat_report, pp_heat WHERE (    (pp_heat.heatid = vr_lf_heat_report.heatid) AND (pp_heat.heatid_cust = vr_lf_heat_report.heatid_cust) And pp_heat.heatid_cust = '" + StrHeatNo + "' )";
                            OracleDataAdapter adp = new OracleDataAdapter(StrQueryString1, OrcConOrc);
                            adp.Fill(dt1SMSII);
                            OrcConOrc.Close();
                            if (dt1SMSII.Rows.Count > 0)
                            {
                                string StrDepartureTime = dt1SMSII.Rows[0][4].ToString();
                                if (StrDepartureTime != "")
                                {
                                    FnInsertSMSIIdb2(StrHeatNo);
                                    FnUpdateFormulaFields(StrHeatNo);
                                    FnUpdateFormulaFieldsDivideByZERO(StrHeatNo);
                                }
                            }

                        }
                    }

                }
            }
            catch
            {
            }
            finally
            {
                if (OrcConOrc != null) OrcConOrc.Close();

            }

        }
        public string FnInsertSMSIIdb1(string StrConarcA)
        {
            string Status = "false";
            try
            {
                DbconSql.Open();
                string StrExecteQuery = "";
                StrExecteQuery = "Insert Into ConarcOpGuide (HeatNo,LMWeight) values('" + StrConarcA + "','185')";
                SqlCommand CmdIns = new SqlCommand(StrExecteQuery, DbconSql);
                CmdIns.ExecuteNonQuery();
                DbconSql.Close();

                OrcConOrc.Open();
                DataTable dt1SMSII = new DataTable();
                string StrQueryString = "SELECT distinct pp_heat_plant.heatid_cust,pp_heat_plant.plant, pd_report.announcetime,pd_heatdata.treatstart_act,pd_report.tappingstarttime,pd_report.tappingendtime, pd_report.taptotapduration, pd_report.l1_dri_total_cons, pd_report.l1_lime_total_cons,pd_report.l1_dolo_total_cons, pd_report.l1_coal_total_cons,pdb_heat_data.o2_blow_dur, pdb_heat_data.total_o2_cons,pd_report.total_elec_egy, pde_heat_data.power_on_dur,pdb_heat_data.blowstarttime,pde_heat_data.arcing_starttime   FROM pp_heat_plant,pd_report,pd_heatdata,       pdb_heat_data, pde_heat_data,       pdf_heat_phase_data, pd_anl  WHERE   (pp_heat_plant.plant = pd_report.plant)        AND (pd_heatdata.plant = pd_report.plant)        AND (pd_heatdata.heatid = pd_report.heatid)        AND (pd_heatdata.treatid = pd_report.treatid)        AND (pd_heatdata.heatid = pdb_heat_data.heatid)        AND (pdb_heat_data.heatid = pde_heat_data.heatid)        AND (pp_heat_plant.heatid = pd_report.heatid)        AND (pde_heat_data.heatid = pdf_heat_phase_data.heatid)        AND (pdf_heat_phase_data.heatid = pd_anl.heatid)    AND ( pp_heat_plant.heatid_cust='" + StrConarcA + "')    AND (pp_heat_plant.plant='CON')";
                OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                adp.Fill(dt1SMSII);
                OrcConOrc.Close();
                if (dt1SMSII.Rows.Count > 0)
                {
                    DateTime dtAnnounceTime = System.DateTime.Now;
                    DateTime dtStartTime = System.DateTime.Now;
                    DateTime dtBlowingStartTime = System.DateTime.Now;
                    DateTime dtArchingStartTime = System.DateTime.Now;
                    String StrAnnounceTime = dt1SMSII.Rows[0][2].ToString().Replace(',', '.');
                    if (StrAnnounceTime != "")
                    {
                        dtAnnounceTime = Convert.ToDateTime(StrAnnounceTime);
                    }
                    String StrStartTime = dt1SMSII.Rows[0][3].ToString().Replace(',', '.');
                    if (StrStartTime != "")
                    {
                        dtStartTime = Convert.ToDateTime(StrStartTime);
                    }
                    //String StrStartTime = dt1SMSII.Rows[0][4].ToString().Replace(',', '.');
                    //if (StrStartTime != "")
                    //{
                    //    dtStartTime = Convert.ToDateTime(StrStartTime);
                    //}
                    String StrHeatTap = dt1SMSII.Rows[0][5].ToString().Replace(',', '.');
                    String StrTapToTap = dt1SMSII.Rows[0][6].ToString();
                    String StrDRI = dt1SMSII.Rows[0][7].ToString();
                    if (StrDRI == "") { StrDRI = "0"; }
                    String StrLIME = dt1SMSII.Rows[0][8].ToString();
                    if (StrLIME == "") { StrLIME = "0"; }
                    String StrDolime = dt1SMSII.Rows[0][9].ToString();
                    if (StrDolime == "") { StrDolime = "0"; }
                    String StrCoke = dt1SMSII.Rows[0][10].ToString();
                    if (StrCoke == "") { StrCoke = "0"; }
                    String StrBlowTime = dt1SMSII.Rows[0][11].ToString();
                    String StrCokeInjection = dt1SMSII.Rows[0][12].ToString();
                    String StrTotalEnery = dt1SMSII.Rows[0][13].ToString();
                    String StrArchingTime = dt1SMSII.Rows[0][14].ToString();
                    String StrBlowingStartTime = dt1SMSII.Rows[0][15].ToString().Replace(',', '.');
                    if (StrBlowingStartTime != "")
                    {
                        dtBlowingStartTime = Convert.ToDateTime(StrBlowingStartTime);
                    }
                    String StrArchingStartTime = dt1SMSII.Rows[0][16].ToString().Replace(',', '.');
                    if (StrArchingStartTime != "")
                    {
                        dtArchingStartTime = Convert.ToDateTime(StrArchingStartTime);
                    }
                    if (dtBlowingStartTime < dtArchingStartTime)
                    {
                        dtStartTime = dtBlowingStartTime;
                    }
                    else
                    {
                        dtStartTime = dtArchingStartTime;
                    }
                    string StrShift = GetShift(dtAnnounceTime.ToString("HH:mm:ss"));
                    DbconSql.Open();
                    SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set Shift='" + StrShift + "',SetupTime=DATEDIFF(minute,'" + dtAnnounceTime + "','" + dtStartTime + "'),HeatTap='" + StrHeatTap + "',Date='" + StrAnnounceTime.Replace(',', '.') + "',TapToTap='" + StrTapToTap + "',CF_DRI=" + StrDRI + ",CF_LIME=" + StrLIME + ",CF_Dolime=" + StrDolime + ",CF_Coke='" + StrCoke + "',ENERGY='" + StrTotalEnery + "',BlowingTime='" + StrBlowTime + "',INJ_Coke='" + StrCokeInjection + "',ArchingTime='" + StrArchingTime + "' where HeatNo='" + StrConarcA + "' ; Update SCRAP_YARD  set EDATE='" + StrAnnounceTime.Replace(',', '.') + "' where HEATNO='" + StrConarcA + "' ; Update  ConarcOpGuide Set CF_DRI=ROUND(CF_DRI/1000,2),CF_LIME=round(CF_LIME/1000,2),CF_Dolime=round(CF_Dolime/1000,2) where HeatNo='" + StrConarcA + "'", DbconSql);
                    Cmd.ExecuteNonQuery();
                    DbconSql.Close();
                }
                return Status = "true";
            }
            catch (Exception ex)
            {
            }

            return Status;
        }
        public string FnInsertSMSIIdb2(string StrConarcA)
        {
            string Status = "false";
            try
            {

                OrcConOrc.Open();
                DataTable dt2SMSII = new DataTable();
                string StrQueryString = "SELECT SUM (vr_lf_heat_report_additionlf.mat_charged) AS totalwire  FROM vr_lf_heat_report_additionlf, pp_heat WHERE ((pp_heat.heatid = vr_lf_heat_report_additionlf.heatid))   AND pp_heat.heatid_cust = '" + StrConarcA + "'   AND vr_lf_heat_report_additionlf.mat_code = '4202' ";
                OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                adp.Fill(dt2SMSII);
                OrcConOrc.Close();
                if (dt2SMSII.Rows.Count > 0)
                {
                    String StrTotalInKg = dt2SMSII.Rows[0][0].ToString();
                    if (StrTotalInKg != "")
                    {
                        DbconSql.Open();
                        SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set AlWireinKg='" + StrTotalInKg + "' where HeatNo='" + StrConarcA + "'", DbconSql);
                        Cmd.ExecuteNonQuery();
                        DbconSql.Close();
                    }
                }

                Status = "true";
            }
            catch (Exception ex)
            {
            }

            return Status;
        }
        public string FnInsertSMSIIdb3(string StrConarcA)
        {
            string Status = "false";
            try
            {
                OrcConOrc.Open();
                DataTable dt3SMSII = new DataTable();
                string StrQueryString = "SELECT pp_heat.heatid, pp_heat.heatid_cust,  pd_report_heat_final_anl.ename,pd_report_heat_final_anl.anl,    pd_report_heat_final_anl.plant,       pd_report_heat_final_anl.treatid  FROM pd_report_heat_final_anl, pp_heat WHERE ((pp_heat.heatid = pd_report_heat_final_anl.heatid)) and pp_heat.heatid_cust='" + StrConarcA + "' AND pd_report_heat_final_anl.plant='CON'";
                OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                adp.Fill(dt3SMSII);
                OrcConOrc.Close();
                if (dt3SMSII.Rows.Count > 0)
                {
                    String StrCCons = dt3SMSII.Rows[3][3].ToString();
                    String StrMnCons = dt3SMSII.Rows[7][3].ToString();
                    String StrSCons = dt3SMSII.Rows[10][3].ToString();
                    String StrPCons = dt3SMSII.Rows[9][3].ToString();
                    //String StrSiCons = dt3SMSII.Rows[16][3].ToString();
                    //String StrN2ppm = dt3SMSII.Rows[9][3].ToString();
                    String StrSiCons = "0";
                    String StrN2ppm = "0";


                    //SqlConnection SqlCon = new SqlConnection(StrConnectionSQL);
                    DbconSql.Open();
                    SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set CHEM_C='" + StrCCons + "',CHEM_Mn='" + StrMnCons + "',CHEM_Si='" + StrSiCons + "',CHEM_S='" + StrSCons + "',CHEM_P='" + StrPCons + "',CHEM_N2ppm='" + StrN2ppm + "'   where HeatNo='" + StrConarcA + "'", DbconSql);
                    Cmd.ExecuteNonQuery();
                    DbconSql.Close();
                }
                return Status = "true";
            }
            catch (Exception ex)
            {
            }
            return Status;
        }
        public string FnInsertSMSIIdb4(string StrConarcA)
        {
            string Status = "false";
            try
            {
                OrcConOrc.Open();
                DataTable dt3SMSII = new DataTable();
                string StrQueryString = "SELECT pd_report.heatid, pd_report.lasttemp, pd_report.departtime,pd_report.LASTTEMPTIME   FROM pd_report, pp_heat WHERE ((pd_report.heatid = pp_heat.heatid)) AND ( pp_heat.heatid_cust='" + StrConarcA + "')    AND (pd_report.plant='CON')";
                OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                adp.Fill(dt3SMSII);
                OrcConOrc.Close();
                if (dt3SMSII.Rows.Count > 0)
                {
                    String StrTappingTime = dt3SMSII.Rows[0][1].ToString();

                    //SqlConnection SqlCon = new SqlConnection(StrConnectionSQL);
                    DbconSql.Open();
                    SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set TappingTemp='" + StrTappingTime + "'  where HeatNo='" + StrConarcA + "'", DbconSql);
                    Cmd.ExecuteNonQuery();
                    DbconSql.Close();
                }
                return Status = "true";
            }
            catch (Exception ex)
            {
            }
            return Status;
        }
        public string FnInsertSMSIIdb5(string StrConarcA)
        {
            string Status = "false";
            try
            {
                OrcConOrc.Open();
                DataTable dt3SMSII = new DataTable();
                string StrQueryString = "SELECT  distinct pp_heat_plant.heatid_cust as HEATID,pp_heat_plant.plant,Sum(pdf_heat_phase_data.top_lance_o2_cons) as TotO2TL ,sum(pdf_heat_phase_data.doorlance_o2_cons) as Tot_O2DOOR,sum(pdf_heat_phase_data.ebtlance_o2_cons) as Tot_EBTlen,sum(pdf_heat_phase_data.doorlance_o2_cons)+sum(pdf_heat_phase_data.ebtlance_o2_cons) as Tot_O2_Cons FROM pp_heat_plant,pdf_heat_phase_data   WHERE   (pp_heat_plant.heatid = pdf_heat_phase_data.heatid) AND ( pp_heat_plant.heatid_cust='" + StrConarcA + "') AND (pp_heat_plant.plant='CON') group by pp_heat_plant.heatid_cust,pp_heat_plant.plant";
                OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                adp.Fill(dt3SMSII);
                OrcConOrc.Close();
                if (dt3SMSII.Rows.Count > 0)
                {
                    String StrTotalO2TL = dt3SMSII.Rows[0][2].ToString();
                    String StrTotalO2DOOR = dt3SMSII.Rows[0][5].ToString();
                    //String StrTotalEBTLen = dt3SMSII.Rows[0][4].ToString();

                    //SqlConnection SqlCon = new SqlConnection(StrConnectionSQL);
                    if (StrTotalO2TL == "")
                    {
                        StrTotalO2TL = "0";
                    }
                    if (StrTotalO2DOOR == "")
                    {
                        StrTotalO2DOOR = "0";
                    }
                    DbconSql.Open();
                    SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set INJ_O2fromTL='" + StrTotalO2TL + "',INJ_O2fromEBTlance_doorlance='" + StrTotalO2DOOR + "'  where HeatNo='" + StrConarcA + "'", DbconSql);
                    Cmd.ExecuteNonQuery();
                    DbconSql.Close();


                }
                return Status = "true";
            }
            catch (Exception ex)
            {
            }
            return Status;
        }
        public string FnInsertSMSIIdb6(string StrConarcA)
        {
            string Status = "false";
            try
            {

                OrcConOrc.Open();
                DataTable dt3SMSII = new DataTable();
                string StrQueryString = "SELECT SUM (vr_lf_heat_report_additionlf.mat_charged) AS Al_ingot_LF  FROM vr_lf_heat_report_additionlf, pp_heat WHERE ((pp_heat.heatid = vr_lf_heat_report_additionlf.heatid))   AND pp_heat.heatid_cust = '" + StrConarcA + "'   AND vr_lf_heat_report_additionlf.mat_code = '4401'";
                OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                adp.Fill(dt3SMSII);
                OrcConOrc.Close();
                if (dt3SMSII.Rows.Count > 0)
                {
                    String StrAlIngotLF = dt3SMSII.Rows[0][0].ToString();
                    if (StrAlIngotLF == "")
                    {
                        StrAlIngotLF = "0";
                    }
                    //SqlConnection SqlCon = new SqlConnection(StrConnectionSQL);
                    DbconSql.Open();
                    SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set AlIngotinLF='" + StrAlIngotLF + "' where HeatNo='" + StrConarcA + "'", DbconSql);
                    Cmd.ExecuteNonQuery();
                    DbconSql.Close();
                }
                return Status = "true";
            }
            catch (Exception ex)
            {
            }
            return Status;
        }
        public string FnInsertSMSIIdb7(string StrConarcA)
        {
            string Status = "false";
            try
            {

                OrcConOrc.Open();
                DataTable dt3SMSII = new DataTable();
                string StrQueryString = "SELECT nvl(pdf_heat_data.hmweight,0)/1000 as HMWt      FROM pdf_heat_data, pp_heat WHERE ((pp_heat.heatid = pdf_heat_data.heatid)        and (pp_heat .heatid_cust= '" + StrConarcA + "')        and (pdf_heat_data.plant='CON'))";
                OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                adp.Fill(dt3SMSII);
                OrcConOrc.Close();
                if (dt3SMSII.Rows.Count > 0)
                {
                    String StrHM = dt3SMSII.Rows[0][0].ToString();
                    if (StrHM == "")
                    {
                        StrHM = "0";
                    }
                    //SqlConnection SqlCon = new SqlConnection(StrConnectionSQL);
                    DbconSql.Open();
                    SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set HM='" + StrHM + "' where HeatNo='" + StrConarcA + "'", DbconSql);
                    Cmd.ExecuteNonQuery();
                    DbconSql.Close();
                }
                return Status = "true";
            }
            catch (Exception ex)
            {
            }
            return Status;
        }
        public string FnInsertSMSIIdb8(string StrConarcA)
        {
            string Status = "false";
            try
            {

                OrcConOrc.Open();
                DataTable dt3SMSII = new DataTable();
                string StrQueryString = "SELECT  distinct (pd_recipe_entry.matweight) as ALIngot   FROM pp_heat, pd_recipe_entry WHERE ((pp_heat.heatid = pd_recipe_entry.heatid)        and (pp_heat.heatid_cust= '" + StrConarcA + "')        AND (pd_recipe_entry.mat_code = '4401')        AND (pd_recipe_entry.plant = 'CON'))";
                OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                adp.Fill(dt3SMSII);
                OrcConOrc.Close();
                if (dt3SMSII.Rows.Count > 0)
                {
                    String StrALIngot = dt3SMSII.Rows[0][0].ToString();
                    if (StrALIngot == "")
                    {
                        StrALIngot = "0";
                    }
                    //SqlConnection SqlCon = new SqlConnection(StrConnectionSQL);
                    DbconSql.Open();
                    SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set AlIngot='" + StrALIngot + "' where HeatNo='" + StrConarcA + "'", DbconSql);
                    Cmd.ExecuteNonQuery();
                    DbconSql.Close();
                }
                return Status = "true";
            }
            catch (Exception ex)
            {
            }
            return Status;
        }
        public string FnInsertSMSIIdb9(string StrConarcA)
        {
            string Status = "false";
            try
            {

                OrcConOrc.Open();
                DataTable dt3SMSII = new DataTable();
                string StrQueryString = "SELECT   distinct sum(pde_heat_data_inj.inj_cons) AS Coke FROM pde_heat_data_inj, pp_heat_plant WHERE ((pp_heat_plant.heatid = pde_heat_data_inj.heatid) AND ( pp_heat_plant.heatid_cust='" + StrConarcA + "') AND (pp_heat_plant.plant='CON') ) AND (pde_heat_data_inj.inj_mat='5400')";
                OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
                adp.Fill(dt3SMSII);
                OrcConOrc.Close();
                if (dt3SMSII.Rows.Count > 0)
                {
                    String StrInjCOKE = dt3SMSII.Rows[0][0].ToString();
                    if (StrInjCOKE == "")
                    {
                        StrInjCOKE = "0";
                    }
                    //SqlConnection SqlCon = new SqlConnection(StrConnectionSQL);
                    DbconSql.Open();
                    string StrSQLQuery = "Update  ConarcOpGuide Set INJ_Coke='" + StrInjCOKE + "' where HeatNo='" + StrConarcA + "'";
                    StrSQLQuery = StrSQLQuery + " ; update ConarcOpGuide set DelayTime=TapToTap-(ArchingTime+BlowingTime+SetupTime+0.04) where HeatNo='" + StrConarcA + "'";
                    SqlCommand Cmd = new SqlCommand(StrSQLQuery, DbconSql);
                    Cmd.ExecuteNonQuery();
                    DbconSql.Close();
                }
                return Status = "true";
            }
            catch (Exception ex)
            {
            }
            return Status;
        }
        public void FnUpdateFormulaFields(string StrHeatNo)
        {
            try
            {
                // string StrQueryString = "update ConarcOpGuide set TotalCharge=ISNULL(B_CASTIRON,0)+ISNULL(B_CRCABundle,0)+ISNULL(B_HBI,0)+ISNULL(B_HMS,0)+ISNULL(B_PI,0)+ISNULL(B_PlantRecycle,0)+ISNULL(B_RE_fc,0)+ISNULL(B_SCPrec_Skull,0)+ISNULL(B_ShreddScrap,0)+ISNULL(CF_DRI,0)+ISNULL(HM,0),Yield=((ISNULL(LMWeight,0)/ISNULL(TotalCharge,0))*100),AlWireinKg=(ISNULL(TotalAl,0)*36 ),PreviousHeatTapped=DATEADD(minute,2,date),KWH_TON=(ISNULL(ENERGY,0)/ISNULL(TotalCharge,0)) where HeatNo='" + StrHeatNo + "'";
                DbconSql.Close();
                string StrQueryString = "update ConarcOpGuide set TotalCharge=ISNULL(B_CASTIRON,0)+ISNULL(B_CRCABundle,0)+ISNULL(B_HBI,0)+ISNULL(B_HMS,0)+ISNULL(B_PI,0)+ISNULL(B_PlantRecycle,0)+ISNULL(B_RE_fc,0)+ISNULL(B_SCPrec_Skull,0)+ISNULL(B_ShreddScrap,0)+(ISNULL(CF_DRI,0))+ISNULL(HM,0)where HeatNo='" + StrHeatNo + "'";
                //StrQueryString = StrQueryString+ " ; update ConarcOpGuide set Yield=((ISNULL(LMWeight,0)/ISNULL(TotalCharge,0))*100)where HeatNo='" + StrHeatNo + "'";
                StrQueryString = StrQueryString + " ;update ConarcOpGuide set AlWirein_M=(ISNULL(AlWireinKg,0)/0.36 )where HeatNo='" + StrHeatNo + "'";
                StrQueryString = StrQueryString + " ;update ConarcOpGuide set PreviousHeatTapped=DATEADD(minute,2,date)where HeatNo='" + StrHeatNo + "'";
                //StrQueryString = StrQueryString +  " ;update ConarcOpGuide set KWH_TON=(ISNULL(ENERGY,0)/ISNULL(TotalCharge,0))  where HeatNo='" + StrHeatNo + "'";

                DbconSql.Open();
                SqlCommand Cmd = new SqlCommand(StrQueryString, DbconSql);
                Cmd.ExecuteNonQuery();
                DbconSql.Close();
            }
            catch (Exception ex)
            {
            }

        }
        public void FnUpdateFormulaFieldsDivideByZERO(string StrHeatNo)
        {
            DbconSql.Open();
            DataTable dt = new DataTable();
            SqlDataAdapter adp = new SqlDataAdapter("select ENERGY,TotalCharge from ConarcOpGuide where HeatNo='" + StrHeatNo + "'", DbconSql);
            adp.Fill(dt);
            DbconSql.Close();
            string StrEnergy = "0";
            string StrTotalCharge = "0";
            if (dt.Rows.Count > 0)
            {
                StrEnergy = dt.Rows[0][0].ToString();
            }
            if (dt.Rows.Count > 0)
            {
                StrTotalCharge = dt.Rows[0][1].ToString();
            }
            string StrQueryString = "";
            if (StrTotalCharge == "0.00" || StrTotalCharge == "")
            {
                StrQueryString = "update ConarcOpGuide set Yield=0 where HeatNo='" + StrHeatNo + "'";
            }
            else
            {
                StrQueryString = "update ConarcOpGuide set Yield=((ISNULL(LMWeight,0)/ISNULL(TotalCharge,0))*100)where HeatNo='" + StrHeatNo + "'";
            }
            if (StrEnergy == "0" || StrEnergy == "" || StrTotalCharge == "0.00")
            {
                StrQueryString = StrQueryString + " ;update ConarcOpGuide set KWH_TON=0  where HeatNo='" + StrHeatNo + "'";
            }
            else
            {
                StrQueryString = StrQueryString + " ;update ConarcOpGuide set KWH_TON=(ISNULL(ENERGY,0)/ISNULL(TotalCharge,0))  where HeatNo='" + StrHeatNo + "'";
            }
            StrQueryString = StrQueryString + " ;update ConarcOpGuide set TotalAl=(ISNULL(AlIngot,0)+ISNULL(AlIngotinLF,0))+ISNULL(AlWireinKg,0)  where HeatNo='" + StrHeatNo + "'";
            StrQueryString = StrQueryString + " ;Update  ConarcOpGuide Set AlCons_Kg_T=isnull(TotalAl,0)/isnull(LMWeight,1) where HeatNo='" + StrHeatNo + "'";

            DbconSql.Open();
            SqlCommand Cmd = new SqlCommand(StrQueryString, DbconSql);
            Cmd.ExecuteNonQuery();
            DbconSql.Close();

        }

        private void timer2_Tick(object sender, EventArgs e)
        {

            DateTime C_datetimeLocal = Convert.ToDateTime(System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            string StrQueryString = "select HeatNo from lgc_heattimetable where  DATEDIFF(minute,C_datetime,'" + C_datetimeLocal + "')>30";
            DataTable dtSCRAPYARD = new DataTable();

            dtSCRAPYARD = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtSCRAPYARD.Rows.Count > 0)
            {
                for (int i = 0; i < dtSCRAPYARD.Rows.Count; i++)
                {
                    string StrHeatNo = dtSCRAPYARD.Rows[i][0].ToString();
                    string Status1 = FnInsertSMSIIdb1(StrHeatNo);
                    //string Status2 = FnInsertSMSIIdb2(StrHeatNo);
                    string Status3 = FnInsertSMSIIdb3(StrHeatNo);
                    string Status4 = FnInsertSMSIIdb4(StrHeatNo);
                    string Status5 = FnInsertSMSIIdb5(StrHeatNo);
                    string Status6 = FnInsertSMSIIdb6(StrHeatNo);
                    string Status7 = FnInsertSMSIIdb7(StrHeatNo);
                    string Status8 = FnInsertSMSIIdb8(StrHeatNo);
                    string Status9 = FnInsertSMSIIdb9(StrHeatNo);
                    if (Status1 == "true" && Status3 == "true" && Status4 == "true" && Status5 == "true" && Status6 == "true" && Status7 == "true" && Status8 == "true")
                    {
                        string Str_lgc_HeatNo_Delete = "delete from lgc_heattimetable where HeatNo='" + StrHeatNo + "'";
                        bool Status = clsObj.DBInsertUpdateDeleteMIS(Str_lgc_HeatNo_Delete);

                    }
                    else
                    {
                        string Str_lgc_HeatNo_Delete = "delete from ConarcOpGuide where HeatNo='" + StrHeatNo + "'";
                        bool Status = clsObj.DBInsertUpdateDeleteMIS(Str_lgc_HeatNo_Delete);

                    }
                    FnUpdateFormulaFields(StrHeatNo);
                    FnUpdateFormulaFieldsDivideByZERO(StrHeatNo);
                    FnUpdateHMWeight_ManualEntry();
                    FnUpdateHMWeight_ManualEntry_ZERO();
                }

            }

        }
        public void FnUpdateHMWeight_ManualEntry()
        {

            string StrQueryString = "select HeatNo from ConarcOpGuide where HMWtasPerHMDS is null and DS_NDS is null and CeloxTips is null and Sample is null and TempTips is null and  TapOxygen is null and ProcessType is null order by SLNo desc";
            DataTable dtSCRAPYARD = new DataTable();

            dtSCRAPYARD = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtSCRAPYARD.Rows.Count > 0)
            {
                for (int i = 0; i < dtSCRAPYARD.Rows.Count; i++)
                {
                    string StrHeatNo = dtSCRAPYARD.Rows[i][0].ToString();
                    if (StrHeatNo != "")
                    {
                        FnUpdateHMWeight_ManualEntry(StrHeatNo);
                    }
                }

            }


        }
        public void FnUpdateHMWeight_ManualEntry_ZERO()
        {

            string StrQueryString = "select HeatNo from ConarcOpGuide where HMWtasPerHMDS =0 order by SLNo desc";
            DataTable dtSCRAPYARD = new DataTable();

            dtSCRAPYARD = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtSCRAPYARD.Rows.Count > 0)
            {
                for (int i = 0; i < dtSCRAPYARD.Rows.Count; i++)
                {
                    string StrHeatNo = dtSCRAPYARD.Rows[i][0].ToString();
                    if (StrHeatNo != "")
                    {
                        FnUpdateHMWeight_ManualEntry(StrHeatNo);
                    }
                }

            }


        }
        public string FnUpdateHMWeight_ManualEntry(string StrHeatNo)
        {
            string Status = "false";
            DbconSql.Open();
            DataTable dt_SC_SMSII = new DataTable();
            string StrQueryString = "select HMwtHMD,DS_NDS,CeloxTips,Sample,TempTips,TapOxygen,ProcessType,HBI,CPC from mnl_HMwtHMD_ConarcProcessData where HeatNo='" + StrHeatNo + "'";
            SqlDataAdapter adp = new SqlDataAdapter(StrQueryString, DbconSql);
            adp.Fill(dt_SC_SMSII);
            DbconSql.Close();
            try
            {
                if (dt_SC_SMSII.Rows.Count > 0)
                {
                    String StrHMwtHMD = dt_SC_SMSII.Rows[0][0].ToString();
                    if (StrHMwtHMD == "") { StrHMwtHMD = "0"; }
                    String StrDS_NDS = dt_SC_SMSII.Rows[0][1].ToString();
                    if (StrDS_NDS == "") { StrDS_NDS = "0"; }
                    String StrCeloxTips = dt_SC_SMSII.Rows[0][2].ToString();
                    if (StrCeloxTips == "") { StrCeloxTips = "0"; }
                    String StrSample = dt_SC_SMSII.Rows[0][3].ToString();
                    if (StrSample == "") { StrSample = "0"; }
                    String StrTempTips = dt_SC_SMSII.Rows[0][4].ToString();
                    if (StrTempTips == "") { StrTempTips = "0"; }
                    String StrTapOxygen = dt_SC_SMSII.Rows[0][5].ToString();
                    if (StrTapOxygen == "") { StrTapOxygen = "0"; }
                    String StrProcessType = dt_SC_SMSII.Rows[0][6].ToString();
                    if (StrProcessType == "") { StrProcessType = "0"; }
                    String StrHBI = dt_SC_SMSII.Rows[0][7].ToString();
                    if (StrHBI == "") { StrHBI = "0"; }
                    String StrCPC = dt_SC_SMSII.Rows[0][8].ToString();
                    if (StrCPC == "") { StrCPC = "0"; }

                    DbconSql.Open();
                    string StrQueryHMWt = "Update  ConarcOpGuide Set HMWtasPerHMDS='" + StrHMwtHMD + "',DS_NDS='" + StrDS_NDS + "',CeloxTips='" + StrCeloxTips + "',Sample='" + StrSample + "',TempTips='" + StrTempTips + "',TapOxygen='" + StrTapOxygen + "',ProcessType='" + StrProcessType + "',B_HBI='" + StrHBI + "',CPC='" + StrCPC + "'  where HeatNo='" + StrHeatNo + "'";
                    StrQueryHMWt = StrQueryHMWt + " ; update ConarcOpGuide set Diff=HMWtasPerHMDS-HM  where heatno='" + StrHeatNo + "'";
                    SqlCommand Cmd = new SqlCommand(StrQueryHMWt, DbconSql);
                    Cmd.ExecuteNonQuery();
                    DbconSql.Close();
                }
            }
            catch
            {
            }
            finally
            {
                if (DbconSql != null) DbconSql.Close();

            }
            return Status = "true";
        }
        private void btnEnterManualA_Click(object sender, EventArgs e)
        {
            string Status1 = FnInsertSMSIIdb1ManualUpdate(txtHeatNo.Text.Trim());
            string Status2 = FnInsertSMSIIdb2(txtHeatNo.Text.Trim());
            string Status3 = FnInsertSMSIIdb3(txtHeatNo.Text.Trim());
            string Status4 = FnInsertSMSIIdb4(txtHeatNo.Text.Trim());
            string Status5 = FnInsertSMSIIdb5(txtHeatNo.Text.Trim());
            string Status6 = FnInsertSMSIIdb6(txtHeatNo.Text.Trim());
            string Status7 = FnInsertSMSIIdb7(txtHeatNo.Text.Trim());
            string Status8 = FnInsertSMSIIdb8(txtHeatNo.Text.Trim());
            string Status9 = FnInsertSMSIIdb9(txtHeatNo.Text.Trim());
            FnUpdateFormulaFields(txtHeatNo.Text.Trim());
            FnUpdateFormulaFieldsDivideByZERO(txtHeatNo.Text.Trim());
        }
        public string FnInsertSMSIIdb1ManualUpdate(string StrHeatNo)
        {
            string Status = "false";


            OrcConOrc.Open();
            DataTable dt1SMSII = new DataTable();
            string StrQueryString = "SELECT distinct pp_heat_plant.heatid_cust,pp_heat_plant.plant, pd_report.announcetime,pd_heatdata.treatstart_act,pd_report.tappingstarttime,pd_report.tappingendtime, pd_report.taptotapduration, pd_report.l1_dri_total_cons, pd_report.l1_lime_total_cons,pd_report.l1_dolo_total_cons, pd_report.l1_coal_total_cons,pdb_heat_data.o2_blow_dur, pdb_heat_data.total_o2_cons,pd_report.total_elec_egy, pde_heat_data.power_on_dur   FROM pp_heat_plant,pd_report,pd_heatdata,       pdb_heat_data, pde_heat_data,       pdf_heat_phase_data, pd_anl  WHERE   (pp_heat_plant.plant = pd_report.plant)        AND (pd_heatdata.plant = pd_report.plant)        AND (pd_heatdata.heatid = pd_report.heatid)        AND (pd_heatdata.treatid = pd_report.treatid)        AND (pd_heatdata.heatid = pdb_heat_data.heatid)        AND (pdb_heat_data.heatid = pde_heat_data.heatid)        AND (pp_heat_plant.heatid = pd_report.heatid)        AND (pde_heat_data.heatid = pdf_heat_phase_data.heatid)        AND (pdf_heat_phase_data.heatid = pd_anl.heatid)    AND ( pp_heat_plant.heatid_cust='" + StrHeatNo + "')    AND (pp_heat_plant.plant='CON')";
            OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
            adp.Fill(dt1SMSII);
            OrcConOrc.Close();
            if (dt1SMSII.Rows.Count > 0)
            {
                DateTime dtAnnounceTime = System.DateTime.Now;
                DateTime dtStartTime = System.DateTime.Now;
                String StrAnnounceTime = dt1SMSII.Rows[0][2].ToString().Replace(',', '.');
                if (StrAnnounceTime != "")
                {
                    dtAnnounceTime = Convert.ToDateTime(StrAnnounceTime);
                }
                String StrStartTime = dt1SMSII.Rows[0][4].ToString().Replace(',', '.');
                if (StrStartTime != "")
                {
                    dtStartTime = Convert.ToDateTime(StrStartTime);
                }
                String StrHeatTap = dt1SMSII.Rows[0][5].ToString().Replace(',', '.');
                String StrTapToTap = dt1SMSII.Rows[0][6].ToString();
                String StrDRI = dt1SMSII.Rows[0][7].ToString();
                if (StrDRI == "") { StrDRI = "0"; }
                String StrLIME = dt1SMSII.Rows[0][8].ToString();
                if (StrLIME == "") { StrLIME = "0"; }
                String StrDolime = dt1SMSII.Rows[0][9].ToString();
                if (StrDolime == "") { StrDolime = "0"; }
                String StrCoke = dt1SMSII.Rows[0][10].ToString();
                if (StrCoke == "") { StrCoke = "0"; }
                String StrBlowTime = dt1SMSII.Rows[0][11].ToString();
                String StrCokeInjection = dt1SMSII.Rows[0][12].ToString();
                String StrTotalEnery = dt1SMSII.Rows[0][13].ToString();
                String StrArchingTime = dt1SMSII.Rows[0][14].ToString();
                string StrShift = GetShift(dtAnnounceTime.ToString("HH:mm:ss"));
                DbconSql.Open();
                SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set Shift='" + StrShift + "',SetupTime=DATEDIFF(minute,'" + dtAnnounceTime + "','" + dtStartTime + "'),HeatTap='" + StrHeatTap + "',Date='" + StrAnnounceTime.Replace(',', '.') + "',TapToTap='" + StrTapToTap + "',CF_DRI=r" + StrDRI + ",CF_LIME=" + StrLIME + ",CF_Dolime=" + StrDolime + ",CF_Coke='" + StrCoke + "',ENERGY='" + StrTotalEnery + "',BlowingTime='" + StrBlowTime + "',INJ_Coke='" + StrCokeInjection + "',ArchingTime='" + StrArchingTime + "' where HeatNo='" + StrHeatNo + "' ; Update SCRAP_YARD  set EDATE='" + StrAnnounceTime.Replace(',', '.') + "' where HEATNO='" + StrHeatNo + "'", DbconSql);
                Cmd.ExecuteNonQuery();
                DbconSql.Close();
            }
            return Status = "true";


            return Status;
        }
        public string FnInsertSMSIIdb1ManualInsert(string StrHeatNo)
        {
            string Status = "false";

            DbconSql.Open();
            string StrExecteQuery = "";
            StrExecteQuery = "Insert Into ConarcOpGuide (HeatNo,LMWeight) values('" + StrHeatNo + "','185')";
            SqlCommand CmdIns = new SqlCommand(StrExecteQuery, DbconSql);
            CmdIns.ExecuteNonQuery();
            DbconSql.Close();

            OrcConOrc.Open();
            DataTable dt1SMSII = new DataTable();
            string StrQueryString = "SELECT distinct pp_heat_plant.heatid_cust,pp_heat_plant.plant, pd_report.announcetime,pd_heatdata.treatstart_act,pd_report.tappingstarttime,pd_report.tappingendtime, pd_report.taptotapduration, pd_report.l1_dri_total_cons, pd_report.l1_lime_total_cons,pd_report.l1_dolo_total_cons, pd_report.l1_coal_total_cons,pdb_heat_data.o2_blow_dur, pdb_heat_data.total_o2_cons,pd_report.total_elec_egy, pde_heat_data.power_on_dur   FROM pp_heat_plant,pd_report,pd_heatdata,       pdb_heat_data, pde_heat_data,       pdf_heat_phase_data, pd_anl  WHERE   (pp_heat_plant.plant = pd_report.plant)        AND (pd_heatdata.plant = pd_report.plant)        AND (pd_heatdata.heatid = pd_report.heatid)        AND (pd_heatdata.treatid = pd_report.treatid)        AND (pd_heatdata.heatid = pdb_heat_data.heatid)        AND (pdb_heat_data.heatid = pde_heat_data.heatid)        AND (pp_heat_plant.heatid = pd_report.heatid)        AND (pde_heat_data.heatid = pdf_heat_phase_data.heatid)        AND (pdf_heat_phase_data.heatid = pd_anl.heatid)    AND ( pp_heat_plant.heatid_cust='" + StrHeatNo + "')    AND (pp_heat_plant.plant='CON')";
            OracleDataAdapter adp = new OracleDataAdapter(StrQueryString, OrcConOrc);
            adp.Fill(dt1SMSII);
            OrcConOrc.Close();
            if (dt1SMSII.Rows.Count > 0)
            {
                DateTime dtAnnounceTime = System.DateTime.Now;
                DateTime dtStartTime = System.DateTime.Now;
                String StrAnnounceTime = dt1SMSII.Rows[0][2].ToString().Replace(',', '.');
                if (StrAnnounceTime != "")
                {
                    dtAnnounceTime = Convert.ToDateTime(StrAnnounceTime);
                }
                String StrStartTime = dt1SMSII.Rows[0][4].ToString().Replace(',', '.');
                if (StrStartTime != "")
                {
                    dtStartTime = Convert.ToDateTime(StrStartTime);
                }
                String StrHeatTap = dt1SMSII.Rows[0][5].ToString().Replace(',', '.');
                String StrTapToTap = dt1SMSII.Rows[0][6].ToString();
                String StrDRI = dt1SMSII.Rows[0][7].ToString();
                if (StrDRI == "") { StrDRI = "0"; }
                String StrLIME = dt1SMSII.Rows[0][8].ToString();
                if (StrLIME == "") { StrLIME = "0"; }
                String StrDolime = dt1SMSII.Rows[0][9].ToString();
                if (StrDolime == "") { StrDolime = "0"; }
                String StrCoke = dt1SMSII.Rows[0][10].ToString();
                if (StrCoke == "") { StrCoke = "0"; }
                String StrBlowTime = dt1SMSII.Rows[0][11].ToString();
                String StrCokeInjection = dt1SMSII.Rows[0][12].ToString();
                String StrTotalEnery = dt1SMSII.Rows[0][13].ToString();
                String StrArchingTime = dt1SMSII.Rows[0][14].ToString();
                string StrShift = GetShift(dtAnnounceTime.ToString("HH:mm:ss"));
                DbconSql.Open();
                SqlCommand Cmd = new SqlCommand("Update  ConarcOpGuide Set Shift='" + StrShift + "',SetupTime=DATEDIFF(minute,'" + dtAnnounceTime + "','" + dtStartTime + "'),HeatTap='" + StrHeatTap + "',Date='" + StrAnnounceTime.Replace(',', '.') + "',TapToTap='" + StrTapToTap + "',CF_DRI='" + StrDRI + "',CF_LIME='" + StrLIME + "',CF_Dolime='" + StrDolime + "',CF_Coke='" + StrCoke + "',ENERGY='" + StrTotalEnery + "',BlowingTime='" + StrBlowTime + "',INJ_Coke='" + StrCokeInjection + "',ArchingTime='" + StrArchingTime + "' where HeatNo='" + StrHeatNo + "' ; Update SCRAP_YARD  set EDATE='" + StrAnnounceTime.Replace(',', '.') + "' where HEATNO='" + StrHeatNo + "'", DbconSql);
                Cmd.ExecuteNonQuery();
                DbconSql.Close();
            }
            return Status = "true";


            return Status;
        }
        private string GetShift(string Time)
        {
            //string DateNew = Time.ToString("HH:mm:ss");
            string DateNew = Time;
            string Shift = "";
            DateNew = DateNew.Substring(0, 2);
            int DateInt = Convert.ToInt32(DateNew);
            if (DateInt >= 6 && DateInt < 14)
            {
                Shift = "A";
            }
            else if (DateInt >= 14 && DateInt < 22)
            {
                Shift = "B";
            }
            else
            {
                Shift = "C";
            }
            return Shift;
        }

        private void btnHeatNoInsert_Click(object sender, EventArgs e)
        {
            string Status1 = FnInsertSMSIIdb1ManualInsert(txtHeatNoInsert.Text.Trim());
            string Status2 = FnInsertSMSIIdb2(txtHeatNoInsert.Text.Trim());
            string Status3 = FnInsertSMSIIdb3(txtHeatNoInsert.Text.Trim());
            string Status4 = FnInsertSMSIIdb4(txtHeatNoInsert.Text.Trim());
            string Status5 = FnInsertSMSIIdb5(txtHeatNoInsert.Text.Trim());
            string Status6 = FnInsertSMSIIdb6(txtHeatNoInsert.Text.Trim());
            string Status7 = FnInsertSMSIIdb7(txtHeatNoInsert.Text.Trim());
            string Status8 = FnInsertSMSIIdb8(txtHeatNoInsert.Text.Trim());
            string Status9 = FnInsertSMSIIdb9(txtHeatNoInsert.Text.Trim());
            FnUpdateFormulaFields(txtHeatNo.Text.Trim());
            FnUpdateFormulaFieldsDivideByZERO(txtHeatNo.Text.Trim());
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            try
            {

                FnInsertCokeOvenSync();
            }
            catch (Exception ex)
            {
                txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
            }

        }
        public void FnInsertCokeOvenSync()
        {

            string StrQueryString = "";
            DbconSql.Open();
            StrQueryString = "select Max(AutoId)'AutoId' from  v_batt_oprn_log_coke";
            SqlDataAdapter adp = new SqlDataAdapter(StrQueryString, DbconSql);
            DataTable dt = new DataTable();
            adp.Fill(dt);
            DbconSql.Close();
            string StrMaxId = dt.Rows[0][0].ToString();

            string StrQueryStringCokeOven = "Select v_batt_oprn_log.Autoid,v_batt_oprn_log.DateTime,v_batt_oprn_log.EquipType ,v_equipmentdetails.EquipmentDesc,v_eventmaster.EventDesc,v_machinedetails.MachineName from v_batt_oprn_log    ,v_equipmentdetails    ,v_machinedetails    ,v_eventmaster WHERE v_batt_oprn_log.EquipItemCode = v_equipmentdetails.EquipDtlCode AND v_batt_oprn_log.EventCode =  v_eventmaster.EventCode AND    v_batt_oprn_log.McnDtlCode =  v_machinedetails.McnDtlCode  AND v_batt_oprn_log.Autoid > " + StrMaxId + " order by  v_batt_oprn_log.Autoid asc";

            try
            {
                DbconMySql.Open();
                MySqlDataAdapter adpMySQL = new MySqlDataAdapter(StrQueryStringCokeOven, DbconMySql);
                DataSet ds = new DataSet();
                adpMySQL.Fill(ds);
                DbconMySql.Close();
                int i = 0;
                if (ds.Tables[0].Rows.Count > 0)
                {
                    DbconSql.Open();

                    for (i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        SqlCommand cmd = new SqlCommand("Insert into v_batt_oprn_log_coke (AutoId,DateTime,EquipType,EquipDesc,EventDesc,MachineName) values('" + ds.Tables[0].Rows[i][0].ToString() + "' ,'" + ds.Tables[0].Rows[i][1].ToString() + "' ,'" + ds.Tables[0].Rows[i][2].ToString() + "','" + ds.Tables[0].Rows[i][3].ToString() + "','" + ds.Tables[0].Rows[i][4].ToString() + "','" + ds.Tables[0].Rows[i][5].ToString() + "')", DbconSql);
                        cmd.ExecuteNonQuery();
                    }
                    DbconSql.Close();
                }
            }
            catch (Exception ex)
            {
                txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
            }
            finally
            {
                DbconSql.Close();
            }

        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            //string Status9 = FnInsertSMSIIdb9("12B3018");
            //FnUpdateTotalAl();
            FnUpdateHMWeight_ManualEntry();
            FnUpdateHMWeight_ManualEntry_ZERO();
        }

        public void FnUpdateSCRAPYARDCheck()
        {

            string StrQueryString = "select HeatNo from Scrap_Yard where ShreddScrap is null and CRCABundle is null and HBI is null and PI is null and SKULL is null and  CASTIRON is null and HMS is null and RE_fc is null and PlantRecycle is null order by SLNo desc";
            DataTable dtSCRAPYARD = new DataTable();

            dtSCRAPYARD = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtSCRAPYARD.Rows.Count > 0)
            {
                for (int i = 0; i < dtSCRAPYARD.Rows.Count; i++)
                {
                    string StrHeatNo = dtSCRAPYARD.Rows[i][0].ToString();
                    if (StrHeatNo != "")
                    {
                        FnUpdateSCRAPYARDCheck(StrHeatNo);
                    }
                }

            }


        }
        public string FnUpdateSCRAPYARDCheck(string StrHeatNo)
        {
            string Status = "false";
            DbconSql.Open();
            DataTable dt_SC_SMSII = new DataTable();
            string StrQueryString = "Select B_SHREDDSCRAP,B_CRCABUNDLE,B_HBI,B_PI,B_SCPrec_Skull,B_CASTIRON,B_HMS,B_RE_FC,B_PLANTRECYCLE  from ConarcOpGuide where HeatNo='" + StrHeatNo + "'";
            SqlDataAdapter adp = new SqlDataAdapter(StrQueryString, DbconSql);
            adp.Fill(dt_SC_SMSII);
            DbconSql.Close();
            if (dt_SC_SMSII.Rows.Count > 0)
            {
                String StrSHREDDSCRAP = dt_SC_SMSII.Rows[0][0].ToString();
                if (StrSHREDDSCRAP == "") { StrSHREDDSCRAP = "0"; }
                String StrCRCABUNDLE = dt_SC_SMSII.Rows[0][1].ToString();
                if (StrCRCABUNDLE == "") { StrCRCABUNDLE = "0"; }
                String StrHBI = dt_SC_SMSII.Rows[0][2].ToString();
                if (StrHBI == "") { StrHBI = "0"; }
                String StrPI = dt_SC_SMSII.Rows[0][3].ToString();
                if (StrPI == "") { StrPI = "0"; }
                String StrSKULL = dt_SC_SMSII.Rows[0][4].ToString();
                if (StrSKULL == "") { StrSKULL = "0"; }
                String StrCASTIRON = dt_SC_SMSII.Rows[0][5].ToString();
                if (StrCASTIRON == "") { StrCASTIRON = "0"; }
                String StrHMS = dt_SC_SMSII.Rows[0][6].ToString();
                if (StrHMS == "") { StrHMS = "0"; }
                String StrRE_FC = dt_SC_SMSII.Rows[0][7].ToString();
                if (StrRE_FC == "") { StrRE_FC = "0"; }
                String StrPLANTRECYCLE = dt_SC_SMSII.Rows[0][8].ToString();
                if (StrPLANTRECYCLE == "") { StrPLANTRECYCLE = "0"; }
                //String StrShiftInchConarc = dt_SC_SMSII.Rows[0][9].ToString();


                //SqlConnection SqlCon = new SqlConnection(StrConnectionSQL);
                DbconSql.Open();



                SqlCommand Cmd = new SqlCommand("Update  Scrap_Yard Set ShreddScrap='" + StrSHREDDSCRAP + "',CRCABundle='" + StrCRCABUNDLE + "',HBI='" + StrHBI + "',PI='" + StrPI + "',SKULL='" + StrSKULL + "',CASTIRON='" + StrCASTIRON + "',HMS='" + StrHMS + "',RE_fc='" + StrRE_FC + "',PlantRecycle='" + StrPLANTRECYCLE + "'  where HeatNo='" + StrHeatNo + "'", DbconSql);
                Cmd.ExecuteNonQuery();
                DbconSql.Close();
            }
            return Status = "true";
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            //UnCheck//
            //FnInsertCASTERIIdbTable_PPC_PROD_ORDER_CASTERI();
            //FnInsertCASTERIIdbTable_PPC_SLAB_ORDER_CASTERI();
            //FnInsertCASTERIIdbTable_PPC_PROD_ORDER_CASTERII();
            //FnInsertCASTERIIdbTable_PPC_SLAB_ORDER_CASTERII();
            //UnCheck//

        }

        public string FnInsertCASTERIIdbTable_PPC_PROD_ORDER_CASTERI()
        {
            string Status = "false";

            string StrQueryString = "SELECT UCode,PROD_ORDERID,PL_HEATID,PL_SEQID,HEAT_SEQ_NO,PL_ARRIVAL_TIME,PL_NET_WEIGHT,GRADE_CODE,CLO_ENABLED,UPD_TIME,PL_LADLE_NO,PL_WIDTH,TWIN_CAST,SEND_STATUS FROM [mnl_PPC_PROD_ORDER] WHERE SEND_STATUS IS NULL";
            DataTable dtMISDB = new DataTable();

            dtMISDB = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtMISDB.Rows.Count > 0)
            {
                for (int i = 0; i < dtMISDB.Rows.Count; i++)
                {


                    string UCode = dtMISDB.Rows[i][0].ToString();
                    string PROD_ORDERID = dtMISDB.Rows[i][1].ToString();
                    string PL_HEATID = dtMISDB.Rows[i][2].ToString();
                    string PL_SEQID = dtMISDB.Rows[i][3].ToString();
                    string HEAT_SEQ_NO = dtMISDB.Rows[i][4].ToString();
                    string PL_ARRIVAL_TIME = dtMISDB.Rows[i][5].ToString();
                    string PL_NET_WEIGHT = dtMISDB.Rows[i][6].ToString();
                    string GRADE_CODE = dtMISDB.Rows[i][7].ToString();
                    string CLO_ENABLED = dtMISDB.Rows[i][8].ToString();
                    string UPD_TIME = dtMISDB.Rows[i][9].ToString();
                    string PL_LADLE_NO = dtMISDB.Rows[i][10].ToString();
                    string PL_WIDTH = dtMISDB.Rows[i][11].ToString();
                    string TWIN_CAST = dtMISDB.Rows[i][12].ToString();


                    OrcConOrcCasterI.Open();
                    string StrOrcCasterInsertQuery = "Insert into mnl_PPC_PROD_ORDER (PROD_ORDERID,PL_HEATID,PL_SEQID,HEAT_SEQ_NO,PL_ARRIVAL_TIME,PL_NET_WEIGHT,GRADE_CODE,CLO_ENABLED,UPD_TIME,PL_LADLE_NO,PL_WIDTH,TWIN_CAST) values ('" + PROD_ORDERID + "','" + PL_HEATID + "','" + PL_SEQID + "','" + HEAT_SEQ_NO + "','" + PL_ARRIVAL_TIME + "','" + PL_NET_WEIGHT + "','" + GRADE_CODE + "','" + CLO_ENABLED + "','" + UPD_TIME + "','" + PL_LADLE_NO + "','" + PL_WIDTH + "','" + TWIN_CAST + "')";
                    OracleCommand cmdCasterI = new OracleCommand(StrOrcCasterInsertQuery, OrcConOrcCasterI);
                    try
                    {
                        cmdCasterI.ExecuteNonQuery();
                        Status = "true";
                    }
                    catch (Exception ex)
                    {
                        Status = "false";
                        txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
                    }

                    OrcConOrcCasterI.Close();
                    if (Status == "true")
                    {
                        DbconSql.Open();
                        SqlCommand cmdMISDB = new SqlCommand("update mnl_PPC_PROD_ORDER set SEND_STATUS=1 WHERE UCode='" + UCode + "'", DbconSql);
                        cmdMISDB.ExecuteNonQuery();
                        DbconSql.Close();
                    }

                }
            }

            return Status;
        }

        public string FnInsertCASTERIIdbTable_PPC_SLAB_ORDER_CASTERI()
        {
            string Status = "false";

            string StrQueryString = "SELECT UCode,PROD_ORDERID,SLAB_ORDERID,SLAB_QUANTITY,WIDTH_HEAD,WIDTH_TAIL,SLAB_LENGTH_MIN,SLAB_LENGTH_AIM,THICKNESS_HEAD,THICKNESS_TAIL,STATUS,UPD_TIME,L3_SLAB_ROUTE,SLAB_TYPE,SLAB_LENGTH_MAX,SALES_ORDER_NO,CUSTOMER_NAME,SEND_STATUS FROM [mnl_PPC_SLAB_ORDER] WHERE SEND_STATUS IS NULL";
            DataTable dtMISDB = new DataTable();

            dtMISDB = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtMISDB.Rows.Count > 0)
            {
                for (int i = 0; i < dtMISDB.Rows.Count; i++)
                {
                    string UCode = dtMISDB.Rows[i][0].ToString();
                    string PROD_ORDERID = dtMISDB.Rows[i][1].ToString();
                    string SLAB_ORDERID = dtMISDB.Rows[i][2].ToString();
                    string SLAB_QUANTITY = dtMISDB.Rows[i][3].ToString();
                    string WIDTH_HEAD = dtMISDB.Rows[i][4].ToString();
                    string WIDTH_TAIL = dtMISDB.Rows[i][5].ToString();
                    string SLAB_LENGTH_MIN = dtMISDB.Rows[i][6].ToString();
                    string SLAB_LENGTH_AIM = dtMISDB.Rows[i][7].ToString();
                    string THICKNESS_HEAD = dtMISDB.Rows[i][8].ToString();
                    string THICKNESS_TAIL = dtMISDB.Rows[i][9].ToString();
                    string STATUS = dtMISDB.Rows[i][10].ToString();
                    //string CLO_PRIORITY = dtMISDB.Rows[i][11].ToString();
                    string UPD_TIME = dtMISDB.Rows[i][11].ToString();
                    string L3_SLAB_ROUTE = dtMISDB.Rows[i][12].ToString();
                    string SLAB_TYPE = dtMISDB.Rows[i][13].ToString();
                    string SLAB_LENGTH_MAX = dtMISDB.Rows[i][14].ToString();
                    //string SUBSLAB_R_WIDTH = dtMISDB.Rows[i][15].ToString();
                    //string SUBSLAB_M_WIDTH = dtMISDB.Rows[i][16].ToString();
                    //string SUBSLAB_L_WIDTH = dtMISDB.Rows[i][17].ToString();
                    //string CLO_SUBSLAB_PREPROC = dtMISDB.Rows[i][18].ToString();
                    //string STRAND_ASSIGN = dtMISDB.Rows[i][19].ToString();
                    string SALES_ORDER_NO = dtMISDB.Rows[i][15].ToString();
                    string CUSTOMER_NAME = dtMISDB.Rows[i][16].ToString();

                    OrcConOrcCasterI.Open();
                    string StrOrcCasterInsertQuery = "Insert into PPC_SLAB_ORDER(PROD_ORDERID,SLAB_ORDERID,SLAB_QUANTITY,WIDTH_HEAD,WIDTH_TAIL,SLAB_LENGTH_MIN,SLAB_LENGTH_AIM,THICKNESS_HEAD,THICKNESS_TAIL,STATUS,UPD_TIME,L3_SLAB_ROUTE,SLAB_TYPE,SLAB_LENGTH_MAX,SALES_ORDER_NO,CUSTOMER_NAME) values ('" + PROD_ORDERID + "','" + SLAB_ORDERID + "','" + SLAB_QUANTITY + "','" + WIDTH_HEAD + "','" + WIDTH_TAIL + "','" + SLAB_LENGTH_MIN + "','" + SLAB_LENGTH_AIM + "','" + THICKNESS_HEAD + "','" + THICKNESS_TAIL + "','" + STATUS + "','" + UPD_TIME + "','" + L3_SLAB_ROUTE + "','" + SLAB_TYPE + "','" + SLAB_LENGTH_MAX + "','" + SALES_ORDER_NO + "','" + CUSTOMER_NAME + "')";
                    OracleCommand cmdCasterI = new OracleCommand(StrOrcCasterInsertQuery, OrcConOrcCasterI);
                    try
                    {
                        cmdCasterI.ExecuteNonQuery();
                        Status = "true";
                    }
                    catch (Exception ex)
                    {
                        Status = "false";
                        txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
                    }
                    OrcConOrcCasterI.Close();
                    if (Status == "true")
                    {
                        DbconSql.Open();
                        SqlCommand cmdMISDB = new SqlCommand("update mnl_PPC_SLAB_ORDER set SEND_STATUS=1 WHERE UCode='" + UCode + "'", DbconSql);
                        cmdMISDB.ExecuteNonQuery();
                        DbconSql.Close();
                    }

                }
            }

            return Status;
        }

        public string FnInsertCASTERIIdbTable_PPC_PROD_ORDER_CASTERII()
        {
            string Status = "false";

            string StrQueryString = "SELECT UCode,PROD_ORDERID,PL_HEATID,PL_SEQID,HEAT_SEQ_NO,PL_ARRIVAL_TIME,PL_NET_WEIGHT,GRADE_CODE,CLO_ENABLED,UPD_TIME,PL_LADLE_NO,PL_WIDTH,TWIN_CAST,SEND_STATUS FROM [mnl_PPC_PROD_ORDER] WHERE SEND_STATUS IS NULL";
            DataTable dtMISDB = new DataTable();

            dtMISDB = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtMISDB.Rows.Count > 0)
            {
                for (int i = 0; i < dtMISDB.Rows.Count; i++)
                {


                    string UCode = dtMISDB.Rows[i][0].ToString();
                    string PROD_ORDERID = dtMISDB.Rows[i][1].ToString();
                    string PL_HEATID = dtMISDB.Rows[i][2].ToString();
                    string PL_SEQID = dtMISDB.Rows[i][3].ToString();
                    string HEAT_SEQ_NO = dtMISDB.Rows[i][4].ToString();
                    string PL_ARRIVAL_TIME = dtMISDB.Rows[i][5].ToString();
                    string PL_NET_WEIGHT = dtMISDB.Rows[i][6].ToString();
                    string GRADE_CODE = dtMISDB.Rows[i][7].ToString();
                    string CLO_ENABLED = dtMISDB.Rows[i][8].ToString();
                    string UPD_TIME = dtMISDB.Rows[i][9].ToString();
                    string PL_LADLE_NO = dtMISDB.Rows[i][10].ToString();
                    string PL_WIDTH = dtMISDB.Rows[i][11].ToString();
                    string TWIN_CAST = dtMISDB.Rows[i][12].ToString();


                    OrcConOrcCasterII.Open();
                    string StrOrcCasterInsertQuery = "Insert into PPC_PROD_ORDER (PROD_ORDERID,PL_HEATID,PL_SEQID,HEAT_SEQ_NO,PL_ARRIVAL_TIME,PL_NET_WEIGHT,GRADE_CODE,CLO_ENABLED,UPD_TIME,PL_LADLE_NO,PL_WIDTH,TWIN_CAST) values ('" + PROD_ORDERID + "','" + PL_HEATID + "','" + PL_SEQID + "','" + HEAT_SEQ_NO + "','" + PL_ARRIVAL_TIME + "','" + PL_NET_WEIGHT + "','" + GRADE_CODE + "','" + CLO_ENABLED + "','" + UPD_TIME + "','" + PL_LADLE_NO + "','" + PL_WIDTH + "','" + TWIN_CAST + "')";
                    OracleCommand cmdCasterII = new OracleCommand(StrOrcCasterInsertQuery, OrcConOrcCasterII);
                    cmdCasterII.ExecuteNonQuery();
                    OrcConOrcCasterII.Close();
                    Status = "true";
                    if (Status == "true")
                    {
                        DbconSql.Open();
                        SqlCommand cmdMISDB = new SqlCommand("update mnl_PPC_PROD_ORDER set SEND_STATUS=1 WHERE UCode='" + UCode + "'", DbconSql);
                        cmdMISDB.ExecuteNonQuery();
                        DbconSql.Close();
                    }

                }
            }

            return Status;
        }

        public string FnInsertCASTERIIdbTable_PPC_SLAB_ORDER_CASTERII()
        {
            string Status = "false";

            string StrQueryString = "SELECT UCode,PROD_ORDERID,SLAB_ORDERID,SLAB_QUANTITY,WIDTH_HEAD,WIDTH_TAIL,SLAB_LENGTH_MIN,SLAB_LENGTH_AIM,THICKNESS_HEAD,THICKNESS_TAIL,STATUS,UPD_TIME,L3_SLAB_ROUTE,SLAB_TYPE,SLAB_LENGTH_MAX,SALES_ORDER_NO,CUSTOMER_NAME,SEND_STATUS FROM [mnl_PPC_SLAB_ORDER] WHERE SEND_STATUS IS NULL";
            DataTable dtMISDB = new DataTable();

            dtMISDB = clsObj.DBSelectQueryMIS_Table(StrQueryString);

            if (dtMISDB.Rows.Count > 0)
            {
                for (int i = 0; i < dtMISDB.Rows.Count; i++)
                {
                    string UCode = dtMISDB.Rows[i][0].ToString();
                    string PROD_ORDERID = dtMISDB.Rows[i][1].ToString();
                    string SLAB_ORDERID = dtMISDB.Rows[i][2].ToString();
                    string SLAB_QUANTITY = dtMISDB.Rows[i][3].ToString();
                    string WIDTH_HEAD = dtMISDB.Rows[i][4].ToString();
                    string WIDTH_TAIL = dtMISDB.Rows[i][5].ToString();
                    string SLAB_LENGTH_MIN = dtMISDB.Rows[i][6].ToString();
                    string SLAB_LENGTH_AIM = dtMISDB.Rows[i][7].ToString();
                    string THICKNESS_HEAD = dtMISDB.Rows[i][8].ToString();
                    string THICKNESS_TAIL = dtMISDB.Rows[i][9].ToString();
                    string STATUS = dtMISDB.Rows[i][10].ToString();
                    //string CLO_PRIORITY = dtMISDB.Rows[i][11].ToString();
                    string UPD_TIME = dtMISDB.Rows[i][11].ToString();
                    string L3_SLAB_ROUTE = dtMISDB.Rows[i][12].ToString();
                    string SLAB_TYPE = dtMISDB.Rows[i][13].ToString();
                    string SLAB_LENGTH_MAX = dtMISDB.Rows[i][14].ToString();
                    //string SUBSLAB_R_WIDTH = dtMISDB.Rows[i][15].ToString();
                    //string SUBSLAB_M_WIDTH = dtMISDB.Rows[i][16].ToString();
                    //string SUBSLAB_L_WIDTH = dtMISDB.Rows[i][17].ToString();
                    //string CLO_SUBSLAB_PREPROC = dtMISDB.Rows[i][18].ToString();
                    //string STRAND_ASSIGN = dtMISDB.Rows[i][19].ToString();
                    string SALES_ORDER_NO = dtMISDB.Rows[i][15].ToString();
                    string CUSTOMER_NAME = dtMISDB.Rows[i][16].ToString();

                    OrcConOrcCasterII.Open();
                    string StrOrcCasterInsertQuery = "Insert into PPC_SLAB_ORDER(PROD_ORDERID,SLAB_ORDERID,SLAB_QUANTITY,WIDTH_HEAD,WIDTH_TAIL,SLAB_LENGTH_MIN,SLAB_LENGTH_AIM,THICKNESS_HEAD,THICKNESS_TAIL,STATUS,UPD_TIME,L3_SLAB_ROUTE,SLAB_TYPE,SLAB_LENGTH_MAX,SALES_ORDER_NO,CUSTOMER_NAME) values ('" + PROD_ORDERID + "','" + SLAB_ORDERID + "','" + SLAB_QUANTITY + "','" + WIDTH_HEAD + "','" + WIDTH_TAIL + "','" + SLAB_LENGTH_MIN + "','" + SLAB_LENGTH_AIM + "','" + THICKNESS_HEAD + "','" + THICKNESS_TAIL + "','" + STATUS + "','" + UPD_TIME + "','" + L3_SLAB_ROUTE + "','" + SLAB_TYPE + "','" + SLAB_LENGTH_MAX + "','" + SALES_ORDER_NO + "','" + CUSTOMER_NAME + "')";
                    OracleCommand cmdCasterII = new OracleCommand(StrOrcCasterInsertQuery, OrcConOrcCasterII);
                    try
                    {
                        cmdCasterII.ExecuteNonQuery();
                        Status = "true";
                    }
                    catch (Exception ex)
                    {
                        Status = "false";
                        txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
                    }
                    OrcConOrcCasterII.Close();

                    if (Status == "true")
                    {
                        DbconSql.Open();
                        SqlCommand cmdMISDB = new SqlCommand("update mnl_PPC_SLAB_ORDER set SEND_STATUS=1 WHERE UCode='" + UCode + "'", DbconSql);
                        cmdMISDB.ExecuteNonQuery();
                        DbconSql.Close();
                    }

                }
            }

            return Status;
        }

        private void LadleOpen_CheckStatus_Timer_Tick(object sender, EventArgs e)
        {

            string StrSTEEL_ID_caster1 = "";
            string StrSTEEL_ID_caster2 = "";
            string StrSTEEL_ID_caster3 = "";
            string StrHEATID_caster1 = "";
            string StrHEATID_caster2 = "";
            string StrHEATID_caster3 = "";

            //string StrQuerySeqIDMax = "select max(SEQ_NO)AS SEQ_NO from pdc_Sequence";
            //string StrQuerySteelID = "SELECT  pdc_heat.STEEL_ID,pdc_heat.HEATID from pdc_heat where  STEEL_ID=(Select max(STEEL_ID) from pdc_heat where pdc_heat.HEAT_STATUS_CODE=2 and pdc_heat.GRADE_CODE  is not null )";
            string StrQuerySteelID = "SELECT  pdc_heat.STEEL_ID,pdc_heat.HEATID from pdc_heat where SUBSTR(Ladle_Open_Time,1,10)='" + DateTime.Now.ToString("yyyy-MM-dd") + "' and HEAT_STATUS_CODE=2 and GRADE_CODE is not null order by Ladle_Open_Time desc";
            //string Query2 = "select heatid_cust from pp_heat where heatid=(select  heatid  from pd_heatdata where plant='CON' and plantno=2 and treatstart_plan=(Select max(treatstart_plan) from pd_heatdata where plant='CON' and plantno=2  ))";
            DBConnections clsObj_LOChkStat = new DBConnections();
            //DataTable dt_MaxSeqNo = new DataTable();
            DataTable dt_LOChkStat = new DataTable();
            DataTable dt_LOChkStatCaster1 = new DataTable();
            DataTable dt_LOChkStatCaster3 = new DataTable();
            //dt_MaxSeqNo = clsObj.DBSelectQueryCASTERII_Table(StrQuerySeqIDMax);
            //if (dt_MaxSeqNo.Rows.Count > 0)
            //{
            //    StrSeqNoMax_caster2 = dt_MaxSeqNo.Rows[0][0].ToString();
            //}
            dt_LOChkStat = clsObj.DBSelectQueryCASTERII_Table(StrQuerySteelID);
            dt_LOChkStatCaster1 = clsObj.DBSelectQueryCASTERI_Table(StrQuerySteelID);
            dt_LOChkStatCaster3 = clsObj.DBSelectQueryCASTERIII_Table(StrQuerySteelID);
            if (dt_LOChkStatCaster1.Rows.Count > 0)
            {

                for (int i = 0; i < dt_LOChkStatCaster1.Rows.Count; i++)
                {
                    StrSTEEL_ID_caster1 = dt_LOChkStatCaster1.Rows[i][0].ToString();
                    StrHEATID_caster1 = dt_LOChkStatCaster1.Rows[i][1].ToString();
                    FnCASTERI_1(StrSTEEL_ID_caster1, StrHEATID_caster1);
                    FnCASTERI_2(StrSTEEL_ID_caster1, StrHEATID_caster1);
                    FnCASTERI_3(StrSTEEL_ID_caster1, StrHEATID_caster1);
                    FnCASTERI_4(StrSTEEL_ID_caster1, StrHEATID_caster1);
                    FnCASTERI_5(StrSTEEL_ID_caster1, StrHEATID_caster1);
                    Fn_HEAT_REPORT_TO_L3_SLAB_CASTERI(StrHEATID_caster1);
                }

            }
            if (dt_LOChkStat.Rows.Count > 0)
            {

                for (int i = 0; i < dt_LOChkStat.Rows.Count; i++)
                {
                    StrSTEEL_ID_caster2 = dt_LOChkStat.Rows[i][0].ToString();
                    StrHEATID_caster2 = dt_LOChkStat.Rows[i][1].ToString();
                    FnCASTERII_1(StrSTEEL_ID_caster2, StrHEATID_caster2);
                    FnCASTERII_2(StrSTEEL_ID_caster2, StrHEATID_caster2);
                    FnCASTERII_3(StrSTEEL_ID_caster2, StrHEATID_caster2);
                    FnCASTERII_4(StrSTEEL_ID_caster2, StrHEATID_caster2);
                    FnCASTERII_5(StrSTEEL_ID_caster2, StrHEATID_caster2);
                    Fn_HEAT_REPORT_TO_L3_SLAB_CASTERII(StrHEATID_caster2);
                }

            }


            if (dt_LOChkStatCaster3.Rows.Count > 0)
            {

                for (int i = 0; i < dt_LOChkStatCaster3.Rows.Count; i++)
                {
                    StrSTEEL_ID_caster3 = dt_LOChkStatCaster3.Rows[i][0].ToString();
                    StrHEATID_caster3 = dt_LOChkStatCaster3.Rows[i][1].ToString();
                    FnCASTERIII_1(StrSTEEL_ID_caster3, StrHEATID_caster3);
                    FnCASTERIII_2(StrSTEEL_ID_caster3, StrHEATID_caster3);
                    FnCASTERIII_3(StrSTEEL_ID_caster3, StrHEATID_caster3);
                    FnCASTERIII_4(StrSTEEL_ID_caster3, StrHEATID_caster3);
                    FnCASTERIII_5(StrSTEEL_ID_caster3, StrHEATID_caster3);
                    Fn_HEAT_REPORT_TO_L3_SLAB_CASTERIII(StrHEATID_caster3);
                }

            }

        }
        public void FnCASTERI_1(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQueryMIS = "select SteelID from HEAT_REPORT_TO_L3 where SteelID='" + StrSTEEL_ID + "'";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtMIS = new DataTable();
            dtMIS = clsObj.DBSelectQueryMIS_Table(StrQueryMIS);
            if (dtMIS.Rows.Count == 0)
            {

                string StrMISInsetQuery = "Insert Into HEAT_REPORT_TO_L3 (SteelID,HeatID) values('" + StrSTEEL_ID + "','" + StrHEATID + "')";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }

            //string StrQuery = "SELECT pdc_heat.heatid, pdc_heat.last_heat_flag, pdc_heat.grade_code,pdc_heat.prod_orderid, pdc_heat.ladle_no, pdc_heat.lancing_flag,pdc_heat.shroud_type, pdc_heat.tund_powder_type,pdc_heat.stopper_rod_type, pdc_heat.nozzle_type,pdc_heat.ladle_open_weight, pdc_heat.ladle_open_time,pdc_heat.ladle_close_time, pdc_heat.heat_seq_no as HeatInSeq,pdc_heat.ladle_arrival_time, pdc_heat.ladle_close_weight,pdc_heat.slab_weight, pdc_heat.head_crop_weight,pdc_heat.tail_crop_weight, pdc_heat.cut_lost_weight,pdc_heat.slag_weight, pdc_heat.yield, pdc_heat.sample_lost_weight,pdc_heat.ladle_open_weight,pdc_heat.ladle_close_weight,pdc_heat.slab_length,PDC_HEAT.AVG_SLAB_WIDTH ,PDC_HEAT.POURING_DURATION,PDC_STRAND.Avg_cast_speed,PD_HEAT.LAST_TEMP,PDC_HEAT.Slab_COUNT,PDC_TUNDISH.TUND_SEQ_NO,PDC_TUNDISH.CAR_NO,PDC_TUNDISH.TUND_CLOSE_TIME,PDC_HEAT.LADLE_CLOSE_TUND_WEIGHT,PDC_HEAT.CAST_WEIGHT FROM pdc_heat,pdc_heat_strand,pdc_strand ,PD_HEAT ,PDC_TUNDISH WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID and pdc_heat_strand.STRAND_STEEL_ID=pdc_strand.STEEL_ID and PDC_TUNDISH.SEQUENCE_STEEL_ID=PDC_HEAT.SEQUENCE_STEEL_ID and PD_HEAT.HEATID=PDC_HEAT.HEATID and PDC_HEAT.HEATID='" + StrHEATID + "'";
            string StrQuery = "SELECT pdc_heat.heatid, pdc_heat.last_heat_flag, pdc_heat.grade_code,pdc_heat.prod_orderid, pdc_heat.ladle_no, pdc_heat.lancing_flag,pdc_heat.shroud_type, pdc_heat.tund_powder_type,pdc_heat.stopper_rod_type, pdc_heat.nozzle_type,pdc_heat.ladle_open_weight, pdc_heat.ladle_open_time,pdc_heat.ladle_close_time, pdc_heat.heat_seq_no as HeatInSeq,pdc_heat.ladle_arrival_time, pdc_heat.ladle_close_weight,pdc_heat.slab_weight, pdc_heat.head_crop_weight,pdc_heat.tail_crop_weight, pdc_heat.cut_lost_weight,pdc_heat.slag_weight, pdc_heat.yield, pdc_heat.sample_lost_weight,pdc_heat.ladle_open_weight,pdc_heat.ladle_close_weight,pdc_heat.slab_length,PDC_HEAT.AVG_SLAB_WIDTH ,PDC_HEAT.POURING_DURATION,PDC_STRAND.Avg_cast_speed,PD_HEAT.LAST_TEMP,PDC_HEAT.Slab_COUNT,PDC_TUNDISH.TUND_SEQ_NO,PDC_TUNDISH.CAR_NO,PDC_TUNDISH.TUND_CLOSE_TIME,PDC_HEAT.LADLE_CLOSE_TUND_WEIGHT,PDC_HEAT.CAST_WEIGHT FROM pdc_heat,pdc_heat_strand,pdc_strand ,PD_HEAT ,PDC_TUNDISH WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID and pdc_heat_strand.STRAND_STEEL_ID=pdc_strand.STEEL_ID and PDC_TUNDISH.SEQUENCE_STEEL_ID=PDC_HEAT.SEQUENCE_STEEL_ID and PD_HEAT.HEATID=PDC_HEAT.HEATID and PDC_HEAT.HEATID='" + StrHEATID + "'";
            //                         ,0            ,1                       ,2                    ,3                    ,4                 ,5                     ,6                   ,7                        ,8                        ,9                    ,10                        ,11                        ,12                      ,13                                ,14                         ,15                           ,16                 ,17                        ,18                        ,19                      ,20                   ,21             ,22                              ,23                    ,24                        ,25                   ,26  slab width          ,27 Pouring Duration       , 28  cast Speed         ,29 Lifting temp   ,30 Cut Slab Count  ,31  Tundish No       , 32 Tundish car No ,33 TC_DateTime             ,34 TundRemWeight                , 35 Net_Ls_Cast      
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERI_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string SteelId = "";//StrSTEEL_ID
                string HeatId = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                if (HeatId == "")
                {
                    HeatId = "0";
                }
                string SequenceNo = "0";
                string HeatInSeq = dtSelect.Rows[0][13].ToString();//pdc_heat.heat_seq_no 13
                if (HeatInSeq == "")
                {
                    HeatInSeq = "0";
                }
                string IsLastHeatInSeq = "0";
                string SteelGradeCode = dtSelect.Rows[0][2].ToString();//pdc_heat.grade_code   2
                if (SteelGradeCode == "")
                {
                    SteelGradeCode = "0";
                }
                string ProdOrderNo = dtSelect.Rows[0][3].ToString();//pdc_heat.prod_orderid    3 
                if (ProdOrderNo == "")
                {
                    ProdOrderNo = "0";
                }
                string CrewId = "0";
                string ShiftId = "0";
                string LadleNo = dtSelect.Rows[0][4].ToString();//pdc_heat.ladle_no 4
                if (LadleNo == "")
                {
                    LadleNo = "0";
                }
                string LadleLifeCount = "0";
                string Lancing = dtSelect.Rows[0][5].ToString();//pdc_heat.lancing_flag 5
                if (Lancing == "")
                {
                    Lancing = "0";
                }
                string ShroudType = dtSelect.Rows[0][6].ToString();//pdc_heat.shroud_type 6
                string TundPowerType = dtSelect.Rows[0][7].ToString();//pdc_heat.tund_powder_type  7
                string MidPowerType = "0";

                string StopperRodType = dtSelect.Rows[0][8].ToString();//pdc_heat.stopper_rod_type  8
                string NozzleType = dtSelect.Rows[0][9].ToString();//pdc_heat.nozzle_type 9
                string TundishNo = "0";
                string TundishCardNo = "0";
                string LadleArrival_DateTime = dtSelect.Rows[0][14].ToString().Replace(',', '.');//pdc_heat.ladle_arrival_time  14
                string LO_DateTime = dtSelect.Rows[0][11].ToString().Replace(',', '.'); ;//pdc_heat.ladle_open_time  11
                string LC_DateTime = dtSelect.Rows[0][12].ToString().Replace(',', '.'); ;//pdc_heat.ladle_close_time  12
                //string LadlePouringDur = DATEDIFF(minute,LC_DateTime,LO_DateTime)
                //string LadlePouringDur = dtSelect.Rows[0][15].ToString();//pdc_heat.ladle_open_time 11 - pdc_heat.ladle_close_time  12
                string LadleWeightNet = "0";
                string LadleRemSteelWght = "";
                string Ladle_Open_Wt = dtSelect.Rows[0][10].ToString();
                string Ladle_Close_Wt = dtSelect.Rows[0][15].ToString();
                if (Ladle_Open_Wt == "")
                {
                    Ladle_Open_Wt = "0";
                }
                if (Ladle_Close_Wt == "")
                {
                    Ladle_Close_Wt = "0";
                }
                //if (dtSelect.Rows[0][10].ToString() != "" || dtSelect.Rows[0][10].ToString() == DBNull.Value)
                //{
                //    Ladle_Open_Wt = Convert.ToInt64(dtSelect.Rows[0][10].ToString());
                //}
                //if (dtSelect.Rows[0][15].ToString() == "" || dtSelect.Rows[0][15].ToString() == null)
                //{
                //    Ladle_Close_Wt = Convert.ToInt64(dtSelect.Rows[0][15].ToString());
                //}
                string PouredHeatWeight = Convert.ToString(Convert.ToInt64(Ladle_Open_Wt) - Convert.ToInt64(Ladle_Close_Wt));//pdc_heat.ladle_open_weight 10 - pdc_heat.ladle_close_weight 15
                string TundRemWeight = "";
                string TotalSlabWeight = dtSelect.Rows[0][16].ToString();//pdc_heat.slab_weight 16
                if (TotalSlabWeight == "")
                {
                    TotalSlabWeight = "0";
                }
                string HeadWeight = dtSelect.Rows[0][17].ToString();//pdc_heat.head_crop_weight  17
                if (HeadWeight == "")
                {
                    HeadWeight = "0";
                }
                string TailWeight = dtSelect.Rows[0][18].ToString();//pdc_heat.tail_crop_weight  18
                if (TailWeight == "")
                {
                    TailWeight = "0";
                }
                string TotalCutLoss = dtSelect.Rows[0][19].ToString();//pdc_heat.cut_lost_weight  19
                if (TotalCutLoss == "")
                {
                    TotalCutLoss = "0";
                }
                string TotalSampleWeight = dtSelect.Rows[0][22].ToString();//pdc_heat.sample_lost_weight  22
                if (TotalSampleWeight == "")
                {
                    TotalSampleWeight = "0";
                }
                string SlagWeight = dtSelect.Rows[0][20].ToString();//pdc_heat.slag_weight 20
                if (SlagWeight == "")
                {
                    SlagWeight = "0";
                }
                string Yield = dtSelect.Rows[0][21].ToString();//pdc_heat.yield 21
                if (Yield == "")
                {
                    Yield = "0";
                }
                string Ld_Int_Wt = dtSelect.Rows[0][23].ToString();//pdc_heat.yield 23
                if (Ld_Int_Wt == "")
                {
                    Ld_Int_Wt = "0";
                }
                string Ld_Fin_Wt = dtSelect.Rows[0][24].ToString();//pdc_heat.yield 24
                if (Ld_Fin_Wt == "")
                {
                    Ld_Fin_Wt = "0";
                }
                string slab_length = dtSelect.Rows[0][25].ToString();//pdc_heat.slab_length 25
                if (slab_length == "")
                {
                    slab_length = "0";
                }
                string slab_Width = dtSelect.Rows[0]["AVG_SLAB_WIDTH"].ToString();//pdc_heat.AVG_slab_Width 26
                if (slab_Width == "")
                {
                    slab_Width = "0";
                }
                string StrPouringDuration = dtSelect.Rows[0]["POURING_DURATION"].ToString();//pdc_heat.POURING_DURATION 27
                if (StrPouringDuration == "")
                {
                    StrPouringDuration = "0";
                }
                string StrAVG_CAST_SPEED = dtSelect.Rows[0]["AVG_CAST_SPEED"].ToString();//pdc_heat.AVG_CAST_SPEED 28
                if (StrAVG_CAST_SPEED == "")
                {
                    StrAVG_CAST_SPEED = "0";
                }
                string StrLifting_Temp = dtSelect.Rows[0]["LAST_TEMP"].ToString();//pd_heat.LAST_TEMP 29
                if (StrLifting_Temp == "")
                {
                    StrLifting_Temp = "0";
                }
                string StrSlabCount = dtSelect.Rows[0]["SLAB_COUNT"].ToString();//pdc_heat.SLAB_COUNT 30
                if (StrSlabCount == "")
                {
                    StrSlabCount = "0";
                }
                string StrTundishNo = dtSelect.Rows[0]["TUND_SEQ_NO"].ToString();//pdc_TUNDISH.TUND_SEQ_NO 31
                if (StrTundishNo == "")
                {
                    StrTundishNo = "0";
                }
                string StrTundishCarNo = dtSelect.Rows[0]["CAR_NO"].ToString();//pdc_TUNDISH.CAR_NO 32
                if (StrTundishCarNo == "")
                {
                    StrTundishCarNo = "0";
                }
                string TC_DateTime = dtSelect.Rows[0]["TUND_CLOSE_TIME"].ToString().Replace(',', '.');// PDC_TUNDISH.TUND_CLOSE_TIME 33

                string StrTundRemWeight = dtSelect.Rows[0]["LADLE_CLOSE_TUND_WEIGHT"].ToString();//PDC_HEAT.LADLE_CLOSE_TUND_WEIGHT 34
                if (StrTundRemWeight == "")
                {
                    StrTundRemWeight = "0";
                }
                string StrNet_Ls_Cast = dtSelect.Rows[0]["CAST_WEIGHT"].ToString();//PDC_HEAT.CAST_WEIGHT 35
                if (StrNet_Ls_Cast == "")
                {
                    StrNet_Ls_Cast = "0";
                }


                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set HeatInSeq='" + HeatInSeq + "',SteelGradeCode='" + SteelGradeCode + "',ProdOrderNo='" + ProdOrderNo + "',LadleNo='" + LadleNo + "',Lancing='" + Lancing + "',ShroudType='" + ShroudType + "',TundPowerType='" + TundPowerType + "',StopperRodType='" + StopperRodType + "',NozzleType='" + NozzleType + "',LadleArrival_DateTime='" + LadleArrival_DateTime + "',LO_DateTime='" + LO_DateTime + "',LC_DateTime='" + LC_DateTime + "',LadlePouringDur=DATEDIFF(minute,LO_DateTime,LC_DateTime),TotalSlabWeight='" + TotalSlabWeight + "',HeadWeight='" + HeadWeight + "',TailWeight='" + TailWeight + "',TotalCutLoss='" + TotalCutLoss + "',TotalSampleWeight='" + TotalSampleWeight + "',SlagWeight='" + SlagWeight + "',Yield='" + Yield + "',Ld_Int_Wt='" + Ld_Int_Wt + "',Ld_Fin_Wt='" + Ld_Fin_Wt + "',Slab_Len='" + slab_length + "',Slab_Width='" + slab_Width + "',AVG_CAST_SPEED='" + StrAVG_CAST_SPEED + "',Lift_Temp='" + StrLifting_Temp + "',Ask_Temp='" + StrLifting_Temp + "',Cut_Slab_No='" + StrSlabCount + "',TundishNo='" + StrTundishNo + "',TundishCarNo='" + StrTundishCarNo + "',TC_DateTime='" + TC_DateTime + "',TundRemWeight='" + StrTundRemWeight + "',Net_Ls_Cast='" + StrNet_Ls_Cast + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);



            }

        }
        public void FnCASTERI_2(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT pdc_heat_shift.shift_code, pdc_heat_shift.crew_id  FROM pdc_heat_shift where pdc_heat_shift.HEAT_STEEL_ID=" + StrSTEEL_ID + "";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERI_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string StrShiftId = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                string StrCrewId = dtSelect.Rows[0][1].ToString();//pdc_heat.heatid   0

                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set ShiftId='" + StrShiftId + "',CrewId='" + StrCrewId + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }
        }
        public void FnCASTERI_3(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT pdc_tundish.tund_no, pdc_tundish.car_no, pdc_tundish.tund_close_weight,pdc_tundish.tund_close_time FROM pdc_tundish where pdc_tundish.steel_id=" + StrSTEEL_ID + "";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERI_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string StrTundishNo = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                string StrTundishCardNo = dtSelect.Rows[0][1].ToString();//pdc_heat.heatid   0
                string StrTundRemWeight = dtSelect.Rows[0][2].ToString();//pdc_heat.heatid   0
                string StrTC_DateTime = dtSelect.Rows[0][3].ToString().Replace(',', '.');//pdc_heat.heatid   0

                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set TundishNo='" + StrTundishNo + "',TundishCardNo='" + StrTundishCardNo + "',TundRemWeight='" + StrTundRemWeight + "',TC_DateTime='" + StrTC_DateTime + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }
        }
        public void FnCASTERI_4(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT CAST_POWDER_TYPE as MldPowderType FROM PDC_STRAND WHERE STEEL_ID = (SELECT STRAND_STEEL_ID FROM PDC_HEAT_STRAND WHERE HEAT_STEEL_ID = " + StrSTEEL_ID + " AND STRAND_NUM = 1)";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERI_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string StrMidPowerType = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set MidPowerType='" + StrMidPowerType + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }
        }
        public void FnCASTERI_5(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT pdc_tundish_temp.meas_no, pdc_tundish_temp.meas_time,pdc_tundish_temp.meas_temp, pdc_tundish_temp.ladle_weight,pdc_tundish_temp.tundish_weight FROM pdc_tundish_temp WHERE STEEL_ID ='" + StrSTEEL_ID + "'";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERI_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                for (int i = 0; i < dtSelect.Rows.Count; i++)
                {
                    string MEAS_NO = dtSelect.Rows[i][0].ToString();
                    string MEAS_TIME = dtSelect.Rows[i][1].ToString();
                    string MEAS_TEMP = dtSelect.Rows[i][2].ToString();
                    string LADLE_WEIGHT = dtSelect.Rows[i][3].ToString();
                    string TUNDISH_WEIGHT = dtSelect.Rows[i][4].ToString();

                    string StrQueryCheck = "select * from Cater2_Tundish_Temp where STEEL_ID='" + StrSTEEL_ID + "' and MEAS_NO='" + MEAS_NO + "'";
                    DataTable dtCheck = new DataTable();
                    dtCheck = clsObjMIS.DBSelectQueryMIS_Table(StrQueryCheck);
                    if (dtCheck.Rows.Count == 0)
                    {
                        string StrMISInsetQuery = "Insert Into  Cater2_Tundish_Temp (STEEL_ID,MEAS_NO,MEAS_TIME,MEAS_TEMP,LADLE_WEIGHT,TUNDISH_WEIGHT) values('" + StrSTEEL_ID + "','" + MEAS_NO + "','" + MEAS_TIME + "','" + MEAS_TEMP + "','" + LADLE_WEIGHT + "','" + TUNDISH_WEIGHT + "')";
                        bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);
                    }
                }
            }
        }

        public void FnCASTERII_1(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQueryMIS = "select SteelID from HEAT_REPORT_TO_L3 where SteelID='" + StrSTEEL_ID + "'";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtMIS = new DataTable();
            dtMIS = clsObj.DBSelectQueryMIS_Table(StrQueryMIS);
            if (dtMIS.Rows.Count == 0)
            {

                string StrMISInsetQuery = "Insert Into HEAT_REPORT_TO_L3 (SteelID,HeatID) values('" + StrSTEEL_ID + "','" + StrHEATID + "')";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }

            string StrQuery = "SELECT pdc_heat.heatid, pdc_heat.last_heat_flag, pdc_heat.grade_code,pdc_heat.prod_orderid, pdc_heat.ladle_no, pdc_heat.lancing_flag,pdc_heat.shroud_type, pdc_heat.tund_powder_type,pdc_heat.stopper_rod_type, pdc_heat.nozzle_type,pdc_heat.ladle_open_weight, pdc_heat.ladle_open_time,pdc_heat.ladle_close_time, pdc_heat.heat_seq_no as HeatInSeq,pdc_heat.ladle_arrival_time, pdc_heat.ladle_close_weight,pdc_heat.slab_weight, pdc_heat.head_crop_weight,pdc_heat.tail_crop_weight, pdc_heat.cut_lost_weight,pdc_heat.slag_weight, pdc_heat.yield, pdc_heat.sample_lost_weight,pdc_heat.ladle_open_weight,pdc_heat.ladle_close_weight,pdc_heat.slab_length,PDC_HEAT.AVG_SLAB_WIDTH ,PDC_HEAT.POURING_DURATION,PDC_STRAND.Avg_cast_speed,PD_HEAT.LAST_TEMP,PDC_HEAT.Slab_COUNT,PDC_TUNDISH.TUND_SEQ_NO,PDC_TUNDISH.CAR_NO,PDC_TUNDISH.TUND_CLOSE_TIME,PDC_HEAT.LADLE_CLOSE_TUND_WEIGHT,PDC_HEAT.CAST_WEIGHT FROM pdc_heat,pdc_heat_strand,pdc_strand ,PD_HEAT ,PDC_TUNDISH WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID and pdc_heat_strand.STRAND_STEEL_ID=pdc_strand.STEEL_ID and PDC_TUNDISH.SEQUENCE_STEEL_ID=PDC_HEAT.SEQUENCE_STEEL_ID and PD_HEAT.HEATID=PDC_HEAT.HEATID and PDC_HEAT.HEATID='" + StrHEATID + "'";
            //                         ,0            ,1                       ,2                    ,3                    ,4                 ,5                     ,6                   ,7                        ,8                        ,9                    ,10                        ,11                        ,12                      ,13                                ,14                         ,15                           ,16                 ,17                        ,18                        ,19                      ,20                   ,21             ,22                              ,23                    ,24                        ,25                   ,26  slab width          ,27 Pouring Duration       , 28  cast Speed         ,29 Lifting temp   ,30 Cut Slab Count  ,31  Tundish No       , 32 Tundish car No ,33 TC_DateTime             ,34 TundRemWeight                , 35 Net_Ls_Cast      
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERII_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string SteelId = "";//StrSTEEL_ID
                string HeatId = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                if (HeatId == "")
                {
                    HeatId = "0";
                }
                string SequenceNo = "0";
                string HeatInSeq = dtSelect.Rows[0][13].ToString();//pdc_heat.heat_seq_no 13
                if (HeatInSeq == "")
                {
                    HeatInSeq = "0";
                }
                string IsLastHeatInSeq = "0";
                string SteelGradeCode = dtSelect.Rows[0][2].ToString();//pdc_heat.grade_code   2
                if (SteelGradeCode == "")
                {
                    SteelGradeCode = "0";
                }
                string ProdOrderNo = dtSelect.Rows[0][3].ToString();//pdc_heat.prod_orderid    3 
                if (ProdOrderNo == "")
                {
                    ProdOrderNo = "0";
                }
                string CrewId = "0";
                string ShiftId = "0";
                string LadleNo = dtSelect.Rows[0][4].ToString();//pdc_heat.ladle_no 4
                if (LadleNo == "")
                {
                    LadleNo = "0";
                }
                string LadleLifeCount = "0";
                string Lancing = dtSelect.Rows[0][5].ToString();//pdc_heat.lancing_flag 5
                if (Lancing == "")
                {
                    Lancing = "0";
                }
                string ShroudType = dtSelect.Rows[0][6].ToString();//pdc_heat.shroud_type 6
                string TundPowerType = dtSelect.Rows[0][7].ToString();//pdc_heat.tund_powder_type  7
                string MidPowerType = "0";
                string StopperRodType = dtSelect.Rows[0][8].ToString();//pdc_heat.stopper_rod_type  8
                string NozzleType = dtSelect.Rows[0][9].ToString();//pdc_heat.nozzle_type 9
                string TundishNo = "0";
                string TundishCardNo = "0";
                string LadleArrival_DateTime = dtSelect.Rows[0][14].ToString().Replace(',', '.');//pdc_heat.ladle_arrival_time  14
                string LO_DateTime = dtSelect.Rows[0][11].ToString().Replace(',', '.'); ;//pdc_heat.ladle_open_time  11
                string LC_DateTime = dtSelect.Rows[0][12].ToString().Replace(',', '.'); ;//pdc_heat.ladle_close_time  12
                //string LadlePouringDur = DATEDIFF(minute,LC_DateTime,LO_DateTime)
                //string LadlePouringDur = dtSelect.Rows[0][15].ToString();//pdc_heat.ladle_open_time 11 - pdc_heat.ladle_close_time  12
                string LadleWeightNet = "0";
                string LadleRemSteelWght = "";
                string Ladle_Open_Wt = dtSelect.Rows[0][10].ToString();
                string Ladle_Close_Wt = dtSelect.Rows[0][15].ToString();
                if (Ladle_Open_Wt == "")
                {
                    Ladle_Open_Wt = "0";
                }
                if (Ladle_Close_Wt == "")
                {
                    Ladle_Close_Wt = "0";
                }
                //if (dtSelect.Rows[0][10].ToString() != "" || dtSelect.Rows[0][10].ToString() == DBNull.Value)
                //{
                //    Ladle_Open_Wt = Convert.ToInt64(dtSelect.Rows[0][10].ToString());
                //}
                //if (dtSelect.Rows[0][15].ToString() == "" || dtSelect.Rows[0][15].ToString() == null)
                //{
                //    Ladle_Close_Wt = Convert.ToInt64(dtSelect.Rows[0][15].ToString());
                //}
                string PouredHeatWeight = Convert.ToString(Convert.ToInt64(Ladle_Open_Wt) - Convert.ToInt64(Ladle_Close_Wt));//pdc_heat.ladle_open_weight 10 - pdc_heat.ladle_close_weight 15
                string TundRemWeight = "";
                string TotalSlabWeight = dtSelect.Rows[0][16].ToString();//pdc_heat.slab_weight 16
                if (TotalSlabWeight == "")
                {
                    TotalSlabWeight = "0";
                }
                string HeadWeight = dtSelect.Rows[0][17].ToString();//pdc_heat.head_crop_weight  17
                if (HeadWeight == "")
                {
                    HeadWeight = "0";
                }
                string TailWeight = dtSelect.Rows[0][18].ToString();//pdc_heat.tail_crop_weight  18
                if (TailWeight == "")
                {
                    TailWeight = "0";
                }
                string TotalCutLoss = dtSelect.Rows[0][19].ToString();//pdc_heat.cut_lost_weight  19
                if (TotalCutLoss == "")
                {
                    TotalCutLoss = "0";
                }
                string TotalSampleWeight = dtSelect.Rows[0][22].ToString();//pdc_heat.sample_lost_weight  22
                if (TotalSampleWeight == "")
                {
                    TotalSampleWeight = "0";
                }
                string SlagWeight = dtSelect.Rows[0][20].ToString();//pdc_heat.slag_weight 20
                if (SlagWeight == "")
                {
                    SlagWeight = "0";
                }
                string Yield = dtSelect.Rows[0][21].ToString();//pdc_heat.yield 21
                if (Yield == "")
                {
                    Yield = "0";
                }
                string Ld_Int_Wt = dtSelect.Rows[0][23].ToString();//pdc_heat.yield 23
                if (Ld_Int_Wt == "")
                {
                    Ld_Int_Wt = "0";
                }
                string Ld_Fin_Wt = dtSelect.Rows[0][24].ToString();//pdc_heat.yield 24
                if (Ld_Fin_Wt == "")
                {
                    Ld_Fin_Wt = "0";
                }
                string slab_length = dtSelect.Rows[0][25].ToString();//pdc_heat.slab_length 25
                if (slab_length == "")
                {
                    slab_length = "0";
                }
                string slab_Width = dtSelect.Rows[0]["AVG_SLAB_WIDTH"].ToString();//pdc_heat.AVG_slab_Width 26
                if (slab_Width == "")
                {
                    slab_Width = "0";
                }
                string StrPouringDuration = dtSelect.Rows[0]["POURING_DURATION"].ToString();//pdc_heat.POURING_DURATION 27
                if (StrPouringDuration == "")
                {
                    StrPouringDuration = "0";
                }
                string StrAVG_CAST_SPEED = dtSelect.Rows[0]["AVG_CAST_SPEED"].ToString();//pdc_heat.AVG_CAST_SPEED 28
                if (StrAVG_CAST_SPEED == "")
                {
                    StrAVG_CAST_SPEED = "0";
                }
                string StrLifting_Temp = dtSelect.Rows[0]["LAST_TEMP"].ToString();//pd_heat.LAST_TEMP 29
                if (StrLifting_Temp == "")
                {
                    StrLifting_Temp = "0";
                }
                string StrSlabCount = dtSelect.Rows[0]["SLAB_COUNT"].ToString();//pdc_heat.SLAB_COUNT 30
                if (StrSlabCount == "")
                {
                    StrSlabCount = "0";
                }
                string StrTundishNo = dtSelect.Rows[0]["TUND_SEQ_NO"].ToString();//pdc_TUNDISH.TUND_SEQ_NO 31
                if (StrTundishNo == "")
                {
                    StrTundishNo = "0";
                }
                string StrTundishCarNo = dtSelect.Rows[0]["CAR_NO"].ToString();//pdc_TUNDISH.CAR_NO 32
                if (StrTundishCarNo == "")
                {
                    StrTundishCarNo = "0";
                }
                string TC_DateTime = dtSelect.Rows[0]["TUND_CLOSE_TIME"].ToString().Replace(',', '.');// PDC_TUNDISH.TUND_CLOSE_TIME 33

                string StrTundRemWeight = dtSelect.Rows[0]["LADLE_CLOSE_TUND_WEIGHT"].ToString();//PDC_HEAT.LADLE_CLOSE_TUND_WEIGHT 34
                if (StrTundRemWeight == "")
                {
                    StrTundRemWeight = "0";
                }
                string StrNet_Ls_Cast = dtSelect.Rows[0]["CAST_WEIGHT"].ToString();//PDC_HEAT.CAST_WEIGHT 35
                if (StrNet_Ls_Cast == "")
                {
                    StrNet_Ls_Cast = "0";
                }


                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set HeatInSeq='" + HeatInSeq + "',SteelGradeCode='" + SteelGradeCode + "',ProdOrderNo='" + ProdOrderNo + "',LadleNo='" + LadleNo + "',Lancing='" + Lancing + "',ShroudType='" + ShroudType + "',TundPowerType='" + TundPowerType + "',StopperRodType='" + StopperRodType + "',NozzleType='" + NozzleType + "',LadleArrival_DateTime='" + LadleArrival_DateTime + "',LO_DateTime='" + LO_DateTime + "',LC_DateTime='" + LC_DateTime + "',LadlePouringDur=DATEDIFF(minute,LO_DateTime,LC_DateTime),TotalSlabWeight='" + TotalSlabWeight + "',HeadWeight='" + HeadWeight + "',TailWeight='" + TailWeight + "',TotalCutLoss='" + TotalCutLoss + "',TotalSampleWeight='" + TotalSampleWeight + "',SlagWeight='" + SlagWeight + "',Yield='" + Yield + "',Ld_Int_Wt='" + Ld_Int_Wt + "',Ld_Fin_Wt='" + Ld_Fin_Wt + "',Slab_Len='" + slab_length + "',Slab_Width='" + slab_Width + "',AVG_CAST_SPEED='" + StrAVG_CAST_SPEED + "',Lift_Temp='" + StrLifting_Temp + "',Ask_Temp='" + StrLifting_Temp + "',Cut_Slab_No='" + StrSlabCount + "',TundishNo='" + StrTundishNo + "',TundishCarNo='" + StrTundishCarNo + "',TC_DateTime='" + TC_DateTime + "',TundRemWeight='" + StrTundRemWeight + "',Net_Ls_Cast='" + StrNet_Ls_Cast + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }

        }
        // Insert Into HEAT_REPORT_TO_L3_SLAB(HeatId,Slab_Seq_No,MARKID_ACT,L2_SLABID,THICKNESS_HEAD,THICKNESS_TAIL,WIDTH_HEAD,WIDTH_TAIL)values()
        public void FnCASTERII_2(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT pdc_heat_shift.shift_code, pdc_heat_shift.crew_id  FROM pdc_heat_shift where pdc_heat_shift.HEAT_STEEL_ID=" + StrSTEEL_ID + "";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERII_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string StrShiftId = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                string StrCrewId = dtSelect.Rows[0][1].ToString();//pdc_heat.heatid   0

                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set ShiftId='" + StrShiftId + "',CrewId='" + StrCrewId + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }
        }
        public void FnCASTERII_3(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT pdc_tundish.tund_no, pdc_tundish.car_no, pdc_tundish.tund_close_weight,pdc_tundish.tund_close_time FROM pdc_tundish where pdc_tundish.steel_id=" + StrSTEEL_ID + "";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERII_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string StrTundishNo = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                string StrTundishCardNo = dtSelect.Rows[0][1].ToString();//pdc_heat.heatid   0
                string StrTundRemWeight = dtSelect.Rows[0][2].ToString();//pdc_heat.heatid   0
                string StrTC_DateTime = dtSelect.Rows[0][3].ToString().Replace(',', '.');//pdc_heat.heatid   0

                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set TundishNo='" + StrTundishNo + "',TundishCardNo='" + StrTundishCardNo + "',TundRemWeight='" + StrTundRemWeight + "',TC_DateTime='" + StrTC_DateTime + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }
        }
        public void FnCASTERII_4(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT CAST_POWDER_TYPE as MldPowderType FROM PDC_STRAND WHERE STEEL_ID = (SELECT STRAND_STEEL_ID FROM PDC_HEAT_STRAND WHERE HEAT_STEEL_ID = " + StrSTEEL_ID + " AND STRAND_NUM = 1)";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERII_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string StrMidPowerType = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set MidPowerType='" + StrMidPowerType + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }
        }
        public void FnCASTERII_5(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT pdc_tundish_temp.meas_no, pdc_tundish_temp.meas_time,pdc_tundish_temp.meas_temp, pdc_tundish_temp.ladle_weight,pdc_tundish_temp.tundish_weight FROM pdc_tundish_temp WHERE STEEL_ID =" + StrSTEEL_ID + "";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERII_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                for (int i = 0; i < dtSelect.Rows.Count; i++)
                {
                    string MEAS_NO = dtSelect.Rows[i][0].ToString();
                    string MEAS_TIME = dtSelect.Rows[i][1].ToString();
                    string MEAS_TEMP = dtSelect.Rows[i][2].ToString();
                    string LADLE_WEIGHT = dtSelect.Rows[i][3].ToString();
                    string TUNDISH_WEIGHT = dtSelect.Rows[i][4].ToString();

                    string StrQueryCheck = "select * from Cater2_Tundish_Temp where STEEL_ID='" + StrSTEEL_ID + "' and MEAS_NO='" + MEAS_NO + "'";
                    DataTable dtCheck = new DataTable();
                    dtCheck = clsObjMIS.DBSelectQueryMIS_Table(StrQueryCheck);
                    if (dtCheck.Rows.Count == 0)
                    {
                        string StrMISInsetQuery = "Insert Into  Cater2_Tundish_Temp (STEEL_ID,MEAS_NO,MEAS_TIME,MEAS_TEMP,LADLE_WEIGHT,TUNDISH_WEIGHT) values('" + StrSTEEL_ID + "','" + MEAS_NO + "','" + MEAS_TIME + "','" + MEAS_TEMP + "','" + LADLE_WEIGHT + "','" + TUNDISH_WEIGHT + "')";
                        bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);
                    }
                }
            }
        }

        public void FnCASTERIII_1(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQueryMIS = "select SteelID from HEAT_REPORT_TO_L3 where SteelID='" + StrSTEEL_ID + "'";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtMIS = new DataTable();
            dtMIS = clsObj.DBSelectQueryMIS_Table(StrQueryMIS);
            if (dtMIS.Rows.Count == 0)
            {

                string StrMISInsetQuery = "Insert Into HEAT_REPORT_TO_L3 (SteelID,HeatID) values('" + StrSTEEL_ID + "','" + StrHEATID + "')";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }

            //string StrQuery = "SELECT pdc_heat.heatid, pdc_heat.last_heat_flag, pdc_heat.grade_code,pdc_heat.prod_orderid, pdc_heat.ladle_no, pdc_heat.lancing_flag,pdc_heat.shroud_type, pdc_heat.tund_powder_type,pdc_heat.stopper_rod_type, pdc_heat.nozzle_type,pdc_heat.ladle_open_weight, pdc_heat.ladle_open_time,pdc_heat.ladle_close_time, pdc_heat.heat_seq_no as HeatInSeq,pdc_heat.ladle_arrival_time, pdc_heat.ladle_close_weight,pdc_heat.slab_weight, pdc_heat.head_crop_weight,pdc_heat.tail_crop_weight, pdc_heat.cut_lost_weight,pdc_heat.slag_weight, pdc_heat.yield, pdc_heat.sample_lost_weight,pdc_heat.ladle_open_weight,pdc_heat.ladle_close_weight,pdc_heat.slab_length,PDC_HEAT.AVG_SLAB_WIDTH ,PDC_HEAT.POURING_DURATION,PDC_STRAND.Avg_cast_speed,PD_HEAT.LAST_TEMP,PDC_HEAT.Slab_COUNT,PDC_TUNDISH.TUND_SEQ_NO,PDC_TUNDISH.CAR_NO,PDC_TUNDISH.TUND_CLOSE_TIME,PDC_HEAT.LADLE_CLOSE_TUND_WEIGHT,PDC_HEAT.CAST_WEIGHT FROM pdc_heat,pdc_heat_strand,pdc_strand ,PD_HEAT ,PDC_TUNDISH WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID and pdc_heat_strand.STRAND_STEEL_ID=pdc_strand.STEEL_ID and PDC_TUNDISH.SEQUENCE_STEEL_ID=PDC_HEAT.SEQUENCE_STEEL_ID and PD_HEAT.HEATID=PDC_HEAT.HEATID and PDC_HEAT.HEATID='" + StrHEATID + "'";
            string StrQuery = "SELECT pdc_heat.heatid, pdc_heat.last_heat_flag, pdc_heat.grade_code,pdc_heat.prod_orderid, pdc_heat.ladle_no, pdc_heat.lancing_flag,pdc_heat.shroud_type, pdc_heat.tund_powder_type,pdc_heat.stopper_rod_type, pdc_heat.nozzle_type,pdc_heat.ladle_open_weight, pdc_heat.ladle_open_time,pdc_heat.ladle_close_time, pdc_heat.heat_seq_no as HeatInSeq,pdc_heat.ladle_arrival_time, pdc_heat.ladle_close_weight,pdc_heat.slab_weight, pdc_heat.head_crop_weight,pdc_heat.tail_crop_weight, pdc_heat.cut_lost_weight,pdc_heat.slag_weight, pdc_heat.yield, pdc_heat.sample_lost_weight,pdc_heat.ladle_open_weight,pdc_heat.ladle_close_weight,pdc_heat.slab_length,PDC_HEAT.AVG_SLAB_WIDTH ,PDC_HEAT.POURING_DURATION,PDC_STRAND.Avg_cast_speed,PD_HEAT.LAST_TEMP,PDC_HEAT.Slab_COUNT,PDC_TUNDISH.TUND_SEQ_NO,PDC_TUNDISH.CAR_NO,PDC_TUNDISH.TUND_CLOSE_TIME,PDC_HEAT.LADLE_CLOSE_TUND_WEIGHT,PDC_HEAT.CAST_WEIGHT FROM pdc_heat,pdc_heat_strand,pdc_strand ,PD_HEAT ,PDC_TUNDISH WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID and pdc_heat_strand.STRAND_STEEL_ID=pdc_strand.STEEL_ID and PDC_TUNDISH.SEQUENCE_STEEL_ID=PDC_HEAT.SEQUENCE_STEEL_ID and PD_HEAT.HEATID=PDC_HEAT.HEATID and PDC_HEAT.HEATID='" + StrHEATID + "'";
            //                         ,0            ,1                       ,2                    ,3                    ,4                 ,5                     ,6                   ,7                        ,8                        ,9                    ,10                        ,11                        ,12                      ,13                                ,14                         ,15                           ,16                 ,17                        ,18                        ,19                      ,20                   ,21             ,22                              ,23                    ,24                        ,25                   ,26  slab width          ,27 Pouring Duration       , 28  cast Speed         ,29 Lifting temp   ,30 Cut Slab Count  ,31  Tundish No       , 32 Tundish car No ,33 TC_DateTime             ,34 TundRemWeight                , 35 Net_Ls_Cast      
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERIII_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string SteelId = "";//StrSTEEL_ID
                string HeatId = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                if (HeatId == "")
                {
                    HeatId = "0";
                }
                string SequenceNo = "0";
                string HeatInSeq = dtSelect.Rows[0][13].ToString();//pdc_heat.heat_seq_no 13
                if (HeatInSeq == "")
                {
                    HeatInSeq = "0";
                }
                string IsLastHeatInSeq = "0";
                string SteelGradeCode = dtSelect.Rows[0][2].ToString();//pdc_heat.grade_code   2
                if (SteelGradeCode == "")
                {
                    SteelGradeCode = "0";
                }
                string ProdOrderNo = dtSelect.Rows[0][3].ToString();//pdc_heat.prod_orderid    3 
                if (ProdOrderNo == "")
                {
                    ProdOrderNo = "0";
                }
                string CrewId = "0";
                string ShiftId = "0";
                string LadleNo = dtSelect.Rows[0][4].ToString();//pdc_heat.ladle_no 4
                if (LadleNo == "")
                {
                    LadleNo = "0";
                }
                string LadleLifeCount = "0";
                string Lancing = dtSelect.Rows[0][5].ToString();//pdc_heat.lancing_flag 5
                if (Lancing == "")
                {
                    Lancing = "0";
                }
                string ShroudType = dtSelect.Rows[0][6].ToString();//pdc_heat.shroud_type 6
                string TundPowerType = dtSelect.Rows[0][7].ToString();//pdc_heat.tund_powder_type  7
                string MidPowerType = "0";
                string StopperRodType = dtSelect.Rows[0][8].ToString();//pdc_heat.stopper_rod_type  8
                string NozzleType = dtSelect.Rows[0][9].ToString();//pdc_heat.nozzle_type 9
                string TundishNo = "0";
                string TundishCardNo = "0";
                string LadleArrival_DateTime = dtSelect.Rows[0][14].ToString().Replace(',', '.');//pdc_heat.ladle_arrival_time  14
                string LO_DateTime = dtSelect.Rows[0][11].ToString().Replace(',', '.'); ;//pdc_heat.ladle_open_time  11
                string LC_DateTime = dtSelect.Rows[0][12].ToString().Replace(',', '.'); ;//pdc_heat.ladle_close_time  12
                //string LadlePouringDur = DATEDIFF(minute,LC_DateTime,LO_DateTime)
                //string LadlePouringDur = dtSelect.Rows[0][15].ToString();//pdc_heat.ladle_open_time 11 - pdc_heat.ladle_close_time  12
                string LadleWeightNet = "0";
                string LadleRemSteelWght = "";
                string Ladle_Open_Wt = dtSelect.Rows[0][10].ToString();
                string Ladle_Close_Wt = dtSelect.Rows[0][15].ToString();
                if (Ladle_Open_Wt == "")
                {
                    Ladle_Open_Wt = "0";
                }
                if (Ladle_Close_Wt == "")
                {
                    Ladle_Close_Wt = "0";
                }
                //if (dtSelect.Rows[0][10].ToString() != "" || dtSelect.Rows[0][10].ToString() == DBNull.Value)
                //{
                //    Ladle_Open_Wt = Convert.ToInt64(dtSelect.Rows[0][10].ToString());
                //}
                //if (dtSelect.Rows[0][15].ToString() == "" || dtSelect.Rows[0][15].ToString() == null)
                //{
                //    Ladle_Close_Wt = Convert.ToInt64(dtSelect.Rows[0][15].ToString());
                //}
                string PouredHeatWeight = Convert.ToString(Convert.ToInt64(Ladle_Open_Wt) - Convert.ToInt64(Ladle_Close_Wt));//pdc_heat.ladle_open_weight 10 - pdc_heat.ladle_close_weight 15
                string TundRemWeight = "";
                string TotalSlabWeight = dtSelect.Rows[0][16].ToString();//pdc_heat.slab_weight 16
                if (TotalSlabWeight == "")
                {
                    TotalSlabWeight = "0";
                }
                string HeadWeight = dtSelect.Rows[0][17].ToString();//pdc_heat.head_crop_weight  17
                if (HeadWeight == "")
                {
                    HeadWeight = "0";
                }
                string TailWeight = dtSelect.Rows[0][18].ToString();//pdc_heat.tail_crop_weight  18
                if (TailWeight == "")
                {
                    TailWeight = "0";
                }
                string TotalCutLoss = dtSelect.Rows[0][19].ToString();//pdc_heat.cut_lost_weight  19
                if (TotalCutLoss == "")
                {
                    TotalCutLoss = "0";
                }
                string TotalSampleWeight = dtSelect.Rows[0][22].ToString();//pdc_heat.sample_lost_weight  22
                if (TotalSampleWeight == "")
                {
                    TotalSampleWeight = "0";
                }
                string SlagWeight = dtSelect.Rows[0][20].ToString();//pdc_heat.slag_weight 20
                if (SlagWeight == "")
                {
                    SlagWeight = "0";
                }
                string Yield = dtSelect.Rows[0][21].ToString();//pdc_heat.yield 21
                if (Yield == "")
                {
                    Yield = "0";
                }
                string Ld_Int_Wt = dtSelect.Rows[0][23].ToString();//pdc_heat.yield 23
                if (Ld_Int_Wt == "")
                {
                    Ld_Int_Wt = "0";
                }
                string Ld_Fin_Wt = dtSelect.Rows[0][24].ToString();//pdc_heat.yield 24
                if (Ld_Fin_Wt == "")
                {
                    Ld_Fin_Wt = "0";
                }
                string slab_length = dtSelect.Rows[0][25].ToString();//pdc_heat.slab_length 25
                if (slab_length == "")
                {
                    slab_length = "0";
                }
                string slab_Width = dtSelect.Rows[0]["AVG_SLAB_WIDTH"].ToString();//pdc_heat.AVG_slab_Width 26
                if (slab_Width == "")
                {
                    slab_Width = "0";
                }
                string StrPouringDuration = dtSelect.Rows[0]["POURING_DURATION"].ToString();//pdc_heat.POURING_DURATION 27
                if (StrPouringDuration == "")
                {
                    StrPouringDuration = "0";
                }
                string StrAVG_CAST_SPEED = dtSelect.Rows[0]["AVG_CAST_SPEED"].ToString();//pdc_heat.AVG_CAST_SPEED 28
                if (StrAVG_CAST_SPEED == "")
                {
                    StrAVG_CAST_SPEED = "0";
                }
                string StrLifting_Temp = dtSelect.Rows[0]["LAST_TEMP"].ToString();//pd_heat.LAST_TEMP 29
                if (StrLifting_Temp == "")
                {
                    StrLifting_Temp = "0";
                }
                string StrSlabCount = dtSelect.Rows[0]["SLAB_COUNT"].ToString();//pdc_heat.SLAB_COUNT 30
                if (StrSlabCount == "")
                {
                    StrSlabCount = "0";
                }
                string StrTundishNo = dtSelect.Rows[0]["TUND_SEQ_NO"].ToString();//pdc_TUNDISH.TUND_SEQ_NO 31
                if (StrTundishNo == "")
                {
                    StrTundishNo = "0";
                }
                string StrTundishCarNo = dtSelect.Rows[0]["CAR_NO"].ToString();//pdc_TUNDISH.CAR_NO 32
                if (StrTundishCarNo == "")
                {
                    StrTundishCarNo = "0";
                }
                string TC_DateTime = dtSelect.Rows[0]["TUND_CLOSE_TIME"].ToString().Replace(',', '.');// PDC_TUNDISH.TUND_CLOSE_TIME 33

                string StrTundRemWeight = dtSelect.Rows[0]["LADLE_CLOSE_TUND_WEIGHT"].ToString();//PDC_HEAT.LADLE_CLOSE_TUND_WEIGHT 34
                if (StrTundRemWeight == "")
                {
                    StrTundRemWeight = "0";
                }
                string StrNet_Ls_Cast = dtSelect.Rows[0]["CAST_WEIGHT"].ToString();//PDC_HEAT.CAST_WEIGHT 35
                if (StrNet_Ls_Cast == "")
                {
                    StrNet_Ls_Cast = "0";
                }


                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set HeatInSeq='" + HeatInSeq + "',SteelGradeCode='" + SteelGradeCode + "',ProdOrderNo='" + ProdOrderNo + "',LadleNo='" + LadleNo + "',Lancing='" + Lancing + "',ShroudType='" + ShroudType + "',TundPowerType='" + TundPowerType + "',StopperRodType='" + StopperRodType + "',NozzleType='" + NozzleType + "',LadleArrival_DateTime='" + LadleArrival_DateTime + "',LO_DateTime='" + LO_DateTime + "',LC_DateTime='" + LC_DateTime + "',LadlePouringDur=DATEDIFF(minute,LO_DateTime,LC_DateTime),TotalSlabWeight='" + TotalSlabWeight + "',HeadWeight='" + HeadWeight + "',TailWeight='" + TailWeight + "',TotalCutLoss='" + TotalCutLoss + "',TotalSampleWeight='" + TotalSampleWeight + "',SlagWeight='" + SlagWeight + "',Yield='" + Yield + "',Ld_Int_Wt='" + Ld_Int_Wt + "',Ld_Fin_Wt='" + Ld_Fin_Wt + "',Slab_Len='" + slab_length + "',Slab_Width='" + slab_Width + "',AVG_CAST_SPEED='" + StrAVG_CAST_SPEED + "',Lift_Temp='" + StrLifting_Temp + "',Ask_Temp='" + StrLifting_Temp + "',Cut_Slab_No='" + StrSlabCount + "',TundishNo='" + StrTundishNo + "',TundishCarNo='" + StrTundishCarNo + "',TC_DateTime='" + TC_DateTime + "',TundRemWeight='" + StrTundRemWeight + "',Net_Ls_Cast='" + StrNet_Ls_Cast + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);



            }

        }
        public void FnCASTERIII_2(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT pdc_heat_shift.shift_code, pdc_heat_shift.crew_id  FROM pdc_heat_shift where pdc_heat_shift.HEAT_STEEL_ID=" + StrSTEEL_ID + "";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERIII_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string StrShiftId = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                string StrCrewId = dtSelect.Rows[0][1].ToString();//pdc_heat.heatid   0

                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set ShiftId='" + StrShiftId + "',CrewId='" + StrCrewId + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }
        }
        public void FnCASTERIII_3(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT pdc_tundish.tund_no, pdc_tundish.car_no, pdc_tundish.tund_close_weight,pdc_tundish.tund_close_time FROM pdc_tundish where pdc_tundish.steel_id=" + StrSTEEL_ID + "";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERIII_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string StrTundishNo = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                string StrTundishCardNo = dtSelect.Rows[0][1].ToString();//pdc_heat.heatid   0
                string StrTundRemWeight = dtSelect.Rows[0][2].ToString();//pdc_heat.heatid   0
                string StrTC_DateTime = dtSelect.Rows[0][3].ToString().Replace(',', '.');//pdc_heat.heatid   0

                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set TundishNo='" + StrTundishNo + "',TundishCardNo='" + StrTundishCardNo + "',TundRemWeight='" + StrTundRemWeight + "',TC_DateTime='" + StrTC_DateTime + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }
        }
        public void FnCASTERIII_4(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT CAST_POWDER_TYPE as MldPowderType FROM PDC_STRAND WHERE STEEL_ID = (SELECT STRAND_STEEL_ID FROM PDC_HEAT_STRAND WHERE HEAT_STEEL_ID = " + StrSTEEL_ID + " AND STRAND_NUM = 1)";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERIII_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                string StrMidPowerType = dtSelect.Rows[0][0].ToString();//pdc_heat.heatid   0
                string StrMISInsetQuery = "Update HEAT_REPORT_TO_L3 set MidPowerType='" + StrMidPowerType + "' where SteelID='" + StrSTEEL_ID + "'";
                bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);

            }
        }
        public void FnCASTERIII_5(string StrSTEEL_ID, string StrHEATID)
        {
            string StrQuery = "SELECT pdc_tundish_temp.meas_no, pdc_tundish_temp.meas_time,pdc_tundish_temp.meas_temp, pdc_tundish_temp.ladle_weight,pdc_tundish_temp.tundish_weight FROM pdc_tundish_temp WHERE STEEL_ID ='" + StrSTEEL_ID + "'";
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            dtSelect = clsObj.DBSelectQueryCASTERIII_Table(StrQuery);
            if (dtSelect.Rows.Count > 0)
            {
                for (int i = 0; i < dtSelect.Rows.Count; i++)
                {
                    string MEAS_NO = dtSelect.Rows[i][0].ToString();
                    string MEAS_TIME = dtSelect.Rows[i][1].ToString();
                    string MEAS_TEMP = dtSelect.Rows[i][2].ToString();
                    string LADLE_WEIGHT = dtSelect.Rows[i][3].ToString();
                    string TUNDISH_WEIGHT = dtSelect.Rows[i][4].ToString();

                    string StrQueryCheck = "select * from Cater2_Tundish_Temp where STEEL_ID='" + StrSTEEL_ID + "' and MEAS_NO='" + MEAS_NO + "'";
                    DataTable dtCheck = new DataTable();
                    dtCheck = clsObjMIS.DBSelectQueryMIS_Table(StrQueryCheck);
                    if (dtCheck.Rows.Count == 0)
                    {
                        string StrMISInsetQuery = "Insert Into  Cater2_Tundish_Temp (STEEL_ID,MEAS_NO,MEAS_TIME,MEAS_TEMP,LADLE_WEIGHT,TUNDISH_WEIGHT) values('" + StrSTEEL_ID + "','" + MEAS_NO + "','" + MEAS_TIME + "','" + MEAS_TEMP + "','" + LADLE_WEIGHT + "','" + TUNDISH_WEIGHT + "')";
                        bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);
                    }
                }
            }
        }

        public void Fn_HEAT_REPORT_TO_L3_SLAB_CASTERI(string StrHEATID)
        {
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            string StrMax_Slab_Seq_No = "select ISNULL(MAX(Slab_Seq_No),0)'Slab_Seq_No' from Heat_Report_to_l3_SLAB where HeatId ='" + StrHEATID + "'";
            DataTable dtCheck = new DataTable();
            dtCheck = clsObjMIS.DBSelectQueryMIS_Table(StrMax_Slab_Seq_No);
            int Max_Slab_Seq_No = 0;
            if (dtCheck.Rows.Count > 0)
            {
                try
                {
                    Max_Slab_Seq_No = Convert.ToInt16(dtCheck.Rows[0][0].ToString());
                }
                catch
                {
                }
            }
            clsObjMIS = new DBConnections();
            dtSelect = new DataTable();
            string StrQuery = "select PDC_HEAT.HEATID,PDC_SLAB.SLAB_SEQ_NO,PDC_SLAB.MARKID_ACT,PDC_SLAB.L2_SLABID,PDC_SLAB_ORDER.THICKNESS_HEAD,PDC_SLAB_ORDER.THICKNESS_TAIL ,PDC_SLAB_ORDER.WIDTH_HEAD,PDC_SLAB_ORDER.WIDTH_TAIL from PDC_SLAB,PDC_SLAB_ORDER,PDC_HEAT where PDC_SLAB.PROD_ORDERID=PDC_HEAT.PROD_ORDERID AND PDC_SLAB.STEEL_ID=PDC_SLAB_ORDER.STEEL_ID AND PDC_HEAT.HEATID='" + StrHEATID + "'  AND PDC_SLAB.SLAB_SEQ_NO>'" + Max_Slab_Seq_No + "'";
            dtSelect = clsObj.DBSelectQueryCASTERI_Table(StrQuery);
            string StrMISInsetQuery = "";
            if (dtSelect.Rows.Count > 0)
            {
                for (int i = 0; i < dtSelect.Rows.Count; i++)
                {
                    string HEATID = dtSelect.Rows[i][0].ToString();
                    string SLAB_SEQ_NO = dtSelect.Rows[i][1].ToString();
                    string MARKID_ACT = dtSelect.Rows[i][2].ToString();
                    string L2_SLABID = dtSelect.Rows[i][3].ToString();
                    string THICKNESS_HEAD = dtSelect.Rows[i][4].ToString();
                    string THICKNESS_TAIL = dtSelect.Rows[i][5].ToString();
                    string WIDTH_HEAD = dtSelect.Rows[i][6].ToString();
                    string WIDTH_TAIL = dtSelect.Rows[i][7].ToString();
                    StrMISInsetQuery = "Insert Into  Heat_Report_to_l3_SLAB (HeatId,Slab_Seq_No,MARKID_ACT,L2_SLABID,THICKNESS_HEAD,THICKNESS_TAIL,WIDTH_HEAD,WIDTH_TAIL) values('" + HEATID + "','" + SLAB_SEQ_NO + "','" + MARKID_ACT + "','" + L2_SLABID + "','" + THICKNESS_HEAD + "','" + THICKNESS_TAIL + "','" + WIDTH_HEAD + "','" + WIDTH_TAIL + "')";
                    bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);
                }
            } // @ +Manish 16052014 for updating mis table.
            clsObjMIS = new DBConnections();
            StrMax_Slab_Seq_No = "select ISNULL(Slab_Seq_No,0)'Slab_Seq_No' from Heat_Report_to_l3_SLAB where HeatId ='" + StrHEATID + "' And MARKID_ACT = ''";
            DataTable dtCheck1 = new DataTable();
            dtCheck1 = clsObjMIS.DBSelectQueryMIS_Table(StrMax_Slab_Seq_No);
            if (dtCheck1.Rows.Count > 0)
            {
                for (int i = 0; i < dtCheck1.Rows.Count; i++)
                {
                    string Slab_Seq_No = dtCheck1.Rows[i][0].ToString();
                    clsObjMIS = new DBConnections();
                    dtSelect = new DataTable();
                    string StrQueryupdate = "select PDC_HEAT.HEATID,PDC_SLAB.SLAB_SEQ_NO,PDC_SLAB.MARKID_ACT,PDC_SLAB.L2_SLABID,PDC_SLAB_ORDER.THICKNESS_HEAD,PDC_SLAB_ORDER.THICKNESS_TAIL ,PDC_SLAB_ORDER.WIDTH_HEAD,PDC_SLAB_ORDER.WIDTH_TAIL from PDC_SLAB,PDC_SLAB_ORDER,PDC_HEAT where PDC_SLAB.PROD_ORDERID=PDC_HEAT.PROD_ORDERID AND PDC_SLAB.STEEL_ID=PDC_SLAB_ORDER.STEEL_ID AND PDC_HEAT.HEATID='" + StrHEATID + "'  AND PDC_SLAB.SLAB_SEQ_NO ='" + Slab_Seq_No + "'";
                    dtSelect = clsObj.DBSelectQueryCASTERI_Table(StrQueryupdate);
                    string StrMISUpdateQuery = "";
                    if (dtSelect.Rows.Count > 0)
                    {
                        for (int j = 0; j < dtSelect.Rows.Count; j++)
                        {
                            string HEATID = dtSelect.Rows[j][0].ToString();
                            string SLAB_SEQ_NO = dtSelect.Rows[j][1].ToString();
                            string MARKID_ACT = dtSelect.Rows[j][2].ToString();
                            string L2_SLABID = dtSelect.Rows[j][3].ToString();
                            string THICKNESS_HEAD = dtSelect.Rows[j][4].ToString();
                            string THICKNESS_TAIL = dtSelect.Rows[j][5].ToString();
                            string WIDTH_HEAD = dtSelect.Rows[j][6].ToString();
                            string WIDTH_TAIL = dtSelect.Rows[j][7].ToString();
                            StrMISUpdateQuery = "Update Heat_Report_to_l3_SLAB set MARKID_ACT='" + MARKID_ACT + "',L2_SLABID='" + L2_SLABID + "',THICKNESS_HEAD='" + THICKNESS_HEAD + "',THICKNESS_TAIL='" + THICKNESS_TAIL + "',WIDTH_HEAD='" + WIDTH_HEAD + "',WIDTH_TAIL='" + WIDTH_TAIL + "' where HeatId='" + HEATID + "' and SLAB_SEQ_NO='" + SLAB_SEQ_NO + "'";
                            bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISUpdateQuery);
                        }
                    }

                }
            }   // @ -Manish @ BSl 
        }
        public void Fn_HEAT_REPORT_TO_L3_SLAB_CASTERII(string StrHEATID)
        {
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            string StrMax_Slab_Seq_No = "select ISNULL(MAX(Slab_Seq_No),0)'Slab_Seq_No' from Heat_Report_to_l3_SLAB where HeatId ='" + StrHEATID + "'";
            DataTable dtCheck = new DataTable();
            dtCheck = clsObjMIS.DBSelectQueryMIS_Table(StrMax_Slab_Seq_No);
            int Max_Slab_Seq_No = 0;
            if (dtCheck.Rows.Count > 0)
            {
                try
                {
                    Max_Slab_Seq_No = Convert.ToInt16(dtCheck.Rows[0][0].ToString());
                }
                catch
                {
                }
            }

            clsObjMIS = new DBConnections();
            dtSelect = new DataTable();
            string StrQuery = "select PDC_HEAT.HEATID,PDC_SLAB.SLAB_SEQ_NO,PDC_SLAB.MARKID_ACT,PDC_SLAB.L2_SLABID,PDC_SLAB_ORDER.THICKNESS_HEAD,PDC_SLAB_ORDER.THICKNESS_TAIL ,PDC_SLAB_ORDER.WIDTH_HEAD,PDC_SLAB_ORDER.WIDTH_TAIL from PDC_SLAB,PDC_SLAB_ORDER,PDC_HEAT where PDC_SLAB.PROD_ORDERID=PDC_HEAT.PROD_ORDERID AND PDC_SLAB.STEEL_ID=PDC_SLAB_ORDER.STEEL_ID AND PDC_HEAT.HEATID='" + StrHEATID + "'  AND PDC_SLAB.SLAB_SEQ_NO>'" + Max_Slab_Seq_No + "'";
            dtSelect = clsObj.DBSelectQueryCASTERII_Table(StrQuery);
            string StrMISInsetQuery = "";
            if (dtSelect.Rows.Count > 0)
            {
                for (int i = 0; i < dtSelect.Rows.Count; i++)
                {
                    string HEATID = dtSelect.Rows[i][0].ToString();
                    string SLAB_SEQ_NO = dtSelect.Rows[i][1].ToString();
                    string MARKID_ACT = dtSelect.Rows[i][2].ToString();
                    string L2_SLABID = dtSelect.Rows[i][3].ToString();
                    string THICKNESS_HEAD = dtSelect.Rows[i][4].ToString();
                    string THICKNESS_TAIL = dtSelect.Rows[i][5].ToString();
                    string WIDTH_HEAD = dtSelect.Rows[i][6].ToString();
                    string WIDTH_TAIL = dtSelect.Rows[i][7].ToString();
                    StrMISInsetQuery = "Insert Into  Heat_Report_to_l3_SLAB (HeatId,Slab_Seq_No,MARKID_ACT,L2_SLABID,THICKNESS_HEAD,THICKNESS_TAIL,WIDTH_HEAD,WIDTH_TAIL) values('" + HEATID + "','" + SLAB_SEQ_NO + "','" + MARKID_ACT + "','" + L2_SLABID + "','" + THICKNESS_HEAD + "','" + THICKNESS_TAIL + "','" + WIDTH_HEAD + "','" + WIDTH_TAIL + "')";
                    bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);
                }
            }
            // +Manish @BSl for updating table with Actula Marking 16052014
            clsObjMIS = new DBConnections();
            StrMax_Slab_Seq_No = "select ISNULL(Slab_Seq_No,0)'Slab_Seq_No' from Heat_Report_to_l3_SLAB where HeatId ='" + StrHEATID + "' And MARKID_ACT = ''";
            DataTable dtCheck1 = new DataTable();
            dtCheck1 = clsObjMIS.DBSelectQueryMIS_Table(StrMax_Slab_Seq_No);
            if (dtCheck1.Rows.Count > 0)
            {
                for (int i = 0; i < dtCheck1.Rows.Count; i++)
                {
                    string Slab_Seq_No = dtCheck1.Rows[i][0].ToString();
                    clsObjMIS = new DBConnections();
                    dtSelect = new DataTable();
                    string StrQueryupdate = "select PDC_HEAT.HEATID,PDC_SLAB.SLAB_SEQ_NO,PDC_SLAB.MARKID_ACT,PDC_SLAB.L2_SLABID,PDC_SLAB_ORDER.THICKNESS_HEAD,PDC_SLAB_ORDER.THICKNESS_TAIL ,PDC_SLAB_ORDER.WIDTH_HEAD,PDC_SLAB_ORDER.WIDTH_TAIL from PDC_SLAB,PDC_SLAB_ORDER,PDC_HEAT where PDC_SLAB.PROD_ORDERID=PDC_HEAT.PROD_ORDERID AND PDC_SLAB.STEEL_ID=PDC_SLAB_ORDER.STEEL_ID AND PDC_HEAT.HEATID='" + StrHEATID + "'  AND PDC_SLAB.SLAB_SEQ_NO ='" + Slab_Seq_No + "'";
                    dtSelect = clsObj.DBSelectQueryCASTERII_Table(StrQueryupdate);
                    string StrMISUpdateQuery = "";
                    if (dtSelect.Rows.Count > 0)
                    {
                        for (int j = 0; j < dtSelect.Rows.Count; j++)
                        {
                            string HEATID = dtSelect.Rows[j][0].ToString();
                            string SLAB_SEQ_NO = dtSelect.Rows[j][1].ToString();
                            string MARKID_ACT = dtSelect.Rows[j][2].ToString();
                            string L2_SLABID = dtSelect.Rows[j][3].ToString();
                            string THICKNESS_HEAD = dtSelect.Rows[j][4].ToString();
                            string THICKNESS_TAIL = dtSelect.Rows[j][5].ToString();
                            string WIDTH_HEAD = dtSelect.Rows[j][6].ToString();
                            string WIDTH_TAIL = dtSelect.Rows[j][7].ToString();
                            StrMISUpdateQuery = "Update Heat_Report_to_l3_SLAB set MARKID_ACT='" + MARKID_ACT + "',L2_SLABID='" + L2_SLABID + "',THICKNESS_HEAD='" + THICKNESS_HEAD + "',THICKNESS_TAIL='" + THICKNESS_TAIL + "',WIDTH_HEAD='" + WIDTH_HEAD + "',WIDTH_TAIL='" + WIDTH_TAIL + "' where HeatId='" + HEATID + "' and SLAB_SEQ_NO='" + SLAB_SEQ_NO + "'";
                            bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISUpdateQuery);
                        }
                    }

                }
            }  // - Manish @ BSL 

        }
        public void Fn_HEAT_REPORT_TO_L3_SLAB_CASTERIII(string StrHEATID)
        {
            DBConnections clsObjMIS = new DBConnections();
            DataTable dtSelect = new DataTable();
            string StrMax_Slab_Seq_No = "select ISNULL(MAX(Slab_Seq_No),0)'Slab_Seq_No' from Heat_Report_to_l3_SLAB where HeatId ='" + StrHEATID + "'";
            DataTable dtCheck = new DataTable();
            dtCheck = clsObjMIS.DBSelectQueryMIS_Table(StrMax_Slab_Seq_No);
            int Max_Slab_Seq_No = 0;
            if (dtCheck.Rows.Count > 0)
            {
                try
                {
                    Max_Slab_Seq_No = Convert.ToInt16(dtCheck.Rows[0][0].ToString());
                }
                catch
                {
                }
            }

            clsObjMIS = new DBConnections();
            dtSelect = new DataTable();
            string StrQuery = "select PDC_HEAT.HEATID,PDC_SLAB.SLAB_SEQ_NO,PDC_SLAB.MARKID_ACT,PDC_SLAB.L2_SLABID,PDC_SLAB_ORDER.THICKNESS_HEAD,PDC_SLAB_ORDER.THICKNESS_TAIL ,PDC_SLAB_ORDER.WIDTH_HEAD,PDC_SLAB_ORDER.WIDTH_TAIL from PDC_SLAB,PDC_SLAB_ORDER,PDC_HEAT where PDC_SLAB.PROD_ORDERID=PDC_HEAT.PROD_ORDERID AND PDC_SLAB.STEEL_ID=PDC_SLAB_ORDER.STEEL_ID AND PDC_HEAT.HEATID='" + StrHEATID + "'  AND PDC_SLAB.SLAB_SEQ_NO>'" + Max_Slab_Seq_No + "'";
            dtSelect = clsObj.DBSelectQueryCASTERIII_Table(StrQuery);
            string StrMISInsetQuery = "";
            if (dtSelect.Rows.Count > 0)
            {
                for (int i = 0; i < dtSelect.Rows.Count; i++)
                {
                    string HEATID = dtSelect.Rows[i][0].ToString();
                    string SLAB_SEQ_NO = dtSelect.Rows[i][1].ToString();
                    string MARKID_ACT = dtSelect.Rows[i][2].ToString();
                    string L2_SLABID = dtSelect.Rows[i][3].ToString();
                    string THICKNESS_HEAD = dtSelect.Rows[i][4].ToString();
                    string THICKNESS_TAIL = dtSelect.Rows[i][5].ToString();
                    string WIDTH_HEAD = dtSelect.Rows[i][6].ToString();
                    string WIDTH_TAIL = dtSelect.Rows[i][7].ToString();
                    StrMISInsetQuery = "Insert Into  Heat_Report_to_l3_SLAB (HeatId,Slab_Seq_No,MARKID_ACT,L2_SLABID,THICKNESS_HEAD,THICKNESS_TAIL,WIDTH_HEAD,WIDTH_TAIL) values('" + HEATID + "','" + SLAB_SEQ_NO + "','" + MARKID_ACT + "','" + L2_SLABID + "','" + THICKNESS_HEAD + "','" + THICKNESS_TAIL + "','" + WIDTH_HEAD + "','" + WIDTH_TAIL + "')";
                    bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISInsetQuery);
                }
            }
            // +manish for updating actual marking string in the mis table date 16052014
            clsObjMIS = new DBConnections();
            StrMax_Slab_Seq_No = "select ISNULL(Slab_Seq_No,0)'Slab_Seq_No' from Heat_Report_to_l3_SLAB where HeatId ='" + StrHEATID + "' And MARKID_ACT = ''";
            DataTable dtCheck1 = new DataTable();
            dtCheck1 = clsObjMIS.DBSelectQueryMIS_Table(StrMax_Slab_Seq_No);
            if (dtCheck1.Rows.Count > 0)
            {
                for (int i = 0; i < dtCheck1.Rows.Count; i++)
                {
                    string Slab_Seq_No = dtCheck1.Rows[i][0].ToString();
                    clsObjMIS = new DBConnections();
                    dtSelect = new DataTable();
                    string StrQueryupdate = "select PDC_HEAT.HEATID,PDC_SLAB.SLAB_SEQ_NO,PDC_SLAB.MARKID_ACT,PDC_SLAB.L2_SLABID,PDC_SLAB_ORDER.THICKNESS_HEAD,PDC_SLAB_ORDER.THICKNESS_TAIL ,PDC_SLAB_ORDER.WIDTH_HEAD,PDC_SLAB_ORDER.WIDTH_TAIL from PDC_SLAB,PDC_SLAB_ORDER,PDC_HEAT where PDC_SLAB.PROD_ORDERID=PDC_HEAT.PROD_ORDERID AND PDC_SLAB.STEEL_ID=PDC_SLAB_ORDER.STEEL_ID AND PDC_HEAT.HEATID='" + StrHEATID + "'  AND PDC_SLAB.SLAB_SEQ_NO ='" + Slab_Seq_No + "'";
                    dtSelect = clsObj.DBSelectQueryCASTERIII_Table(StrQueryupdate);
                    string StrMISUpdateQuery = "";
                    if (dtSelect.Rows.Count > 0)
                    {
                        for (int j = 0; j < dtSelect.Rows.Count; j++)
                        {
                            string HEATID = dtSelect.Rows[j][0].ToString();
                            string SLAB_SEQ_NO = dtSelect.Rows[j][1].ToString();
                            string MARKID_ACT = dtSelect.Rows[j][2].ToString();
                            string L2_SLABID = dtSelect.Rows[j][3].ToString();
                            string THICKNESS_HEAD = dtSelect.Rows[j][4].ToString();
                            string THICKNESS_TAIL = dtSelect.Rows[j][5].ToString();
                            string WIDTH_HEAD = dtSelect.Rows[j][6].ToString();
                            string WIDTH_TAIL = dtSelect.Rows[j][7].ToString();
                            StrMISUpdateQuery = "Update Heat_Report_to_l3_SLAB set MARKID_ACT='" + MARKID_ACT + "',L2_SLABID='" + L2_SLABID + "',THICKNESS_HEAD='" + THICKNESS_HEAD + "',THICKNESS_TAIL='" + THICKNESS_TAIL + "',WIDTH_HEAD='" + WIDTH_HEAD + "',WIDTH_TAIL='" + WIDTH_TAIL + "' where HeatId='" + HEATID + "' and SLAB_SEQ_NO='" + SLAB_SEQ_NO + "'";
                            bool Status = clsObjMIS.DBInsertUpdateDeleteMIS(StrMISUpdateQuery);
                        }
                    }

                }
            }   // - Manish @bsl 

        }

        private void CastingFinished_CheckStatus_Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                string StrSTEEL_ID_caster1 = "";
                string StrSTEEL_ID_caster2 = "";
                string StrSTEEL_ID_caster3 = "";
                string StrHEATID_caster1 = "";
                string StrHEATID_caster2 = "";
                string StrHEATID_caster3 = "";
                string StrQuerySteelID = "SELECT  STEEL_ID,HEATID from ( SELECT pdc_heat.STEEL_ID,pdc_heat.HEATID, RANK() OVER (ORDER BY pdc_heat.Ladle_Open_Time DESC) rowno FROM pdc_heat  where SUBSTR(pdc_heat.Ladle_Open_Time,1,10)='" + DateTime.Now.ToString("yyyy-MM-dd") + "' and pdc_heat.HEAT_STATUS_CODE=5 and pdc_heat.GRADE_CODE is not null order by pdc_heat.Ladle_Open_Time desc) where  rowno =1";
                //string StrQuerySteelID = "SELECT  pdc_heat.STEEL_ID,pdc_heat.HEATID from pdc_heat where SUBSTR(Ladle_Open_Time,1,10)='" + DateTime.Now.ToString("yyyy-MM-dd") + "' and HEAT_STATUS_CODE=5 and GRADE_CODE is not null order by Ladle_Open_Time desc";
                //string StrQuerySteelID = "SELECT  pdc_heat.STEEL_ID,pdc_heat.HEATID from pdc_heat where  STEEL_ID=(Select max(STEEL_ID) from pdc_heat where pdc_heat.HEAT_STATUS_CODE=5 and pdc_heat.GRADE_CODE  is not null )";
                //string Query2 = "select heatid_cust from pp_heat where heatid=(select  heatid  from pd_heatdata where plant='CON' and plantno=2 and treatstart_plan=(Select max(treatstart_plan) from pd_heatdata where plant='CON' and plantno=2  ))";
                DBConnections clsObj_LOChkStat = new DBConnections();
                DataTable dt_LOChkStat = new DataTable();
                DataTable dt_LOChkStatCaster1 = new DataTable();
                DataTable dt_LOChkStatCaster3 = new DataTable();
                dt_LOChkStat = clsObj.DBSelectQueryCASTERII_Table(StrQuerySteelID);
                dt_LOChkStatCaster1 = clsObj.DBSelectQueryCASTERI_Table(StrQuerySteelID);
                dt_LOChkStatCaster3 = clsObj.DBSelectQueryCASTERIII_Table(StrQuerySteelID);
                if (dt_LOChkStat.Rows.Count > 0)
                {

                    for (int i = 0; i < dt_LOChkStat.Rows.Count; i++)
                    {
                        StrSTEEL_ID_caster2 = dt_LOChkStat.Rows[i][0].ToString();
                        StrHEATID_caster2 = dt_LOChkStat.Rows[i][1].ToString();
                        FnCASTERII_1(StrSTEEL_ID_caster2, StrHEATID_caster2);
                        FnCASTERII_2(StrSTEEL_ID_caster2, StrHEATID_caster2);
                        FnCASTERII_3(StrSTEEL_ID_caster2, StrHEATID_caster2);
                        FnCASTERII_4(StrSTEEL_ID_caster2, StrHEATID_caster2);
                        FnCASTERII_5(StrSTEEL_ID_caster2, StrHEATID_caster2);
                        Fn_HEAT_REPORT_TO_L3_SLAB_CASTERII(StrHEATID_caster2);
                    }

                }
                if (dt_LOChkStatCaster1.Rows.Count > 0)
                {

                    for (int i = 0; i < dt_LOChkStatCaster1.Rows.Count; i++)
                    {
                        StrSTEEL_ID_caster1 = dt_LOChkStatCaster1.Rows[i][0].ToString();
                        StrHEATID_caster1 = dt_LOChkStatCaster1.Rows[i][1].ToString();
                        FnCASTERI_1(StrSTEEL_ID_caster1, StrHEATID_caster1);
                        FnCASTERI_2(StrSTEEL_ID_caster1, StrHEATID_caster1);
                        FnCASTERI_3(StrSTEEL_ID_caster1, StrHEATID_caster1);
                        FnCASTERI_4(StrSTEEL_ID_caster1, StrHEATID_caster1);
                        FnCASTERI_5(StrSTEEL_ID_caster1, StrHEATID_caster1);
                        Fn_HEAT_REPORT_TO_L3_SLAB_CASTERI(StrHEATID_caster1);
                    }

                }
                if (dt_LOChkStatCaster3.Rows.Count > 0)
                {

                    for (int i = 0; i < dt_LOChkStatCaster3.Rows.Count; i++)
                    {
                        StrSTEEL_ID_caster3 = dt_LOChkStatCaster3.Rows[i][0].ToString();
                        StrHEATID_caster3 = dt_LOChkStatCaster3.Rows[i][1].ToString();
                        FnCASTERIII_1(StrSTEEL_ID_caster3, StrHEATID_caster3);
                        FnCASTERIII_2(StrSTEEL_ID_caster3, StrHEATID_caster3);
                        FnCASTERIII_3(StrSTEEL_ID_caster3, StrHEATID_caster3);
                        FnCASTERIII_4(StrSTEEL_ID_caster3, StrHEATID_caster3);
                        FnCASTERIII_5(StrSTEEL_ID_caster3, StrHEATID_caster3);
                        Fn_HEAT_REPORT_TO_L3_SLAB_CASTERIII(StrHEATID_caster3);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("CastingFinished_CheckStatus_Timer_Tick" + ex.ToString());
                txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
            }

        }

        private void SMCPLANTSTATUS_Tick(object sender, EventArgs e)
        {
            FnSMC_Caster();
            FnSMC_Conarc();
            FnSMC_HMD();
            FnSMC_LF();
            FnSMC_RH();
        }
        public void FnFinAimChem()
        {
            //Final Chemistry Inseration to MIS Database
            DataTable dt_MIS_Final_Heat_Chem = new DataTable();
            //string StrQueryMISFinChem = "SELECT a.heatid hid,a.sample_no,a.taken_time,a.sample_loc,a.grade_code,nvl(a.LIQUID_TEMP,0) LIQ,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'C' AND ae.analysis_id = a.analysis_id),0) c,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Si' AND ae.analysis_id = a.analysis_id),0) si,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Mn' AND ae.analysis_id = a.analysis_id),0) mn,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'P' AND ae.analysis_id = a.analysis_id),0) p,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'S' AND ae.analysis_id = a.analysis_id),0) s,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Al' AND ae.analysis_id = a.analysis_id),0) al,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Als' AND ae.analysis_id = a.analysis_id),0) als,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Cu'  AND ae.analysis_id = a.analysis_id),0) cu,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Cr' AND ae.analysis_id = a.analysis_id),0) cr,nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae WHERE element_name = 'Mo' AND ae.analysis_id = a.analysis_id),0) mo,nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae WHERE element_name = 'Ni' AND ae.analysis_id = a.analysis_id),0) ni,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'V' AND ae.analysis_id = a.analysis_id),0) v,nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae WHERE element_name = 'Ti' AND ae.analysis_id = a.analysis_id),0) ti,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Nb' AND ae.analysis_id = a.analysis_id),0) nb,nvl ((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Ca'  AND ae.analysis_id = a.analysis_id),0) ca,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'W'  AND ae.analysis_id = a.analysis_id),0) w,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Sn'  AND ae.analysis_id = a.analysis_id),0) sn,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'As' AND ae.analysis_id = a.analysis_id),0) asr,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Te' AND ae.analysis_id = a.analysis_id),0) te,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Bi' AND ae.analysis_id = a.analysis_id),0) bi,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'B' AND ae.analysis_id = a.analysis_id),0) b,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Pb' AND ae.analysis_id = a.analysis_id),0) Pb,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Mg' AND ae.analysis_id = a.analysis_id),0) Mg,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'N' AND ae.analysis_id = a.analysis_id),0) N,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Ve' AND ae.analysis_id = a.analysis_id),0) Ve,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Co' AND ae.analysis_id = a.analysis_id),0) Co,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Ce' AND ae.analysis_id = a.analysis_id),0) Ce,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Sb' AND ae.analysis_id = a.analysis_id),0) Sb,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Zr' AND ae.analysis_id = a.analysis_id),0) Zr,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'O' AND ae.analysis_id = a.analysis_id),0) O,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'H' AND ae.analysis_id = a.analysis_id),0) H FROM pd_analysis a where (1=1) and a.sample_loc = 'TUN' AND a.heatid ='" + StrHEATID + "'";
            string StrQueryMISFinChem = "select *from(SELECT a.heatid hid,a.sample_no,a.taken_time,a.sample_loc,a.grade_code,nvl(a.LIQUID_TEMP,0) LIQ,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'C' AND ae.analysis_id = a.analysis_id),0) c,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Si' AND ae.analysis_id = a.analysis_id),0) si,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Mn' AND ae.analysis_id = a.analysis_id),0) mn,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'P' AND ae.analysis_id = a.analysis_id),0) p,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'S' AND ae.analysis_id = a.analysis_id),0) s,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Al' AND ae.analysis_id = a.analysis_id),0) al,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Al_S' AND ae.analysis_id = a.analysis_id),0) als,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Cu'  AND ae.analysis_id = a.analysis_id),0) cu,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Cr' AND ae.analysis_id = a.analysis_id),0) cr,nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae WHERE element_name = 'Mo' AND ae.analysis_id = a.analysis_id),0) mo,nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae WHERE element_name = 'Ni' AND ae.analysis_id = a.analysis_id),0) ni,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'V' AND ae.analysis_id = a.analysis_id),0) v,nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae WHERE element_name = 'Ti' AND ae.analysis_id = a.analysis_id),0) ti,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Nb' AND ae.analysis_id = a.analysis_id),0) nb,nvl ((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Ca'  AND ae.analysis_id = a.analysis_id),0) ca,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'W'  AND ae.analysis_id = a.analysis_id),0) w,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Sn'  AND ae.analysis_id = a.analysis_id),0) sn,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'As' AND ae.analysis_id = a.analysis_id),0) asr,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Te' AND ae.analysis_id = a.analysis_id),0) te,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Bi' AND ae.analysis_id = a.analysis_id),0) bi,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'B' AND ae.analysis_id = a.analysis_id),0) b,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Pb' AND ae.analysis_id = a.analysis_id),0) Pb,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Mg' AND ae.analysis_id = a.analysis_id),0) Mg,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'N' AND ae.analysis_id = a.analysis_id),0) N,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Ve' AND ae.analysis_id = a.analysis_id),0) Ve,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Co' AND ae.analysis_id = a.analysis_id),0) Co,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Ce' AND ae.analysis_id = a.analysis_id),0) Ce,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Sb' AND ae.analysis_id = a.analysis_id),0) Sb,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'Zr' AND ae.analysis_id = a.analysis_id),0) Zr,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'O' AND ae.analysis_id = a.analysis_id),0) O,nvl((SELECT MAX (actual_value) FROM pd_analysis_entry ae WHERE element_name = 'H' AND ae.analysis_id = a.analysis_id),0) H FROM pd_analysis a where (1=1) and a.sample_loc = 'TUN' and  SUBSTR(a.taken_time ,1,10)='" + DateTime.Now.ToString("yyyy-MM-dd") + "' order by a.taken_time desc) where rownum<=5";
            dt_MIS_Final_Heat_Chem = clsObj.DBSelectQueryCASTERII_Table(StrQueryMISFinChem);
            if (dt_MIS_Final_Heat_Chem.Rows.Count > 0)
            {
                DataTable dt_Mis_FinChem = new DataTable();
                for (int i = 0; i < dt_MIS_Final_Heat_Chem.Rows.Count; i++)
                {
                    string StrQueryMISCaster = "select *from SMS_FinalChemistry where HeatId='" + dt_MIS_Final_Heat_Chem.Rows[i]["hid"].ToString() + "' and SAMPLE_NO='" + dt_MIS_Final_Heat_Chem.Rows[i]["sample_no"].ToString() + "'";
                    dt_Mis_FinChem = clsObj.DBSelectQueryMIS_Table(StrQueryMISCaster);
                    if (dt_Mis_FinChem.Rows.Count == 0)
                    {
                        string StrInsertQuery = "INSERT INTO SMS_FinalChemistry(HeatID,SAMPLE_NO,TAKEN_TIME,SAMPLE_LOC,GRADE,LIQ,C,Si,Mn,P,S,Al,ALS,Cu," +
                            "Cr,Mo,Ni,V,Ti,Nb,Ca,W,Sn,Asr,Te,Bi,B,Pb,Mg,N,Ve,CO,Ce,Sb,Zr,O,H) VALUES('" + dt_MIS_Final_Heat_Chem.Rows[i]["hid"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["sample_no"].ToString() +
                            "','" + dt_MIS_Final_Heat_Chem.Rows[i]["taken_time"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["sample_loc"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["grade_code"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["LIQ"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["C"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["SI"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["MN"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["P"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["S"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["AL"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["ALS"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["CU"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["CR"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["MO"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["NI"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["V"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["TI"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["NB"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["CA"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["W"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["SN"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["ASR"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["TE"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["BI"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["B"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["PB"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["MG"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["N"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["VE"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["CO"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["CE"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["SB"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["ZR"].ToString() + "','" + dt_MIS_Final_Heat_Chem.Rows[i]["O"].ToString() + "','" +
                            dt_MIS_Final_Heat_Chem.Rows[i]["H"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "UPDATE SMS_FinalChemistry SET HeatID ='" + dt_MIS_Final_Heat_Chem.Rows[i]["hid"].ToString() + "',SAMPLE_NO='" + dt_MIS_Final_Heat_Chem.Rows[i]["sample_no"].ToString() + "',C ='" + dt_MIS_Final_Heat_Chem.Rows[i]["C"].ToString() + "',Mn ='" + dt_MIS_Final_Heat_Chem.Rows[i]["MN"].ToString() + "',Si ='" + dt_MIS_Final_Heat_Chem.Rows[i]["SI"].ToString() + "',S ='" + dt_MIS_Final_Heat_Chem.Rows[i]["S"].ToString() + "',P='" + dt_MIS_Final_Heat_Chem.Rows[i]["P"].ToString() +
                            "',Ni ='" + dt_MIS_Final_Heat_Chem.Rows[i]["NI"].ToString() + "',Cr ='" + dt_MIS_Final_Heat_Chem.Rows[i]["CR"].ToString() + "',Mo ='" + dt_MIS_Final_Heat_Chem.Rows[i]["MO"].ToString() + "',Cu ='" + dt_MIS_Final_Heat_Chem.Rows[i]["CU"].ToString() + "',Sn ='" + dt_MIS_Final_Heat_Chem.Rows[i]["SN"].ToString() + "',Al='" + dt_MIS_Final_Heat_Chem.Rows[i]["AL"].ToString() +
                            "',ALS ='" + dt_MIS_Final_Heat_Chem.Rows[i]["ALS"].ToString() + "',Ti='" + dt_MIS_Final_Heat_Chem.Rows[i]["TI"].ToString() + "',B='" + dt_MIS_Final_Heat_Chem.Rows[i]["B"].ToString() + "',V ='" + dt_MIS_Final_Heat_Chem.Rows[i]["V"].ToString() + "',Nb='" + dt_MIS_Final_Heat_Chem.Rows[i]["NB"].ToString() +
                            "',Ca ='" + dt_MIS_Final_Heat_Chem.Rows[i]["CA"].ToString() + "',N ='" + dt_MIS_Final_Heat_Chem.Rows[i]["N"].ToString() + "' WHERE HeatID='" + dt_MIS_Final_Heat_Chem.Rows[i]["hid"].ToString() + "' and SAMPLE_NO='" + dt_MIS_Final_Heat_Chem.Rows[i]["sample_no"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void FnSMC_Caster()
        {

            string StrHEATID = "";
            string StrStatus = "";
            //string StrQueryOrclCaster = "SELECT Distinct DECODE(p.HEATID||p.TREATID,'CCS_current_HEAT_ID'||'CCS_current_TREAT_ID','Y','N') AS ACT,de.PLANT||de.PLANTNO AS UNIT,p.HEATID_CUST HEAT_ID,ph.PRODORDERID,grade.STEELGRADECODE,grade.STEELGRADECODEDESC,status.HEATSTATUS STATUS,status.HEATSTATUSDESC STATUS_DESC,de.REMSTEEL,de.TEMP FROM PP_HEAT_PLANT p,PP_HEAT ph,(SELECT g.STEELGRADECODE, g.STEELGRADECODEDESC FROM PP_GRADE g) grade,(SELECT UNIQUE s.ORD_STAT_DESC,s.ORD_STAT_NO,s.ORD_STAT_CODE,o.PRODORDERID,o.STEELGRADECODE FROM GC_PRD_STAT s, PP_ORDER o WHERE o.ORD_STAT_NO = s.ORD_STAT_NO) ord,(SELECT st.HEATSTATORDER,st.HEATSTATUS,st.HEATSTATUSDESC FROM GC_HEAT_STAT st) status,PD_PLANT_CCS de WHERE de.PLANT='CCS'AND p.EXPIRATIONDATE='VALID'AND p.metal_type= 'Steel'AND ph.metal_type= 'Steel'AND ph.HEATID_CUST= p.HEATID_CUST AND grade.STEELGRADECODE = ord.STEELGRADECODE AND ord.PRODORDERId= ph.PRODORDERID AND de.HEATID_CUST= p.HEATID_CUST AND status.HEATSTATORDER = ph.HEATSTATORDER AND ph.HEATSTATORDER IN (91, 92, 93)";
            //string StrQueryOrclCaster = "SELECT distinct  de.PLANT  ||de.PLANTNO AS UNIT,  p.HEATID_CUST HEAT_ID,  grade.STEELGRADECODE,  status.HEATSTATUSDESC STATUS_DESC,  de.HEATSTATUS,  de.REMSTEEL,  de.requiredtime,  (de.castspeeds1) / 1000 AS CasterSpeed,  de.remsteel_tundish,  de.CASTERSTATUS1,    de.mouldwidths1l FROM PP_HEAT_PLANT p,   PP_HEAT ph,  (SELECT g.STEELGRADECODE, g.STEELGRADECODEDESC FROM PP_GRADE g  ) grade,  (SELECT UNIQUE s.ORD_STAT_DESC,    s.ORD_STAT_NO,    s.ORD_STAT_CODE,    o.PRODORDERID,    o.STEELGRADECODE   FROM GC_PRD_STAT s,     PP_ORDER o   WHERE o.ORD_STAT_NO = s.ORD_STAT_NO   ) ord,   (SELECT st.HEATSTATORDER,     st.HEATSTATUS,     st.HEATSTATUSDESC   FROM GC_HEAT_STAT st   ) status,   PD_PLANT_CCS de WHERE de.PLANT           ='CCS' AND p.EXPIRATIONDATE     ='VALID' AND p.metal_type         = 'Steel' AND ph.metal_type        = 'Steel' AND ph.HEATID_CUST       = p.HEATID_CUST AND grade.STEELGRADECODE = ord.STEELGRADECODE AND ord.PRODORDERId      = ph.PRODORDERID AND de.HEATID_CUST       = p.HEATID_CUST AND status.HEATSTATORDER = ph.HEATSTATORDER AND ph.HEATSTATORDER    IN (91, 92, 93)";
            //string StrQueryOrclCaster = "SELECT de.PLANT  || de.PLANTNO AS UNIT,   de.heatid_cust,  grade.STEELGRADECODE,  de.HEATSTATUS,  de.REMSTEEL,  de.REQUIREDTIME,  (de.CASTSPEEDS1) / 1000 AS CasterSpeed,  de.REMSTEEL_TUNDISH,  de.CASTERSTATUS1,    de.MOULDWIDTHS1L FROM PP_HEAT ph,  (SELECT g.STEELGRADECODE, g.STEELGRADECODEDESC FROM PP_GRADE g  ) grade,  (SELECT UNIQUE s.ORD_STAT_DESC,    s.ORD_STAT_NO,    s.ORD_STAT_CODE,    o.PRODORDERID,    o.STEELGRADECODE  FROM GC_PRD_STAT s,    PP_ORDER o  WHERE o.ORD_STAT_NO = s.ORD_STAT_NO  ) ord,  PD_PLANT_CCS de WHERE grade.STEELGRADECODE = ord.STEELGRADECODE AND ord.PRODORDERID        = ph.PRODORDERID AND ph.HEATID_CUST         = de.HEATID_CUST AND (de.PLANT              = 'CCS' AND ph.METAL_TYPE          = 'Steel' AND de.casterstatus1   IN (9, 10, 11, 12, 13, 14, 15, 16, 17)) order by de.plantno asc";
            //string StrQueryOrclCaster = "SELECT de.PLANT  || de.PLANTNO AS UNIT,de.heatid_cust,(select 'BSCO6' from dual) as STEELGRADECODE,de.HEATSTATUS,de.REMSTEEL,de.REQUIREDTIME,(de.CASTSPEEDS1) / 1000 AS CasterSpeed,de.REMSTEEL_TUNDISH,de.CASTERSTATUS1,de.MOULDWIDTHS1L FROM PD_PLANT_CCS de where (de.PLANT              = 'CCS' AND de.casterstatus1   IN (9, 10, 11, 12, 13, 14, 15, 16, 17)) order by de.plantno asc";
            //@MANISH ADDED SEQUENC LENGTH AND CAST LENGTH
            string StrQueryOrclCaster = "SELECT de.PLANT  || de.PLANTNO AS UNIT,de.heatid_cust,(select 'BSCO6' from dual) as STEELGRADECODE,de.HEATSTATUS,de.REMSTEEL,de.REQUIREDTIME,(de.CASTSPEEDS1) / 1000 AS CasterSpeed,de.REMSTEEL_TUNDISH,de.CASTERSTATUS1,de.MOULDWIDTHS1L,de.CAST_LENGTH as CastLength,de.NOHEATSEQPLAN FROM PD_PLANT_CCS de where (de.PLANT              = 'CCS' AND de.casterstatus1   IN (9, 10, 11, 12, 13, 14, 15, 16, 17)) order by de.plantno asc";
            DataTable dt_Orcl_Caster = new DataTable();
            dt_Orcl_Caster = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclCaster);
            if (dt_Orcl_Caster.Rows.Count > 0)
            {
                for (int i = 0; i < dt_Orcl_Caster.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_Caster.Rows[i][1].ToString();
                    StrStatus = dt_Orcl_Caster.Rows[i][3].ToString();
                    //if (StrStatus == "5")
                    //{
                    //    FnFinAimChem(StrHEATID, StrStatus);
                    //}
                    DataTable dt_MIS_Caster = new DataTable();
                    string StrQueryMISCaster = "select HEAT_ID,HEATSTATUS,SLNo from OP_CASTER_STATUS WHERE HEAT_ID='" + StrHEATID + "' AND HEATSTATUS='" + StrStatus + "'";
                    dt_MIS_Caster = clsObj.DBSelectQueryMIS_Table(StrQueryMISCaster);
                    if (dt_MIS_Caster.Rows.Count == 0)
                    {
                        //string StrInsertQuery = "insert into OP_CASTER_STATUS(UNIT,HEAT_ID,STEL_GRADE_CODE,STATUS_DESC,HEATSTATUS,REM_STEEL,REQUIREDTIME,CASTERSPEED,REMSTEEL_TUNDISH,CASTERSTATUS1,MOULDWIDTHS1L)values('" + dt_Orcl_Caster.Rows[i][0].ToString() + "','" + dt_Orcl_Caster.Rows[i][1].ToString() + "','" + dt_Orcl_Caster.Rows[i][2].ToString() + "','','" + dt_Orcl_Caster.Rows[i][3].ToString() + "','" + dt_Orcl_Caster.Rows[i][4].ToString() + "','" + dt_Orcl_Caster.Rows[i][5].ToString().Replace(",", ".") + "','" + dt_Orcl_Caster.Rows[i][6].ToString() + "','" + dt_Orcl_Caster.Rows[i][7].ToString() + "','" + dt_Orcl_Caster.Rows[i][8].ToString() + "','" + dt_Orcl_Caster.Rows[i][9].ToString() + "')";
                        string StrInsertQuery = "insert into OP_CASTER_STATUS(UNIT,HEAT_ID,STEL_GRADE_CODE,STATUS_DESC,HEATSTATUS,REM_STEEL,REQUIREDTIME,CASTERSPEED,REMSTEEL_TUNDISH,CASTERSTATUS1,MOULDWIDTHS1L,CAST_LENGTH,NOHEATSEQPLAN)values('" + dt_Orcl_Caster.Rows[i][0].ToString() + "','" + dt_Orcl_Caster.Rows[i][1].ToString() + "','" + dt_Orcl_Caster.Rows[i][2].ToString() + "','','" + dt_Orcl_Caster.Rows[i][3].ToString() + "','" + dt_Orcl_Caster.Rows[i][4].ToString() + "','" + dt_Orcl_Caster.Rows[i][5].ToString().Replace(",", ".") + "','" + dt_Orcl_Caster.Rows[i][6].ToString() + "','" + dt_Orcl_Caster.Rows[i][7].ToString() + "','" + dt_Orcl_Caster.Rows[i][8].ToString() + "','" + dt_Orcl_Caster.Rows[i][9].ToString() + "','" + dt_Orcl_Caster.Rows[i][10].ToString() + "','" + dt_Orcl_Caster.Rows[i][11].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                    else
                    {
                        //string StrInsertQuery = "Update OP_CASTER_STATUS set UNIT='" + dt_Orcl_Caster.Rows[i][0].ToString() + "',HEAT_ID='" + dt_Orcl_Caster.Rows[i][1].ToString() + "',STEL_GRADE_CODE='" + dt_Orcl_Caster.Rows[i][2].ToString() + "',STATUS_DESC='',HEATSTATUS='" + dt_Orcl_Caster.Rows[i][3].ToString() + "',REM_STEEL='" + dt_Orcl_Caster.Rows[i][4].ToString() + "',REQUIREDTIME='" + dt_Orcl_Caster.Rows[i][5].ToString().Replace(",", ".") + "',CASTERSPEED='" + dt_Orcl_Caster.Rows[i][6].ToString() + "',REMSTEEL_TUNDISH='" + dt_Orcl_Caster.Rows[i][7].ToString() + "',CASTERSTATUS1='" + dt_Orcl_Caster.Rows[i][8].ToString() + "',MOULDWIDTHS1L='" + dt_Orcl_Caster.Rows[i][9].ToString() + "' where SLNo='" + dt_MIS_Caster.Rows[0][2].ToString() + "'";
                        string StrInsertQuery = "Update OP_CASTER_STATUS set UNIT='" + dt_Orcl_Caster.Rows[i][0].ToString() + "',HEAT_ID='" + dt_Orcl_Caster.Rows[i][1].ToString() + "',STEL_GRADE_CODE='" + dt_Orcl_Caster.Rows[i][2].ToString() + "',STATUS_DESC='',HEATSTATUS='" + dt_Orcl_Caster.Rows[i][3].ToString() + "',REM_STEEL='" + dt_Orcl_Caster.Rows[i][4].ToString() + "',REQUIREDTIME='" + dt_Orcl_Caster.Rows[i][5].ToString().Replace(",", ".") + "',CASTERSPEED='" + dt_Orcl_Caster.Rows[i][6].ToString() + "',REMSTEEL_TUNDISH='" + dt_Orcl_Caster.Rows[i][7].ToString() + "',CASTERSTATUS1='" + dt_Orcl_Caster.Rows[i][8].ToString() + "',MOULDWIDTHS1L='" + dt_Orcl_Caster.Rows[i][9].ToString() + " ',CAST_LENGTH='" + dt_Orcl_Caster.Rows[i][10].ToString() + "',NOHEATSEQPLAN='" + dt_Orcl_Caster.Rows[i][11].ToString() + "' where SLNo='" + dt_MIS_Caster.Rows[0][2].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery); //@ -MANISH 30-05-2014
                    }
                }


            }

        }
        private void FnSMC_Conarc()
        {

            string StrTREATID = "";
            string StrHEATID = "";
            string StrStatus = "";
            //string StrQueryOrclConarc = "SELECT Distinct DECODE(p.HEATID  ||p.TREATID,'CON_current_HEAT_ID'  ||'CON_current_TREAT_ID','Y','N') AS ACT,  p.PLANT  ||p.PLANTNO AS UNIT,  p.HEATID_CUST HEAT_ID,  p.TREATID_CUST TREATID,  ph.PRODORDERID,  grade.STEELGRADECODE,  grade.STEELGRADECODEDESC,  ord.HEATING_MODE,  ord.ORD_STAT_CODE ORD_STAT,  ord.ORD_STAT_DESC ORD_STAT_DESC,  status.HEATSTATUS STATUS,  status.HEATSTATUSDESC STATUS_DESC,  SUBSTR(DECODE(p.ACTTREATSTART, NULL, p.PLANTREATSTART, p.ACTTREATSTART),0,19) START_T,  nvl(DECODE(p.ACTTREATSTART, NULL, '0', '1'),0) START_T_B,  p.ACTTREATEND AS END_T,  nvl(DECODE(p.ACTTREATEND, NULL, '0', '1'),0) END_T_B,  nvl(DECODE(p.ACTTREATEND, NULL, phase.STEELMASS, rep.FINALWEIGHT),0) WEIGHT,  nvl(DECODE(p.ACTTREATEND, NULL, '1', '0'),0) WEIGHT_B,  nvl(DECODE(p.ACTTREATEND, NULL, phase.TEMP, rep.FINALTEMPCALC),0) TEMP,  nvl(DECODE(p.ACTTREATEND, NULL, '1', '0'),0) TEMP_B,  nvl(DECODE(p.ACTTREATEND, NULL, oxygen.RESULTVALUE, de.total_o2_cons),0) OXYGEN,  nvl(DECODE(p.ACTTREATEND, NULL, '1', '0'),0) OXYGEN_B,  nvl(DECODE(p.ACTTREATEND, NULL, energy.RESULTVALUE, de.TOTAL_ELEC_EGY),0) ENERGY,  nvl(DECODE(p.ACTTREATEND, NULL, '1', '0'),0) ENERGY_B FROM PP_HEAT_PLANT p,  PP_HEAT ph,  (SELECT g.STEELGRADECODE, g.STEELGRADECODEDESC FROM PP_GRADE g  ) grade,  (SELECT UNIQUE s.ORD_STAT_DESC,    s.ORD_STAT_NO,    s.ORD_STAT_CODE,    o.PRODORDERID,    o.STEELGRADECODE,    o.heating_mode  FROM GC_PRD_STAT s,    PP_ORDER o  WHERE o.ORD_STAT_NO = s.ORD_STAT_NO  ) ord,(SELECT st.HEATSTATORDER,  st.HEATSTATUS,  st.HEATSTATUSDESC FROM GC_HEAT_STAT st) status,(SELECT h.FINALWEIGHT,  h.HEATID,  h.TREATID,  h.FINALTEMPCALC FROM PD_REPORT h WHERE h.PLANT= 'CON') rep,(SELECT r.HEATID,  r.TREATID,  r.STEELMASS,  r.TEMP FROM PD_PHASE_RES r WHERE r.PLANT     = 'CON' AND r.RES_PHASENO =  (SELECT p.RES_PHASENO  FROM GC_PHASE p  WHERE p.PHASENAME = 'Tap'  AND p.PLANT       = 'CON'  )) phase,(SELECT pl.HEATID,  pl.TREATID,  pl.RESULTVALUE FROM PD_PHASE_RES_PLANT pl WHERE pl.PLANT     = 'CON' AND pl.RES_PHASENO =  (SELECT p.RES_PHASENO  FROM GC_PHASE p  WHERE p.PHASENAME = 'Tap'  AND p.PLANT       = 'CON'  ) AND pl.RESULTVALNO =  (SELECT rp.RESULTVALNO  FROM GC_PHASE_RES_PLANT rp  WHERE rp.PLANT       = 'CON'  AND rp.RESULTVARNAME = 'Energy'  )) energy,(SELECT pl.HEATID,  pl.TREATID,  pl.RESULTVALUE FROM PD_PHASE_RES_PLANT pl WHERE pl.PLANT     = 'CON' AND pl.RES_PHASENO =  (SELECT p.RES_PHASENO  FROM GC_PHASE p  WHERE p.PHASENAME = 'Tap'  AND p.PLANT       = 'CON'  )AND pl.RESULTVALNO =  (SELECT rp.RESULTVALNO  FROM GC_PHASE_RES_PLANT rp  WHERE rp.PLANT       = 'CON'  AND rp.RESULTVARNAME = 'O2Fl_L'  )) oxygen,PDE_HEAT_DATA de WHERE p.PLANT = 'CON' AND p.EXPIRATIONDATE ='VALID' AND p.metal_type = 'Steel' AND ph.metal_type = 'Steel' AND ph.HEATID_CUST(+) = p.HEATID_CUST AND grade.STEELGRADECODE(+) = ord.STEELGRADECODE AND ord.PRODORDERId(+) = ph.PRODORDERID AND phase.HEATID(+)= p.HEATID AND phase.TREATID(+) = p.TREATID AND rep.HEATID(+) = p.HEATID AND rep.TREATID(+) = p.TREATID AND energy.HEATID(+) = p.HEATID AND energy.TREATID(+) = p.TREATID AND oxygen.HEATID(+) = p.HEATID AND oxygen.TREATID(+) = p.TREATID AND de.HEATID(+) = p.HEATID AND de.TREATID(+) = p.TREATID AND de.plant = p.plant AND status.HEATSTATORDER(+) = ph.HEATSTATORDER AND p.ACTTREATSTART is not null  and p.ACTTREATEND is null  ";
            //string StrQueryOrclConarc = "SELECT p.PLANT  || p.PLANTNO AS UNIT,  p.HEATID_CUST HEAT_ID,  p.TREATID_CUST TREATID,  grade.STEELGRADECODE as GRADE,  ord.HEATING_MODE, status.HEATSTATUS STATUS,  status.HEATSTATUSDESC STATUS_DESC,  SUBSTR(DECODE(p.ACTTREATSTART, NULL, p.PLANTREATSTART, p.ACTTREATSTART), 0, 19) START_T,   SUBSTR(DECODE(p.ACTTREATEND, NULL, p.ACTTREATEND, p.ACTTREATEND), 0, 19) END_T,  de.ARCING_STARTTIME,  de.ARCING_ENDTIME, de.POWER_ON_DUR, de.TOTAL_ELEC_EGY, PDB_HEAT_DATA.BLOWSTARTTIME, PDB_HEAT_DATA.BLOWENDTIME, PDB_HEAT_DATA.O2_BLOW_DUR,          PDB_HEAT_DATA.TOTAL_O2_CONS AS TOTAL_OXYGEN_CONS FROM PP_HEAT_PLANT p,  PP_HEAT ph,  (SELECT g.STEELGRADECODE, g.STEELGRADECODEDESC FROM PP_GRADE g  ) grade,  (SELECT UNIQUE s.ORD_STAT_DESC,    s.ORD_STAT_NO,    s.ORD_STAT_CODE,    o.PRODORDERID,    o.STEELGRADECODE,    o.HEATING_MODE  FROM GC_PRD_STAT s,    PP_ORDER o  WHERE o.ORD_STAT_NO = s.ORD_STAT_NO  ) ord,  (SELECT st.HEATSTATORDER,    st.HEATSTATUS,    st.HEATSTATUSDESC  FROM GC_HEAT_STAT st  ) status,  PDE_HEAT_DATA de,  PDB_HEAT_DATA WHERE ph.HEATID_CUST(+)     = p.HEATID_CUST AND grade.STEELGRADECODE(+) = ord.STEELGRADECODE AND ord.PRODORDERID(+)      = ph.PRODORDERID AND de.HEATID(+)            = p.HEATID AND de.TREATID(+)           = p.TREATID AND de.PLANT                = p.PLANT AND status.HEATSTATORDER(+) = ph.HEATSTATORDER AND PDB_HEAT_DATA.PLANT     = de.PLANT AND PDB_HEAT_DATA.HEATID    = de.HEATID AND PDB_HEAT_DATA.TREATID   = de.TREATID AND (p.PLANT                = 'CON' AND p.EXPIRATIONDATE        = 'VALID' AND p.METAL_TYPE            = 'Steel' AND ph.METAL_TYPE           = 'Steel' AND p.ACTTREATSTART        IS NOT NULL AND p.ACTTREATEND          IS NULL)";
            //@ Manish added tapestart and tap end time dated 10/07/2014 above is running sql
            string StrQueryOrclConarc = "SELECT p.PLANT  || p.PLANTNO AS UNIT,  p.HEATID_CUST HEAT_ID,  p.TREATID_CUST TREATID,  grade.STEELGRADECODE as GRADE,  ord.HEATING_MODE, status.HEATSTATUS STATUS,  status.HEATSTATUSDESC STATUS_DESC,  SUBSTR(DECODE(p.ACTTREATSTART, NULL, p.PLANTREATSTART, p.ACTTREATSTART), 0, 19) START_T,   SUBSTR(DECODE(p.ACTTREATEND, NULL, p.ACTTREATEND, p.ACTTREATEND), 0, 19) END_T,  de.ARCING_STARTTIME,  de.ARCING_ENDTIME, de.POWER_ON_DUR, de.TOTAL_ELEC_EGY, PDB_HEAT_DATA.BLOWSTARTTIME, PDB_HEAT_DATA.BLOWENDTIME, PDB_HEAT_DATA.O2_BLOW_DUR,          PDB_HEAT_DATA.TOTAL_O2_CONS AS TOTAL_OXYGEN_CONS,Phase.Tappingstarttime As Tap_Start,phase.Tappingendtime  As Tap_end FROM PP_HEAT_PLANT p,  PP_HEAT ph,  (SELECT g.STEELGRADECODE, g.STEELGRADECODEDESC FROM PP_GRADE g  ) grade,  (SELECT UNIQUE s.ORD_STAT_DESC,    s.ORD_STAT_NO,    s.ORD_STAT_CODE,    o.PRODORDERID,    o.STEELGRADECODE,    o.HEATING_MODE  FROM GC_PRD_STAT s,    PP_ORDER o  WHERE o.ORD_STAT_NO = s.ORD_STAT_NO  ) ord,  (SELECT st.HEATSTATORDER,    st.HEATSTATUS,    st.HEATSTATUSDESC  FROM GC_HEAT_STAT st  ) status,(SELECT pdr.Heatid,Pdr.Treatid,Pdr.Tappingstarttime,Pdr.Tappingendtime From Pd_Report Pdr WHERE Pdr.Plant = 'CON') phase,  PDE_HEAT_DATA de,  PDB_HEAT_DATA WHERE ph.HEATID_CUST(+)     = p.HEATID_CUST AND grade.STEELGRADECODE(+) = ord.STEELGRADECODE AND ord.PRODORDERID(+)      = ph.PRODORDERID AND de.HEATID(+)            = p.HEATID AND de.TREATID(+)           = p.TREATID AND de.PLANT                = p.PLANT AND status.HEATSTATORDER(+) = ph.HEATSTATORDER AND PDB_HEAT_DATA.PLANT     = de.PLANT AND PDB_HEAT_DATA.HEATID    = de.HEATID AND PDB_HEAT_DATA.TREATID   = de.TREATID And Phase.Heatid(+) = P.Heatid And phase.Treatid(+)        = P.Treatid AND (p.PLANT                = 'CON' AND p.EXPIRATIONDATE        = 'VALID' AND p.METAL_TYPE            = 'Steel' AND ph.METAL_TYPE           = 'Steel' AND p.ACTTREATSTART        IS NOT NULL AND p.ACTTREATEND          IS NULL)";
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclConarc);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                string StrUpdateDefault = "; UPDATE OP_CONARC_STATUS set END_T=null where END_T='1900-01-01 00:00:00.000';UPDATE OP_CONARC_STATUS set ARCING_START_TIME=null where ARCING_START_TIME='1900-01-01 00:00:00.000';UPDATE OP_CONARC_STATUS set ARCING_END_TIME=null where ARCING_END_TIME='1900-01-01 00:00:00.000';UPDATE OP_CONARC_STATUS set BLOW_START_TIME=null where BLOW_START_TIME='1900-01-01 00:00:00.000';UPDATE OP_CONARC_STATUS set BLOW_END_TIME=null where BLOW_END_TIME='1900-01-01 00:00:00.000'";
                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_Conarc.Rows[i]["HEAT_ID"].ToString();
                    StrTREATID = dt_Orcl_Conarc.Rows[i]["TREATID"].ToString();
                    StrStatus = Convert.ToInt32(dt_Orcl_Conarc.Rows[i]["STATUS"]).ToString();
                    //if (dt_Orcl_Conarc.Rows[i]["END_T"].ToString() == "")
                    //{
                    //    string StrEND_T = DBNull.Value.ToString();
                    //}
                    DataTable dt_MIS_Conarc = new DataTable();
                    string StrQueryMISConarc = "select SLNo,HEAT_ID,TREAT_ID,STATUS from OP_CONARC_STATUS WHERE HEAT_ID='" + StrHEATID + "' AND TREAT_ID='" + StrTREATID + "' AND STATUS='" + StrStatus + "'";
                    dt_MIS_Conarc = clsObj.DBSelectQueryMIS_Table(StrQueryMISConarc);
                    if (dt_MIS_Conarc.Rows.Count == 0)
                    {
                        //string StrInsertQuery = "insert into OP_CONARC_STATUS(UNIT,HEAT_ID,TREAT_ID,STEL_GRADE_CODE,HEATING_MODE,STATUS,STATUS_DESC,START_T,END_T,ARCING_START_TIME,ARCING_END_TIME,POWER_ON_DUR,TOTAL_ELECT_EGY,BLOW_START_TIME,BLOW_END_TIME,O2_BLOW_DUR)values('" + dt_Orcl_Conarc.Rows[i][0].ToString() + "','" + dt_Orcl_Conarc.Rows[i][1].ToString() + "','" + dt_Orcl_Conarc.Rows[i][2].ToString() + "','" + dt_Orcl_Conarc.Rows[i][3].ToString() + "','" + dt_Orcl_Conarc.Rows[i][4].ToString() + "','" + dt_Orcl_Conarc.Rows[i][5].ToString() + "','" + dt_Orcl_Conarc.Rows[i][6].ToString() + "','" + dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][8].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][9].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][10].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][11].ToString() + "','" + dt_Orcl_Conarc.Rows[i][12].ToString() + "','" + dt_Orcl_Conarc.Rows[i][13].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][14].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][15].ToString() + "')";
                        string StrInsertQuery = "insert into OP_CONARC_STATUS(UNIT,HEAT_ID,TREAT_ID,STEL_GRADE_CODE,HEATING_MODE,STATUS,STATUS_DESC,START_T,END_T,ARCING_START_TIME,ARCING_END_TIME,POWER_ON_DUR,TOTAL_ELECT_EGY,BLOW_START_TIME,BLOW_END_TIME,O2_BLOW_DUR,Tap_Start,Tap_end)values('" + dt_Orcl_Conarc.Rows[i][0].ToString() + "','" + dt_Orcl_Conarc.Rows[i][1].ToString() + "','" + dt_Orcl_Conarc.Rows[i][2].ToString() + "','" + dt_Orcl_Conarc.Rows[i][3].ToString() + "','" + dt_Orcl_Conarc.Rows[i][4].ToString() + "','" + dt_Orcl_Conarc.Rows[i][5].ToString() + "','" + dt_Orcl_Conarc.Rows[i][6].ToString() + "','" + dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][8].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][9].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][10].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][11].ToString() + "','" + dt_Orcl_Conarc.Rows[i][12].ToString() + "','" + dt_Orcl_Conarc.Rows[i][13].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][14].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i][15].ToString() + "','" + dt_Orcl_Conarc.Rows[i]["Tap_Start"].ToString().Replace(",", ".") + "','" + dt_Orcl_Conarc.Rows[i]["Tap_end"].ToString().Replace(",", ".") + "')";
                        StrInsertQuery = StrInsertQuery + StrUpdateDefault;
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                    else
                    {
                        //string StrUpdateQuery = "Update OP_CONARC_STATUS set UNIT='" + dt_Orcl_Conarc.Rows[i][0].ToString() + "',HEAT_ID='" + dt_Orcl_Conarc.Rows[i][1].ToString() + "',TREAT_ID='" + dt_Orcl_Conarc.Rows[i][2].ToString() + "',STEL_GRADE_CODE='" + dt_Orcl_Conarc.Rows[i][3].ToString() + "',HEATING_MODE='" + dt_Orcl_Conarc.Rows[i][4].ToString() + "',STATUS='" + dt_Orcl_Conarc.Rows[i][5].ToString() + "',STATUS_DESC='" + dt_Orcl_Conarc.Rows[i][6].ToString() + "',START_T='" + dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".") + "',END_T='" + dt_Orcl_Conarc.Rows[i][8].ToString().Replace(",", ".") + "',ARCING_START_TIME='" + dt_Orcl_Conarc.Rows[i][9].ToString().Replace(",", ".") + "',ARCING_END_TIME='" + dt_Orcl_Conarc.Rows[i][10].ToString().Replace(",", ".") + "',POWER_ON_DUR='" + dt_Orcl_Conarc.Rows[i][11].ToString() + "',TOTAL_ELECT_EGY='" + dt_Orcl_Conarc.Rows[i][12].ToString() + "',BLOW_START_TIME='" + dt_Orcl_Conarc.Rows[i][13].ToString().Replace(",", ".") + "',BLOW_END_TIME='" + dt_Orcl_Conarc.Rows[i][14].ToString().Replace(",", ".") + "',O2_BLOW_DUR='" + dt_Orcl_Conarc.Rows[i][15].ToString() + "' where SLNo='" + dt_MIS_Conarc.Rows[0][0].ToString() + "'";
                        string StrUpdateQuery = "Update OP_CONARC_STATUS set UNIT='" + dt_Orcl_Conarc.Rows[i][0].ToString() + "',HEAT_ID='" + dt_Orcl_Conarc.Rows[i][1].ToString() + "',TREAT_ID='" + dt_Orcl_Conarc.Rows[i][2].ToString() + "',STEL_GRADE_CODE='" + dt_Orcl_Conarc.Rows[i][3].ToString() + "',HEATING_MODE='" + dt_Orcl_Conarc.Rows[i][4].ToString() + "',STATUS='" + dt_Orcl_Conarc.Rows[i][5].ToString() + "',STATUS_DESC='" + dt_Orcl_Conarc.Rows[i][6].ToString() + "',START_T='" + dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".") + "',END_T='" + dt_Orcl_Conarc.Rows[i][8].ToString().Replace(",", ".") + "',ARCING_START_TIME='" + dt_Orcl_Conarc.Rows[i][9].ToString().Replace(",", ".") + "',ARCING_END_TIME='" + dt_Orcl_Conarc.Rows[i][10].ToString().Replace(",", ".") + "',POWER_ON_DUR='" + dt_Orcl_Conarc.Rows[i][11].ToString() + "',TOTAL_ELECT_EGY='" + dt_Orcl_Conarc.Rows[i][12].ToString() + "',BLOW_START_TIME='" + dt_Orcl_Conarc.Rows[i][13].ToString().Replace(",", ".") + "',BLOW_END_TIME='" + dt_Orcl_Conarc.Rows[i][14].ToString().Replace(",", ".") + "',O2_BLOW_DUR='" + dt_Orcl_Conarc.Rows[i][15].ToString() + "',Tap_Start='" + dt_Orcl_Conarc.Rows[i]["Tap_Start"].ToString().Replace(",", ".") + "',Tap_end='" + dt_Orcl_Conarc.Rows[i]["Tap_end"].ToString().Replace(",", ".") + "' where SLNo='" + dt_MIS_Conarc.Rows[0][0].ToString().Replace(",", ".") + "'";
                        StrUpdateQuery = StrUpdateQuery + StrUpdateDefault;
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);

                    }
                }


            }

        }
        private void FnSMC_HMD()
        {

            string StrHEATID = "";
            string StrTREATID = "";
            string StrStatus = "";
            string StrQueryOrclHMD = "SELECT pp.PLANT  || pp.PLANTNO   AS UNIT,  pp.HEATID_CUST  AS HEAT_ID,  pp.TREATID_CUST AS TREATID,  heat.HEATSTATUS STATUS,  heat.HEATSTATUSDESC STATUS_DESC,  SUBSTR(php.HM_WEIGHT_START, 0, 19) WEIGHT,  php.TREAT_TYPE_OPER as TREATMENT_TYPE,     pp.ACTTREATSTART,  php.INJECT_START_TIME,  php.INJECT_END_TIME,   pp.ACTTREATEND,  nvl(php.TEMP_INITIAL,0) as TEMP_INITIAL,  nvl(php.AIM_S_GRADE,0) as AIM_S_GRADE,  nvl(php.HM_TEMP_END,0) as HM_TEMP_END FROM PP_HEAT ph,  PDH_HEAT_DATA php,  PP_HEAT_PLANT pp,  (SELECT s.HEATSTATUSDESC, s.HEATSTATORDER, s.HEATSTATUS FROM GC_HEAT_STAT s   ) heat WHERE pp.HEATID_CUST   = ph.HEATID_CUST AND heat.HEATSTATORDER = ph.HEATSTATORDER AND pp.HEATID          = php.HEATID(+) AND pp.TREATID         = php.TREATID(+) AND pp.PLANT           = php.PLANT(+) AND (pp.EXPIRATIONDATE = 'VALID' AND pp.METAL_TYPE      = 'HM' AND ph.METAL_TYPE      = 'HM') AND pp.ACTTREATEND        IS NULL";
            DataTable dt_Orcl_HMD = new DataTable();
            dt_Orcl_HMD = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclHMD);
            if (dt_Orcl_HMD.Rows.Count > 0)
            {
                string StrUpdateDefault = ";UPDATE OP_HMD_STATUS set ACT_TREAT_START=null where ACT_TREAT_START='1900-01-01 00:00:00.000' ; UPDATE OP_HMD_STATUS set INJECT_START_TIME=null where INJECT_START_TIME='1900-01-01 00:00:00.000' ; UPDATE OP_HMD_STATUS set INJECT_END_TIME=null where INJECT_END_TIME='1900-01-01 00:00:00.000' ; UPDATE OP_HMD_STATUS set ACT_TREAT_END=null where ACT_TREAT_END='1900-01-01 00:00:00.000'";
                for (int i = 0; i < dt_Orcl_HMD.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_HMD.Rows[i][1].ToString();
                    StrTREATID = dt_Orcl_HMD.Rows[i][2].ToString();
                    StrStatus = dt_Orcl_HMD.Rows[i][3].ToString();
                    DataTable dt_MIS_HMD = new DataTable();
                    string StrQueryMISHMD = "select SLNo,HEAT_ID,TREAT_ID,STATUS from OP_HMD_STATUS WHERE HEAT_ID='" + StrHEATID + "' AND TREAT_ID='" + StrTREATID + "' AND STATUS='" + StrStatus + "'";
                    dt_MIS_HMD = clsObj.DBSelectQueryMIS_Table(StrQueryMISHMD);
                    if (dt_MIS_HMD.Rows.Count == 0)
                    {
                        string StrInsertQuery = "insert into OP_HMD_STATUS(UNIT,HEAT_ID,TREAT_ID,STATUS,STATUS_DESC,WEIGHT,TREATMENT_TYPE,ACT_TREAT_START,INJECT_START_TIME,INJECT_END_TIME,ACT_TREAT_END,TEMP_INITIAL,AIM_S_GRAGE,HM_TEMP_END)values('" + dt_Orcl_HMD.Rows[i][0].ToString() + "','" + dt_Orcl_HMD.Rows[i][1].ToString() + "','" + dt_Orcl_HMD.Rows[i][2].ToString() + "','" + dt_Orcl_HMD.Rows[i][3].ToString() + "','" + dt_Orcl_HMD.Rows[i][4].ToString() + "','" + dt_Orcl_HMD.Rows[i][5].ToString() + "','" + dt_Orcl_HMD.Rows[i][6].ToString() + "','" + dt_Orcl_HMD.Rows[i][7].ToString().Replace(",", ".") + "','" + dt_Orcl_HMD.Rows[i][8].ToString().Replace(",", ".") + "','" + dt_Orcl_HMD.Rows[i][9].ToString().Replace(",", ".") + "','" + dt_Orcl_HMD.Rows[i][10].ToString().Replace(",", ".") + "','" + dt_Orcl_HMD.Rows[i][11].ToString() + "','" + dt_Orcl_HMD.Rows[i][12].ToString() + "','" + dt_Orcl_HMD.Rows[i][13].ToString() + "')";
                        StrInsertQuery = StrInsertQuery + StrUpdateDefault;
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                    else
                    {
                        string StrUpdateQuery = "Update OP_HMD_STATUS set UNIT='" + dt_Orcl_HMD.Rows[i][0].ToString() + "',HEAT_ID='" + dt_Orcl_HMD.Rows[i][1].ToString() + "',TREAT_ID='" + dt_Orcl_HMD.Rows[i][2].ToString() + "',STATUS='" + dt_Orcl_HMD.Rows[i][3].ToString() + "',STATUS_DESC='" + dt_Orcl_HMD.Rows[i][4].ToString() + "',WEIGHT='" + dt_Orcl_HMD.Rows[i][5].ToString() + "',TREATMENT_TYPE='" + dt_Orcl_HMD.Rows[i][6].ToString() + "',ACT_TREAT_START='" + dt_Orcl_HMD.Rows[i][7].ToString().Replace(",", ".") + "',INJECT_START_TIME='" + dt_Orcl_HMD.Rows[i][8].ToString().Replace(",", ".") + "',INJECT_END_TIME='" + dt_Orcl_HMD.Rows[i][9].ToString().Replace(",", ".") + "',ACT_TREAT_END='" + dt_Orcl_HMD.Rows[i][10].ToString().Replace(",", ".") + "',TEMP_INITIAL='" + dt_Orcl_HMD.Rows[i][11].ToString() + "',AIM_S_GRAGE='" + dt_Orcl_HMD.Rows[i][12].ToString() + "',HM_TEMP_END='" + dt_Orcl_HMD.Rows[i][13].ToString() + "' where SLNo='" + dt_MIS_HMD.Rows[0][0].ToString() + "'";
                        StrUpdateQuery = StrUpdateQuery + StrUpdateDefault;
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);

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
            // @MAnish 08-07-2014 changed select statment of ord block for actual grade  --> (SELECT UNIQUE s.ORD_STAT_DESC,     s.ORD_STAT_NO,     s.ORD_STAT_CODE,     o.PRODORDERID,     o.STEELGRADECODE   FROM GC_PRD_STAT s,     PP_ORDER o   WHERE o.ORD_STAT_NO = s.ORD_STAT_NO   ) ord 
            string StrQueryOrclLF = "SELECT p.PLANT  || p.PLANTNO AS UNIT,  p.HEATID_CUST HEAT_ID,  p.TREATID_CUST TREATID,  grade.STEELGRADECODE,  status.HEATSTATUS,  status.HEATSTATUSDESC STATUS_DESC,  SUBSTR(DECODE(p.ACTTREATSTART, NULL, p.PLANTREATSTART, p.ACTTREATSTART), 0, 19) START_T,  SUBSTR(DECODE(p.ACTTREATEND, NULL, p.ACTTREATEND, p.ACTTREATEND), 0, 19) END_T,  de.LADLENO,  de.FIRST_POWER_ON,  de.LAST_POWER_OFF,  de.POWER_ON_DUR FROM PP_HEAT_PLANT p,   PP_HEAT ph,   (SELECT g.STEELGRADECODE, g.STEELGRADECODEDESC FROM PP_GRADE g   ) grade,   (SELECT UNIQUE s.ORD_STAT_DESC,     s.ORD_STAT_NO,     s.ORD_STAT_CODE,     o.PRODORDERID,     Hd.Steelgradecode_Act as STEELGRADECODE   FROM GC_PRD_STAT s,     PP_ORDER o, Pd_Heatdata hd   WHERE o.ORD_STAT_NO = s.ORD_STAT_NO and O.Prodorderid = Hd.Prodorderid_Plan   ) ord,   (SELECT st.HEATSTATORDER,     st.HEATSTATUS,     st.HEATSTATUSDESC   FROM GC_HEAT_STAT st   ) status,   PDL_HEAT_DATA de WHERE ph.HEATID_CUST(+)     = p.HEATID_CUST AND grade.STEELGRADECODE(+) = ord.STEELGRADECODE AND ord.PRODORDERID(+)      = ph.PRODORDERID AND de.HEATID(+)            = p.HEATID AND de.TREATID(+)           = p.TREATID AND status.HEATSTATORDER(+) = ph.HEATSTATORDER AND (p.EXPIRATIONDATE        = 'VALID' AND p.METAL_TYPE            = 'Steel' AND ph.METAL_TYPE           = 'Steel' AND SUBSTR(p.ACTTREATSTART,1,10)='" + System.DateTime.Now.Date.ToString("yyyy-MM-dd") + "') ORDER BY p.acttreatstart DESC";
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
                    StrStartDate = dt_Orcl_LF.Rows[i]["START_T"].ToString();
                    DataTable dt_MIS_LF = new DataTable();
                    string StrQueryMISLF = "select SLNo,HEAT_ID,TREAT_ID,STATUS,START_T from OP_LF_STATUS WHERE HEAT_ID='" + StrHEATID + "' AND TREAT_ID='" + StrTREATID + "' AND STATUS='" + StrStatus + "' and convert(varchar(10),START_T,120)=convert(varchar(10),'" + StrStartDate + "',120)";
                    dt_MIS_LF = clsObj.DBSelectQueryMIS_Table(StrQueryMISLF);
                    if (dt_MIS_LF.Rows.Count == 0)
                    {
                        string StrInsertQuery = "insert into OP_LF_STATUS(UNIT,HEAT_ID,TREAT_ID,STEL_GRADE_CODE,STATUS,STATUS_DESC,START_T,END_T,LADLENO,FIRST_POWER_ON,LAST_POWER_OFF,POWER_ON_DUR)values('" + dt_Orcl_LF.Rows[i][0].ToString() + "','" + dt_Orcl_LF.Rows[i][1].ToString() + "','" + dt_Orcl_LF.Rows[i][2].ToString() + "','" + dt_Orcl_LF.Rows[i][3].ToString() + "','" + dt_Orcl_LF.Rows[i][4].ToString() + "','" + dt_Orcl_LF.Rows[i][5].ToString() + "','" + dt_Orcl_LF.Rows[i][6].ToString().Replace(",", ".") + "','" + dt_Orcl_LF.Rows[i][7].ToString().Replace(",", ".") + "','" + dt_Orcl_LF.Rows[i][8].ToString() + "','" + dt_Orcl_LF.Rows[i][9].ToString().Replace(",", ".") + "','" + dt_Orcl_LF.Rows[i][10].ToString().Replace(",", ".") + "','" + dt_Orcl_LF.Rows[i][11].ToString() + "')";
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
        private void FnSMC_RH()
        {

            string StrHEATID = "";
            string StrTREATID = "";
            string StrStatus = "";
            string StrQueryOrclRH = "SELECT Distinct p.PLANT  ||p.PLANTNO AS UNIT,  p.HEATID_CUST HEAT_ID,  p.TREATID_CUST TREATID,  ph.PRODORDERID,  grade.STEELGRADECODE,  grade.STEELGRADECODEDESC,  ord.ORD_STAT_CODE ORD_STAT,  ord.ORD_STAT_DESC ORD_STAT_DESC,  status.HEATSTATUS STATUS,  status.HEATSTATUSDESC STATUS_DESC,  SUBSTR(DECODE(p.ACTTREATSTART, NULL, p.PLANTREATSTART, p.ACTTREATSTART),0,19) START_T,  DECODE(p.ACTTREATSTART, NULL, '0', '1') START_T_B,  SUBSTR(DECODE(p.ACTTREATEND, NULL, p.PLANTREATEND, p.ACTTREATEND),0,19) END_T,  DECODE(p.ACTTREATEND, NULL, '0', '1') END_T_B,  DECODE(p.ACTTREATEND, NULL, phase.STEELMASS, rep.FINALWEIGHT) WEIGHT,  DECODE(p.ACTTREATEND, NULL, '1', '0') WEIGHT_B,  DECODE(p.ACTTREATEND, NULL, phase.TEMP, rep.FINALTEMPCALC) TEMP,  DECODE(p.ACTTREATEND, NULL, '1', '0') TEMP_B,  DECODE(p.ACTTREATEND, NULL, oxygen.RESULTVALUE, de.total_o2_cons) OXYGEN,  DECODE(p.ACTTREATEND, NULL, '1', '0') OXYGEN_B FROM PP_HEAT_PLANT p,  PP_HEAT ph,  (SELECT g.STEELGRADECODE, g.STEELGRADECODEDESC FROM PP_GRADE g  ) grade,  (SELECT UNIQUE s.ORD_STAT_DESC,    s.ORD_STAT_NO,    s.ORD_STAT_CODE,    o.PRODORDERID,    o.STEELGRADECODE  FROM GC_PRD_STAT s,    PP_ORDER o  WHERE o.ORD_STAT_NO = s.ORD_STAT_NO  ) ord,  (SELECT st.HEATSTATORDER,    st.HEATSTATUS,    st.HEATSTATUSDESC  FROM GC_HEAT_STAT st  ) status,  (SELECT h.FINALWEIGHT,    h.HEATID,    h.TREATID,    h.FINALTEMPCALC  FROM PD_REPORT h  WHERE h.PLANT= 'RH'  ) rep,  (SELECT r.HEATID,    r.TREATID,    r.STEELMASS,    r.TEMP  FROM PD_PHASE_RES r  WHERE r.PLANT     = 'RH'  AND r.RES_PHASENO =    (SELECT p.RES_PHASENO    FROM GC_PHASE p    WHERE p.PHASENAME = 'Alloying'    AND p.PLANT       = 'RH'    )  ) phase,  (SELECT pl.HEATID,    pl.TREATID,    pl.RESULTVALUE  FROM PD_PHASE_RES_PLANT pl  WHERE pl.PLANT     = 'RH'  AND pl.RES_PHASENO =   (SELECT p.RES_PHASENO    FROM GC_PHASE p    WHERE p.PHASENAME = 'Alloying'    AND p.PLANT       = 'RH'    )  AND pl.RESULTVALNO =    (SELECT rp.RESULTVALNO    FROM GC_PHASE_RES_PLANT rp    WHERE rp.PLANT       = 'RH'    AND rp.RESULTVARNAME = 'O2Fl_L'    )  ) oxygen,  PDR_HEAT_DATA de WHERE p.PLANT               = 'RH'AND p.EXPIRATIONDATE        ='VALID'AND p.metal_type            = 'Steel'AND ph.metal_type           = 'Steel'AND ph.HEATID_CUST(+)       = p.HEATID_CUST AND grade.STEELGRADECODE(+) = ord.STEELGRADECODE AND ord.PRODORDERId(+)      = ph.PRODORDERID AND phase.HEATID(+)         = p.HEATID AND phase.TREATID(+)        = p.TREATID AND rep.HEATID(+)           = p.HEATID AND rep.TREATID(+)          = p.TREATID AND oxygen.HEATID(+)        = p.HEATID AND oxygen.TREATID(+)       = p.TREATID AND de.HEATID(+)            = p.HEATID AND de.TREATID(+)           = p.TREATID AND status.HEATSTATORDER(+) = ph.HEATSTATORDER AND p.ACTTREATSTART        IS NOT NULL AND p.ACTTREATEND          IS NULL";
            DataTable dt_Orcl_RH = new DataTable();
            dt_Orcl_RH = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclRH);
            if (dt_Orcl_RH.Rows.Count > 0)
            {
                for (int i = 0; i < dt_Orcl_RH.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_RH.Rows[i][1].ToString();
                    StrTREATID = dt_Orcl_RH.Rows[i][2].ToString();
                    StrStatus = dt_Orcl_RH.Rows[i][8].ToString();
                    DataTable dt_MIS_RH = new DataTable();
                    string StrQueryMISRH = "select HEAT_ID,TREAT_ID,STATUS from OP_RH_STATUS WHERE HEAT_ID='" + StrHEATID + "' AND TREAT_ID='" + StrTREATID + "' AND STATUS='" + StrStatus + "'";
                    dt_MIS_RH = clsObj.DBSelectQueryMIS_Table(StrQueryMISRH);
                    if (dt_MIS_RH.Rows.Count == 0)
                    {
                        string StrInsertQuery = "insert into OP_RH_STATUS(UNIT,HEAT_ID,TREAT_ID,PROD_ORDER_ID,STEL_GRADE_CODE,STEL_GRADE_CODE_DESC,ORD_STAT,ORD_STAT_DESC,STATUS,STATUS_DESC,START_T,START_T_B,END_T,END_T_B,WEIGHT,WEIGHT_B,TEMP,TEMP_B,OXYGEN,OXYGEN_B)values('" + dt_Orcl_RH.Rows[i][0].ToString() + "','" + dt_Orcl_RH.Rows[i][1].ToString() + "','" + dt_Orcl_RH.Rows[i][2].ToString() + "','" + dt_Orcl_RH.Rows[i][3].ToString() + "','" + dt_Orcl_RH.Rows[i][4].ToString() + "','" + dt_Orcl_RH.Rows[i][5].ToString() + "','" + dt_Orcl_RH.Rows[i][6].ToString() + "','" + dt_Orcl_RH.Rows[i][7].ToString() + "','" + dt_Orcl_RH.Rows[i][8].ToString() + "','" + dt_Orcl_RH.Rows[i][9].ToString() + "','" + dt_Orcl_RH.Rows[i][10].ToString() + "','" + dt_Orcl_RH.Rows[i][11].ToString() + "','" + dt_Orcl_RH.Rows[i][12].ToString() + "','" + dt_Orcl_RH.Rows[i][13].ToString() + "','" + dt_Orcl_RH.Rows[i][14].ToString() + "','" + dt_Orcl_RH.Rows[i][15].ToString() + "','" + dt_Orcl_RH.Rows[i][16].ToString() + "','" + dt_Orcl_RH.Rows[i][17].ToString() + "','" + dt_Orcl_RH.Rows[i][18].ToString() + "','" + dt_Orcl_RH.Rows[i][19].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }

        }
        # region "SMS HEAT TRACKING SYSTEM"

        private void smsHeatTracking_Tick(object sender, EventArgs e)
        {
            string StrPLANT = "";
            string StrPLANTNO = "";
            string StrMIS_PLANTNO = "";
            string StrDate = System.DateTime.Now.ToString("yyyy-MM-dd");
            StrPLANT = "CON"; StrPLANTNO = "1"; StrMIS_PLANTNO = "3";
            FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            StrPLANT = "CON"; StrPLANTNO = "2"; StrMIS_PLANTNO = "4";
            FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            StrPLANT = "LF"; StrPLANTNO = "1"; StrMIS_PLANTNO = "9";
            FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            StrPLANT = "LF"; StrPLANTNO = "2"; StrMIS_PLANTNO = "10";
            FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            ///////////////StrPLANT = "LF"; StrPLANTNO = "3"; StrMIS_PLANTNO = "11";
            ///////////////FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            //FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            StrPLANT = "HMD"; StrPLANTNO = "1"; StrMIS_PLANTNO = "1";
            FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            //FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            StrPLANT = "HMD"; StrPLANTNO = "2"; StrMIS_PLANTNO = "2";
            FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            //FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            StrPLANT = "CASTER-I"; StrPLANTNO = "1"; StrMIS_PLANTNO = "15";
            FnSMS_CasterI_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            FnSMS_CasterI_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            StrPLANT = "CASTER-II"; StrPLANTNO = "2"; StrMIS_PLANTNO = "16";
            FnSMS_CasterII_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            FnSMS_CasterII_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            StrPLANT = "CASTER-III"; StrPLANTNO = "3"; StrMIS_PLANTNO = "17";
            FnSMS_CasterIII_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            FnSMS_CasterIII_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
            //string StrUpdateQuery = "update  smsHeatTracker set actEnd='24' where actStart > actEnd and DateStamp<'" + StrDate + "'";
            //bool StrUStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
            FnSMS_SMCDB_HeatBlow_Arc(StrDate);
            FnSMS_SMCDB_HeatBlow_Arc_LF(StrDate);

        }
        private void FnSMS_CasterI_UpdateHeatEndTime(string StrPLANT, string StrPLANTNO, string StrMIS_PLANTNO, string StrDate)
        {
            string StrUptQuery = "UPDATE t   SET t.actEnnd = c.actStrrt,t.actEnd=c.actEnd,t.actEnds=c.actEnd FROM smsHeatTracker t INNER JOIN (SELECT a.plantNo,a.heatNo,a.DateStamp,a.actStrt as actStart,b.actStrt,dateadd(minute,-1,b.actStrt) as actStrrt ,replace(convert(char(5), dateadd(minute,-1,b.actStrt), 108),':','.') as actEnd FROM ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt) , * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) b right JOIN ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt)+1, * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) a ON b.Id = a.Id) c ON  c.heatNo=t.heatNo and c.plantNo=t.plantNo and c.DateStamp=t.DateStamp and c.actStart=t.actStrt";
            bool StrUpdtStatus = clsObj.DBInsertUpdateDeleteMIS(StrUptQuery);
            return;

            DataTable dt_MIS_TimeUpdt = new DataTable();
            //string StrQueryEndTmeUpdate = "select heatNo from smsHeatTracker where convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + System.DateTime.Now.ToString("yyyy-MM-dd") + "',120) and  (actEnd=0 or actEnd is null)  and plantNo='" + StrMIS_PLANTNO + "'";
            string StrQueryEndTmeUpdate = "select heatNo,DateStamp from smsHeatTracker where  actEnnd is null and plantNo='" + StrMIS_PLANTNO + "'";
            dt_MIS_TimeUpdt = clsObj.DBSelectQueryMIS_Table(StrQueryEndTmeUpdate);
            if (dt_MIS_TimeUpdt.Rows.Count <= 0)
            {
                return;
            }
            for (int j = 0; j < dt_MIS_TimeUpdt.Rows.Count; j++)
            {
                //string StrQueryCasterI = "SELECT heatid as HeatID_Cust,ladle_no as Laddle_no,ladle_open_time  AS Cast_Start_time,ladle_close_time AS Cast_End_time,SUBSTR(ladle_open_time,1,10) as Act_Start_Dt FROM pdc_heat where heatid='" + dt_MIS_TimeUpdt.Rows[j][0].ToString() + "'";
                string StrQueryCasterI = "SELECT heatid as HeatID_Cust,ladle_no as Laddle_no,ladle_open_time  AS Cast_Start_time,ladle_close_time AS Cast_End_time,SUBSTR(ladle_open_time,1,10) as Act_Start_Dt,SUBSTR(ladle_Close_time,1,10) as Act_End_Dt FROM pdc_heat where heatid='" + dt_MIS_TimeUpdt.Rows[j][0].ToString() + "' and SUBSTR(ladle_open_time,1,10)='" + Convert.ToDateTime(dt_MIS_TimeUpdt.Rows[j][1].ToString()).ToString("yyyy-MM-dd") + "' and (ladle_Close_time is not null)";
                DataTable dt_Orcl_Conarc = new DataTable();
                dt_Orcl_Conarc = clsObj.DBSelectQueryCASTERI_Table(StrQueryCasterI);
                if (dt_Orcl_Conarc.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                    {
                        string StrHEATID = "";
                        string StrActStart = null;
                        string StrActEnd = null;
                        string StrActStartDate = null;
                        string StrActEndDate = null;
                        object StrLadleNo = DBNull.Value;
                        object StrActStartDte = DBNull.Value;
                        object StrActEndDte = DBNull.Value;
                        StrHEATID = dt_Orcl_Conarc.Rows[i]["HeatID_Cust"].ToString();
                        if (dt_Orcl_Conarc.Rows[i]["Laddle_no"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Laddle_no"].ToString() != string.Empty)
                        {
                            StrLadleNo = dt_Orcl_Conarc.Rows[i]["Laddle_no"].ToString();
                        }
                        if (dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString() != string.Empty)
                        {
                            StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                            StrActStartDate = dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString().Replace(",", ".");
                        }
                        if (dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString() != string.Empty)
                        {
                            StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                            StrActEndDate = dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString().Replace(",", ".");
                        }
                        if (dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != string.Empty)
                        {
                            StrActStartDte = dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString();
                        }
                        else
                        { StrActStartDte = DBNull.Value; }
                        if (dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString() != string.Empty)
                        {
                            StrActEndDte = dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString();
                        }
                        else
                        { StrActEndDte = DBNull.Value; }
                        string StrUpdate_Insert_Query = "";
                        //StrUpdate_Insert_Query = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,DateStamp,LadleNo,actStrt,actEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','0','" + StrActEnd + "','','" + StrActEndDte + "','" + StrLadleNo + "','" + StrActStartDate + "','" + StrActEndDate + "')";
                        StrUpdate_Insert_Query = "update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',LadleNo='" + StrLadleNo + "',actStrt='" + StrActStartDate + "',actEnnd='" + StrActEndDate + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "' and DateStamp='" + StrActStartDte + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdate_Insert_Query);

                    }
                }
            }

        }
        private void FnSMS_CasterII_UpdateHeatEndTime(string StrPLANT, string StrPLANTNO, string StrMIS_PLANTNO, string StrDate)
        {
            string StrUptQuery = "UPDATE t   SET t.actEnnd = c.actStrrt,t.actEnd=c.actEnd,t.actEnds=c.actEnd FROM smsHeatTracker t INNER JOIN (SELECT a.plantNo,a.heatNo,a.DateStamp,a.actStrt as actStart,b.actStrt,dateadd(minute,-1,b.actStrt) as actStrrt ,replace(convert(char(5), dateadd(minute,-1,b.actStrt), 108),':','.') as actEnd FROM ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt) , * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) b right JOIN ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt)+1, * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) a ON b.Id = a.Id) c ON  c.heatNo=t.heatNo and c.plantNo=t.plantNo and c.DateStamp=t.DateStamp and c.actStart=t.actStrt";
            bool StrUpdtStatus = clsObj.DBInsertUpdateDeleteMIS(StrUptQuery);
            return;
            DataTable dt_MIS_TimeUpdt = new DataTable();
            //string StrQueryEndTmeUpdate = "select heatNo from smsHeatTracker where convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + System.DateTime.Now.ToString("yyyy-MM-dd") + "',120) and  (actEnd=0 or actEnd is null)  and plantNo='" + StrMIS_PLANTNO + "'";
            string StrQueryEndTmeUpdate = "select heatNo,DateStamp from smsHeatTracker where actEnnd is null and plantNo='" + StrMIS_PLANTNO + "'";
            dt_MIS_TimeUpdt = clsObj.DBSelectQueryMIS_Table(StrQueryEndTmeUpdate);
            if (dt_MIS_TimeUpdt.Rows.Count <= 0)
            {
                return;
            }
            for (int j = 0; j < dt_MIS_TimeUpdt.Rows.Count; j++)
            {
                //string StrQueryCasterI = "SELECT heatid as HeatID_Cust,ladle_no as Laddle_no,ladle_open_time  AS Cast_Start_time,ladle_close_time AS Cast_End_time,SUBSTR(ladle_open_time,1,10) as Act_Start_Dt FROM pdc_heat where heatid='" + dt_MIS_TimeUpdt.Rows[j][0].ToString() + "'";
                string StrQueryCasterI = "SELECT heatid as HeatID_Cust,ladle_no as Laddle_no,ladle_open_time  AS Cast_Start_time,ladle_close_time AS Cast_End_time,SUBSTR(ladle_open_time,1,10) as Act_Start_Dt,SUBSTR(ladle_Close_time,1,10) as Act_End_Dt FROM pdc_heat where heatid='" + dt_MIS_TimeUpdt.Rows[j][0].ToString() + "' and SUBSTR(ladle_open_time,1,10)='" + Convert.ToDateTime(dt_MIS_TimeUpdt.Rows[j][1].ToString()).ToString("yyyy-MM-dd") + "' and (ladle_Close_time is not null)";
                DataTable dt_Orcl_Conarc = new DataTable();
                dt_Orcl_Conarc = clsObj.DBSelectQueryCASTERII_Table(StrQueryCasterI);
                if (dt_Orcl_Conarc.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                    {
                        string StrHEATID = "";
                        string StrActStart = null;
                        string StrActEnd = null;
                        string StrActStartDate = null;
                        string StrActEndDate = null;
                        object StrLadleNo = DBNull.Value;
                        object StrActStartDte = DBNull.Value;
                        object StrActEndDte = DBNull.Value;
                        StrHEATID = dt_Orcl_Conarc.Rows[i]["HeatID_Cust"].ToString();
                        if (dt_Orcl_Conarc.Rows[i]["Laddle_no"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Laddle_no"].ToString() != string.Empty)
                        {
                            StrLadleNo = dt_Orcl_Conarc.Rows[i]["Laddle_no"].ToString();
                        }
                        if (dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString() != string.Empty)
                        {
                            StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                            StrActStartDate = dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString().Replace(",", ".");
                        }
                        if (dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString() != string.Empty)
                        {
                            StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                            StrActEndDate = dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString().Replace(",", ".");
                        }
                        if (dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != string.Empty)
                        {
                            StrActStartDte = dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString();
                        }
                        else
                        { StrActStartDte = DBNull.Value; }
                        if (dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString() != string.Empty)
                        {
                            StrActEndDte = dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString();
                        }
                        else
                        { StrActEndDte = DBNull.Value; }
                        string StrUpdate_Insert_Query = "";
                        //StrUpdate_Insert_Query = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,DateStamp,LadleNo,actStrt,actEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','0','" + StrActEnd + "','','" + StrActEndDte + "','" + StrLadleNo + "','" + StrActStartDate + "','" + StrActEndDate + "')";
                        StrUpdate_Insert_Query = "update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',LadleNo='" + StrLadleNo + "',actStrt='" + StrActStartDate + "',actEnnd='" + StrActEndDate + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "' and DateStamp='" + StrActStartDte + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdate_Insert_Query);

                    }
                }
            }

        }
        private void FnSMS_CasterIII_UpdateHeatEndTime(string StrPLANT, string StrPLANTNO, string StrMIS_PLANTNO, string StrDate)
        {
            string StrUptQuery = "UPDATE t   SET t.actEnnd = c.actStrrt,t.actEnd=c.actEnd,t.actEnds=c.actEnd FROM smsHeatTracker t INNER JOIN (SELECT a.plantNo,a.heatNo,a.DateStamp,a.actStrt as actStart,b.actStrt,dateadd(minute,-1,b.actStrt) as actStrrt ,replace(convert(char(5), dateadd(minute,-1,b.actStrt), 108),':','.') as actEnd FROM ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt) , * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) b right JOIN ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt)+1, * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) a ON b.Id = a.Id) c ON  c.heatNo=t.heatNo and c.plantNo=t.plantNo and c.DateStamp=t.DateStamp and c.actStart=t.actStrt";
            bool StrUpdtStatus = clsObj.DBInsertUpdateDeleteMIS(StrUptQuery);
            return;

            DataTable dt_MIS_TimeUpdt = new DataTable();
            //string StrQueryEndTmeUpdate = "select heatNo from smsHeatTracker where convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + System.DateTime.Now.ToString("yyyy-MM-dd") + "',120) and  (actEnd=0 or actEnd is null)  and plantNo='" + StrMIS_PLANTNO + "'";
            string StrQueryEndTmeUpdate = "select heatNo,DateStamp from smsHeatTracker where actEnnd is null and plantNo='" + StrMIS_PLANTNO + "'";
            dt_MIS_TimeUpdt = clsObj.DBSelectQueryMIS_Table(StrQueryEndTmeUpdate);
            if (dt_MIS_TimeUpdt.Rows.Count <= 0)
            {
                return;
            }
            for (int j = 0; j < dt_MIS_TimeUpdt.Rows.Count; j++)
            {
                //string StrQueryCasterI = "SELECT heatid as HeatID_Cust,ladle_no as Laddle_no,ladle_open_time  AS Cast_Start_time,ladle_close_time AS Cast_End_time,SUBSTR(ladle_open_time,1,10) as Act_Start_Dt FROM pdc_heat where heatid='" + dt_MIS_TimeUpdt.Rows[j][0].ToString() + "'";
                string StrQueryCasterI = "SELECT heatid as HeatID_Cust,ladle_no as Laddle_no,ladle_open_time  AS Cast_Start_time,ladle_close_time AS Cast_End_time,SUBSTR(ladle_open_time,1,10) as Act_Start_Dt,SUBSTR(ladle_Close_time,1,10) as Act_End_Dt FROM pdc_heat where heatid='" + dt_MIS_TimeUpdt.Rows[j][0].ToString() + "' and SUBSTR(ladle_open_time,1,10)='" + Convert.ToDateTime(dt_MIS_TimeUpdt.Rows[j][1].ToString()).ToString("yyyy-MM-dd") + "' and (ladle_Close_time is not null)";
                DataTable dt_Orcl_Conarc = new DataTable();
                dt_Orcl_Conarc = clsObj.DBSelectQueryCASTERIII_Table(StrQueryCasterI);
                if (dt_Orcl_Conarc.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                    {
                        string StrHEATID = "";
                        string StrActStart = null;
                        string StrActEnd = null;
                        string StrActStartDate = null;
                        string StrActEndDate = null;
                        object StrLadleNo = DBNull.Value;
                        object StrActStartDte = DBNull.Value;
                        object StrActEndDte = DBNull.Value;
                        StrHEATID = dt_Orcl_Conarc.Rows[i]["HeatID_Cust"].ToString();
                        if (dt_Orcl_Conarc.Rows[i]["Laddle_no"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Laddle_no"].ToString() != string.Empty)
                        {
                            StrLadleNo = dt_Orcl_Conarc.Rows[i]["Laddle_no"].ToString();
                        }
                        if (dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString() != string.Empty)
                        {
                            StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                            StrActStartDate = dt_Orcl_Conarc.Rows[i]["Cast_Start_time"].ToString().Replace(",", ".");
                        }
                        if (dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString() != string.Empty)
                        {
                            StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                            StrActEndDate = dt_Orcl_Conarc.Rows[i]["Cast_End_time"].ToString().Replace(",", ".");
                        }
                        if (dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != string.Empty)
                        {
                            StrActStartDte = dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString();
                        }
                        else
                        { StrActStartDte = DBNull.Value; }
                        if (dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString() != string.Empty)
                        {
                            StrActEndDte = dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString();
                        }
                        else
                        { StrActEndDte = DBNull.Value; }
                        string StrUpdate_Insert_Query = "";
                        //StrUpdate_Insert_Query = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,DateStamp,LadleNo,actStrt,actEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','0','" + StrActEnd + "','','" + StrActEndDte + "','" + StrLadleNo + "','" + StrActStartDate + "','" + StrActEndDate + "')";
                        StrUpdate_Insert_Query = "update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',LadleNo='" + StrLadleNo + "',actStrt='" + StrActStartDate + "',actEnnd='" + StrActEndDate + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "' and DateStamp='" + StrActStartDte + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdate_Insert_Query);

                    }
                }
            }

        }

        private void FnSMS_CasterIII_HeatTracking(string StrPLANT, string StrPLANTNO, string StrMIS_PLANTNO, string StrDate)
        {

            
            //Old Code @01_08_2014 For Revert Back
            //string StrQueryCasterIII = "Select Pdc_Heat.Heatid As Heatid_Cust,Pdc_Heat.Ladle_No As Laddle_No,Pdc_Heat.Ladle_Open_Time As Cast_Start_Time,Pdc_Heat.Ladle_Close_Time As Cast_End_Time,Substr(Pdc_Heat.Ladle_Open_Time,1,10) As Act_Start_Dt,replace(Pdc_Heat.Ladle_Open_Time,',','.') As ActStrt,replace(Pdc_Heat.Ladle_Close_Time,',','.') As ActEnnd,Pdc_Heat.Grade_Code As Grade, CONCAT('L+',Pdc_Heat.Heat_Seq_No -1 ) As Heat_In_Seq,Pdc_Strand.Avg_Cast_Speed As Avg_Cast_Speed,Pdc_Strand.Cast_Len_Bgn As Cast_Start_Length,Pdc_Strand.Cast_Len_End As Cast_End_Length,CASE WHEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) > 0 THEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) ELSE NULL END as Casted_Length From Pdc_Heat,Pdc_Strand,pdc_heat_strand WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID And Pdc_Heat_Strand.Strand_Steel_Id=Pdc_Strand.Steel_Id AND SUBSTR(ladle_open_time,1,10)='" + StrDate + "' ORDER BY Ladle_Open_Time DESC";
            string StrQueryCasterIII = "Select * from (Select Pdc_Heat.Heatid As Heatid_Cust,Pdc_Heat.Ladle_No As Laddle_No,Pdc_Heat.Ladle_Open_Time As Cast_Start_Time,Pdc_Heat.Ladle_Close_Time As Cast_End_Time,Substr(Pdc_Heat.Ladle_Open_Time,1,10) As Act_Start_Dt,replace(Pdc_Heat.Ladle_Open_Time,',','.') As ActStrt,replace(Pdc_Heat.Ladle_Close_Time,',','.') As ActEnnd,Pdc_Heat.Grade_Code As Grade, CONCAT('L+',Pdc_Heat.Heat_Seq_No -1 ) As Heat_In_Seq,Pdc_Strand.Avg_Cast_Speed As Avg_Cast_Speed,Pdc_Strand.Cast_Len_Bgn As Cast_Start_Length,Pdc_Strand.Cast_Len_End As Cast_End_Length,CASE WHEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) > 0 THEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) ELSE NULL END as Casted_Length From Pdc_Heat,Pdc_Strand,pdc_heat_strand WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID And Pdc_Heat_Strand.Strand_Steel_Id=Pdc_Strand.Steel_Id AND SUBSTR(ladle_open_time,1,10)<='" + StrDate + "' ORDER BY Ladle_Open_Time DESC) where rownum <=5 ";
            
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryCASTERIII_Table(StrQueryCasterIII);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    string StrHEATID = "";
                    string StrActStart = null;
                    string StrActEnd = null;
                    object StrLadleNo = DBNull.Value;
                    object StrActStartDte = DBNull.Value;
                    object StrActStrt = DBNull.Value;
                    object StrActEnnd = DBNull.Value;
                    object StrGrade = DBNull.Value;
                    string StrAvgCastSpeed = null;
                    object StrSeqNo = DBNull.Value;
                    string StrCastLength = null;
                    object TundishTemp = DBNull.Value;
                    StrHEATID = dt_Orcl_Conarc.Rows[i][0].ToString();
                    if (dt_Orcl_Conarc.Rows[i][1].ToString() != null && dt_Orcl_Conarc.Rows[i][1].ToString() != string.Empty)
                    {
                        StrLadleNo = dt_Orcl_Conarc.Rows[i][1].ToString();
                    }
                    if (dt_Orcl_Conarc.Rows[i][2].ToString() != null && dt_Orcl_Conarc.Rows[i][2].ToString() != string.Empty)
                    {
                        StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][2].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                    }
                    if (dt_Orcl_Conarc.Rows[i][3].ToString() != null && dt_Orcl_Conarc.Rows[i][3].ToString() != string.Empty)
                    {
                        StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][3].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                    }
                    if (dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != string.Empty)
                    {
                        StrActStartDte = dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString();
                    }
                    else
                    { StrActStartDte = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["ACTSTRT"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ACTSTRT"].ToString() != string.Empty)
                    {
                        StrActStrt = dt_Orcl_Conarc.Rows[i]["ACTSTRT"].ToString();
                    }
                    else
                    { StrActStrt = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["ACTENND"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ACTENND"].ToString() != string.Empty)
                    {
                        StrActEnnd = dt_Orcl_Conarc.Rows[i]["ACTENND"].ToString();
                    }
                    else
                    { StrActEnnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["GRADE"].ToString() != null && dt_Orcl_Conarc.Rows[i]["GRADE"].ToString() != string.Empty)
                    {
                        StrGrade = dt_Orcl_Conarc.Rows[i]["GRADE"].ToString();
                    }
                    else
                    { StrGrade = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["HEAT_IN_SEQ"].ToString() != null && dt_Orcl_Conarc.Rows[i]["HEAT_IN_SEQ"].ToString() != string.Empty)
                    {
                        StrSeqNo = dt_Orcl_Conarc.Rows[i]["HEAT_IN_SEQ"].ToString();
                    }
                    else
                    { StrSeqNo = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["AVG_CAST_SPEED"].ToString() != null && dt_Orcl_Conarc.Rows[i]["AVG_CAST_SPEED"].ToString() != string.Empty)
                    {
                        StrAvgCastSpeed = dt_Orcl_Conarc.Rows[i]["AVG_CAST_SPEED"].ToString();
                    }
                    else
                    { StrAvgCastSpeed = "null"; }
                    if (dt_Orcl_Conarc.Rows[i]["CASTED_LENGTH"].ToString() != null && dt_Orcl_Conarc.Rows[i]["CASTED_LENGTH"].ToString() != string.Empty)
                    {
                        StrCastLength = dt_Orcl_Conarc.Rows[i]["CASTED_LENGTH"].ToString();
                    }
                    else
                    { StrCastLength = "null"; }

                    //++++++++++++++++Tundish Temp Calc+++++++++++++++++
                    DataTable tundish = new DataTable();
                    string selectTemp = "SELECT pdc_tundish_temp.meas_temp as Temp FROM pdc_heat,pdc_tundish_temp WHERE pdc_heat.steel_id = pdc_tundish_temp.steel_id  And Pdc_Tundish_Temp.Meas_No = '1' AND Pdc_Heat.Heatid     = '" + StrHEATID + "'";
                    tundish = clsObj.DBSelectQueryCASTERIII_Table(selectTemp);
                    foreach (DataRow dr in tundish.Rows)
                    {

                        TundishTemp = Convert.ToInt32(dr["Temp"].ToString());
                    }
                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select heatNo from smsHeatTracker where heatNo='" + StrHEATID + "' and plantNo='" + StrMIS_PLANTNO + "'";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update smsHeatTracker set LadleNo='" + StrLadleNo + "',actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',actEnds='" + StrActEnd + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "',Grade='" + StrGrade + "',AvgCastSpeed=" + StrAvgCastSpeed + ",SeqNo='" + StrSeqNo + "',CastLength=" + StrCastLength + ",LastTemp='" + TundishTemp + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "'  and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120)";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,DateStamp,LadleNo,actEnds,actStrt,actEnnd,Grade,AvgCastSpeed,SeqNo,CastLength,LastTemp)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + StrActStart + "','" + StrActEnd + "','','" + StrActStartDte + "','" + StrLadleNo + "','" + StrActEnd + "','" + StrActStrt + "','" + StrActEnnd + "','" + StrGrade + "'," + StrAvgCastSpeed + ",'" + StrSeqNo + "'," + StrCastLength + ",'" + TundishTemp + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void FnSMS_CasterII_HeatTracking(string StrPLANT, string StrPLANTNO, string StrMIS_PLANTNO, string StrDate)
        {

            
            //@old query@01_08_2014 for RevertBack
            //string StrQueryCasterII = "Select Pdc_Heat.Heatid As Heatid_Cust,Pdc_Heat.Ladle_No As Laddle_No,Pdc_Heat.Ladle_Open_Time As Cast_Start_Time,Pdc_Heat.Ladle_Close_Time As Cast_End_Time,Substr(Pdc_Heat.Ladle_Open_Time,1,10) As Act_Start_Dt,replace(Pdc_Heat.Ladle_Open_Time,',','.') As ActStrt,replace(Pdc_Heat.Ladle_Close_Time,',','.') As ActEnnd,Pdc_Heat.Grade_Code As Grade, CONCAT('L+',Pdc_Heat.Heat_Seq_No -1 ) As Heat_In_Seq,Pdc_Strand.Avg_Cast_Speed As Avg_Cast_Speed,Pdc_Strand.Cast_Len_Bgn As Cast_Start_Length,Pdc_Strand.Cast_Len_End As Cast_End_Length,CASE WHEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) > 0 THEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) ELSE NULL END as Casted_Length From Pdc_Heat,Pdc_Strand,pdc_heat_strand WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID And Pdc_Heat_Strand.Strand_Steel_Id=Pdc_Strand.Steel_Id AND SUBSTR(ladle_open_time,1,10)='" + StrDate + "' ORDER BY Ladle_Open_Time DESC";
            string StrQueryCasterII = "Select * from (Select Pdc_Heat.Heatid As Heatid_Cust,Pdc_Heat.Ladle_No As Laddle_No,Pdc_Heat.Ladle_Open_Time As Cast_Start_Time,Pdc_Heat.Ladle_Close_Time As Cast_End_Time,Substr(Pdc_Heat.Ladle_Open_Time,1,10) As Act_Start_Dt,replace(Pdc_Heat.Ladle_Open_Time,',','.') As ActStrt,replace(Pdc_Heat.Ladle_Close_Time,',','.') As ActEnnd,Pdc_Heat.Grade_Code As Grade, CONCAT('L+',Pdc_Heat.Heat_Seq_No -1 ) As Heat_In_Seq,Pdc_Strand.Avg_Cast_Speed As Avg_Cast_Speed,Pdc_Strand.Cast_Len_Bgn As Cast_Start_Length,Pdc_Strand.Cast_Len_End As Cast_End_Length,CASE WHEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) > 0 THEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) ELSE NULL END as Casted_Length From Pdc_Heat,Pdc_Strand,pdc_heat_strand WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID And Pdc_Heat_Strand.Strand_Steel_Id=Pdc_Strand.Steel_Id AND SUBSTR(ladle_open_time,1,10)<='" + StrDate + "' ORDER BY Ladle_Open_Time DESC) where rownum <=5 ";
          
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryCASTERII_Table(StrQueryCasterII);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    string StrHEATID = "";
                    string StrActStart = null;
                    string StrActEnd = null;
                    object StrLadleNo = DBNull.Value;
                    object StrActStartDte = DBNull.Value;
                    string StrActStrt = null;
                    string StrActEnnd = null;
                    object StrGrade = DBNull.Value;
                    string StrAvgCastSpeed = null;
                    object StrSeqNo = DBNull.Value;
                    string StrCastLength = null;
                    object TundishTemp = DBNull.Value;
                    StrHEATID = dt_Orcl_Conarc.Rows[i][0].ToString();
                    if (dt_Orcl_Conarc.Rows[i][1].ToString() != null && dt_Orcl_Conarc.Rows[i][1].ToString() != string.Empty)
                    {
                        StrLadleNo = dt_Orcl_Conarc.Rows[i][1].ToString();
                    }
                    if (dt_Orcl_Conarc.Rows[i][2].ToString() != null && dt_Orcl_Conarc.Rows[i][2].ToString() != string.Empty)
                    {
                        StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][2].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                    }
                    if (dt_Orcl_Conarc.Rows[i][3].ToString() != null && dt_Orcl_Conarc.Rows[i][3].ToString() != string.Empty)
                    {
                        StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][3].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                    }
                    if (dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != string.Empty)
                    {
                        StrActStartDte = dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString();
                    }
                    else
                    { StrActStartDte = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["ACTSTRT"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ACTSTRT"].ToString() != string.Empty)
                    {
                        StrActStrt = dt_Orcl_Conarc.Rows[i]["ACTSTRT"].ToString();
                    }
                    else
                    { StrActStrt = null; }
                    if (dt_Orcl_Conarc.Rows[i]["ACTENND"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ACTENND"].ToString() != string.Empty)
                    {
                        StrActEnnd = dt_Orcl_Conarc.Rows[i]["ACTENND"].ToString();
                        // +Manish check for ancient laddle close time 1900
                        if (StrActEnnd.Substring(0, 4).Equals("1900"))
                        {
                            StrActEnnd = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }
                    else
                    { StrActEnnd = null; }
                    if (dt_Orcl_Conarc.Rows[i]["GRADE"].ToString() != null && dt_Orcl_Conarc.Rows[i]["GRADE"].ToString() != string.Empty)
                    {
                        StrGrade = dt_Orcl_Conarc.Rows[i]["GRADE"].ToString();
                    }
                    else
                    { StrGrade = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["HEAT_IN_SEQ"].ToString() != null && dt_Orcl_Conarc.Rows[i]["HEAT_IN_SEQ"].ToString() != string.Empty)
                    {
                        StrSeqNo = dt_Orcl_Conarc.Rows[i]["HEAT_IN_SEQ"].ToString();
                    }
                    else
                    { StrSeqNo = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["AVG_CAST_SPEED"].ToString() != null && dt_Orcl_Conarc.Rows[i]["AVG_CAST_SPEED"].ToString() != string.Empty)
                    {
                        StrAvgCastSpeed = dt_Orcl_Conarc.Rows[i]["AVG_CAST_SPEED"].ToString();
                    }
                    else
                    { StrAvgCastSpeed = "null"; }
                    if (dt_Orcl_Conarc.Rows[i]["CASTED_LENGTH"].ToString() != null && dt_Orcl_Conarc.Rows[i]["CASTED_LENGTH"].ToString() != string.Empty)
                    {
                        StrCastLength = dt_Orcl_Conarc.Rows[i]["CASTED_LENGTH"].ToString();
                    }
                    else
                    { StrCastLength = "null"; }
                    //++++++++++++++++Tundish Temp Calc+++++++++++++++++
                    DataTable tundish = new DataTable();
                    string selectTemp = "SELECT pdc_tundish_temp.meas_temp as Temp FROM pdc_heat,pdc_tundish_temp WHERE pdc_heat.steel_id = pdc_tundish_temp.steel_id  And Pdc_Tundish_Temp.Meas_No = '1' AND Pdc_Heat.Heatid     = '" + StrHEATID + "'";
                    tundish = clsObj.DBSelectQueryCASTERII_Table(selectTemp);
                    foreach (DataRow dr in tundish.Rows)
                    {
                        TundishTemp = Convert.ToInt32(dr["Temp"].ToString());
                    }
                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select heatNo from smsHeatTracker where heatNo='" + StrHEATID + "' and plantNo='" + StrMIS_PLANTNO + "'";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update smsHeatTracker set LadleNo='" + StrLadleNo + "',actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',actEnds='" + StrActEnd + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "',Grade='" + StrGrade + "',AvgCastSpeed=" + StrAvgCastSpeed + ",SeqNo='" + StrSeqNo + "',CastLength=" + StrCastLength + ",LastTemp='" + TundishTemp + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "'  and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120)";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,DateStamp,LadleNo,actEnds,actStrt,actEnnd,Grade,AvgCastSpeed,SeqNo,CastLength,LastTemp)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + StrActStart + "','" + StrActEnd + "','','" + StrActStartDte + "','" + StrLadleNo + "','" + StrActEnd + "','" + StrActStrt + "','" + StrActEnnd + "','" + StrGrade + "'," + StrAvgCastSpeed + ",'" + StrSeqNo + "'," + StrCastLength + ",'" + TundishTemp + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }

        private void FnSMS_CasterI_HeatTracking(string StrPLANT, string StrPLANTNO, string StrMIS_PLANTNO, string StrDate)
        {

              //Original for RevertBack @01_08_2014
            //string StrQueryCasterI = "Select Pdc_Heat.Heatid As Heatid_Cust,Pdc_Heat.Ladle_No As Laddle_No,Pdc_Heat.Ladle_Open_Time As Cast_Start_Time,Pdc_Heat.Ladle_Close_Time As Cast_End_Time,Substr(Pdc_Heat.Ladle_Open_Time,1,10) As Act_Start_Dt,replace(Pdc_Heat.Ladle_Open_Time,',','.') As ActStrt,replace(Pdc_Heat.Ladle_Close_Time,',','.') As ActEnnd,Pdc_Heat.Grade_Code As Grade, CONCAT('L+',Pdc_Heat.Heat_Seq_No -1 ) As Heat_In_Seq,Pdc_Strand.Avg_Cast_Speed As Avg_Cast_Speed,Pdc_Strand.Cast_Len_Bgn As Cast_Start_Length,Pdc_Strand.Cast_Len_End As Cast_End_Length,CASE WHEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) > 0 THEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) ELSE NULL END as Casted_Length From Pdc_Heat,Pdc_Strand,pdc_heat_strand WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID And Pdc_Heat_Strand.Strand_Steel_Id=Pdc_Strand.Steel_Id AND SUBSTR(ladle_open_time,1,10)='" + StrDate + "' ORDER BY Ladle_Open_Time DESC";
            string StrQueryCasterI = "Select * from (Select Pdc_Heat.Heatid As Heatid_Cust,Pdc_Heat.Ladle_No As Laddle_No,Pdc_Heat.Ladle_Open_Time As Cast_Start_Time,Pdc_Heat.Ladle_Close_Time As Cast_End_Time,Substr(Pdc_Heat.Ladle_Open_Time,1,10) As Act_Start_Dt,replace(Pdc_Heat.Ladle_Open_Time,',','.') As ActStrt,replace(Pdc_Heat.Ladle_Close_Time,',','.') As ActEnnd,Pdc_Heat.Grade_Code As Grade, CONCAT('L+',Pdc_Heat.Heat_Seq_No -1 ) As Heat_In_Seq,Pdc_Strand.Avg_Cast_Speed As Avg_Cast_Speed,Pdc_Strand.Cast_Len_Bgn As Cast_Start_Length,Pdc_Strand.Cast_Len_End As Cast_End_Length,CASE WHEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) > 0 THEN NVL(Pdc_Strand.Cast_Len_End,0) - NVL(Pdc_Strand.Cast_Len_Bgn,0) ELSE NULL END as Casted_Length From Pdc_Heat,Pdc_Strand,pdc_heat_strand WHERE pdc_heat.STEEL_ID=pdc_heat_strand.HEAT_STEEL_ID And Pdc_Heat_Strand.Strand_Steel_Id=Pdc_Strand.Steel_Id AND SUBSTR(ladle_open_time,1,10)<='" + StrDate + "' ORDER BY Ladle_Open_Time DESC) where rownum <=5 ";
                        
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryCASTERI_Table(StrQueryCasterI);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    string StrHEATID = "";
                    string StrActStart = null;
                    string StrActEnd = null;
                    object StrLadleNo = DBNull.Value;
                    object StrActStartDte = DBNull.Value;
                    string StrActStrt = null;
                    string StrActEnnd = null;
                    object StrGrade = DBNull.Value;
                    string StrAvgCastSpeed = null;
                    object StrSeqNo = DBNull.Value;
                    string StrCastLength = null;
                    object TundishTemp = DBNull.Value;
                    StrHEATID = dt_Orcl_Conarc.Rows[i][0].ToString();
                    if (dt_Orcl_Conarc.Rows[i][1].ToString() != null && dt_Orcl_Conarc.Rows[i][1].ToString() != string.Empty)
                    {
                        StrLadleNo = dt_Orcl_Conarc.Rows[i][1].ToString();
                    }
                    if (dt_Orcl_Conarc.Rows[i][2].ToString() != null && dt_Orcl_Conarc.Rows[i][2].ToString() != string.Empty)
                    {
                        StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][2].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                    }
                    if (dt_Orcl_Conarc.Rows[i][3].ToString() != null && dt_Orcl_Conarc.Rows[i][3].ToString() != string.Empty)
                    {
                        StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][3].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                    }
                    if (dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString() != string.Empty)
                    {
                        StrActStartDte = dt_Orcl_Conarc.Rows[i]["Act_Start_Dt"].ToString();
                    }
                    else
                    { StrActStartDte = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["ACTSTRT"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ACTSTRT"].ToString() != string.Empty)
                    {
                        StrActStrt = dt_Orcl_Conarc.Rows[i]["ACTSTRT"].ToString();
                    }
                    else
                    { StrActStrt = null; }
                    if (dt_Orcl_Conarc.Rows[i]["ACTENND"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ACTENND"].ToString() != string.Empty)
                    {
                        StrActEnnd = dt_Orcl_Conarc.Rows[i]["ACTENND"].ToString();
                        // +Manish check for ancient laddle close time 1900
                        if (StrActEnnd.Substring(0, 4).Equals("1900"))
                        {
                            StrActEnnd = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }
                    else
                    { StrActEnnd = null; }
                    if (dt_Orcl_Conarc.Rows[i]["GRADE"].ToString() != null && dt_Orcl_Conarc.Rows[i]["GRADE"].ToString() != string.Empty)
                    {
                        StrGrade = dt_Orcl_Conarc.Rows[i]["GRADE"].ToString();
                    }
                    else
                    { StrGrade = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["HEAT_IN_SEQ"].ToString() != null && dt_Orcl_Conarc.Rows[i]["HEAT_IN_SEQ"].ToString() != string.Empty)
                    {
                        StrSeqNo = dt_Orcl_Conarc.Rows[i]["HEAT_IN_SEQ"].ToString();
                    }
                    else
                    { StrSeqNo = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["AVG_CAST_SPEED"].ToString() != null && dt_Orcl_Conarc.Rows[i]["AVG_CAST_SPEED"].ToString() != string.Empty)
                    {
                        StrAvgCastSpeed = dt_Orcl_Conarc.Rows[i]["AVG_CAST_SPEED"].ToString();
                    }
                    else
                    { StrAvgCastSpeed = "null"; }
                    if (dt_Orcl_Conarc.Rows[i]["CASTED_LENGTH"].ToString() != null && dt_Orcl_Conarc.Rows[i]["CASTED_LENGTH"].ToString() != string.Empty)
                    {
                        StrCastLength = dt_Orcl_Conarc.Rows[i]["CASTED_LENGTH"].ToString();
                    }
                    else
                    { StrCastLength = "null"; }
                    //++++++++++++++++Tundish Temp Calc+++++++++++++++++
                    DataTable tundish = new DataTable();
                    string selectTemp = "SELECT pdc_tundish_temp.meas_temp as Temp FROM pdc_heat,pdc_tundish_temp WHERE pdc_heat.steel_id = pdc_tundish_temp.steel_id  And Pdc_Tundish_Temp.Meas_No = '1' AND Pdc_Heat.Heatid     = '" + StrHEATID + "'";
                    tundish = clsObj.DBSelectQueryCASTERI_Table(selectTemp);
                    foreach (DataRow dr in tundish.Rows)
                    {
                        TundishTemp = Convert.ToInt32(dr["Temp"].ToString());
                    }
                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select heatNo from smsHeatTracker where heatNo='" + StrHEATID + "' and plantNo='" + StrMIS_PLANTNO + "'";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update smsHeatTracker set LadleNo='" + StrLadleNo + "',actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',actEnds='" + StrActEnd + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "',Grade='" + StrGrade + "',AvgCastSpeed=" + StrAvgCastSpeed + ",SeqNo='" + StrSeqNo + "',CastLength=" + StrCastLength + ",LastTemp='" + TundishTemp + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "'  and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120)";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,DateStamp,LadleNo,actEnds,actStrt,actEnnd,Grade,AvgCastSpeed,SeqNo,CastLength,LastTemp)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + StrActStart + "','" + StrActEnd + "','','" + StrActStartDte + "','" + StrLadleNo + "','" + StrActEnd + "','" + StrActStrt + "','" + StrActEnnd + "','" + StrGrade + "'," + StrAvgCastSpeed + ",'" + StrSeqNo + "'," + StrCastLength + ",'" + TundishTemp + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }

        private void FnSMS_SMCDB_HeatTracking(string StrPLANT, string StrPLANTNO, string StrMIS_PLANTNO, string StrDate)
        {
            //System.DateTime.Now.ToString("yyyy-MM-dd")
            //Old Query To Revert Back @01_08_2014
             //string StrQueryOrclCONARC = "SELECT  hp.HEATID_CUST AS HEAT_ID,hp.PLANT ||hp.PLANTNO    AS PLANT,hp.TREATID_CUST AS TREAT_ID,hp.PLANLADLETYPE ||hp.PLANLADLENO AS LADLE_ID, decode(replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.')) AS PLAN_START_TIME,decode(replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'))   AS PLAN_END_TIME, replace(SUBSTR(ACTTREATSTART,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hp.ACTTREATSTART,1,10) as Act_Start_Dt,replace(hp.ACTTREATSTART,',','.') AS ACTSTRT,replace(hp.ACTTREATEND,',','.') AS ACTENND,Hd.Steelgradecode_Act as Grade,Hp.Heatid as int_heatid FROM PP_HEAT_PLANT hp left outer join PD_HEATDATA hd on hp.HEATID=hd.HEATID   where hd.PLANT='" + StrPLANT + "' and  hp.EXPIRATIONDATE='VALID' and SUBSTR(hp.ACTTREATSTART,1,10)='" + StrDate + "' and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrPLANTNO + "' ORDER BY hp.HEATID DESC ";
             //string StrQueryOrclCONARC = "SELECT * FROM (SELECT  hp.HEATID_CUST AS HEAT_ID,hp.PLANT ||hp.PLANTNO    AS PLANT,hp.TREATID_CUST AS TREAT_ID,hp.PLANLADLETYPE ||hp.PLANLADLENO AS LADLE_ID, decode(replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.')) AS PLAN_START_TIME,decode(replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'))   AS PLAN_END_TIME, replace(SUBSTR(ACTTREATSTART,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hp.ACTTREATSTART,1,10) as Act_Start_Dt,replace(hp.ACTTREATSTART,',','.') AS ACTSTRT,replace(hp.ACTTREATEND,',','.') AS ACTENND,Hd.Steelgradecode_Act as Grade,Hp.Heatid as int_heatid FROM PP_HEAT_PLANT hp left outer join PD_HEATDATA hd on hp.HEATID=hd.HEATID   where hd.PLANT='" + StrPLANT + "' and  hp.EXPIRATIONDATE='VALID' and SUBSTR(hp.ACTTREATSTART,1,10)<='" + StrDate + "' and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrPLANTNO + "' ORDER BY hp.HEATID DESC ) where rownum <= 5";
            string StrQueryOrclCONARC = "SELECT * from (SELECT PP.HEATID_CUST AS HEAT_ID,PP.PLANT||PP.PLANTNO    AS PLANT,PP.TREATID_CUST AS TREAT_ID,PP.PLANLADLETYPE||PP.PLANLADLENO AS LADLE_ID, DECODE(REPLACE(SUBSTR(PP.PLANTREATSTART,11,6) ,':','.'),0,NULL,REPLACE(SUBSTR(PP.PLANTREATSTART,11,6) ,':','.')) AS PLAN_START_TIME,DECODE(REPLACE(SUBSTR(PP.PLANTREATEND,11,6) ,':','.'),0,NULL,REPLACE(SUBSTR(PP.PLANTREATEND,11,6) ,':','.')) AS PLAN_END_TIME,REPLACE(SUBSTR(ACTTREATSTART,11,6) ,':','.') AS ACT_START_DATE,Decode(Replace(Substr(Pp.Acttreatend,11,6) ,':','.'),0,Null,Replace(Substr(Pp.Acttreatend,11,6) ,':','.') ) As Act_End_Date ,Pd.Ladleno_In as ladleno,SUBSTR(pp.ACTTREATSTART,1,10) as Act_Start_Dt,replace(pp.ACTTREATSTART,',','.') AS ACTSTRT,replace(pp.ACTTREATEND,',','.') AS ACTENND,Pd.Steelgradecode as Grade,Pd.Heatid as int_heatid,Pd.Lasttemp As Lasttemp,Pd.Total_Elec_Egy As Total_Energy FROM PP_HEAT_PLANT pp,Pd_Report Pd Where Pd.Heatid = Pp.Heatid And Pd.Treatid = Pp.Treatid And Pd.Plant = Pp.Plant And Pd.Plant = '" + StrPLANT + "' and pd.PLANTNO = '" + StrPLANTNO + "' and pp.EXPIRATIONDATE = 'VALID' And Substr(Pp.Acttreatstart,1,10) <='" + StrDate + "'  order by pp.ACTTREATSTART desc, pd.HEATID desc, pd.TREATID desc) where rownum <= 10";
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclCONARC);
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
                object StrActStrt = DBNull.Value;
                object StrActEnnd = DBNull.Value;
                object StrGrade = DBNull.Value;
                //+++++++++++++++Manish For last temp in conarc and LHF +++++++++++++++++++++
                object intHeatId = DBNull.Value;
                double TotalEnergy = 0;
                double lastTemp = 0;
                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_Conarc.Rows[i][0].ToString();
                    StrTREATID = dt_Orcl_Conarc.Rows[i][2].ToString();

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

                    if (dt_Orcl_Conarc.Rows[i]["ActStrt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ActStrt"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActStrt = dt_Orcl_Conarc.Rows[i]["ActStrt"].ToString();
                    }
                    else
                    { StrActStrt = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["ActEnnd"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ActEnnd"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrActEnnd = dt_Orcl_Conarc.Rows[i]["ActEnnd"].ToString();
                    }
                    else
                    { StrActEnnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["Grade"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Grade"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrGrade = dt_Orcl_Conarc.Rows[i]["Grade"].ToString();
                    }
                    else
                    { StrGrade = DBNull.Value; }
                    //++++++++++++++++++Last Temp Implementation +++++++++++++++
                    //if (dt_Orcl_Conarc.Rows[i]["int_heatid"].ToString() != null && dt_Orcl_Conarc.Rows[i]["int_heatid"].ToString() != string.Empty)
                    //{
                    //    intHeatId = dt_Orcl_Conarc.Rows[i]["int_heatid"].ToString();
                    //    // helping datatable for conatining temp
                    //    DataTable temp = new DataTable();
                    //    string SelectTemp = "select TEMP from Pd_Anl where HEATID = '" + intHeatId + "' and PROBETYPENO = '1' order by Rectime desc";
                    //    temp = clsObj.DBSelectQueryStatusScreen_Table(SelectTemp);
                    //    if (temp.Rows.Count > 0)
                    //    {
                    //        lastTemp = Convert.ToInt32(temp.Rows[0]["TEMP"].ToString());
                    //    }
                    //    //------------------End of Last Temp------------------------------------------- 
                    //}
                    //else
                    //{ intHeatId = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["Lasttemp"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Lasttemp"].ToString() != string.Empty)
                    {
                        lastTemp = Convert.ToDouble(dt_Orcl_Conarc.Rows[i]["Lasttemp"].ToString());
                    }

                    if (dt_Orcl_Conarc.Rows[i]["Total_Energy"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Total_Energy"].ToString() != string.Empty)
                    {
                        TotalEnergy = Convert.ToDouble(dt_Orcl_Conarc.Rows[i]["Total_Energy"].ToString());
                    }

                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select heatNo from smsHeatTracker where heatNo='" + StrHEATID + "' and plantNo='" + StrMIS_PLANTNO + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120) and actStart='" + StrActStart + "'";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count > 0)
                    {
                        //string StrUpdateQuery = "Update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',plndStart='" + StrPlndStart + "',plndEnd='" + StrPlndEnd + "',LadleNo='" + StrLadleNo + "',actEnds='" + StrActEnd + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "',Grade='" + StrGrade + "',LastTemp='" + lastTemp + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120) and actStart='" + StrActStart + "'; update smsHeatTracker set actEnnd=null where actEnnd='1900-01-01 00:00:00.000'";
                        string StrUpdateQuery = "Update smsHeatTracker set actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',plndStart='" + StrPlndStart + "',plndEnd='" + StrPlndEnd + "',LadleNo='" + StrLadleNo + "',actEnds='" + StrActEnd + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "',Grade='" + StrGrade + "',LastTemp='" + lastTemp + "',Tot_Energy='" + TotalEnergy + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120) and actStart='" + StrActStart + "'; update smsHeatTracker set actEnnd=null where actEnnd='1900-01-01 00:00:00.000'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        //string StrInsertQuery = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,plndStart,plndEnd,DateStamp,LadleNo,actEnds,actStrt,actEnnd,Grade,LastTemp)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + StrActStart + "','" + StrActEnd + "','','" + StrPlndStart + "','" + StrPlndEnd + "','" + StrActStartDte + "','" + StrLadleNo + "','" + StrActEnd + "','" + StrActStrt + "','" + StrActEnnd + "','" + StrGrade + "','" + lastTemp + "')";
                        string StrInsertQuery = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,plndStart,plndEnd,DateStamp,LadleNo,actEnds,actStrt,actEnnd,Grade,LastTemp,Tot_Energy)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + StrActStart + "','" + StrActEnd + "','','" + StrPlndStart + "','" + StrPlndEnd + "','" + StrActStartDte + "','" + StrLadleNo + "','" + StrActEnd + "','" + StrActStrt + "','" + StrActEnnd + "','" + StrGrade + "','" + lastTemp + "','" + TotalEnergy + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void FnSMS_SMCDB_HeatBlow_Arc_LF(string StrDate)
        {
            //StrDate = "2014-04-20";
            string StrMIS_PLANTNO = "";
            //string StrQueryOrclCONARC = "select PDR.PLANT || PDR.PLANTNO as Plant,HT.HEATID_CUST as HEATID,decode(replace(SUBSTR(to_char(VRLF.POWERONTIME, 'yyyy-mm-dd hh24:mi:ss'),11,6) ,':','.'),0,null,replace(SUBSTR(to_char(VRLF.POWERONTIME, 'yyyy-mm-dd hh24:mi:ss'),11,6) ,':','.')) AS baStart,decode(replace(SUBSTR(to_char(VRLF.POWEROFFTIME, 'yyyy-mm-dd hh24:mi:ss'),11,6) ,':','.'),0,null,replace(SUBSTR(to_char(VRLF.POWEROFFTIME, 'yyyy-mm-dd hh24:mi:ss'),11,6) ,':','.')) AS baEnd,'Arcing Time' as blow_arc,to_char(VRLF.POWERONTIME,'YYYY-MM-DD') as DteStamp,'2' as baStat,VRLF.POWERONTIME as baStrt,VRLF.POWEROFFTIME as baEnnd from VR_LF_HEAT_REPORT_POWER VRLF,PD_REPORT PDR,PP_HEAT_PLANT HT where  PDR.HEATID=VRLF.HEATID  AND HT.HEATID=PDR.HEATID and to_char(VRLF.POWERONTIME,'YYYY-MM-DD') = '" + StrDate + "'";
            string StrQueryOrclCONARC = "SELECT PDR.PLANT || PDR.PLANTNO AS Plant,HT.HEATID_CUST AS HEATID,DECODE(REPLACE(SUBSTR(TO_CHAR(VRLF.POWERONTIME, 'yyyy-mm-dd hh24:mi:ss'),11,6) ,':','.'),0,NULL,REPLACE(SUBSTR(TO_CHAR(VRLF.POWERONTIME, 'yyyy-mm-dd hh24:mi:ss'),11,6) ,':','.'))   AS baStart,DECODE(REPLACE(SUBSTR(TO_CHAR(VRLF.POWEROFFTIME, 'yyyy-mm-dd hh24:mi:ss'),11,6) ,':','.'),0,NULL,REPLACE(SUBSTR(TO_CHAR(VRLF.POWEROFFTIME, 'yyyy-mm-dd hh24:mi:ss'),11,6) ,':','.')) AS baEnd,'Arcing Time'  AS blow_arc,  TO_CHAR(VRLF.POWERONTIME,'YYYY-MM-DD') AS DteStamp,  '2' AS baStat,  VRLF.POWERONTIME  AS baStrt,  VRLF.POWEROFFTIME AS baEnnd FROM VR_LF_HEAT_REPORT_POWER VRLF,   PD_REPORT PDR,   Pp_Heat_Plant Ht WHERE PDR.HEATID =VRLF.HEATID   And Ht.Heatid=Pdr.Heatid ANd Ht.Plant =Pdr.Plant And Ht.Plant ='LF' And Ht.Treatid =Pdr.Treatid AND TO_CHAR(VRLF.POWERONTIME,'YYYY-MM-DD') ='" + StrDate + "'";
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclCONARC);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                string StrHEATID = "";
                object StrBADesc = "";
                object StrPlant = "";
                object baStart = DBNull.Value;
                object baEnd = DBNull.Value;
                object DteStamp = DBNull.Value;
                object baStat = DBNull.Value;
                object baStrt = DBNull.Value;
                object baEnnd = DBNull.Value;

                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_Conarc.Rows[i]["HEATID"].ToString();

                    if (dt_Orcl_Conarc.Rows[i]["baStart"].ToString() != null && dt_Orcl_Conarc.Rows[i]["baStart"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baStart = dt_Orcl_Conarc.Rows[i]["baStart"].ToString();
                    }
                    else
                    { baStart = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["baEnd"].ToString() != null && dt_Orcl_Conarc.Rows[i]["baEnd"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baEnd = dt_Orcl_Conarc.Rows[i]["baEnd"].ToString();
                    }
                    else
                    { baEnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["blow_arc"].ToString() != null && dt_Orcl_Conarc.Rows[i]["blow_arc"].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrBADesc = dt_Orcl_Conarc.Rows[i]["blow_arc"].ToString();
                    }
                    else
                    { StrBADesc = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["DTESTAMP"].ToString() != null && dt_Orcl_Conarc.Rows[i]["DTESTAMP"].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        DteStamp = dt_Orcl_Conarc.Rows[i]["DTESTAMP"].ToString();
                    }
                    else
                    { DteStamp = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["baStrt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["baStrt"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baStrt = dt_Orcl_Conarc.Rows[i]["baStrt"].ToString();
                    }
                    else
                    { baStrt = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["baEnnd"].ToString() != null && dt_Orcl_Conarc.Rows[i]["baEnnd"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baEnnd = dt_Orcl_Conarc.Rows[i]["baEnnd"].ToString();
                    }
                    else
                    { baEnnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != null && dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrPlant = dt_Orcl_Conarc.Rows[i]["PLANT"].ToString();
                    }
                    else
                    { StrPlant = DBNull.Value; }
                    if (StrPlant.ToString() == "LF1")
                    {
                        StrMIS_PLANTNO = "9";
                    }
                    else
                    {
                        StrMIS_PLANTNO = "10";
                    }
                    if (StrBADesc.ToString() == "Blowing Time")
                    {
                        baStat = "1";
                    }
                    else
                    {
                        baStat = "2";
                    }
                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select HeatNo from smsHeatBlowArc where Plant='" + StrPlant + "' and blow_arc='" + StrBADesc + "' and baStart='" + baStart + "' and baEnd='" + baEnd + "' and plantNo='" + StrMIS_PLANTNO + "' and convert(varchar(10),DteStamp,120)=convert(varchar(10),'" + DteStamp + "',120)";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count <= 0)
                    {
                        if (baStart != DBNull.Value && baEnd != DBNull.Value)
                        {
                            string StrInsertQuery = "insert smsHeatBlowArc(plantNo,heatNo,baStart,baEnd,blow_arc,DteStamp,Plant,baStat,batStrt,batEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + baStart + "','" + baEnd + "','" + StrBADesc + "','" + DteStamp + "','" + StrPlant + "','" + baStat + "','" + baStrt + "','" + baEnnd + "')";
                            bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                        }

                    }
                }
            }
        }
        private void FnSMS_SMCDB_HeatBlow_Arc(string StrDate)
        {
            //StrDate = "2014-04-20";
            string StrMIS_PLANTNO = "";
            //string StrQueryOrclCONARC = "SELECT DISTINCT pdd.CODE,gdf.FUNCTIONDESC,SUBSTR(pdd.DELAYSTART,1,10) as DteStamp, decode(replace(SUBSTR(pdd.DELAYSTART,11,6) ,':','.'),0,null,replace(SUBSTR(pdd.DELAYSTART,11,6) ,':','.')) AS Start_Time, decode(replace(SUBSTR(pdd.DELAYEND,11,6) ,':','.'),0,null,replace(SUBSTR(pdd.DELAYEND,11,6) ,':','.')) AS End_Time,C.DELAYDESC AS BA_DESC, PP_HEAT_PLANT.HEATID_CUST,pdd.PLANT || pdd.PLANTNO as Plant,pdd.PLANTNO,replace(pdd.DELAYSTART,',','.') as bastrt,replace(pdd.DELAYEND,',','.') as baEnnd FROM PD_DELAYS pdd,PD_REPORT pdr_delay,GT_DELAY_CODE C,GC_DELAY_FUNC gdf,PP_HEAT_PLANT WHERE pdd.DELAYSTART >= pdr_delay.ANNOUNCETIME AND pdd.DELAYSTART <= pdr_delay.TAPPINGENDTIME AND pdr_delay.PLANTNO = pdd.PLANTNO AND C.CODE = pdd.CODE AND C.FUNC = gdf.FUNC AND pdr_delay.HEATID = PP_HEAT_PLANT.HEATID AND PP_HEAT_PLANT.TREATID = pdr_delay.TREATID AND PP_HEAT_PLANT.PLANT = pdr_delay.PLANT AND pdd.CODE = 'P002' AND (pdr_delay.PLANT = 'CON' AND pdd.EXPIRATIONDATE = 'VALID') AND SUBSTR(pdr_delay.treatstart_act,1,10)='" + StrDate + "' union all SELECT DISTINCT pdd.CODE,  gdf.FUNCTIONDESC,SUBSTR(pdd.DELAYSTART,1,10) as DteStamp, decode(replace(SUBSTR(pdd.DELAYSTART,11,6) ,':','.'),0,null,replace(SUBSTR(pdd.DELAYSTART,11,6) ,':','.')) AS Start_Time , decode(replace(SUBSTR(pdd.DELAYEND,11,6) ,':','.'),0,null,replace(SUBSTR(pdd.DELAYEND,11,6) ,':','.')) AS End_Time  ,  C.DELAYDESC  AS BA_DESC,  PP_HEAT_PLANT.HEATID_CUST,  pdd.PLANT || pdd.PLANTNO as Plant,pdd.PLANTNO,replace(pdd.DELAYSTART,',','.') as bastrt,replace(pdd.DELAYEND,',','.') as baEnnd FROM PD_DELAYS pdd,  PD_REPORT pdr_delay,  GT_DELAY_CODE C,  GC_DELAY_FUNC gdf,  PP_HEAT_PLANT WHERE pdd.DELAYSTART     >= pdr_delay.ANNOUNCETIME AND pdd.DELAYSTART       <= pdr_delay.TAPPINGENDTIME AND pdr_delay.PLANTNO     = pdd.PLANTNO AND C.CODE                = pdd.CODE AND C.FUNC                = gdf.FUNC AND pdr_delay.HEATID      = PP_HEAT_PLANT.HEATID AND PP_HEAT_PLANT.TREATID = pdr_delay.TREATID AND PP_HEAT_PLANT.PLANT   = pdr_delay.PLANT AND pdd.CODE             = 'P001' AND (pdr_delay.PLANT       = 'CON' AND pdd.EXPIRATIONDATE    = 'VALID') AND SUBSTR(pdr_delay.treatstart_act,1,10)='" + StrDate + "'";
            string StrQueryOrclCONARC = "select * from (SELECT DISTINCT pdd.CODE,gdf.FUNCTIONDESC,SUBSTR(pdd.DELAYSTART,1,10) as DteStamp, decode(replace(SUBSTR(pdd.DELAYSTART,11,6) ,':','.'),0,null,replace(SUBSTR(pdd.DELAYSTART,11,6) ,':','.')) AS Start_Time, decode(replace(SUBSTR(pdd.DELAYEND,11,6) ,':','.'),0,null,replace(SUBSTR(pdd.DELAYEND,11,6) ,':','.')) AS End_Time,C.DELAYDESC AS BA_DESC, PP_HEAT_PLANT.HEATID_CUST,pdd.PLANT || pdd.PLANTNO as Plant,pdd.PLANTNO,replace(pdd.DELAYSTART,',','.') as bastrt,replace(pdd.DELAYEND,',','.') as baEnnd FROM PD_DELAYS pdd,PD_REPORT pdr_delay,GT_DELAY_CODE C,GC_DELAY_FUNC gdf,PP_HEAT_PLANT WHERE pdd.DELAYSTART >= pdr_delay.ANNOUNCETIME AND pdd.DELAYSTART <= pdr_delay.TAPPINGENDTIME AND pdr_delay.PLANTNO = pdd.PLANTNO AND C.CODE = pdd.CODE AND C.FUNC = gdf.FUNC AND pdr_delay.HEATID = PP_HEAT_PLANT.HEATID AND PP_HEAT_PLANT.TREATID = pdr_delay.TREATID AND PP_HEAT_PLANT.PLANT = pdr_delay.PLANT AND pdd.CODE in ('P001','P002','P003') AND (pdr_delay.PLANT = 'CON' AND pdd.EXPIRATIONDATE = 'VALID') AND SUBSTR(pdr_delay.treatstart_act,1,10)='" + StrDate + "') where rownum <= 7"; 
            DataTable dt_Orcl_Conarc = new DataTable();
            dt_Orcl_Conarc = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclCONARC);
            if (dt_Orcl_Conarc.Rows.Count > 0)
            {
                string StrHEATID = "";
                object StrBADesc = "";
                object StrPlant = "";
                object Start_Time = DBNull.Value;
                object End_Time = DBNull.Value;
                object DteStamp = DBNull.Value;
                object baStat = DBNull.Value;
                object baStrt = DBNull.Value;
                object baEnnd = DBNull.Value;
                int IdCode = -1;

                for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                {
                    StrHEATID = dt_Orcl_Conarc.Rows[i]["HEATID_CUST"].ToString();
                    string Pcode = dt_Orcl_Conarc.Rows[i]["CODE"].ToString();
                    IdCode = Convert.ToInt32(Pcode.Substring(Pcode.Length - 1));

                    if (dt_Orcl_Conarc.Rows[i]["START_TIME"].ToString() != null && dt_Orcl_Conarc.Rows[i]["START_TIME"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        Start_Time = dt_Orcl_Conarc.Rows[i]["START_TIME"].ToString();
                    }
                    else
                    { Start_Time = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["END_TIME"].ToString() != null && dt_Orcl_Conarc.Rows[i]["END_TIME"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        End_Time = dt_Orcl_Conarc.Rows[i]["END_TIME"].ToString();
                    }
                    else
                    { End_Time = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["BA_DESC"].ToString() != null && dt_Orcl_Conarc.Rows[i]["BA_DESC"].ToString() != string.Empty)
                    {
                        //StrPlndStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][4].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrBADesc = dt_Orcl_Conarc.Rows[i]["BA_DESC"].ToString();
                    }
                    else
                    { StrBADesc = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["DTESTAMP"].ToString() != null && dt_Orcl_Conarc.Rows[i]["DTESTAMP"].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        DteStamp = dt_Orcl_Conarc.Rows[i]["DTESTAMP"].ToString();
                    }
                    else
                    { DteStamp = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["baStrt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["baStrt"].ToString() != string.Empty)
                    {
                        //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baStrt = dt_Orcl_Conarc.Rows[i]["baStrt"].ToString();
                    }
                    else
                    { baStrt = DBNull.Value; }

                    if (dt_Orcl_Conarc.Rows[i]["baEnnd"].ToString() != null && dt_Orcl_Conarc.Rows[i]["baEnnd"].ToString() != string.Empty)
                    {
                        //StrActEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][7].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        baEnnd = dt_Orcl_Conarc.Rows[i]["baEnnd"].ToString();
                    }
                    else
                    { baEnnd = DBNull.Value; }
                    if (dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != null && dt_Orcl_Conarc.Rows[i]["PLANT"].ToString() != string.Empty)
                    {
                        //StrPlndEnd = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][5].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrPlant = dt_Orcl_Conarc.Rows[i]["PLANT"].ToString();
                    }
                    else
                    { StrPlant = DBNull.Value; }
                    if (StrPlant.ToString() == "CON1")
                    {
                        StrMIS_PLANTNO = "3";
                    }
                    else
                    {
                        StrMIS_PLANTNO = "4";
                    }
                    //if (StrBADesc.ToString() == "Blowing Time")
                    //{
                    //    baStat = "1";
                    //}
                    //else
                    //{
                    //    baStat = "2";
                    //}
                    switch (IdCode)
                    {
                        case 1: baStat = "1"; break;
                        case 2: baStat = "2"; break;
                        case 3: baStat = "3"; break;
                    }
                    DataTable dt_MIS_CONARC = new DataTable();
                    string StrQueryCONARC = "select HeatNo from smsHeatBlowArc where Plant='" + StrPlant + "' and blow_arc='" + StrBADesc + "' and baStart='" + Start_Time + "' and baEnd='" + End_Time + "' and plantNo='" + StrMIS_PLANTNO + "' and convert(varchar(10),DteStamp,120)=convert(varchar(10),'" + DteStamp + "',120)";
                    dt_MIS_CONARC = clsObj.DBSelectQueryMIS_Table(StrQueryCONARC);
                    if (dt_MIS_CONARC.Rows.Count <= 0)
                    {
                        if (Start_Time != DBNull.Value && End_Time != DBNull.Value)
                        {
                            string StrInsertQuery = "insert smsHeatBlowArc(plantNo,heatNo,baStart,baEnd,blow_arc,DteStamp,Plant,baStat,batStrt,batEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','" + Start_Time + "','" + End_Time + "','" + StrBADesc + "','" + DteStamp + "','" + StrPlant + "','" + baStat + "','" + baStrt + "','" + baEnnd + "')";
                            bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                        }

                    }
                }
            }
        }
        private void FnSMS_SMCDB_UpdateHeatEndTime(string StrPLANT, string StrPLANTNO, string StrMIS_PLANTNO, string StrDate)
        {
            string StrUptQuery = "UPDATE t   SET t.actEnnd = c.actStrrt,t.actEnd=c.actEnd,t.actEnds=c.actEnd FROM smsHeatTracker t INNER JOIN (SELECT a.plantNo,a.heatNo,a.DateStamp,a.actStrt as actStart,b.actStrt,dateadd(minute,-1,b.actStrt) as actStrrt ,replace(convert(char(5), dateadd(minute,-1,b.actStrt), 108),':','.') as actEnd FROM ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt) , * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) b right JOIN ( SELECT id = ROW_NUMBER ( ) OVER (ORDER BY actStrt)+1, * FROM smsHeatTracker where Plantno=" + StrMIS_PLANTNO + " and actStrt is not null  and actEnnd is null) a ON b.Id = a.Id) c ON  c.heatNo=t.heatNo and c.plantNo=t.plantNo and c.DateStamp=t.DateStamp and c.actStart=t.actStrt";
            bool StrUpdtStatus = clsObj.DBInsertUpdateDeleteMIS(StrUptQuery);
            return;

            DataTable dt_MIS_TimeUpdt = new DataTable();
            //string StrQueryEndTmeUpdate = "select heatNo from smsHeatTracker where convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + System.DateTime.Now.ToString("yyyy-MM-dd") + "',120) and (actEnd=0 or actEnd is null)  and plantNo='" + StrMIS_PLANTNO + "'";
            string StrQueryEndTmeUpdate = "select heatNo,DateStamp from smsHeatTracker where actEnnd is null and plantNo='" + StrMIS_PLANTNO + "'";

            dt_MIS_TimeUpdt = clsObj.DBSelectQueryMIS_Table(StrQueryEndTmeUpdate);
            if (dt_MIS_TimeUpdt.Rows.Count <= 0)
            {
                return;
            }
            for (int j = 0; j < dt_MIS_TimeUpdt.Rows.Count; j++)
            {
                //string StrQueryOrclCONARC = "SELECT  hp.HEATID_CUST AS HEAT_ID,hp.PLANT ||hp.PLANTNO    AS PLANT,hp.TREATID_CUST AS TREAT_ID,hp.PLANLADLETYPE ||hp.PLANLADLENO AS LADLE_ID, decode(replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.')) AS PLAN_START_TIME,decode(replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'))   AS PLAN_END_TIME, replace(SUBSTR(ACTTREATSTART,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.') ) AS ACT_END_DATE  ,hd.ladleno,SUBSTR(hp.ACTTREATSTART,1,10) as Act_Start_Dt,hp.HEATID FROM PP_HEAT_PLANT hp left outer join PD_HEATDATA hd on hp.HEATID=hd.HEATID   where hd.PLANT='" + StrPLANT + "' and  hp.EXPIRATIONDATE='VALID' and SUBSTR(hp.ACTTREATSTART,1,10)='2014-04-15' and hp.PLANT='" + StrPLANT + "'and hp.PLANTNO='" + StrPLANTNO + "' and hp.HEATID_CUST='" + dt_MIS_TimeUpdt.Rows[j][0].ToString() + "'  ORDER BY hp.HEATID DESC ";
                //string StrQueryOrclCONARC = "SELECT  hp.HEATID_CUST AS HEAT_ID,hp.PLANT ||hp.PLANTNO    AS PLANT,hp.TREATID_CUST AS TREAT_ID,hp.PLANLADLETYPE ||hp.PLANLADLENO AS LADLE_ID,decode(replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.')) AS PLAN_START_TIME,decode(replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'))   AS PLAN_END_TIME, replace(SUBSTR(ACTTREATSTART,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.'),0,null, replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.') ) AS ACT_END_DATE ,hd.ladleno,SUBSTR(hp.ACTTREATSTART,1,10) as Act_Start_Dt,SUBSTR(hp.ACTTREATEND,1,10) as Act_End_Dt,hp.HEATID FROM PP_HEAT_PLANT hp left outer join PD_HEATDATA hd on hp.HEATID=hd.HEATID   where hd.PLANT='" + StrPLANT + "' and hp.HEATID_CUST='" + dt_MIS_TimeUpdt.Rows[j][0].ToString() + "' and hp.EXPIRATIONDATE='VALID' and hp.PLANTNO='" + StrPLANTNO + "' and hp.PLANT='" + StrPLANT + "' ORDER BY hp.HEATID DESC ";
                string StrQueryOrclCONARC = "SELECT  hp.HEATID_CUST AS HEAT_ID,hp.PLANT ||hp.PLANTNO    AS PLANT,hp.TREATID_CUST AS TREAT_ID,hp.PLANLADLETYPE ||hp.PLANLADLENO AS LADLE_ID,decode(replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATSTART,11,6) ,':','.')) AS PLAN_START_TIME,decode(replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'),0,null,replace(SUBSTR(hp.PLANTREATEND,11,6) ,':','.'))   AS PLAN_END_TIME, replace(SUBSTR(ACTTREATSTART,11,6) ,':','.') AS ACT_START_DATE, decode(replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.'),0,null, replace(SUBSTR(hp.ACTTREATEND,11,6) ,':','.') ) AS ACT_END_DATE,hd.ladleno,hp.HEATID,SUBSTR(hp.ACTTREATSTART,1,10) as Act_Start_Dt,SUBSTR(hp.ACTTREATEND,1,10) as Act_End_Dt,replace(hp.ACTTREATSTART,',','.') AS ACTSTRT,replace(hp.ACTTREATEND,',','.') AS ACTENND,Hd.Steelgradecode_Act as Grade,Hp.Heatid as int_heatid FROM PP_HEAT_PLANT hp left outer join PD_HEATDATA hd on hp.HEATID=hd.HEATID   where hd.PLANT='" + StrPLANT + "' and SUBSTR(hp.ACTTREATSTART,1,10)='" + Convert.ToDateTime(dt_MIS_TimeUpdt.Rows[j][1]).ToString("yyyy-MM-dd") + "' and  hp.HEATID_CUST='" + dt_MIS_TimeUpdt.Rows[j][0].ToString() + "' and hp.EXPIRATIONDATE='VALID'  and hp.PLANT='" + StrPLANT + "'  and (hp.ACTTREATEND is not null) and hp.PLANTNO='" + StrPLANTNO + "' ORDER BY hp.HEATID DESC ";
                DataTable dt_Orcl_Conarc = new DataTable();
                dt_Orcl_Conarc = clsObj.DBSelectQueryStatusScreen_Table(StrQueryOrclCONARC);
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
                    object StrActEndDte = DBNull.Value;
                    object StrActStrt = DBNull.Value;
                    object StrActEnnd = DBNull.Value;
                    for (int i = 0; i < dt_Orcl_Conarc.Rows.Count; i++)
                    {


                        StrHEATID = dt_Orcl_Conarc.Rows[i][0].ToString();
                        StrTREATID = dt_Orcl_Conarc.Rows[i][2].ToString();

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
                        if (dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString() != string.Empty)
                        {
                            //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                            StrActEndDte = dt_Orcl_Conarc.Rows[i]["Act_End_Dt"].ToString();
                        }
                        else
                        { StrActEndDte = DBNull.Value; }
                        if (dt_Orcl_Conarc.Rows[i]["ActStrt"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ActStrt"].ToString() != string.Empty)
                        {
                            //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                            StrActStrt = dt_Orcl_Conarc.Rows[i]["ActStrt"].ToString();
                        }
                        else
                        { StrActStrt = DBNull.Value; }
                        if (dt_Orcl_Conarc.Rows[i]["ActEnnd"].ToString() != null && dt_Orcl_Conarc.Rows[i]["ActEnnd"].ToString() != string.Empty)
                        {
                            //StrActStart = Convert.ToDateTime(dt_Orcl_Conarc.Rows[i][6].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                            StrActEnnd = dt_Orcl_Conarc.Rows[i]["ActEnnd"].ToString();
                        }
                        else
                        { StrActEnnd = DBNull.Value; }
                        string StrUpdate_Insert_Query = "";
                        if (StrActEnd.ToString() != "00.00")
                        {
                            //StrUpdate_Insert_Query = "insert smsHeatTracker(plantNo,heatNo,actStart,actEnd,duration,plndStart,plndEnd,DateStamp,LadleNo,actStrt,actEnnd)values('" + StrMIS_PLANTNO + "','" + StrHEATID + "','0','" + StrActEnd + "','','" + StrPlndStart + "','" + StrPlndEnd + "','" + StrActEndDte + "','" + StrLadleNo + "','" + StrActStrt + "','" + StrActEnnd + "')";
                            StrUpdate_Insert_Query = "Update smsHeatTracker set LadleNo='" + StrLadleNo + "',actStart='" + StrActStart + "',actEnd='" + StrActEnd + "',plndStart='" + StrPlndStart + "',plndEnd='" + StrPlndEnd + "',actStrt='" + StrActStrt + "',actEnnd='" + StrActEnnd + "' where plantNo='" + StrMIS_PLANTNO + "' and heatNo='" + StrHEATID + "' and convert(varchar(10),DateStamp,120)=convert(varchar(10),'" + StrActStartDte + "',120) and actStart='" + StrActStart + "'";// and DateStamp=Convert(varchar(10),'" + System.DateTime.Now.ToString() + "',120)";
                        }
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdate_Insert_Query);
                    }
                }
            }
        }
        # endregion


        # region "CAST SEQUENCE REPORT"

        private void timer_CastSequence_Tick(object sender, EventArgs e)
        {
            try
            {
                string StrCastSeq_caster2_Max = "";
                string StrCastSeqMaxID = "select max(SEQ_NO) as SEQ_NO from pdc_Sequence";
                DBConnections clsObj_CastSeq = new DBConnections();
                DataTable dt_CastSeq = new DataTable();
                dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
                if (dt_CastSeq.Rows.Count > 0)
                {

                    StrCastSeq_caster2_Max = dt_CastSeq.Rows[0][0].ToString();
                    Fn_Caster1_CastSeq_1(StrCastSeq_caster2_Max);
                    Fn_Caster1_CastSeq_2(StrCastSeq_caster2_Max);
                    Fn_Caster1_CastSeq_3(StrCastSeq_caster2_Max);
                    Fn_Caster1_CastSeq_4(StrCastSeq_caster2_Max);
                    Fn_Caster1_CastSeq_5(StrCastSeq_caster2_Max);

                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("CastingFinished_CheckStatus_Timer_Tick" + ex.ToString());
                txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
            }
            try
            {
                string StrCastSeq_caster2_Max = "";
                string StrCastSeqMaxID = "select max(SEQ_NO) as SEQ_NO from pdc_Sequence";
                DBConnections clsObj_CastSeq = new DBConnections();
                DataTable dt_CastSeq = new DataTable();
                dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
                if (dt_CastSeq.Rows.Count > 0)
                {

                    StrCastSeq_caster2_Max = dt_CastSeq.Rows[0][0].ToString();
                    Fn_Caster2_CastSeq_1(StrCastSeq_caster2_Max);
                    Fn_Caster2_CastSeq_2(StrCastSeq_caster2_Max);
                    Fn_Caster2_CastSeq_3(StrCastSeq_caster2_Max);
                    Fn_Caster2_CastSeq_4(StrCastSeq_caster2_Max);
                    Fn_Caster2_CastSeq_5(StrCastSeq_caster2_Max);

                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("CastingFinished_CheckStatus_Timer_Tick" + ex.ToString());
                txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
            }
            try
            {
                string StrCastSeq_caster2_Max = "";
                string StrCastSeqMaxID = "select max(SEQ_NO) as SEQ_NO from pdc_Sequence";
                DBConnections clsObj_CastSeq = new DBConnections();
                DataTable dt_CastSeq = new DataTable();
                dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
                if (dt_CastSeq.Rows.Count > 0)
                {

                    StrCastSeq_caster2_Max = dt_CastSeq.Rows[0][0].ToString();
                    Fn_Caster3_CastSeq_1(StrCastSeq_caster2_Max);
                    Fn_Caster3_CastSeq_2(StrCastSeq_caster2_Max);
                    Fn_Caster3_CastSeq_3(StrCastSeq_caster2_Max);
                    Fn_Caster3_CastSeq_4(StrCastSeq_caster2_Max);
                    Fn_Caster3_CastSeq_5(StrCastSeq_caster2_Max);

                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("CastingFinished_CheckStatus_Timer_Tick" + ex.ToString());
                txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
            }
        }

        private void Fn_Caster1_CastSeq_1(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT steel_id AS steelid,seq_no AS sequenceno,cast_start_time AS caststart FROM pdc_sequence WHERE seq_no='" + StrCastSeq_caster2_Max + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSEQUENCENO = "";
                    string StrCASTSTART = "";
                    StrSTEELID = dt_CastSeq.Rows[i][0].ToString();
                    StrSEQUENCENO = dt_CastSeq.Rows[i][1].ToString();
                    if (dt_CastSeq.Rows[i][2].ToString() != null && dt_CastSeq.Rows[i][2].ToString() != string.Empty)
                    {
                        //StrCASTSTART = Convert.ToDateTime(dt_CastSeq.Rows[i][2].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrCASTSTART = Convert.ToDateTime(dt_CastSeq.Rows[i][2].ToString().Replace(",", ".")).ToString();
                    }
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SeqNo from CastSeqReportHead where SeqNo='" + StrSEQUENCENO + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHead set SeqNo='" + StrSEQUENCENO + "',SteelID='" + StrSTEELID + "',Start='" + StrCASTSTART + "' where SeqNo='" + StrSEQUENCENO + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert CastSeqReportHead(SeqNo,SteelID,Start)values('" + StrSEQUENCENO + "','" + StrSTEELID + "','" + StrCASTSTART + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void Fn_Caster1_CastSeq_2(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT h.steel_id AS SteelId,  MAX (h.ladle_open_time) AS prod_time,MAX (h.heatid) AS heat_no,  MAX (h.grade_code) AS grade_code,  MAX (prod_orderid) AS prod_orderid,  MAX (seq.seq_no) AS SeqNo,  MAX (h.ladle_open_weight) / 1000   AS ladle_weight,  SUM (s.cast_len_end  - s.cast_len_bgn) AS cast_length,  MIN (s.heat_in_mold_time) AS cast_start_time,  MAX (hshift.crew_id)  AS crew,  MAX (s.heat_out_mold_time) AS cast_end_time,  MAX (h.tund_powder_type)  AS tund_powder_type,  MAX (h.tund_powder_weight) AS tund_powder_weight,  MAX (s.cast_powder_type) AS mold_powder_type,  MAX (s.cast_powder_weight) AS mold_powder_weight,  MAX (TRUNC(v.yield,2)) AS yield,  MAX (param.value) AS casterid,MAX (s.cast_len_bgn) AS cast_len_bgn,  MAX (s.cast_len_end) AS cast_len_end FROM pdc_heat h,  pdc_sequence seq,  pdc_strand s,  pdc_heat_strand hs,  pdc_heat_shift hshift,  v_heat_weight v,  gdc_param param WHERE seq.steel_id(+) = h.sequence_steel_id AND s.steel_id(+) = hs.strand_steel_id AND hs.heat_steel_id(+)     = h.steel_id AND hshift.heat_steel_id(+) = h.steel_id AND v.steel_id(+) = h.steel_id AND param.par_context = 'CASTER' AND param.par_name = 'CASTERID' AND seq.seq_no = '" + StrCastSeq_caster2_Max + "' GROUP BY h.steel_id ORDER BY h.steel_id ";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSEQUENCENO = "";
                    string StrProdTime = "", StrCASTStart = "", StrCASTEnd = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSEQUENCENO = dt_CastSeq.Rows[i]["SeqNo"].ToString();
                    if (dt_CastSeq.Rows[i]["PROD_TIME"].ToString() != null && dt_CastSeq.Rows[i]["PROD_TIME"].ToString() != string.Empty)
                    {
                        StrProdTime = Convert.ToDateTime(dt_CastSeq.Rows[i]["PROD_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    if (dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != null && dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != string.Empty)
                    {
                        StrCASTStart = Convert.ToDateTime(dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    if (dt_CastSeq.Rows[i]["CAST_END_TIME"].ToString() != null && dt_CastSeq.Rows[i]["CAST_END_TIME"].ToString() != string.Empty)
                    {
                        StrCASTEnd = Convert.ToDateTime(dt_CastSeq.Rows[i]["CAST_END_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrSEQUENCENO + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHead set Prod_Time='" + StrProdTime + "',Heat_No='" + dt_CastSeq.Rows[i]["Heat_No"].ToString() + "',Grade_Code='" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "',Prod_OrderId='" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "',Ladle_Weight='" + dt_CastSeq.Rows[i]["Ladle_Weight"].ToString() + "',Cast_Length='" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "',Cast_Start_Time='" + StrCASTStart + "',Crew='" + dt_CastSeq.Rows[i]["Crew"].ToString() + "',Cast_End_Time='" + StrCASTEnd + "',Tund_Powder_Type='" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "',Tund_Powder_Weight='" + dt_CastSeq.Rows[i]["Tund_Powder_Weight"].ToString() + "',Mold_Powder_Type='" + dt_CastSeq.Rows[i]["Mold_Powder_Type"].ToString() + "',Mold_Powder_Weight='" + dt_CastSeq.Rows[i]["Mold_Powder_Weight"].ToString() + "',Yield='" + dt_CastSeq.Rows[i]["Yield"].ToString() + "',CasterId='" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "',Cast_Len_Bgn='" + dt_CastSeq.Rows[i]["Cast_Len_Bgn"].ToString() + "',Cast_Len_End='" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + dt_CastSeq.Rows[i]["SeqNo"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert CastSeqReportHeadDtl(SteelId,Prod_Time,Heat_No,Grade_Code,Prod_OrderId,SeqNo,Ladle_Weight,Cast_Length,Cast_Start_Time,Crew,Cast_End_Time,Tund_Powder_Type,Tund_Powder_Weight,Mold_Powder_Type,Mold_Powder_Weight,Yield,CasterId,Cast_Len_Bgn,Cast_Len_End)values('" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "','" + StrProdTime + "','" + dt_CastSeq.Rows[i]["Heat_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "','" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "','" + dt_CastSeq.Rows[i]["SeqNo"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "','" + StrCASTStart + "','" + dt_CastSeq.Rows[i]["Crew"].ToString() + "','" + StrCASTEnd + "','" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Powder_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold_Powder_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold_Powder_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Yield"].ToString() + "','" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_Bgn"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void Fn_Caster1_CastSeq_3(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT p.heat_steel_id as STEELID,SUM (piece_weight_calc) / 1000 AS crop_weight FROM pdc_piece p,pdc_heat h,pdc_sequence seq WHERE h.steel_id  = p.heat_steel_id AND seq.steel_id  = h.sequence_steel_id AND p.piece_type IN (1, 2) AND seq.seq_no    = '" + StrCastSeq_caster2_Max + "' GROUP BY p.heat_steel_id";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrCrop_Weight = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrCrop_Weight = dt_CastSeq.Rows[i]["crop_weight"].ToString();
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrCastSeq_caster2_Max + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHeadDtl set Crop_Weight='" + StrCrop_Weight + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + StrCastSeq_caster2_Max + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_CastSeq_4(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT p.heat_steel_id as STEELID,SUM (sample_weight) / 1000 AS sample_weight FROM pdc_piece p,pdc_heat h,pdc_sequence seq WHERE h.steel_id = p.heat_steel_id AND seq.steel_id = h.sequence_steel_id AND seq.seq_no   = '" + StrCastSeq_caster2_Max + "' GROUP BY p.heat_steel_id";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSample_Weight = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSample_Weight = dt_CastSeq.Rows[i]["Sample_weight"].ToString();
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrCastSeq_caster2_Max + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHeadDtl set Sample_Weight='" + StrSample_Weight + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + StrCastSeq_caster2_Max + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_CastSeq_5(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT p.heat_steel_id as STEELID,  SUM (piece_length) / 1000 AS slab_length,  SUM (DECODE(NVL(piece_weight_meas,0), 0, piece_weight_calc , piece_weight_meas)) / 1000 AS slab_weight,  COUNT (*) AS slab_count FROM pdc_piece p,pdc_heat h,pdc_sequence seq WHERE (p.piece_type = 0 OR p.piece_type = 4) AND h.steel_id = p.heat_steel_id AND seq.steel_id = h.sequence_steel_id AND seq.seq_no = '" + StrCastSeq_caster2_Max + "' GROUP BY p.heat_steel_id";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSlab_Length = "", StrSlab_Weight = "", StrSlab_Count = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSlab_Length = dt_CastSeq.Rows[i]["slab_length"].ToString();
                    StrSlab_Weight = dt_CastSeq.Rows[i]["slab_Weight"].ToString();
                    StrSlab_Count = dt_CastSeq.Rows[i]["slab_Count"].ToString();
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrCastSeq_caster2_Max + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHeadDtl set slab_length='" + StrSlab_Length + "',slab_Weight='" + StrSlab_Weight + "',slab_Count='" + StrSlab_Count + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + StrCastSeq_caster2_Max + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                }
            }
        }


        private void Fn_Caster2_CastSeq_1(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT steel_id AS steelid,seq_no AS sequenceno,cast_start_time AS caststart FROM pdc_sequence WHERE seq_no='" + StrCastSeq_caster2_Max + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSEQUENCENO = "";
                    string StrCASTSTART = "";
                    StrSTEELID = dt_CastSeq.Rows[i][0].ToString();
                    StrSEQUENCENO = dt_CastSeq.Rows[i][1].ToString();
                    if (dt_CastSeq.Rows[i][2].ToString() != null && dt_CastSeq.Rows[i][2].ToString() != string.Empty)
                    {
                        //StrCASTSTART = Convert.ToDateTime(dt_CastSeq.Rows[i][2].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrCASTSTART = Convert.ToDateTime(dt_CastSeq.Rows[i][2].ToString().Replace(",", ".")).ToString();
                    }
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SeqNo from CastSeqReportHead where SeqNo='" + StrSEQUENCENO + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHead set SeqNo='" + StrSEQUENCENO + "',SteelID='" + StrSTEELID + "',Start='" + StrCASTSTART + "' where SeqNo='" + StrSEQUENCENO + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert CastSeqReportHead(SeqNo,SteelID,Start)values('" + StrSEQUENCENO + "','" + StrSTEELID + "','" + StrCASTSTART + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void Fn_Caster2_CastSeq_2(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT h.steel_id AS SteelId,  MAX (h.ladle_open_time) AS prod_time,MAX (h.heatid) AS heat_no,  MAX (h.grade_code) AS grade_code,  MAX (prod_orderid) AS prod_orderid,  MAX (seq.seq_no) AS SeqNo,  MAX (h.ladle_open_weight) / 1000   AS ladle_weight,  SUM (s.cast_len_end  - s.cast_len_bgn) AS cast_length,  MIN (s.heat_in_mold_time) AS cast_start_time,  MAX (hshift.crew_id)  AS crew,  MAX (s.heat_out_mold_time) AS cast_end_time,  MAX (h.tund_powder_type)  AS tund_powder_type,  MAX (h.tund_powder_weight) AS tund_powder_weight,  MAX (s.cast_powder_type) AS mold_powder_type,  MAX (s.cast_powder_weight) AS mold_powder_weight,  MAX (TRUNC(v.yield,2)) AS yield,  MAX (param.value) AS casterid,MAX (s.cast_len_bgn) AS cast_len_bgn,  MAX (s.cast_len_end) AS cast_len_end FROM pdc_heat h,  pdc_sequence seq,  pdc_strand s,  pdc_heat_strand hs,  pdc_heat_shift hshift,  v_heat_weight v,  gdc_param param WHERE seq.steel_id(+) = h.sequence_steel_id AND s.steel_id(+) = hs.strand_steel_id AND hs.heat_steel_id(+)     = h.steel_id AND hshift.heat_steel_id(+) = h.steel_id AND v.steel_id(+) = h.steel_id AND param.par_context = 'CASTER' AND param.par_name = 'CASTERID' AND seq.seq_no = '" + StrCastSeq_caster2_Max + "' GROUP BY h.steel_id ORDER BY h.steel_id ";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSEQUENCENO = "";
                    string StrProdTime = "", StrCASTStart = "", StrCASTEnd = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSEQUENCENO = dt_CastSeq.Rows[i]["SeqNo"].ToString();
                    if (dt_CastSeq.Rows[i]["PROD_TIME"].ToString() != null && dt_CastSeq.Rows[i]["PROD_TIME"].ToString() != string.Empty)
                    {
                        StrProdTime = Convert.ToDateTime(dt_CastSeq.Rows[i]["PROD_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    if (dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != null && dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != string.Empty)
                    {
                        StrCASTStart = Convert.ToDateTime(dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    if (dt_CastSeq.Rows[i]["CAST_END_TIME"].ToString() != null && dt_CastSeq.Rows[i]["CAST_END_TIME"].ToString() != string.Empty)
                    {
                        StrCASTEnd = Convert.ToDateTime(dt_CastSeq.Rows[i]["CAST_END_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrSEQUENCENO + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHead set Prod_Time='" + StrProdTime + "',Heat_No='" + dt_CastSeq.Rows[i]["Heat_No"].ToString() + "',Grade_Code='" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "',Prod_OrderId='" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "',Ladle_Weight='" + dt_CastSeq.Rows[i]["Ladle_Weight"].ToString() + "',Cast_Length='" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "',Cast_Start_Time='" + StrCASTStart + "',Crew='" + dt_CastSeq.Rows[i]["Crew"].ToString() + "',Cast_End_Time='" + StrCASTEnd + "',Tund_Powder_Type='" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "',Tund_Powder_Weight='" + dt_CastSeq.Rows[i]["Tund_Powder_Weight"].ToString() + "',Mold_Powder_Type='" + dt_CastSeq.Rows[i]["Mold_Powder_Type"].ToString() + "',Mold_Powder_Weight='" + dt_CastSeq.Rows[i]["Mold_Powder_Weight"].ToString() + "',Yield='" + dt_CastSeq.Rows[i]["Yield"].ToString() + "',CasterId='" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "',Cast_Len_Bgn='" + dt_CastSeq.Rows[i]["Cast_Len_Bgn"].ToString() + "',Cast_Len_End='" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + dt_CastSeq.Rows[i]["SeqNo"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert CastSeqReportHeadDtl(SteelId,Prod_Time,Heat_No,Grade_Code,Prod_OrderId,SeqNo,Ladle_Weight,Cast_Length,Cast_Start_Time,Crew,Cast_End_Time,Tund_Powder_Type,Tund_Powder_Weight,Mold_Powder_Type,Mold_Powder_Weight,Yield,CasterId,Cast_Len_Bgn,Cast_Len_End)values('" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "','" + StrProdTime + "','" + dt_CastSeq.Rows[i]["Heat_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "','" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "','" + dt_CastSeq.Rows[i]["SeqNo"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "','" + StrCASTStart + "','" + dt_CastSeq.Rows[i]["Crew"].ToString() + "','" + StrCASTEnd + "','" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Powder_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold_Powder_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold_Powder_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Yield"].ToString() + "','" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_Bgn"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void Fn_Caster2_CastSeq_3(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT p.heat_steel_id as STEELID,SUM (piece_weight_calc) / 1000 AS crop_weight FROM pdc_piece p,pdc_heat h,pdc_sequence seq WHERE h.steel_id  = p.heat_steel_id AND seq.steel_id  = h.sequence_steel_id AND p.piece_type IN (1, 2) AND seq.seq_no    = '" + StrCastSeq_caster2_Max + "' GROUP BY p.heat_steel_id";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrCrop_Weight = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrCrop_Weight = dt_CastSeq.Rows[i]["crop_weight"].ToString();
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrCastSeq_caster2_Max + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHeadDtl set Crop_Weight='" + StrCrop_Weight + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + StrCastSeq_caster2_Max + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_CastSeq_4(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT p.heat_steel_id as STEELID,SUM (sample_weight) / 1000 AS sample_weight FROM pdc_piece p,pdc_heat h,pdc_sequence seq WHERE h.steel_id = p.heat_steel_id AND seq.steel_id = h.sequence_steel_id AND seq.seq_no   = '" + StrCastSeq_caster2_Max + "' GROUP BY p.heat_steel_id";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSample_Weight = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSample_Weight = dt_CastSeq.Rows[i]["Sample_weight"].ToString();
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrCastSeq_caster2_Max + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHeadDtl set Sample_Weight='" + StrSample_Weight + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + StrCastSeq_caster2_Max + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_CastSeq_5(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT p.heat_steel_id as STEELID,  SUM (piece_length) / 1000 AS slab_length,  SUM (DECODE(NVL(piece_weight_meas,0), 0, piece_weight_calc , piece_weight_meas)) / 1000 AS slab_weight,  COUNT (*) AS slab_count FROM pdc_piece p,pdc_heat h,pdc_sequence seq WHERE (p.piece_type = 0 OR p.piece_type = 4) AND h.steel_id = p.heat_steel_id AND seq.steel_id = h.sequence_steel_id AND seq.seq_no = '" + StrCastSeq_caster2_Max + "' GROUP BY p.heat_steel_id";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSlab_Length = "", StrSlab_Weight = "", StrSlab_Count = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSlab_Length = dt_CastSeq.Rows[i]["slab_length"].ToString();
                    StrSlab_Weight = dt_CastSeq.Rows[i]["slab_Weight"].ToString();
                    StrSlab_Count = dt_CastSeq.Rows[i]["slab_Count"].ToString();
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrCastSeq_caster2_Max + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHeadDtl set slab_length='" + StrSlab_Length + "',slab_Weight='" + StrSlab_Weight + "',slab_Count='" + StrSlab_Count + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + StrCastSeq_caster2_Max + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                }
            }
        }

        private void Fn_Caster3_CastSeq_1(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT steel_id AS steelid,seq_no AS sequenceno,cast_start_time AS caststart FROM pdc_sequence WHERE seq_no='" + StrCastSeq_caster2_Max + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSEQUENCENO = "";
                    string StrCASTSTART = "";
                    StrSTEELID = dt_CastSeq.Rows[i][0].ToString();
                    StrSEQUENCENO = dt_CastSeq.Rows[i][1].ToString();
                    if (dt_CastSeq.Rows[i][2].ToString() != null && dt_CastSeq.Rows[i][2].ToString() != string.Empty)
                    {
                        //StrCASTSTART = Convert.ToDateTime(dt_CastSeq.Rows[i][2].ToString().Replace(",", ".")).ToString("HH:mm").Replace(":", ".");
                        StrCASTSTART = Convert.ToDateTime(dt_CastSeq.Rows[i][2].ToString().Replace(",", ".")).ToString();
                    }
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SeqNo from CastSeqReportHead where SeqNo='" + StrSEQUENCENO + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHead set SeqNo='" + StrSEQUENCENO + "',SteelID='" + StrSTEELID + "',Start='" + StrCASTSTART + "' where SeqNo='" + StrSEQUENCENO + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert CastSeqReportHead(SeqNo,SteelID,Start)values('" + StrSEQUENCENO + "','" + StrSTEELID + "','" + StrCASTSTART + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void Fn_Caster3_CastSeq_2(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT h.steel_id AS SteelId,  MAX (h.ladle_open_time) AS prod_time,MAX (h.heatid) AS heat_no,  MAX (h.grade_code) AS grade_code,  MAX (prod_orderid) AS prod_orderid,  MAX (seq.seq_no) AS SeqNo,  MAX (h.ladle_open_weight) / 1000   AS ladle_weight,  SUM (s.cast_len_end  - s.cast_len_bgn) AS cast_length,  MIN (s.heat_in_mold_time) AS cast_start_time,  MAX (hshift.crew_id)  AS crew,  MAX (s.heat_out_mold_time) AS cast_end_time,  MAX (h.tund_powder_type)  AS tund_powder_type,  MAX (h.tund_powder_weight) AS tund_powder_weight,  MAX (s.cast_powder_type) AS mold_powder_type,  MAX (s.cast_powder_weight) AS mold_powder_weight,  MAX (TRUNC(v.yield,2)) AS yield,  MAX (param.value) AS casterid,MAX (s.cast_len_bgn) AS cast_len_bgn,  MAX (s.cast_len_end) AS cast_len_end FROM pdc_heat h,  pdc_sequence seq,  pdc_strand s,  pdc_heat_strand hs,  pdc_heat_shift hshift,  v_heat_weight v,  gdc_param param WHERE seq.steel_id(+) = h.sequence_steel_id AND s.steel_id(+) = hs.strand_steel_id AND hs.heat_steel_id(+)     = h.steel_id AND hshift.heat_steel_id(+) = h.steel_id AND v.steel_id(+) = h.steel_id AND param.par_context = 'CASTER' AND param.par_name = 'CASTERID' AND seq.seq_no = '" + StrCastSeq_caster2_Max + "' GROUP BY h.steel_id ORDER BY h.steel_id ";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSEQUENCENO = "";
                    string StrProdTime = "", StrCASTStart = "", StrCASTEnd = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSEQUENCENO = dt_CastSeq.Rows[i]["SeqNo"].ToString();
                    if (dt_CastSeq.Rows[i]["PROD_TIME"].ToString() != null && dt_CastSeq.Rows[i]["PROD_TIME"].ToString() != string.Empty)
                    {
                        StrProdTime = Convert.ToDateTime(dt_CastSeq.Rows[i]["PROD_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    if (dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != null && dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != string.Empty)
                    {
                        StrCASTStart = Convert.ToDateTime(dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    if (dt_CastSeq.Rows[i]["CAST_END_TIME"].ToString() != null && dt_CastSeq.Rows[i]["CAST_END_TIME"].ToString() != string.Empty)
                    {
                        StrCASTEnd = Convert.ToDateTime(dt_CastSeq.Rows[i]["CAST_END_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrSEQUENCENO + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHead set Prod_Time='" + StrProdTime + "',Heat_No='" + dt_CastSeq.Rows[i]["Heat_No"].ToString() + "',Grade_Code='" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "',Prod_OrderId='" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "',Ladle_Weight='" + dt_CastSeq.Rows[i]["Ladle_Weight"].ToString() + "',Cast_Length='" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "',Cast_Start_Time='" + StrCASTStart + "',Crew='" + dt_CastSeq.Rows[i]["Crew"].ToString() + "',Cast_End_Time='" + StrCASTEnd + "',Tund_Powder_Type='" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "',Tund_Powder_Weight='" + dt_CastSeq.Rows[i]["Tund_Powder_Weight"].ToString() + "',Mold_Powder_Type='" + dt_CastSeq.Rows[i]["Mold_Powder_Type"].ToString() + "',Mold_Powder_Weight='" + dt_CastSeq.Rows[i]["Mold_Powder_Weight"].ToString() + "',Yield='" + dt_CastSeq.Rows[i]["Yield"].ToString() + "',CasterId='" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "',Cast_Len_Bgn='" + dt_CastSeq.Rows[i]["Cast_Len_Bgn"].ToString() + "',Cast_Len_End='" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + dt_CastSeq.Rows[i]["SeqNo"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert CastSeqReportHeadDtl(SteelId,Prod_Time,Heat_No,Grade_Code,Prod_OrderId,SeqNo,Ladle_Weight,Cast_Length,Cast_Start_Time,Crew,Cast_End_Time,Tund_Powder_Type,Tund_Powder_Weight,Mold_Powder_Type,Mold_Powder_Weight,Yield,CasterId,Cast_Len_Bgn,Cast_Len_End)values('" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "','" + StrProdTime + "','" + dt_CastSeq.Rows[i]["Heat_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "','" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "','" + dt_CastSeq.Rows[i]["SeqNo"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "','" + StrCASTStart + "','" + dt_CastSeq.Rows[i]["Crew"].ToString() + "','" + StrCASTEnd + "','" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Powder_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold_Powder_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold_Powder_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Yield"].ToString() + "','" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_Bgn"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void Fn_Caster3_CastSeq_3(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT p.heat_steel_id as STEELID,SUM (piece_weight_calc) / 1000 AS crop_weight FROM pdc_piece p,pdc_heat h,pdc_sequence seq WHERE h.steel_id  = p.heat_steel_id AND seq.steel_id  = h.sequence_steel_id AND p.piece_type IN (1, 2) AND seq.seq_no    = '" + StrCastSeq_caster2_Max + "' GROUP BY p.heat_steel_id";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrCrop_Weight = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrCrop_Weight = dt_CastSeq.Rows[i]["crop_weight"].ToString();
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrCastSeq_caster2_Max + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHeadDtl set Crop_Weight='" + StrCrop_Weight + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + StrCastSeq_caster2_Max + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_CastSeq_4(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT p.heat_steel_id as STEELID,SUM (sample_weight) / 1000 AS sample_weight FROM pdc_piece p,pdc_heat h,pdc_sequence seq WHERE h.steel_id = p.heat_steel_id AND seq.steel_id = h.sequence_steel_id AND seq.seq_no   = '" + StrCastSeq_caster2_Max + "' GROUP BY p.heat_steel_id";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSample_Weight = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSample_Weight = dt_CastSeq.Rows[i]["Sample_weight"].ToString();
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrCastSeq_caster2_Max + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHeadDtl set Sample_Weight='" + StrSample_Weight + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + StrCastSeq_caster2_Max + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_CastSeq_5(string StrCastSeq_caster2_Max)
        {
            string StrCastSeqMaxID = "SELECT p.heat_steel_id as STEELID,  SUM (piece_length) / 1000 AS slab_length,  SUM (DECODE(NVL(piece_weight_meas,0), 0, piece_weight_calc , piece_weight_meas)) / 1000 AS slab_weight,  COUNT (*) AS slab_count FROM pdc_piece p,pdc_heat h,pdc_sequence seq WHERE (p.piece_type = 0 OR p.piece_type = 4) AND h.steel_id = p.heat_steel_id AND seq.steel_id = h.sequence_steel_id AND seq.seq_no = '" + StrCastSeq_caster2_Max + "' GROUP BY p.heat_steel_id";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSlab_Length = "", StrSlab_Weight = "", StrSlab_Count = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSlab_Length = dt_CastSeq.Rows[i]["slab_length"].ToString();
                    StrSlab_Weight = dt_CastSeq.Rows[i]["slab_Weight"].ToString();
                    StrSlab_Count = dt_CastSeq.Rows[i]["slab_Count"].ToString();
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from CastSeqReportHeadDtl where SeqNo='" + StrCastSeq_caster2_Max + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update CastSeqReportHeadDtl set slab_length='" + StrSlab_Length + "',slab_Weight='" + StrSlab_Weight + "',slab_Count='" + StrSlab_Count + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + StrCastSeq_caster2_Max + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                }
            }
        }
        # endregion


        # region "HEAT REPORT"
        private void timer_HeatReport_Tick(object sender, EventArgs e)
        {

            try
            {
                FnCasterI_HeatReport();
                FnCasterII_HeatReport();
                FnCasterIII_HeatReport();
            }

            catch (Exception ex)
            {
                MessageBox.Show("CastingFinished_CheckStatus_Timer_Tick" + ex.ToString());
                txtError.Text = ex.ToString() + " AT - " + System.DateTime.Now.ToString();
            }


        }

        private void FnCasterI_HeatReport()
        {
            string StrHeatRpt_caster1_MaxID = "";
            string StrHeatRpt_caster1_Steel_ID = "";
            //string StrHeatRptMaxID = "SELECT STEEL_ID,SEQUENCE,HEAT_ID,IN_SEQ,HEAT_SEQUNCE_NUMBER,SEQ_STEEL_ID,PLAN_NO,STEEL_GRADE,CAST_START_TIME,STATUS,HEAT_STATUS_CODE from ( SELECT STEEL_ID  AS STEEL_ID,SEQ_NO AS SEQUENCE,HEATID AS HEAT_ID,HEAT_SEQ_NO  AS IN_SEQ,SEQUENCE_STEEL_ID AS HEAT_SEQUNCE_NUMBER,SEQUENCE_STEEL_ID AS SEQ_STEEL_ID,PROD_ORDERID      AS PLAN_NO,GRADE_CODE  AS STEEL_GRADE,LADLE_OPEN_TIME   AS CAST_START_TIME,HEAT_STATUS_DESC  AS STATUS,HEAT_STATUS_CODE, RANK() OVER (ORDER BY Ladle_Open_Time DESC) rowno FROM V_HEAT_DATA WHERE (1            =1) AND SEQ_NO =(select max(SEQ_NO) from V_HEAT_DATA) AND HEAT_STATUS_CODE>1 AND SUBSTR(Ladle_Open_Time,1,10)='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "' ORDER BY LADLE_OPEN_TIME DESC) where  rowno =1";
            string StrHeatRptMaxID = "SELECT STEEL_ID  AS STEEL_ID,SEQ_NO AS SEQUENCE,HEATID AS HEAT_ID,HEAT_SEQ_NO  AS IN_SEQ,SEQUENCE_STEEL_ID AS HEAT_SEQUNCE_NUMBER,SEQUENCE_STEEL_ID AS SEQ_STEEL_ID,PROD_ORDERID      AS PLAN_NO,GRADE_CODE        AS STEEL_GRADE,LADLE_OPEN_TIME   AS CAST_START_TIME,HEAT_STATUS_DESC  AS STATUS,HEAT_STATUS_CODE FROM V_HEAT_DATA WHERE (1            =1) AND SEQ_NO =(select max(SEQ_NO) from V_HEAT_DATA) AND HEAT_STATUS_CODE>1 ORDER BY LADLE_OPEN_TIME DESC";

            DBConnections clsObj_HeatRpt = new DBConnections();
            DataTable dt_HeatRpt = new DataTable();
            dt_HeatRpt = clsObj.DBSelectQueryCASTERI_Table(StrHeatRptMaxID);
            if (dt_HeatRpt.Rows.Count > 0)
            {
                for (int i = 0; i < dt_HeatRpt.Rows.Count; i++)
                {
                    StrHeatRpt_caster1_MaxID = dt_HeatRpt.Rows[i]["SEQUENCE"].ToString();
                    StrHeatRpt_caster1_Steel_ID = dt_HeatRpt.Rows[i]["STEEL_ID"].ToString();
                    Fn_Caster1_HeatRpt_HeatReportHead(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatReportLabAnalysis(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatReportTempReport(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatRpt_PlandataReport(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatRptAnalysisReport(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatRptCutDataReport(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatRptHeader(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatRpt_HeatReportSR(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatRpt_SubMoldStrand(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatRpt_SubPracticeTable(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatRpt_AnalysisElement(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                    Fn_Caster1_HeatRpt_HeatRpt_HeatEquipmentData(StrHeatRpt_caster1_MaxID, StrHeatRpt_caster1_Steel_ID);
                }
            }
        }
        private void FnCasterII_HeatReport()
        {
            string StrHeatRpt_caster2_MaxID = "";
            string StrHeatRpt_caster2_Steel_ID = "";
            //string StrHeatRptMaxID = "SELECT STEEL_ID,SEQUENCE,HEAT_ID,IN_SEQ,HEAT_SEQUNCE_NUMBER,SEQ_STEEL_ID,PLAN_NO,STEEL_GRADE,CAST_START_TIME,STATUS,HEAT_STATUS_CODE from ( SELECT STEEL_ID  AS STEEL_ID,SEQ_NO AS SEQUENCE,HEATID AS HEAT_ID,HEAT_SEQ_NO  AS IN_SEQ,SEQUENCE_STEEL_ID AS HEAT_SEQUNCE_NUMBER,SEQUENCE_STEEL_ID AS SEQ_STEEL_ID,PROD_ORDERID      AS PLAN_NO,GRADE_CODE  AS STEEL_GRADE,LADLE_OPEN_TIME   AS CAST_START_TIME,HEAT_STATUS_DESC  AS STATUS,HEAT_STATUS_CODE, RANK() OVER (ORDER BY Ladle_Open_Time DESC) rowno FROM V_HEAT_DATA WHERE (1            =1) AND SEQ_NO =(select max(SEQ_NO) from V_HEAT_DATA) AND HEAT_STATUS_CODE>1 AND SUBSTR(Ladle_Open_Time,1,10)='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "' ORDER BY LADLE_OPEN_TIME DESC) where  rowno =1";
            string StrHeatRptMaxID = "SELECT STEEL_ID  AS STEEL_ID,SEQ_NO AS SEQUENCE,HEATID AS HEAT_ID,HEAT_SEQ_NO  AS IN_SEQ,SEQUENCE_STEEL_ID AS HEAT_SEQUNCE_NUMBER,SEQUENCE_STEEL_ID AS SEQ_STEEL_ID,PROD_ORDERID      AS PLAN_NO,GRADE_CODE        AS STEEL_GRADE,LADLE_OPEN_TIME   AS CAST_START_TIME,HEAT_STATUS_DESC  AS STATUS,HEAT_STATUS_CODE FROM V_HEAT_DATA WHERE (1            =1) AND SEQ_NO =(select max(SEQ_NO) from V_HEAT_DATA) AND HEAT_STATUS_CODE>1 ORDER BY LADLE_OPEN_TIME DESC";
            DBConnections clsObj_HeatRpt = new DBConnections();
            DataTable dt_HeatRpt = new DataTable();
            dt_HeatRpt = clsObj.DBSelectQueryCASTERII_Table(StrHeatRptMaxID);
            if (dt_HeatRpt.Rows.Count > 0)
            {
                for (int i = 0; i < dt_HeatRpt.Rows.Count; i++)
                {
                    StrHeatRpt_caster2_MaxID = dt_HeatRpt.Rows[i]["SEQUENCE"].ToString();
                    StrHeatRpt_caster2_Steel_ID = dt_HeatRpt.Rows[i]["STEEL_ID"].ToString();
                    Fn_Caster2_HeatRpt_HeatReportHead(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatReportLabAnalysis(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatReportTempReport(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatRpt_PlandataReport(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatRptAnalysisReport(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatRptCutDataReport(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatRptHeader(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatRpt_HeatReportSR(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatRpt_SubMoldStrand(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatRpt_SubPracticeTable(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatRpt_AnalysisElement(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                    Fn_Caster2_HeatRpt_HeatRpt_HeatEquipmentData(StrHeatRpt_caster2_MaxID, StrHeatRpt_caster2_Steel_ID);
                }
            }
        }
        private void FnCasterIII_HeatReport()
        {
            string StrHeatRpt_caster3_MaxID = "";
            string StrHeatRpt_caster3_Steel_ID = "";
            //string StrHeatRptMaxID = "SELECT STEEL_ID,SEQUENCE,HEAT_ID,IN_SEQ,HEAT_SEQUNCE_NUMBER,SEQ_STEEL_ID,PLAN_NO,STEEL_GRADE,CAST_START_TIME,STATUS,HEAT_STATUS_CODE from ( SELECT STEEL_ID  AS STEEL_ID,SEQ_NO AS SEQUENCE,HEATID AS HEAT_ID,HEAT_SEQ_NO  AS IN_SEQ,SEQUENCE_STEEL_ID AS HEAT_SEQUNCE_NUMBER,SEQUENCE_STEEL_ID AS SEQ_STEEL_ID,PROD_ORDERID      AS PLAN_NO,GRADE_CODE  AS STEEL_GRADE,LADLE_OPEN_TIME   AS CAST_START_TIME,HEAT_STATUS_DESC  AS STATUS,HEAT_STATUS_CODE, RANK() OVER (ORDER BY Ladle_Open_Time DESC) rowno FROM V_HEAT_DATA WHERE (1            =1) AND SEQ_NO =(select max(SEQ_NO) from V_HEAT_DATA) AND HEAT_STATUS_CODE>1 AND SUBSTR(Ladle_Open_Time,1,10)='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "' ORDER BY LADLE_OPEN_TIME DESC) where  rowno =1";
            string StrHeatRptMaxID = "SELECT STEEL_ID  AS STEEL_ID,SEQ_NO AS SEQUENCE,HEATID AS HEAT_ID,HEAT_SEQ_NO  AS IN_SEQ,SEQUENCE_STEEL_ID AS HEAT_SEQUNCE_NUMBER,SEQUENCE_STEEL_ID AS SEQ_STEEL_ID,PROD_ORDERID      AS PLAN_NO,GRADE_CODE        AS STEEL_GRADE,LADLE_OPEN_TIME   AS CAST_START_TIME,HEAT_STATUS_DESC  AS STATUS,HEAT_STATUS_CODE FROM V_HEAT_DATA WHERE (1            =1) AND SEQ_NO =(select max(SEQ_NO) from V_HEAT_DATA) AND HEAT_STATUS_CODE>1 ORDER BY LADLE_OPEN_TIME DESC";
            DBConnections clsObj_HeatRpt = new DBConnections();
            DataTable dt_HeatRpt = new DataTable();
            dt_HeatRpt = clsObj.DBSelectQueryCASTERIII_Table(StrHeatRptMaxID);
            if (dt_HeatRpt.Rows.Count > 0)
            {
                for (int i = 0; i < dt_HeatRpt.Rows.Count; i++)
                {
                    StrHeatRpt_caster3_MaxID = dt_HeatRpt.Rows[i]["SEQUENCE"].ToString();
                    StrHeatRpt_caster3_Steel_ID = dt_HeatRpt.Rows[i]["STEEL_ID"].ToString();
                    Fn_Caster3_HeatRpt_HeatReportHead(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatReportLabAnalysis(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatReportTempReport(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatRpt_PlandataReport(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatRptAnalysisReport(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatRptCutDataReport(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatRptHeader(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatRpt_HeatReportSR(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatRpt_SubMoldStrand(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatRpt_SubPracticeTable(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatRpt_AnalysisElement(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                    Fn_Caster3_HeatRpt_HeatRpt_HeatEquipmentData(StrHeatRpt_caster3_MaxID, StrHeatRpt_caster3_Steel_ID);
                }
            }
        }

        private void Fn_Caster1_HeatRpt_HeatReportHead(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = "SELECT STEEL_ID AS STEELID,SEQ_NO  AS SeqNo,HEATID  AS HEATID,HEAT_SEQ_NO AS InSeq,SEQUENCE_STEEL_ID AS Heat_Seq_No,SEQUENCE_STEEL_ID AS SEQ_STEEL_ID,PROD_ORDERID      AS PLAN_NO,GRADE_CODE        AS STEEL_GRADE,LADLE_OPEN_TIME   AS CAST_START_TIME,HEAT_STATUS_DESC  AS STATUS,HEAT_STATUS_CODE FROM V_HEAT_DATA WHERE (1=1)AND SEQ_NO='" + StrCastSeq_caster1_MaxID + "' AND HEAT_STATUS_CODE>2 ORDER BY LADLE_OPEN_TIME DESC";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSEQUENCENO = "";
                    string StrCASTSTART = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSEQUENCENO = dt_CastSeq.Rows[i]["SeqNo"].ToString();
                    if (dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != null && dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != string.Empty)
                    {
                        StrCASTSTART = Convert.ToDateTime(dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SeqNo from HeatReportHead where SeqNo='" + StrSEQUENCENO + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatReportHead set HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "',InSeq='" + dt_CastSeq.Rows[i]["InSeq"].ToString() + "',Heat_Seq_No='" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "',Seq_Steel_Id='" + dt_CastSeq.Rows[i]["Seq_Steel_Id"].ToString() + "',Plan_No='" + dt_CastSeq.Rows[i]["Plan_No"].ToString() + "',Steel_Grade='" + dt_CastSeq.Rows[i]["Steel_Grade"].ToString() + "',Cast_Strat_Time='" + StrCASTSTART + "',Status='" + dt_CastSeq.Rows[i]["Status"].ToString() + "',Heat_Status_Code='" + dt_CastSeq.Rows[i]["Heat_Status_Code"].ToString() + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + dt_CastSeq.Rows[i]["SeqNo"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatReportHead(SteelId,SeqNo,HeatId,InSeq,Heat_Seq_No,Seq_Steel_Id,Plan_No,Steel_Grade,Cast_Strat_Time,Status,Heat_Status_Code)values('" + StrSTEELID + "','" + StrSEQUENCENO + "','" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "','" + dt_CastSeq.Rows[i]["InSeq"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Seq_Steel_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Plan_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Steel_Grade"].ToString() + "','" + StrCASTSTART + "','" + dt_CastSeq.Rows[i]["Status"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Status_Code"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatReportLabAnalysis(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = "SELECT se.grade_code,gs.grade_code,gs.CALC_LIQ_TEMP aimliq,nvl(SUM (DECODE (se.element_name, 'C', min_conc)),0) minc,  nvl(SUM (DECODE (se.element_name, 'Si', min_conc)),0) minsi, nvl(SUM (DECODE (se.element_name, 'Mn', min_conc)),0) minmn, nvl(SUM (DECODE (se.element_name, 'P', min_conc)),0) minp, nvl(SUM (DECODE (se.element_name, 'S', min_conc)),0) mins, nvl(SUM (DECODE (se.element_name, 'Al', min_conc)),0) minal, nvl(SUM (DECODE (se.element_name, 'Al_S', min_conc)),0) minals, nvl(SUM (DECODE (se.element_name, 'Cu', min_conc)),0) mincu, nvl(SUM (DECODE (se.element_name, 'Cr', min_conc)),0) mincr, nvl(SUM (DECODE (se.element_name, 'Mo', min_conc)),0) minmo, nvl(SUM (DECODE (se.element_name, 'Ni', min_conc)),0) minni, nvl(SUM (DECODE (se.element_name, 'V', min_conc)),0) minv, nvl(SUM (DECODE (se.element_name, 'Ti', min_conc)),0) minti, nvl(SUM (DECODE (se.element_name, 'Nb', min_conc)),0) minnb, nvl(SUM (DECODE (se.element_name, 'Ca', min_conc)),0) minca,nvl(SUM (DECODE (se.element_name, 'W', min_conc)),0) minw, nvl(SUM (DECODE (se.element_name, 'Sn', min_conc)),0) minsn, nvl(SUM (DECODE (se.element_name, 'As', min_conc)),0) minas, nvl(SUM (DECODE (se.element_name, 'Te', min_conc)),0) minte, nvl(SUM (DECODE (se.element_name, 'Bi', min_conc)),0) minbi, nvl(SUM (DECODE (se.element_name, 'B', min_conc)),0) minb, nvl(SUM (DECODE (se.element_name, 'Pb', min_conc)),0) minpb, nvl(SUM (DECODE (se.element_name, 'Mg', min_conc)),0) minmg, nvl(SUM (DECODE (se.element_name, 'N', min_conc)),0) minn,nvl(SUM (DECODE (se.element_name, 'Ve', min_conc)),0) minve, nvl(SUM (DECODE (se.element_name, 'Co', min_conc)),0) minco, nvl(SUM (DECODE (se.element_name, 'Ce', min_conc)),0) mince, nvl(SUM (DECODE (se.element_name, 'Sb', min_conc)),0) minsb, nvl(SUM (DECODE (se.element_name, 'Zr', min_conc)),0) minzr, nvl(SUM (DECODE (se.element_name, 'O', min_conc)),0) mino, nvl(SUM (DECODE (se.element_name, 'H', min_conc)),0) minh, nvl(SUM (DECODE (se.element_name, 'C', max_conc)),0) maxc, nvl(SUM (DECODE (se.element_name, 'Si', max_conc)),0) maxsi, nvl(SUM (DECODE (se.element_name, 'Mn', max_conc)),0) maxmn, nvl(SUM (DECODE (se.element_name, 'P', max_conc)),0) maxp, nvl(SUM (DECODE (se.element_name, 'S', max_conc)),0) maxs, nvl(SUM (DECODE (se.element_name, 'Al', max_conc)),0) maxal, nvl(SUM (DECODE (se.element_name, 'Al_S', max_conc)),0) maxals, nvl(SUM (DECODE (se.element_name, 'Cu', max_conc)),0) maxcu, nvl(SUM (DECODE (se.element_name, 'Cr', max_conc)),0) maxcr, nvl(SUM (DECODE (se.element_name, 'Mo', max_conc)),0) maxmo, nvl(SUM (DECODE (se.element_name, 'Ni', max_conc)),0) maxni, nvl(SUM (DECODE (se.element_name, 'V', max_conc)),0) maxv, nvl(SUM (DECODE (se.element_name, 'Ti', max_conc)),0) maxti, nvl(SUM (DECODE (se.element_name, 'Nb', max_conc)),0) maxnb, nvl(SUM (DECODE (se.element_name, 'Ca', max_conc)),0) maxca, nvl(SUM (DECODE (se.element_name, 'W', max_conc)),0) maxw, nvl(SUM (DECODE (se.element_name, 'Sn', max_conc)),0) maxsn, nvl(SUM (DECODE (se.element_name, 'As', max_conc)),0) maxas, nvl(SUM (DECODE (se.element_name, 'Te', max_conc)),0) maxte, nvl(SUM (DECODE (se.element_name, 'Bi', max_conc)),0) maxbi, nvl(SUM (DECODE (se.element_name, 'B', max_conc)),0) maxb, nvl(SUM (DECODE (se.element_name, 'Pb', max_conc)),0) maxpb, nvl(SUM (DECODE (se.element_name, 'Mg', max_conc)),0) maxmg, nvl(SUM (DECODE (se.element_name, 'N', max_conc)),0) maxn, nvl(SUM (DECODE (se.element_name, 'Ve', max_conc)),0) maxve, nvl(SUM (DECODE (se.element_name, 'Co', max_conc)),0) maxco, nvl(SUM (DECODE (se.element_name, 'Ce', max_conc)),0) maxce, nvl(SUM (DECODE (se.element_name, 'Sb', max_conc)),0) maxsb, nvl(SUM (DECODE (se.element_name, 'Zr', max_conc)),0) maxzr, nvl(SUM (DECODE (se.element_name, 'O', max_conc)),0) maxo, nvl(SUM (DECODE (se.element_name, 'H', max_conc)),0) maxh, nvl(SUM (DECODE (se.element_name, 'C', aim_conc)),0) aimc, nvl(SUM (DECODE (se.element_name, 'Si', aim_conc)),0) aimsi, nvl(SUM (DECODE (se.element_name, 'Mn', aim_conc)),0) aimmn, nvl(SUM (DECODE (se.element_name, 'P', aim_conc)),0) aimp, nvl(SUM (DECODE (se.element_name, 'S', aim_conc)),0) aims, nvl(SUM (DECODE (se.element_name, 'Al', aim_conc)),0) aimal, nvl(SUM (DECODE (se.element_name, 'Al_S', aim_conc)),0) aimals, nvl(SUM (DECODE (se.element_name, 'Cu', aim_conc)),0) aimcu, nvl(SUM (DECODE (se.element_name, 'Cr', aim_conc)),0) aimcr, nvl(SUM (DECODE (se.element_name, 'Mo', aim_conc)),0) aimmo, nvl(SUM (DECODE (se.element_name, 'Ni', aim_conc)),0) aimni, nvl(SUM (DECODE (se.element_name, 'V', aim_conc)),0) aimv, nvl(SUM (DECODE (se.element_name, 'Ti', aim_conc)),0) aimti, nvl(SUM (DECODE (se.element_name, 'Nb', aim_conc)),0) aimnb,nvl(SUM (DECODE (se.element_name, 'Ca', aim_conc)),0) aimca, nvl(SUM (DECODE (se.element_name, 'W', aim_conc)),0) aimw, nvl(SUM (DECODE (se.element_name, 'Sn', aim_conc)),0) aimsn, nvl(SUM (DECODE (se.element_name, 'As', aim_conc)),0) aimas,nvl(SUM (DECODE (se.element_name, 'Te', aim_conc)),0) aimte, nvl(SUM (DECODE (se.element_name, 'Bi', aim_conc)),0) aimbi,nvl(SUM (DECODE (se.element_name, 'B', aim_conc)),0) aimb, nvl(SUM (DECODE (se.element_name, 'Pb', aim_conc)),0) aimpb, nvl(SUM (DECODE (se.element_name, 'Mg', aim_conc)),0) aimmg, nvl(SUM (DECODE (se.element_name, 'N', aim_conc)),0) aimn, nvl(SUM (DECODE (se.element_name, 'Ve', aim_conc)),0) aimve, nvl(SUM (DECODE (se.element_name, 'Co', aim_conc)),0) aimco, nvl(SUM (DECODE (se.element_name, 'Ce', aim_conc)),0) aimce, nvl(SUM (DECODE (se.element_name, 'Sb', aim_conc)),0) aimsb, nvl(SUM (DECODE (se.element_name, 'Zr', aim_conc)),0) aimzr, nvl(SUM (DECODE (se.element_name, 'O', aim_conc)),0) aimo, nvl(SUM (DECODE (se.element_name, 'H', aim_conc)),0) aimh FROM gcc_spec_analysis se, gcc_grade_spec gs WHERE ((gs.grade_code = se.grade_code)AND (se.grade_code    =  (SELECT grade_code FROM pdc_heat WHERE (steel_id = '" + StrCastSeq_caster1_SteelID + "')  ))) GROUP BY (se.grade_code, gs.grade_code,gs.CALC_LIQ_TEMP)";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatReportLabAnalysis where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatReportLabAnalysis set LIQ_min='" + dt_CastSeq.Rows[i]["aimliq"].ToString() + "',C_min='" + dt_CastSeq.Rows[i]["minc"].ToString() + "',Si_min='" + dt_CastSeq.Rows[i]["minSi"].ToString() + "',Mn_min='" + dt_CastSeq.Rows[i]["minMn"].ToString() + "',P_min='" + dt_CastSeq.Rows[i]["minP"].ToString() + "',S_min='" + dt_CastSeq.Rows[i]["minS"].ToString() + "',Al_min='" + dt_CastSeq.Rows[i]["minAl"].ToString() + "',Al_S_min='" + dt_CastSeq.Rows[i]["minAls"].ToString() + "',Cu_min='" + dt_CastSeq.Rows[i]["minCu"].ToString() + "',Cr_min='" + dt_CastSeq.Rows[i]["minCr"].ToString() + "',Mo_min='" + dt_CastSeq.Rows[i]["minMo"].ToString() + "',Ni_min='" + dt_CastSeq.Rows[i]["minNi"].ToString() + "',V_min='" + dt_CastSeq.Rows[i]["minV"].ToString() + "',Ti_min='" + dt_CastSeq.Rows[i]["minTi"].ToString() + "',Nb_min='" + dt_CastSeq.Rows[i]["minNb"].ToString() + "',Ca_min='" + dt_CastSeq.Rows[i]["MinCa"].ToString() + "',W_min='" + dt_CastSeq.Rows[i]["minW"].ToString() + "',Sn_min='" + dt_CastSeq.Rows[i]["minSn"].ToString() + "',As_min='" + dt_CastSeq.Rows[i]["minAs"].ToString() + "',Te_min='" + dt_CastSeq.Rows[i]["minTe"].ToString() + "',Bi_min='" + dt_CastSeq.Rows[i]["minBi"].ToString() + "',B_min='" + dt_CastSeq.Rows[i]["minB"].ToString() + "',Pb_min='" + dt_CastSeq.Rows[i]["minPb"].ToString() + "',Mg_min='" + dt_CastSeq.Rows[i]["minMg"].ToString() + "',N_min='" + dt_CastSeq.Rows[i]["minN"].ToString() + "',Ve_min='" + dt_CastSeq.Rows[i]["minVe"].ToString() + "',Co_min='" + dt_CastSeq.Rows[i]["minCo"].ToString() + "',Ce_min='" + dt_CastSeq.Rows[i]["minCe"].ToString() + "',Sb_min='" + dt_CastSeq.Rows[i]["minSb"].ToString() + "',Zr_min='" + dt_CastSeq.Rows[i]["minZr"].ToString() + "',O_min='" + dt_CastSeq.Rows[i]["minO"].ToString() + "',H_min='" + dt_CastSeq.Rows[i]["minH"].ToString() + "',C_max='" + dt_CastSeq.Rows[i]["maxC"].ToString() + "',Si_max='" + dt_CastSeq.Rows[i]["maxSi"].ToString() + "',Mn_max='" + dt_CastSeq.Rows[i]["maxMn"].ToString() + "',P_max='" + dt_CastSeq.Rows[i]["maxP"].ToString() + "',S_max='" + dt_CastSeq.Rows[i]["maxS"].ToString() + "',Al_max='" + dt_CastSeq.Rows[i]["maxAl"].ToString() + "',Al_S_max='" + dt_CastSeq.Rows[i]["maxAlS"].ToString() + "',Cu_max='" + dt_CastSeq.Rows[i]["maxCu"].ToString() + "',Cr_max='" + dt_CastSeq.Rows[i]["maxCr"].ToString() + "',Mo_max='" + dt_CastSeq.Rows[i]["maxMo"].ToString() + "',Ni_max='" + dt_CastSeq.Rows[i]["maxNi"].ToString() + "',V_max='" + dt_CastSeq.Rows[i]["maxV"].ToString() + "',Ti_max='" + dt_CastSeq.Rows[i]["maxTi"].ToString() + "',Nb_max='" + dt_CastSeq.Rows[i]["maxNb"].ToString() + "',Ca_max='" + dt_CastSeq.Rows[i]["maxCa"].ToString() + "',W_max='" + dt_CastSeq.Rows[i]["maxW"].ToString() + "',Sn_max='" + dt_CastSeq.Rows[i]["maxSn"].ToString() + "',As_max='" + dt_CastSeq.Rows[i]["maxAs"].ToString() + "',Te_max='" + dt_CastSeq.Rows[i]["maxTe"].ToString() + "',Bi_max='" + dt_CastSeq.Rows[i]["maxBi"].ToString() + "',B_max='" + dt_CastSeq.Rows[i]["maxB"].ToString() + "',Pb_max='" + dt_CastSeq.Rows[i]["maxPb"].ToString() + "',Mg_max='" + dt_CastSeq.Rows[i]["maxMg"].ToString() + "',N_max='" + dt_CastSeq.Rows[i]["maxN"].ToString() + "',Ve_max='" + dt_CastSeq.Rows[i]["maxVe"].ToString() + "',Co_max='" + dt_CastSeq.Rows[i]["maxCo"].ToString() + "',Ce_max='" + dt_CastSeq.Rows[i]["maxCe"].ToString() + "',Sb_max='" + dt_CastSeq.Rows[i]["maxSb"].ToString() + "',Zr_max='" + dt_CastSeq.Rows[i]["maxZr"].ToString() + "',O_max='" + dt_CastSeq.Rows[i]["maxO"].ToString() + "',H_max='" + dt_CastSeq.Rows[i]["maxH"].ToString() + "',C_aim='" + dt_CastSeq.Rows[i]["aimC"].ToString() + "',Si_aim='" + dt_CastSeq.Rows[i]["aimSi"].ToString() + "',Mn_aim='" + dt_CastSeq.Rows[i]["aimMn"].ToString() + "',P_aim='" + dt_CastSeq.Rows[i]["aimP"].ToString() + "',S_aim='" + dt_CastSeq.Rows[i]["aimS"].ToString() + "',Al_aim='" + dt_CastSeq.Rows[i]["aimAl"].ToString() + "',Al_S_aim='" + dt_CastSeq.Rows[i]["aimAlS"].ToString() + "',Cu_aim='" + dt_CastSeq.Rows[i]["aimCu"].ToString() + "',Cr_aim='" + dt_CastSeq.Rows[i]["aimCr"].ToString() + "',Mo_aim='" + dt_CastSeq.Rows[i]["aimMo"].ToString() + "',Ni_aim='" + dt_CastSeq.Rows[i]["aimNi"].ToString() + "',V_aim='" + dt_CastSeq.Rows[i]["aimV"].ToString() + "',Ti_aim='" + dt_CastSeq.Rows[i]["aimTi"].ToString() + "',Nb_aim='" + dt_CastSeq.Rows[i]["aimNb"].ToString() + "',Ca_aim='" + dt_CastSeq.Rows[i]["aimCa"].ToString() + "',W_aim='" + dt_CastSeq.Rows[i]["aimW"].ToString() + "',Sn_aim='" + dt_CastSeq.Rows[i]["aimSn"].ToString() + "',As_aim='" + dt_CastSeq.Rows[i]["aimAs"].ToString() + "',Te_aim='" + dt_CastSeq.Rows[i]["aimTe"].ToString() + "',Bi_aim='" + dt_CastSeq.Rows[i]["aimBi"].ToString() + "',B_aim='" + dt_CastSeq.Rows[i]["aimB"].ToString() + "',Pb_aim='" + dt_CastSeq.Rows[i]["aimPb"].ToString() + "',Mg_aim='" + dt_CastSeq.Rows[i]["aimMg"].ToString() + "',N_aim='" + dt_CastSeq.Rows[i]["aimN"].ToString() + "',Ve_aim='" + dt_CastSeq.Rows[i]["aimVe"].ToString() + "',Co_aim='" + dt_CastSeq.Rows[i]["aimCo"].ToString() + "',Ce_aim='" + dt_CastSeq.Rows[i]["aimCe"].ToString() + "',Sb_aim='" + dt_CastSeq.Rows[i]["aimSb"].ToString() + "',Zr_aim='" + dt_CastSeq.Rows[i]["aimZr"].ToString() + "',O_aim='" + dt_CastSeq.Rows[i]["aimO"].ToString() + "',H_aim='" + dt_CastSeq.Rows[i]["aimH"].ToString() + "' where SteelId='" + StrCastSeq_caster1_SteelID + "' and SeqNo='" + StrCastSeq_caster1_MaxID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatReportLabAnalysis(SeqNo,SteelId,LIQ_min,C_min,Si_min,Mn_min,P_min,S_min,Al_min,Al_S_min,Cu_min,Cr_min,Mo_min,Ni_min,V_min,Ti_min,Nb_min,Ca_min,W_min,Sn_min,As_min,Te_min,Bi_min,B_min,Pb_min,Mg_min,N_min,Ve_min,Co_min,Ce_min,Sb_min,Zr_min,O_min,H_min,C_max,Si_max,Mn_max,P_max,S_max,Al_max,Al_S_max,Cu_max,Cr_max,Mo_max,Ni_max,V_max,Ti_max,Nb_max,Ca_max,W_max,Sn_max,As_max,Te_max,Bi_max,B_max,Pb_max,Mg_max,N_max,Ve_max,Co_max,Ce_max,Sb_max,Zr_max,O_max,H_max,C_aim,Si_aim,Mn_aim,P_aim,S_aim,Al_aim,Al_S_aim,Cu_aim,Cr_aim,Mo_aim,Ni_aim,V_aim,Ti_aim,Nb_aim,Ca_aim,W_aim,Sn_aim,As_aim,Te_aim,Bi_aim,B_aim,Pb_aim,Mg_aim,N_aim,Ve_aim,Co_aim,Ce_aim,Sb_aim,Zr_aim,O_aim,H_aim)values('" + StrCastSeq_caster1_MaxID + "','" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["aimliq"].ToString() + "','" + dt_CastSeq.Rows[i]["minC"].ToString() + "','" + dt_CastSeq.Rows[i]["minSi"].ToString() + "','" + dt_CastSeq.Rows[i]["minMn"].ToString() + "','" + dt_CastSeq.Rows[i]["minP"].ToString() + "','" + dt_CastSeq.Rows[i]["minS"].ToString() + "','" + dt_CastSeq.Rows[i]["minAl"].ToString() + "','" + dt_CastSeq.Rows[i]["minAlS"].ToString() + "','" + dt_CastSeq.Rows[i]["minCu"].ToString() + "','" + dt_CastSeq.Rows[i]["minCr"].ToString() + "','" + dt_CastSeq.Rows[i]["minMo"].ToString() + "','" + dt_CastSeq.Rows[i]["minNi"].ToString() + "','" + dt_CastSeq.Rows[i]["minV"].ToString() + "','" + dt_CastSeq.Rows[i]["minTi"].ToString() + "','" + dt_CastSeq.Rows[i]["minNb"].ToString() + "','" + dt_CastSeq.Rows[i]["minCa"].ToString() + "','" + dt_CastSeq.Rows[i]["minW"].ToString() + "','" + dt_CastSeq.Rows[i]["minsn"].ToString() + "','" + dt_CastSeq.Rows[i]["minAs"].ToString() + "','" + dt_CastSeq.Rows[i]["minTe"].ToString() + "','" + dt_CastSeq.Rows[i]["minBi"].ToString() + "','" + dt_CastSeq.Rows[i]["minB"].ToString() + "','" + dt_CastSeq.Rows[i]["minPb"].ToString() + "','" + dt_CastSeq.Rows[i]["minMg"].ToString() + "','" + dt_CastSeq.Rows[i]["minN"].ToString() + "','" + dt_CastSeq.Rows[i]["minVe"].ToString() + "','" + dt_CastSeq.Rows[i]["minCo"].ToString() + "','" + dt_CastSeq.Rows[i]["minCe"].ToString() + "','" + dt_CastSeq.Rows[i]["minSb"].ToString() + "','" + dt_CastSeq.Rows[i]["minZr"].ToString() + "','" + dt_CastSeq.Rows[i]["minO"].ToString() + "','" + dt_CastSeq.Rows[i]["minH"].ToString() + "','" + dt_CastSeq.Rows[i]["maxC"].ToString() + "','" + dt_CastSeq.Rows[i]["maxSi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxMn"].ToString() + "','" + dt_CastSeq.Rows[i]["maxP"].ToString() + "','" + dt_CastSeq.Rows[i]["maxS"].ToString() + "','" + dt_CastSeq.Rows[i]["maxAl"].ToString() + "','" + dt_CastSeq.Rows[i]["maxAlS"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCu"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCr"].ToString() + "','" + dt_CastSeq.Rows[i]["maxMo"].ToString() + "','" + dt_CastSeq.Rows[i]["maxNi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxV"].ToString() + "','" + dt_CastSeq.Rows[i]["maxTi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxNb"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCa"].ToString() + "','" + dt_CastSeq.Rows[i]["maxW"].ToString() + "','" + dt_CastSeq.Rows[i]["maxSn"].ToString() + "','" + dt_CastSeq.Rows[i]["maxAs"].ToString() + "','" + dt_CastSeq.Rows[i]["maxTe"].ToString() + "','" + dt_CastSeq.Rows[i]["maxBi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxB"].ToString() + "','" + dt_CastSeq.Rows[i]["maxPb"].ToString() + "','" + dt_CastSeq.Rows[i]["maxMg"].ToString() + "','" + dt_CastSeq.Rows[i]["maxN"].ToString() + "','" + dt_CastSeq.Rows[i]["maxVe"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCo"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCe"].ToString() + "','" + dt_CastSeq.Rows[i]["maxSb"].ToString() + "','" + dt_CastSeq.Rows[i]["maxZr"].ToString() + "','" + dt_CastSeq.Rows[i]["maxO"].ToString() + "','" + dt_CastSeq.Rows[i]["maxH"].ToString() + "','" + dt_CastSeq.Rows[i]["aimC"].ToString() + "','" + dt_CastSeq.Rows[i]["aimSi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimMn"].ToString() + "','" + dt_CastSeq.Rows[i]["aimP"].ToString() + "','" + dt_CastSeq.Rows[i]["aimS"].ToString() + "','" + dt_CastSeq.Rows[i]["aimAl"].ToString() + "','" + dt_CastSeq.Rows[i]["aimAlS"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCu"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCr"].ToString() + "','" + dt_CastSeq.Rows[i]["aimMo"].ToString() + "','" + dt_CastSeq.Rows[i]["aimNi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimV"].ToString() + "','" + dt_CastSeq.Rows[i]["aimTi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimNb"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCa"].ToString() + "','" + dt_CastSeq.Rows[i]["aimW"].ToString() + "','" + dt_CastSeq.Rows[i]["aimSn"].ToString() + "','" + dt_CastSeq.Rows[i]["aimAs"].ToString() + "','" + dt_CastSeq.Rows[i]["aimTe"].ToString() + "','" + dt_CastSeq.Rows[i]["aimBi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimB"].ToString() + "','" + dt_CastSeq.Rows[i]["aimPb"].ToString() + "','" + dt_CastSeq.Rows[i]["aimMg"].ToString() + "','" + dt_CastSeq.Rows[i]["aimN"].ToString() + "','" + dt_CastSeq.Rows[i]["aimVe"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCo"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCe"].ToString() + "','" + dt_CastSeq.Rows[i]["aimSb"].ToString() + "','" + dt_CastSeq.Rows[i]["aimZr"].ToString() + "','" + dt_CastSeq.Rows[i]["aimO"].ToString() + "','" + dt_CastSeq.Rows[i]["aimH"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatReportTempReport(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = "SELECT (pdc_tundish_temp.meas_temp - gcc_grade_spec.calc_liq_temp) AS overheat,pdc_tundish_temp.meas_temp as Temp,pdc_tundish_temp.meas_time as Time FROM gcc_grade_spec,pdc_heat,pdc_tundish_temp WHERE pdc_heat.grade_code = gcc_grade_spec.grade_code AND pdc_heat.steel_id     = pdc_tundish_temp.steel_id AND pdc_heat.steel_id     = '" + StrCastSeq_caster1_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatReportTempReport where SteelId='" + StrCastSeq_caster1_SteelID + "' and Time='" + (dt_CastSeq.Rows[i]["Time"].ToString().Replace(",", ".")).ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatReportTempReport set Temp='" + dt_CastSeq.Rows[i]["Temp"].ToString() + "',Temp='" + dt_CastSeq.Rows[i]["OverHeat"].ToString() + "' where SteelId='" + StrCastSeq_caster1_SteelID + "' and Time='" + (dt_CastSeq.Rows[i]["Time"].ToString().Replace(",", ".")).ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatReportTempReport(SeqNo,SteelId,Time,Temp,OverHeat)values('" + StrCastSeq_caster1_MaxID + "','" + StrCastSeq_caster1_SteelID + "','" + (dt_CastSeq.Rows[i]["Time"].ToString().Replace(",", ".")).ToString() + "','" + dt_CastSeq.Rows[i]["Temp"].ToString() + "','" + dt_CastSeq.Rows[i]["OverHeat"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatRpt_PlandataReport(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = "  SELECT pdc_slab.l2_slabid,nvl(pdc_slab_order.length_min,0) as length_min,nvl(pdc_slab_order.length_aim,0) as length_aim, nvl(pdc_slab_order.length_max,0) as length_max, nvl(pdc_slab_order.width_head,0) as width_head, nvl(pdc_slab_order.width_tail,0) as width_tail, nvl(pdc_piece.heat_steel_id,0) as heat_steel_id, nvl(pdc_slab.slab_seq_no,0) as slab_seq_no, nvl(pdc_piece.piece_seq_no,0) as piece_seq_no FROM l2ccs.pdc_piece pdc_piece,l2ccs.pdc_slab_order pdc_slab_order,l2ccs.pdc_slab pdc_slab  WHERE (pdc_slab_order.steel_id(+) = pdc_slab.steel_id) AND (pdc_piece.steel_id           = pdc_slab.steel_id(+)) AND (pdc_piece.heat_steel_id      = '" + StrCastSeq_caster1_SteelID + "')";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_PlandataReport where SteelId='" + StrCastSeq_caster1_SteelID + "' and L2_SlabId='" + dt_CastSeq.Rows[i]["l2_slabid"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_PlandataReport set L2_SlabId='" + dt_CastSeq.Rows[i]["L2_SlabId"].ToString() + "',Length_min='" + dt_CastSeq.Rows[i]["Length_min"].ToString() + "',Length_aim='" + dt_CastSeq.Rows[i]["Length_aim"].ToString() + "',Length_max='" + dt_CastSeq.Rows[i]["Length_max"].ToString() + "',Width_Head='" + dt_CastSeq.Rows[i]["Width_Head"].ToString() + "',Width_Tail='" + dt_CastSeq.Rows[i]["Width_Tail"].ToString() + "',Slab_Seq_No='" + dt_CastSeq.Rows[i]["Slab_Seq_No"].ToString() + "',Piece_Seq_No='" + dt_CastSeq.Rows[i]["Piece_Seq_No"].ToString() + "' where SteelId='" + StrCastSeq_caster1_SteelID + "' and SeqNo='" + StrCastSeq_caster1_MaxID + "'  and L2_SlabId='" + dt_CastSeq.Rows[i]["l2_slabid"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_PlandataReport(SeqNo,SteelId,L2_SlabId,Length_min,Length_aim,Length_max,Width_Head,Width_Tail,Slab_Seq_No,Piece_Seq_No)values('" + StrCastSeq_caster1_MaxID + "','" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["L2_SlabId"].ToString() + "','" + dt_CastSeq.Rows[i]["Length_min"].ToString() + "','" + dt_CastSeq.Rows[i]["Length_aim"].ToString() + "','" + dt_CastSeq.Rows[i]["Length_max"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_Head"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_Tail"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_Seq_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Piece_Seq_No"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatRptAnalysisReport(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = "SELECT a.heatid heatid,  a.sample_no,  a.taken_time,  a.sample_loc,  a.LIQUID_TEMP LIQ,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'C'  AND ae.analysis_id = a.analysis_id  ),0) c,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Si'  AND ae.analysis_id = a.analysis_id  ) ,0)si,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Mn'  AND ae.analysis_id = a.analysis_id  ),0) mn,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'P'  AND ae.analysis_id = a.analysis_id  ),0) p,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'S'  AND ae.analysis_id = a.analysis_id  ),0) s,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Al'  AND ae.analysis_id = a.analysis_id  ),0) al,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Als'  AND ae.analysis_id = a.analysis_id  ),0) als,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Cu'  AND ae.analysis_id = a.analysis_id  ),0) cu,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Cr'  AND ae.analysis_id = a.analysis_id  ),0) cr,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Mo'  AND ae.analysis_id = a.analysis_id  ),0) mo,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ni'  AND ae.analysis_id = a.analysis_id  ),0) ni,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'V'  AND ae.analysis_id = a.analysis_id  ),0) v,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ti'  AND ae.analysis_id = a.analysis_id  ),0) ti,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Nb'  AND ae.analysis_id = a.analysis_id  ),0) nb,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ca'  AND ae.analysis_id = a.analysis_id  ),0) ca,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'W'  AND ae.analysis_id = a.analysis_id  ),0) w,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Sn'  AND ae.analysis_id = a.analysis_id  ),0) sn,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'As'  AND ae.analysis_id = a.analysis_id  ),0) AS a_s,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Te'  AND ae.analysis_id = a.analysis_id  ) ,0)te,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Bi'  AND ae.analysis_id = a.analysis_id  ) ,0)bi,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'B'  AND ae.analysis_id = a.analysis_id  ),0) b,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Pb'  AND ae.analysis_id = a.analysis_id  ),0) Pb,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Mg'  AND ae.analysis_id = a.analysis_id  ),0) Mg,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'N'  AND ae.analysis_id = a.analysis_id  ),0) N,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ve'  AND ae.analysis_id = a.analysis_id  ),0) Ve,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Co'  AND ae.analysis_id = a.analysis_id  ),0) Co,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ce'  AND ae.analysis_id = a.analysis_id  ),0) Ce,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Sb'  AND ae.analysis_id = a.analysis_id  ),0) Sb,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Zr'  AND ae.analysis_id = a.analysis_id  ),0) Zr,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'O'  AND ae.analysis_id = a.analysis_id  ),0) O,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'H'  AND ae.analysis_id = a.analysis_id  ),0) H FROM pd_analysis a,  pdc_heat h WHERE (a.heatid = h.heatid) AND (h.steel_id ='" + StrCastSeq_caster1_SteelID + "')";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRptAnalysisReport where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "' and heatid='" + dt_CastSeq.Rows[i]["heatid"].ToString() + "' and Taken_Time='" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRptAnalysisReport set Sample_No='" + dt_CastSeq.Rows[i]["Sample_No"].ToString() + "',Taken_Time='" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "',Sample_LOC='" + dt_CastSeq.Rows[i]["Sample_LOC"].ToString() + "',LIQ='" + dt_CastSeq.Rows[i]["LIQ"].ToString() + "',C='" + dt_CastSeq.Rows[i]["C"].ToString() + "',Si='" + dt_CastSeq.Rows[i]["Si"].ToString() + "',Mn='" + dt_CastSeq.Rows[i]["Mn"].ToString() + "',P='" + dt_CastSeq.Rows[i]["P"].ToString() + "',S='" + dt_CastSeq.Rows[i]["S"].ToString() + "',Al='" + dt_CastSeq.Rows[i]["Al"].ToString() + "',AlS='" + dt_CastSeq.Rows[i]["AlS"].ToString() + "',Cu='" + dt_CastSeq.Rows[i]["Cu"].ToString() + "',Cr='" + dt_CastSeq.Rows[i]["Cr"].ToString() + "',Mo='" + dt_CastSeq.Rows[i]["Mo"].ToString() + "',Ni='" + dt_CastSeq.Rows[i]["Ni"].ToString() + "',V='" + dt_CastSeq.Rows[i]["V"].ToString() + "',Ti='" + dt_CastSeq.Rows[i]["Ti"].ToString() + "',NB='" + dt_CastSeq.Rows[i]["Nb"].ToString() + "',Ca='" + dt_CastSeq.Rows[i]["Ca"].ToString() + "',W='" + dt_CastSeq.Rows[i]["W"].ToString() + "',Sn='" + dt_CastSeq.Rows[i]["Sn"].ToString() + "',A_S='" + dt_CastSeq.Rows[i]["A_S"].ToString() + "',TE='" + dt_CastSeq.Rows[i]["TE"].ToString() + "',Bi='" + dt_CastSeq.Rows[i]["Bi"].ToString() + "',B='" + dt_CastSeq.Rows[i]["B"].ToString() + "',Pb='" + dt_CastSeq.Rows[i]["Pb"].ToString() + "',Mg='" + dt_CastSeq.Rows[i]["Mg"].ToString() + "',N='" + dt_CastSeq.Rows[i]["N"].ToString() + "',Ve='" + dt_CastSeq.Rows[i]["Ve"].ToString() + "',Co='" + dt_CastSeq.Rows[i]["Co"].ToString() + "',Ce='" + dt_CastSeq.Rows[i]["Ce"].ToString() + "',Sb='" + dt_CastSeq.Rows[i]["Sb"].ToString() + "',Zr='" + dt_CastSeq.Rows[i]["Zr"].ToString() + "',O='" + dt_CastSeq.Rows[i]["O"].ToString() + "',H='" + dt_CastSeq.Rows[i]["H"].ToString() + "' where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "' and heatid='" + dt_CastSeq.Rows[i]["heatid"].ToString() + "' and Taken_Time='" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRptAnalysisReport(SeqNo,SteelId,HeatId,Sample_No,Taken_Time,Sample_LOC,LIQ,C,Si,Mn,P,S,Al,AlS,Cu,Cr,Mo,Ni,V,Ti,NB,Ca,W,Sn,A_S,TE,Bi,B,Pb,Mg,N,Ve,Co,Ce,Sb,Zr,O,H)values('" + StrCastSeq_caster1_MaxID + "','" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "','" + dt_CastSeq.Rows[i]["Sample_No"].ToString() + "','" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Sample_LOC"].ToString() + "','" + dt_CastSeq.Rows[i]["LIQ"].ToString() + "','" + dt_CastSeq.Rows[i]["C"].ToString() + "','" + dt_CastSeq.Rows[i]["Si"].ToString() + "','" + dt_CastSeq.Rows[i]["Mn"].ToString() + "','" + dt_CastSeq.Rows[i]["P"].ToString() + "','" + dt_CastSeq.Rows[i]["S"].ToString() + "','" + dt_CastSeq.Rows[i]["Al"].ToString() + "','" + dt_CastSeq.Rows[i]["AlS"].ToString() + "','" + dt_CastSeq.Rows[i]["Cu"].ToString() + "','" + dt_CastSeq.Rows[i]["Cr"].ToString() + "','" + dt_CastSeq.Rows[i]["Mo"].ToString() + "','" + dt_CastSeq.Rows[i]["Ni"].ToString() + "','" + dt_CastSeq.Rows[i]["V"].ToString() + "','" + dt_CastSeq.Rows[i]["Ti"].ToString() + "','" + dt_CastSeq.Rows[i]["NB"].ToString() + "','" + dt_CastSeq.Rows[i]["Ca"].ToString() + "','" + dt_CastSeq.Rows[i]["W"].ToString() + "','" + dt_CastSeq.Rows[i]["Sn"].ToString() + "','" + dt_CastSeq.Rows[i]["A_S"].ToString() + "','" + dt_CastSeq.Rows[i]["TE"].ToString() + "','" + dt_CastSeq.Rows[i]["Bi"].ToString() + "','" + dt_CastSeq.Rows[i]["B"].ToString() + "','" + dt_CastSeq.Rows[i]["Pb"].ToString() + "','" + dt_CastSeq.Rows[i]["Mg"].ToString() + "','" + dt_CastSeq.Rows[i]["N"].ToString() + "','" + dt_CastSeq.Rows[i]["Ve"].ToString() + "','" + dt_CastSeq.Rows[i]["Co"].ToString() + "','" + dt_CastSeq.Rows[i]["Ce"].ToString() + "','" + dt_CastSeq.Rows[i]["Sb"].ToString() + "','" + dt_CastSeq.Rows[i]["Zr"].ToString() + "','" + dt_CastSeq.Rows[i]["O"].ToString() + "','" + dt_CastSeq.Rows[i]["H"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatRptCutDataReport(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = " SELECT p.strand_no,sl.slab_seq_no   AS slab_no,sl.l2_slabid  AS marking_no,pne.value/1000 AS cold_length,pt.piece_type_desc   AS piece_type,p.head_cast_len   AS cast_length_cut,so.length_min   AS min_length,so.length_aim   AS aim_length,so.length_max  AS max_length,(p.piece_length/1000 - NVL (scrap.scrap_length, 0)) AS good_length,p.piece_length/1000  AS act_length,(DECODE(NVL(piece_weight_meas,0), 0, piece_weight_calc , piece_weight_meas) - NVL (scrap.scrap_mass, 0)) / 1000  AS good_weight,NVL (scrap.scrap_mass, 0) / 1000  AS scrap_weight,p.head_thickness,p.head_width,p.mixzone_begin AS mixzone_begin,p.mixzone_end   AS mixzone_end,p.heat_boundary_pos1, p.sample_weight FROM pdc_piece p,pdc_slab sl,pdc_slab_order so,pdc_heat h,gdc_piece_type pt,pdc_number_entry pne,(SELECT steel_id,SUM (scrap_mass) AS scrap_mass,(SUM (scrap_end) - SUM (scrap_bgn)) AS scrap_length FROM pdc_scrap  GROUP BY steel_id  ) scrap WHERE (p.steel_id    = sl.steel_id(+)) AND (p.steel_id      = so.steel_id(+)) AND (p.steel_id      = scrap.steel_id(+)) AND (p.piece_type    = pt.piece_type(+)) AND (p.heat_steel_id = h.steel_id) AND pne.val_name     = 'ColdLen' AND pne.steel_id     = p.steel_id AND (h.steel_id = '" + StrCastSeq_caster1_SteelID + "') ORDER BY p.head_cast_len ASC";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRptCutDataReport where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "' and Marking_No='" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRptCutDataReport set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Slab_No='" + dt_CastSeq.Rows[i]["Slab_No"].ToString() + "',Marking_No='" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "',Cold_Length='" + dt_CastSeq.Rows[i]["Cold_Length"].ToString() + "',Piece_Type='" + dt_CastSeq.Rows[i]["Piece_Type"].ToString() + "',Cast_Length_Cut='" + dt_CastSeq.Rows[i]["Cast_Length_Cut"].ToString() + "',Min_Length='" + dt_CastSeq.Rows[i]["Min_Length"].ToString() + "',Aim_Length='" + dt_CastSeq.Rows[i]["Aim_Length"].ToString() + "',Max_Length='" + dt_CastSeq.Rows[i]["Max_Length"].ToString() + "',Good_Length='" + dt_CastSeq.Rows[i]["Good_Length"].ToString() + "',Act_Length='" + dt_CastSeq.Rows[i]["Act_Length"].ToString() + "',Good_Weight='" + dt_CastSeq.Rows[i]["Good_Weight"].ToString() + "',scrap_Weight='" + dt_CastSeq.Rows[i]["scrap_Weight"].ToString() + "',Head_Thickness='" + dt_CastSeq.Rows[i]["Head_Thickness"].ToString() + "',Head_Width='" + dt_CastSeq.Rows[i]["Head_Width"].ToString() + "',MixZone_Begin='" + dt_CastSeq.Rows[i]["MixZone_Begin"].ToString() + "',MixZone_End='" + dt_CastSeq.Rows[i]["MixZone_End"].ToString() + "',Heat_Boundary_POS1='" + dt_CastSeq.Rows[i]["Heat_Boundary_POS1"].ToString() + "',Sample_Weight='" + dt_CastSeq.Rows[i]["sample_weight"].ToString() + "' where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "' and Marking_No='" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRptCutDataReport(SeqNo,SteelId,Strand_No,Slab_No,Marking_No,Cold_Length,Piece_Type,Cast_Length_Cut,Min_Length,Aim_Length,Max_Length,Good_Length,Act_Length,Good_Weight,scrap_Weight,Head_Thickness,Head_Width,MixZone_Begin,MixZone_End,Heat_Boundary_POS1,Sample_Weight)values('" + StrCastSeq_caster1_MaxID + "','" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Cold_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Piece_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Length_Cut"].ToString() + "','" + dt_CastSeq.Rows[i]["Min_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Aim_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Max_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Good_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Act_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Good_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["scrap_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Head_Thickness"].ToString() + "','" + dt_CastSeq.Rows[i]["Head_Width"].ToString() + "','" + dt_CastSeq.Rows[i]["MixZone_Begin"].ToString() + "','" + dt_CastSeq.Rows[i]["MixZone_End"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Boundary_POS1"].ToString() + "','" + dt_CastSeq.Rows[i]["sample_weight"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatRptHeader(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = "SELECT  ROUND(nvl(v.slab_weight_net,0),3) AS slab_weight,  nvl(heat.CUT_LOST_WEIGHT,0)   AS cut_lost_weight,  nvl(v.count_slabs,0)          AS slab_count, nvl(heat.avg_slab_width,0) as avg_slab_width, nvl(heat.heatid,0) as heatid,  nvl(heat.heat_seq_no,0) as heat_seq_no,  nvl(heat.grade_code,0) as grade_code, cheat.shift_foreman  AS shift_foreman,  cheat.casting_foreman AS casting_foreman,  heat.prod_orderid, nvl(heat.ladle_no,0) as ladle_no,  nvl(v.ladle_open_weight,0)  AS ladle_open_weight,  nvl(v.ladle_close_weight,0) AS ladle_close_weight,  nvl(v.total_cast_weight,0)  AS cast_weight,  nvl(seq.seq_no,0) as seq_no,  ROUND(nvl(v.yield,2),2) AS yield,  hshift.crew_id   AS crewid,   nvl(heat.treat_counter,0)  AS treat_counter,   nvl(heat.slab_length,0)  AS slab_length,   heat.ladle_close_time  AS ladle_close_time,   nvl(heat.pouring_duration,0)  AS pouring_duration,   heat.ladle_open_time AS ladle_open_time,   nvl(heat.ladle_open_time,0) AS prod_date,   heat.ladle_arrival_time AS ladle_arrival_time,   nvl(heat.tund_powder_type,0)  AS tund_powder_type,   nvl(heat.tund_powder_weight,0) AS tund_powder_weight,  tund_open.tund_open_time AS tund_open_time,   tund_close.tund_close_time,     -1     AS tund_life, nvl(tund_open.tund_open_weight    / 1000 ,0)AS tund_open_weight, nvl( tund_close.tund_close_weight  /1000 ,0) AS tund_close_weight, nvl(DECODE (NVL(tund_open.tund_no,-1), -1, tund_close.tund_no, tund_open.tund_no,0),0) tund_no, nvl(DECODE (NVL(tund_open.car_no, -1), -1, tund_close.car_no, tund_open.car_no,0),0) car_no,  nvl(v.sample_weight,0) AS sample_lost_weight,  nvl(v.scale_loss,0)    AS scale_lost_weight, nvl( heat.ladle_tare_weight,0) as ladle_tare_weight,  nvl((std.cast_len_end     -std.cast_len_bgn) ,0)                           AS total_cast_length,  ROUND(nvl((v.scrap_weight + v.head_crop_weight + v.tail_crop_weight),0),3) AS total_scrap_weight,  param.value                                                         AS casterid FROM pdc_heat heat,  pdc_strand std,  pdc_heat_strand hs,  pdc_heat_shift hshift,  pdc_customer_heat cheat,  pdc_sequence seq,  v_heat_tund_open tund_open,  v_heat_tund_close tund_close,  v_heat_weight v,  gdc_param param WHERE heat.sequence_steel_id    = seq.steel_id AND hs.strand_steel_id          = std.steel_id AND hs.heat_steel_id            = heat.steel_id AND hshift.heat_steel_id(+)     = heat.steel_id AND tund_open.heat_steel_id(+)  = heat.steel_id AND tund_close.heat_steel_id(+) = heat.steel_id AND cheat.steel_id(+)           = heat.steel_id AND v.steel_id(+)               = heat.steel_id AND param.par_context           = 'CASTER' AND param.par_name              = 'CASTERID' AND heat.steel_id= '" + StrCastSeq_caster1_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRptHeader where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "' and HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRptHeader set Slab_Weight ='" + dt_CastSeq.Rows[i]["Slab_Weight"].ToString() + "',Cut_Lost_Weight ='" + dt_CastSeq.Rows[i]["Cut_Lost_Weight"].ToString() + "',Slab_Count='" + dt_CastSeq.Rows[i]["Slab_Count"].ToString() + "',Avg_Slab_Width='" + dt_CastSeq.Rows[i]["Avg_Slab_Width"].ToString() + "',HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "',Heat_Seq_No='" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "',Grade_Code='" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "',Shift_Foreman='" + dt_CastSeq.Rows[i]["Shift_Foreman"].ToString() + "',Casting_Foreman='" + dt_CastSeq.Rows[i]["Casting_Foreman"].ToString() + "',Prod_OrderId='" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "',Ladle_No='" + dt_CastSeq.Rows[i]["Ladle_No"].ToString() + "',Ladle_Open_Weight='" + dt_CastSeq.Rows[i]["Ladle_Open_Weight"].ToString() + "',Ladle_Close_Weight='" + dt_CastSeq.Rows[i]["Ladle_Close_Weight"].ToString() + "',Cast_Weight='" + dt_CastSeq.Rows[i]["Cast_Weight"].ToString() + "',Yield='" + dt_CastSeq.Rows[i]["Yield"].ToString() + "',CrewId='" + dt_CastSeq.Rows[i]["CrewId"].ToString() + "',Treat_Counter='" + dt_CastSeq.Rows[i]["Treat_Counter"].ToString() + "',Slab_Length='" + dt_CastSeq.Rows[i]["Slab_Length"].ToString() + "',Ladle_Close_Time='" + (dt_CastSeq.Rows[i]["Ladle_Close_Time"].ToString()).Replace(",", ".").ToString() + "',Pouring_Duration='" + dt_CastSeq.Rows[i]["Pouring_Duration"].ToString() + "',Ladle_Open_Time='" + (dt_CastSeq.Rows[i]["Ladle_Open_Time"].ToString()).Replace(",", ".").ToString() + "',Prod_date='" + (dt_CastSeq.Rows[i]["Prod_date"].ToString()).Replace(",", ".").ToString() + "',Ladle_Arrival_Time='" + (dt_CastSeq.Rows[i]["Ladle_Arrival_Time"].ToString()).Replace(",", ".").ToString() + "',Tund_Powder_Type='" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "',Tund_Powder_weight='" + dt_CastSeq.Rows[i]["Tund_Powder_weight"].ToString() + "',Tund_Open_Time='" + (dt_CastSeq.Rows[i]["Tund_Open_Time"].ToString()).Replace(",", ".").ToString() + "',Tund_Close_Time='" + (dt_CastSeq.Rows[i]["Tund_Close_Time"].ToString()).Replace(",", ".").ToString() + "',Tund_Life='" + dt_CastSeq.Rows[i]["Tund_Life"].ToString() + "',Tund_Open_Weight='" + dt_CastSeq.Rows[i]["Tund_Open_Weight"].ToString() + "',Tund_Close_Weight='" + dt_CastSeq.Rows[i]["Tund_Close_Weight"].ToString() + "',Tund_No='" + dt_CastSeq.Rows[i]["Tund_No"].ToString() + "',Car_No='" + dt_CastSeq.Rows[i]["Car_No"].ToString() + "',Sample_Lost_Weight='" + dt_CastSeq.Rows[i]["Sample_Lost_Weight"].ToString() + "',Scale_Lost_Weight='" + dt_CastSeq.Rows[i]["Scale_Lost_Weight"].ToString() + "',Ladle_Tare_Weight='" + dt_CastSeq.Rows[i]["Ladle_Tare_Weight"].ToString() + "',Total_Cast_Length='" + dt_CastSeq.Rows[i]["Total_Cast_Length"].ToString() + "',Total_Scrap_Weight='" + dt_CastSeq.Rows[i]["Total_Scrap_Weight"].ToString() + "',CasterId='" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "' where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "' and HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRptHeader(SeqNo,SteelId,Slab_Weight,Cut_Lost_Weight,Slab_Count,Avg_Slab_Width,HeatId,Heat_Seq_No,Grade_Code,Shift_Foreman,Casting_Foreman,Prod_OrderId,Ladle_No,Ladle_Open_Weight,Ladle_Close_Weight,Cast_Weight,Yield,CrewId,Treat_Counter,Slab_Length,Ladle_Close_Time,Pouring_Duration,Ladle_Open_Time,Prod_date,Ladle_Arrival_Time,Tund_Powder_Type,Tund_Powder_weight,Tund_Open_Time,Tund_Close_Time,Tund_Life,Tund_Open_Weight,Tund_Close_Weight,Tund_No,Car_No,Sample_Lost_Weight,Scale_Lost_Weight,Ladle_Tare_Weight,Total_Cast_Length,Total_Scrap_Weight,CasterId)values('" + StrCastSeq_caster1_MaxID + "','" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["slab_weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Cut_Lost_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_Count"].ToString() + "','" + dt_CastSeq.Rows[i]["Avg_Slab_Width"].ToString() + "','" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "','" + dt_CastSeq.Rows[i]["Shift_Foreman"].ToString() + "','" + dt_CastSeq.Rows[i]["Casting_Foreman"].ToString() + "','" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Open_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Close_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Yield"].ToString() + "','" + dt_CastSeq.Rows[i]["CrewId"].ToString() + "','" + dt_CastSeq.Rows[i]["Treat_Counter"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_Length"].ToString() + "','" + (dt_CastSeq.Rows[i]["Ladle_Close_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Pouring_Duration"].ToString() + "','" + (dt_CastSeq.Rows[i]["Ladle_Open_Time"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Prod_date"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Ladle_Arrival_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Powder_weight"].ToString() + "','" + (dt_CastSeq.Rows[i]["Tund_Open_Time"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Tund_Close_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Life"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Open_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Close_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Car_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Sample_Lost_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Scale_Lost_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Tare_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Total_Cast_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Total_Scrap_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatRpt_HeatReportSR(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = "SELECT s.strand_num AS strand_no,  s.avg_cast_speed,  s.temp_speed_tab_num AS copt_temp_tab_num,  s.s_speed_tab_num    AS copt_s_tab_num,  s.smn_speed_tab_num  AS copt_mns_tab_num,  s.tw_speed_tab_num   AS copt_tw_tab_num,  s.hmo_tab_num,  s.softred_tab_num AS hsa_tab_num,  s.gen_tab_num,  s.dsc_tab_num,  s.bops_tab_num AS mms_tab_num,  s.moldadj_tab_num,  s.sgrade_speed_tab_num AS copt_sgrade_tab_num,  s.mlc_tab_num,  s.cast_len_bgn       AS cast_len_begin,  s.cast_len_end       AS cast_len_end,  s.spray_plan_num     AS swplan_tab_num,  clo_tab_num          AS clo_tab_num,  s.clo_sample_tab_num AS clo_smp_tab_num,s.cast_powder_type as Mould_Pow_Type,s.cast_powder_weight as Mould_Pow_Amnt  FROM pdc_strand s,  pdc_heat_strand WHERE pdc_heat_strand.strand_steel_id = s.steel_id AND pdc_heat_strand.heat_steel_id='" + StrCastSeq_caster1_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatReportSR where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatReportSR set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Avg_Cast_Speed='" + dt_CastSeq.Rows[i]["Avg_Cast_Speed"].ToString() + "',Copt_Temp_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_Temp_Tab_Num"].ToString() + "',Copt_S_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_S_Tab_Num"].ToString() + "',Copt_MNS_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_MNS_Tab_Num"].ToString() + "',Copt_TW_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_TW_Tab_Num"].ToString() + "',HMO_Tab_Num='" + dt_CastSeq.Rows[i]["HMO_Tab_Num"].ToString() + "',HSA_Tab_Num='" + dt_CastSeq.Rows[i]["HSA_Tab_Num"].ToString() + "',Gen_Tab_Num='" + dt_CastSeq.Rows[i]["Gen_Tab_Num"].ToString() + "',DSC_Tab_Num='" + dt_CastSeq.Rows[i]["DSC_Tab_Num"].ToString() + "',MMS_Tab_Num='" + dt_CastSeq.Rows[i]["MMS_Tab_Num"].ToString() + "',MOLDADJ_Tab_Num='" + dt_CastSeq.Rows[i]["MOLDADJ_Tab_Num"].ToString() + "',Copt_SGrade_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_SGrade_Tab_Num"].ToString() + "',MLC_Tab_Num='" + dt_CastSeq.Rows[i]["MLC_Tab_Num"].ToString() + "',Cast_Len_Begin='" + dt_CastSeq.Rows[i]["Cast_Len_Begin"].ToString() + "',Cast_Len_End='" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "',SWPlan_Tab_Num='" + dt_CastSeq.Rows[i]["SWPlan_Tab_Num"].ToString() + "',CLO_Tab_Num='" + dt_CastSeq.Rows[i]["CLO_Tab_Num"].ToString() + "',CLO_Smp_Tab_Num='" + dt_CastSeq.Rows[i]["CLO_Smp_Tab_Num"].ToString() + "',Mould_Pow_Type='" + dt_CastSeq.Rows[i]["Mould_Pow_Type"].ToString() + "',Mould_Pow_Amnt='" + dt_CastSeq.Rows[i]["Mould_Pow_Amnt"].ToString() + "' where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatReportSR(SeqNo,SteelId,Strand_No,Avg_Cast_Speed,Copt_Temp_Tab_Num,Copt_S_Tab_Num,Copt_MNS_Tab_Num,Copt_TW_Tab_Num,HMO_Tab_Num,HSA_Tab_Num,Gen_Tab_Num,DSC_Tab_Num,MMS_Tab_Num,MOLDADJ_Tab_Num,Copt_SGrade_Tab_Num,MLC_Tab_Num,Cast_Len_Begin,Cast_Len_End,SWPlan_Tab_Num,CLO_Tab_Num,CLO_Smp_Tab_Num,Mould_Pow_Type,Mould_Pow_Amnt)values('" + StrCastSeq_caster1_MaxID + "','" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Avg_Cast_Speed"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_Temp_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_S_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_MNS_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_TW_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["HMO_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["HSA_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Gen_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["DSC_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["MMS_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["MOLDADJ_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_SGrade_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["MLC_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_Begin"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "','" + dt_CastSeq.Rows[i]["SWPlan_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["CLO_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["CLO_Smp_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Mould_Pow_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Mould_Pow_Amnt"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatRpt_SubMoldStrand(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = "SELECT s.strand_num    AS strand_no,  s.cast_len_bgn       AS cast_start_len,  s.cast_len_end       AS cast_end_len,  s.heat_in_mold_time  AS cast_start_time,  s.heat_out_mold_time AS cast_end_time, nvl(s.mold_no,0) as mold_no FROM pdc_strand s,  pdc_heat_strand hs WHERE s.steel_id     = hs.strand_steel_id AND hs.heat_steel_id='" + StrCastSeq_caster1_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_SubMoldStrand where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_SubMoldStrand set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Cast_Start_Len='" + dt_CastSeq.Rows[i]["Cast_Start_Len"].ToString() + "',Cast_End_Len='" + dt_CastSeq.Rows[i]["Cast_End_Len"].ToString() + "',Cast_Start_Time='" + (dt_CastSeq.Rows[i]["Cast_Start_Time"].ToString()).Replace(",", ".").ToString() + "',Cast_End_Time='" + (dt_CastSeq.Rows[i]["Cast_End_Time"].ToString()).Replace(",", ".").ToString() + "',Mold_No='" + dt_CastSeq.Rows[i]["Mold_No"].ToString() + "' where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_SubMoldStrand(SeqNo,SteelId,Strand_No,Cast_Start_Len,Cast_End_Len,Cast_Start_Time,Cast_End_Time,Mold_No)values('" + StrCastSeq_caster1_MaxID + "','" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Start_Len"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_End_Len"].ToString() + "','" + (dt_CastSeq.Rows[i]["Cast_Start_Time"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Cast_End_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Mold_No"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatRpt_SubPracticeTable(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = "SELECT s.strand_num  AS strand_no,  s.gen_tab_num  AS gen,  s.moldadj_tab_num   AS mold,  s.hmo_tab_num          AS hmo,  s.softred_tab_num      AS hsa,  s.mlc_tab_num          AS mlc,  s.spray_plan_num       AS cooling,  s.bops_tab_num         AS mms,  s.sgrade_speed_tab_num AS copt_speed,  s.temp_speed_tab_num   AS copt_temp,  s.tw_speed_tab_num     AS copt_tw,  s.s_speed_tab_num      AS copt_s,  s.smn_speed_tab_num    AS copt_mns,  s.clo_sample_tab_num   AS clo_smp,  clo_tab_num            AS clo,  s.dsc_tab_num          AS dsc,  0                      AS qual FROM pdc_strand s,  pdc_heat_strand WHERE pdc_heat_strand.strand_steel_id = s.steel_id AND pdc_heat_strand.heat_steel_id='" + StrCastSeq_caster1_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_SubPracticeTable where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_SubPracticeTable set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Gen='" + dt_CastSeq.Rows[i]["Gen"].ToString() + "',Mold='" + dt_CastSeq.Rows[i]["Mold"].ToString() + "',HMO='" + dt_CastSeq.Rows[i]["HMO"].ToString() + "',HSA='" + dt_CastSeq.Rows[i]["HSA"].ToString() + "',MLC='" + dt_CastSeq.Rows[i]["MLC"].ToString() + "',Cooling='" + dt_CastSeq.Rows[i]["Cooling"].ToString() + "',MMS='" + dt_CastSeq.Rows[i]["MMS"].ToString() + "',Copt_Speed='" + dt_CastSeq.Rows[i]["Copt_Speed"].ToString() + "',Copt_Temp='" + dt_CastSeq.Rows[i]["Copt_Temp"].ToString() + "',Copt_TW='" + dt_CastSeq.Rows[i]["Copt_TW"].ToString() + "',Copt_S='" + dt_CastSeq.Rows[i]["Copt_S"].ToString() + "',Copt_MNS='" + dt_CastSeq.Rows[i]["Copt_MNS"].ToString() + "',Clo_SMP='" + dt_CastSeq.Rows[i]["Clo_SMP"].ToString() + "',CLO='" + dt_CastSeq.Rows[i]["CLO"].ToString() + "',DSC='" + dt_CastSeq.Rows[i]["DSC"].ToString() + "',QUAL='" + dt_CastSeq.Rows[i]["QUAL"].ToString() + "' where SeqNo='" + StrCastSeq_caster1_MaxID + "' and SteelId='" + StrCastSeq_caster1_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_SubPracticeTable(SeqNo,SteelId,Strand_No,Gen,Mold,HMO,HSA,MLC,Cooling,MMS,Copt_Speed,Copt_Temp,Copt_TW,Copt_S,Copt_MNS,Clo_SMP,CLO,DSC,QUAL)values('" + StrCastSeq_caster1_MaxID + "','" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Gen"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold"].ToString() + "','" + dt_CastSeq.Rows[i]["HMO"].ToString() + "','" + dt_CastSeq.Rows[i]["HSA"].ToString() + "','" + dt_CastSeq.Rows[i]["MLC"].ToString() + "','" + dt_CastSeq.Rows[i]["Cooling"].ToString() + "','" + dt_CastSeq.Rows[i]["MMS"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_Speed"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_Temp"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_TW"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_S"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_MNS"].ToString() + "','" + dt_CastSeq.Rows[i]["Clo_SMP"].ToString() + "','" + dt_CastSeq.Rows[i]["CLO"].ToString() + "','" + dt_CastSeq.Rows[i]["DSC"].ToString() + "','" + dt_CastSeq.Rows[i]["QUAL"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatRpt_AnalysisElement(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            string StrCastSeqMaxID = "SELECT gcc_element_cat.element_name AS chem_abbr,  gcc_element_cat.display_order     AS display_order,  gcc_element_cat.element_type      AS element_type FROM gcc_element_cat ORDER BY display_order";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select * from HeatRpt_AnalysisElement where DISPLAY_ORDER='" + dt_CastSeq.Rows[i]["DISPLAY_ORDER"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_AnalysisElement set CHEM_ABBR='" + dt_CastSeq.Rows[i]["CHEM_ABBR"].ToString() + "',ELEMENT_TYPE='" + dt_CastSeq.Rows[i]["ELEMENT_TYPE"].ToString() + "' where DISPLAY_ORDER='" + dt_CastSeq.Rows[i]["DISPLAY_ORDER"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_AnalysisElement(CHEM_ABBR,DISPLAY_ORDER,ELEMENT_TYPE)values('" + dt_CastSeq.Rows[i]["CHEM_ABBR"].ToString() + "','" + dt_CastSeq.Rows[i]["DISPLAY_ORDER"].ToString() + "','" + dt_CastSeq.Rows[i]["ELEMENT_TYPE"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster1_HeatRpt_HeatRpt_HeatEquipmentData(string StrCastSeq_caster1_MaxID, string StrCastSeq_caster1_SteelID)
        {
            //Ladle_Shroud_Id,Ladle_Shroud_Life,Ladle_Shroud_Count,Nozzle_Id,Nozzle_Life,Nozzle_Count,Tundish_Id,Tundish_Life,Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxID = "SELECT MAX(eq.customer_id)       AS ladle_shroud_id,  TRUNC(MAX(co.counter_value),2) AS ladle_shroud_life,  COUNT(eq.customer_id)-1        AS ladle_shroud_count FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster1_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'SHROUD' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'HEAT'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Ladle_Shroud_Id='" + dt_CastSeq.Rows[i]["Ladle_Shroud_Id"].ToString() + "',Ladle_Shroud_Life='" + dt_CastSeq.Rows[i]["Ladle_Shroud_Life"].ToString() + "',Ladle_Shroud_Count='" + dt_CastSeq.Rows[i]["Ladle_Shroud_Count"].ToString() + "' where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Ladle_Shroud_Id,Ladle_Shroud_Life,Ladle_Shroud_Count)values('" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["Ladle_Shroud_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Shroud_Life"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Shroud_Count"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }

            //Nozzle_Id,Nozzle_Life,Nozzle_Count,Tundish_Id,Tundish_Life,Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDNozzle = "SELECT MAX(eq.customer_id)       AS nozzle_id,  TRUNC(MAX(co.counter_value),2) AS nozzle_life,  COUNT(eq.customer_id)-1        AS nozzle_count FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster1_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'NOZZLE' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'HEAT'";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxIDNozzle);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Nozzle_Id='" + dt_CastSeq.Rows[i]["Nozzle_Id"].ToString() + "',Nozzle_Life='" + dt_CastSeq.Rows[i]["Nozzle_Life"].ToString() + "',Nozzle_Count='" + dt_CastSeq.Rows[i]["Nozzle_Count"].ToString() + "' where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Nozzle_Id,Nozzle_Life,Nozzle_Count)values('" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["Nozzle_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Nozzle_Life"].ToString() + "','" + dt_CastSeq.Rows[i]["Nozzle_Count"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
            //Tundish_Id,Tundish_Life,Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDTundish = "SELECT eq.customer_id       AS tundish_id,  TRUNC(co.counter_value,2) AS tundish_life FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster1_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'TUNDISH' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'TON' ";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxIDTundish);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set tundish_id='" + dt_CastSeq.Rows[i]["tundish_id"].ToString() + "',Tundish_Life='" + dt_CastSeq.Rows[i]["Tundish_Life"].ToString() + "' where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Tundish_Id,Tundish_Life)values('" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["Tundish_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Tundish_Life"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
            //Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDMold = "SELECT eq.customer_id       AS mold_id,  TRUNC(co.counter_value,2) AS mold_life FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster1_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'MOLD' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'HEAT' ";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxIDMold);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Mold_Id='" + dt_CastSeq.Rows[i]["Mold_Id"].ToString() + "',Mold_Life='" + dt_CastSeq.Rows[i]["Mold_Life"].ToString() + "' where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Mold_Id,Mold_Life)values('" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["Mold_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold_Life"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }

            //Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDStnd = "SELECT strand_no,  cast_length,  event_type,  width_toc,  width_boc,  taper_left,  taper_right FROM pdc_width_change WHERE heat_steel_id = '" + StrCastSeq_caster1_SteelID + "' AND event_type     IN (1,3) ";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERI_Table(StrCastSeqMaxIDStnd);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Cast_Length='" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "',Event_Type='" + dt_CastSeq.Rows[i]["Event_Type"].ToString() + "',Width_TOC='" + dt_CastSeq.Rows[i]["Width_TOC"].ToString() + "',Width_BOC='" + dt_CastSeq.Rows[i]["Width_BOC"].ToString() + "',Taper_Left='" + dt_CastSeq.Rows[i]["Taper_Left"].ToString() + "',Taper_Right='" + dt_CastSeq.Rows[i]["Taper_Right"].ToString() + "' where SteelId='" + StrCastSeq_caster1_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right)values('" + StrCastSeq_caster1_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Event_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_TOC"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_BOC"].ToString() + "','" + dt_CastSeq.Rows[i]["Taper_Left"].ToString() + "','" + dt_CastSeq.Rows[i]["Taper_Right"].ToString() + "')";


                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }

        private void Fn_Caster2_HeatRpt_HeatReportHead(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = "SELECT STEEL_ID AS STEELID,SEQ_NO  AS SeqNo,HEATID  AS HEATID,HEAT_SEQ_NO AS InSeq,SEQUENCE_STEEL_ID AS Heat_Seq_No,SEQUENCE_STEEL_ID AS SEQ_STEEL_ID,PROD_ORDERID      AS PLAN_NO,GRADE_CODE        AS STEEL_GRADE,LADLE_OPEN_TIME   AS CAST_START_TIME,HEAT_STATUS_DESC  AS STATUS,HEAT_STATUS_CODE FROM V_HEAT_DATA WHERE (1=1)AND SEQ_NO='" + StrCastSeq_caster2_MaxID + "' AND HEAT_STATUS_CODE>2 ORDER BY LADLE_OPEN_TIME DESC";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSEQUENCENO = "";
                    string StrCASTSTART = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSEQUENCENO = dt_CastSeq.Rows[i]["SeqNo"].ToString();
                    if (dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != null && dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != string.Empty)
                    {
                        StrCASTSTART = Convert.ToDateTime(dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SeqNo from HeatReportHead where SeqNo='" + StrSEQUENCENO + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatReportHead set HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "',InSeq='" + dt_CastSeq.Rows[i]["InSeq"].ToString() + "',Heat_Seq_No='" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "',Seq_Steel_Id='" + dt_CastSeq.Rows[i]["Seq_Steel_Id"].ToString() + "',Plan_No='" + dt_CastSeq.Rows[i]["Plan_No"].ToString() + "',Steel_Grade='" + dt_CastSeq.Rows[i]["Steel_Grade"].ToString() + "',Cast_Strat_Time='" + StrCASTSTART + "',Status='" + dt_CastSeq.Rows[i]["Status"].ToString() + "',Heat_Status_Code='" + dt_CastSeq.Rows[i]["Heat_Status_Code"].ToString() + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + dt_CastSeq.Rows[i]["SeqNo"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatReportHead(SteelId,SeqNo,HeatId,InSeq,Heat_Seq_No,Seq_Steel_Id,Plan_No,Steel_Grade,Cast_Strat_Time,Status,Heat_Status_Code)values('" + StrSTEELID + "','" + StrSEQUENCENO + "','" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "','" + dt_CastSeq.Rows[i]["InSeq"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Seq_Steel_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Plan_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Steel_Grade"].ToString() + "','" + StrCASTSTART + "','" + dt_CastSeq.Rows[i]["Status"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Status_Code"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatReportLabAnalysis(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = "SELECT se.grade_code,gs.grade_code,gs.CALC_LIQ_TEMP aimliq,nvl(SUM (DECODE (se.element_name, 'C', min_conc)),0) minc,  nvl(SUM (DECODE (se.element_name, 'Si', min_conc)),0) minsi, nvl(SUM (DECODE (se.element_name, 'Mn', min_conc)),0) minmn, nvl(SUM (DECODE (se.element_name, 'P', min_conc)),0) minp, nvl(SUM (DECODE (se.element_name, 'S', min_conc)),0) mins, nvl(SUM (DECODE (se.element_name, 'Al', min_conc)),0) minal, nvl(SUM (DECODE (se.element_name, 'Al_S', min_conc)),0) minals, nvl(SUM (DECODE (se.element_name, 'Cu', min_conc)),0) mincu, nvl(SUM (DECODE (se.element_name, 'Cr', min_conc)),0) mincr, nvl(SUM (DECODE (se.element_name, 'Mo', min_conc)),0) minmo, nvl(SUM (DECODE (se.element_name, 'Ni', min_conc)),0) minni, nvl(SUM (DECODE (se.element_name, 'V', min_conc)),0) minv, nvl(SUM (DECODE (se.element_name, 'Ti', min_conc)),0) minti, nvl(SUM (DECODE (se.element_name, 'Nb', min_conc)),0) minnb, nvl(SUM (DECODE (se.element_name, 'Ca', min_conc)),0) minca,nvl(SUM (DECODE (se.element_name, 'W', min_conc)),0) minw, nvl(SUM (DECODE (se.element_name, 'Sn', min_conc)),0) minsn, nvl(SUM (DECODE (se.element_name, 'As', min_conc)),0) minas, nvl(SUM (DECODE (se.element_name, 'Te', min_conc)),0) minte, nvl(SUM (DECODE (se.element_name, 'Bi', min_conc)),0) minbi, nvl(SUM (DECODE (se.element_name, 'B', min_conc)),0) minb, nvl(SUM (DECODE (se.element_name, 'Pb', min_conc)),0) minpb, nvl(SUM (DECODE (se.element_name, 'Mg', min_conc)),0) minmg, nvl(SUM (DECODE (se.element_name, 'N', min_conc)),0) minn,nvl(SUM (DECODE (se.element_name, 'Ve', min_conc)),0) minve, nvl(SUM (DECODE (se.element_name, 'Co', min_conc)),0) minco, nvl(SUM (DECODE (se.element_name, 'Ce', min_conc)),0) mince, nvl(SUM (DECODE (se.element_name, 'Sb', min_conc)),0) minsb, nvl(SUM (DECODE (se.element_name, 'Zr', min_conc)),0) minzr, nvl(SUM (DECODE (se.element_name, 'O', min_conc)),0) mino, nvl(SUM (DECODE (se.element_name, 'H', min_conc)),0) minh, nvl(SUM (DECODE (se.element_name, 'C', max_conc)),0) maxc, nvl(SUM (DECODE (se.element_name, 'Si', max_conc)),0) maxsi, nvl(SUM (DECODE (se.element_name, 'Mn', max_conc)),0) maxmn, nvl(SUM (DECODE (se.element_name, 'P', max_conc)),0) maxp, nvl(SUM (DECODE (se.element_name, 'S', max_conc)),0) maxs, nvl(SUM (DECODE (se.element_name, 'Al', max_conc)),0) maxal, nvl(SUM (DECODE (se.element_name, 'Al_S', max_conc)),0) maxals, nvl(SUM (DECODE (se.element_name, 'Cu', max_conc)),0) maxcu, nvl(SUM (DECODE (se.element_name, 'Cr', max_conc)),0) maxcr, nvl(SUM (DECODE (se.element_name, 'Mo', max_conc)),0) maxmo, nvl(SUM (DECODE (se.element_name, 'Ni', max_conc)),0) maxni, nvl(SUM (DECODE (se.element_name, 'V', max_conc)),0) maxv, nvl(SUM (DECODE (se.element_name, 'Ti', max_conc)),0) maxti, nvl(SUM (DECODE (se.element_name, 'Nb', max_conc)),0) maxnb, nvl(SUM (DECODE (se.element_name, 'Ca', max_conc)),0) maxca, nvl(SUM (DECODE (se.element_name, 'W', max_conc)),0) maxw, nvl(SUM (DECODE (se.element_name, 'Sn', max_conc)),0) maxsn, nvl(SUM (DECODE (se.element_name, 'As', max_conc)),0) maxas, nvl(SUM (DECODE (se.element_name, 'Te', max_conc)),0) maxte, nvl(SUM (DECODE (se.element_name, 'Bi', max_conc)),0) maxbi, nvl(SUM (DECODE (se.element_name, 'B', max_conc)),0) maxb, nvl(SUM (DECODE (se.element_name, 'Pb', max_conc)),0) maxpb, nvl(SUM (DECODE (se.element_name, 'Mg', max_conc)),0) maxmg, nvl(SUM (DECODE (se.element_name, 'N', max_conc)),0) maxn, nvl(SUM (DECODE (se.element_name, 'Ve', max_conc)),0) maxve, nvl(SUM (DECODE (se.element_name, 'Co', max_conc)),0) maxco, nvl(SUM (DECODE (se.element_name, 'Ce', max_conc)),0) maxce, nvl(SUM (DECODE (se.element_name, 'Sb', max_conc)),0) maxsb, nvl(SUM (DECODE (se.element_name, 'Zr', max_conc)),0) maxzr, nvl(SUM (DECODE (se.element_name, 'O', max_conc)),0) maxo, nvl(SUM (DECODE (se.element_name, 'H', max_conc)),0) maxh, nvl(SUM (DECODE (se.element_name, 'C', aim_conc)),0) aimc, nvl(SUM (DECODE (se.element_name, 'Si', aim_conc)),0) aimsi, nvl(SUM (DECODE (se.element_name, 'Mn', aim_conc)),0) aimmn, nvl(SUM (DECODE (se.element_name, 'P', aim_conc)),0) aimp, nvl(SUM (DECODE (se.element_name, 'S', aim_conc)),0) aims, nvl(SUM (DECODE (se.element_name, 'Al', aim_conc)),0) aimal, nvl(SUM (DECODE (se.element_name, 'Al_S', aim_conc)),0) aimals, nvl(SUM (DECODE (se.element_name, 'Cu', aim_conc)),0) aimcu, nvl(SUM (DECODE (se.element_name, 'Cr', aim_conc)),0) aimcr, nvl(SUM (DECODE (se.element_name, 'Mo', aim_conc)),0) aimmo, nvl(SUM (DECODE (se.element_name, 'Ni', aim_conc)),0) aimni, nvl(SUM (DECODE (se.element_name, 'V', aim_conc)),0) aimv, nvl(SUM (DECODE (se.element_name, 'Ti', aim_conc)),0) aimti, nvl(SUM (DECODE (se.element_name, 'Nb', aim_conc)),0) aimnb,nvl(SUM (DECODE (se.element_name, 'Ca', aim_conc)),0) aimca, nvl(SUM (DECODE (se.element_name, 'W', aim_conc)),0) aimw, nvl(SUM (DECODE (se.element_name, 'Sn', aim_conc)),0) aimsn, nvl(SUM (DECODE (se.element_name, 'As', aim_conc)),0) aimas,nvl(SUM (DECODE (se.element_name, 'Te', aim_conc)),0) aimte, nvl(SUM (DECODE (se.element_name, 'Bi', aim_conc)),0) aimbi,nvl(SUM (DECODE (se.element_name, 'B', aim_conc)),0) aimb, nvl(SUM (DECODE (se.element_name, 'Pb', aim_conc)),0) aimpb, nvl(SUM (DECODE (se.element_name, 'Mg', aim_conc)),0) aimmg, nvl(SUM (DECODE (se.element_name, 'N', aim_conc)),0) aimn, nvl(SUM (DECODE (se.element_name, 'Ve', aim_conc)),0) aimve, nvl(SUM (DECODE (se.element_name, 'Co', aim_conc)),0) aimco, nvl(SUM (DECODE (se.element_name, 'Ce', aim_conc)),0) aimce, nvl(SUM (DECODE (se.element_name, 'Sb', aim_conc)),0) aimsb, nvl(SUM (DECODE (se.element_name, 'Zr', aim_conc)),0) aimzr, nvl(SUM (DECODE (se.element_name, 'O', aim_conc)),0) aimo, nvl(SUM (DECODE (se.element_name, 'H', aim_conc)),0) aimh FROM gcc_spec_analysis se, gcc_grade_spec gs WHERE ((gs.grade_code = se.grade_code)AND (se.grade_code    =  (SELECT grade_code FROM pdc_heat WHERE (steel_id = '" + StrCastSeq_caster2_SteelID + "')  ))) GROUP BY (se.grade_code, gs.grade_code,gs.CALC_LIQ_TEMP)";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatReportLabAnalysis where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatReportLabAnalysis set LIQ_min='" + dt_CastSeq.Rows[i]["aimliq"].ToString() + "',C_min='" + dt_CastSeq.Rows[i]["minc"].ToString() + "',Si_min='" + dt_CastSeq.Rows[i]["minSi"].ToString() + "',Mn_min='" + dt_CastSeq.Rows[i]["minMn"].ToString() + "',P_min='" + dt_CastSeq.Rows[i]["minP"].ToString() + "',S_min='" + dt_CastSeq.Rows[i]["minS"].ToString() + "',Al_min='" + dt_CastSeq.Rows[i]["minAl"].ToString() + "',Al_S_min='" + dt_CastSeq.Rows[i]["minAls"].ToString() + "',Cu_min='" + dt_CastSeq.Rows[i]["minCu"].ToString() + "',Cr_min='" + dt_CastSeq.Rows[i]["minCr"].ToString() + "',Mo_min='" + dt_CastSeq.Rows[i]["minMo"].ToString() + "',Ni_min='" + dt_CastSeq.Rows[i]["minNi"].ToString() + "',V_min='" + dt_CastSeq.Rows[i]["minV"].ToString() + "',Ti_min='" + dt_CastSeq.Rows[i]["minTi"].ToString() + "',Nb_min='" + dt_CastSeq.Rows[i]["minNb"].ToString() + "',Ca_min='" + dt_CastSeq.Rows[i]["MinCa"].ToString() + "',W_min='" + dt_CastSeq.Rows[i]["minW"].ToString() + "',Sn_min='" + dt_CastSeq.Rows[i]["minSn"].ToString() + "',As_min='" + dt_CastSeq.Rows[i]["minAs"].ToString() + "',Te_min='" + dt_CastSeq.Rows[i]["minTe"].ToString() + "',Bi_min='" + dt_CastSeq.Rows[i]["minBi"].ToString() + "',B_min='" + dt_CastSeq.Rows[i]["minB"].ToString() + "',Pb_min='" + dt_CastSeq.Rows[i]["minPb"].ToString() + "',Mg_min='" + dt_CastSeq.Rows[i]["minMg"].ToString() + "',N_min='" + dt_CastSeq.Rows[i]["minN"].ToString() + "',Ve_min='" + dt_CastSeq.Rows[i]["minVe"].ToString() + "',Co_min='" + dt_CastSeq.Rows[i]["minCo"].ToString() + "',Ce_min='" + dt_CastSeq.Rows[i]["minCe"].ToString() + "',Sb_min='" + dt_CastSeq.Rows[i]["minSb"].ToString() + "',Zr_min='" + dt_CastSeq.Rows[i]["minZr"].ToString() + "',O_min='" + dt_CastSeq.Rows[i]["minO"].ToString() + "',H_min='" + dt_CastSeq.Rows[i]["minH"].ToString() + "',C_max='" + dt_CastSeq.Rows[i]["maxC"].ToString() + "',Si_max='" + dt_CastSeq.Rows[i]["maxSi"].ToString() + "',Mn_max='" + dt_CastSeq.Rows[i]["maxMn"].ToString() + "',P_max='" + dt_CastSeq.Rows[i]["maxP"].ToString() + "',S_max='" + dt_CastSeq.Rows[i]["maxS"].ToString() + "',Al_max='" + dt_CastSeq.Rows[i]["maxAl"].ToString() + "',Al_S_max='" + dt_CastSeq.Rows[i]["maxAlS"].ToString() + "',Cu_max='" + dt_CastSeq.Rows[i]["maxCu"].ToString() + "',Cr_max='" + dt_CastSeq.Rows[i]["maxCr"].ToString() + "',Mo_max='" + dt_CastSeq.Rows[i]["maxMo"].ToString() + "',Ni_max='" + dt_CastSeq.Rows[i]["maxNi"].ToString() + "',V_max='" + dt_CastSeq.Rows[i]["maxV"].ToString() + "',Ti_max='" + dt_CastSeq.Rows[i]["maxTi"].ToString() + "',Nb_max='" + dt_CastSeq.Rows[i]["maxNb"].ToString() + "',Ca_max='" + dt_CastSeq.Rows[i]["maxCa"].ToString() + "',W_max='" + dt_CastSeq.Rows[i]["maxW"].ToString() + "',Sn_max='" + dt_CastSeq.Rows[i]["maxSn"].ToString() + "',As_max='" + dt_CastSeq.Rows[i]["maxAs"].ToString() + "',Te_max='" + dt_CastSeq.Rows[i]["maxTe"].ToString() + "',Bi_max='" + dt_CastSeq.Rows[i]["maxBi"].ToString() + "',B_max='" + dt_CastSeq.Rows[i]["maxB"].ToString() + "',Pb_max='" + dt_CastSeq.Rows[i]["maxPb"].ToString() + "',Mg_max='" + dt_CastSeq.Rows[i]["maxMg"].ToString() + "',N_max='" + dt_CastSeq.Rows[i]["maxN"].ToString() + "',Ve_max='" + dt_CastSeq.Rows[i]["maxVe"].ToString() + "',Co_max='" + dt_CastSeq.Rows[i]["maxCo"].ToString() + "',Ce_max='" + dt_CastSeq.Rows[i]["maxCe"].ToString() + "',Sb_max='" + dt_CastSeq.Rows[i]["maxSb"].ToString() + "',Zr_max='" + dt_CastSeq.Rows[i]["maxZr"].ToString() + "',O_max='" + dt_CastSeq.Rows[i]["maxO"].ToString() + "',H_max='" + dt_CastSeq.Rows[i]["maxH"].ToString() + "',C_aim='" + dt_CastSeq.Rows[i]["aimC"].ToString() + "',Si_aim='" + dt_CastSeq.Rows[i]["aimSi"].ToString() + "',Mn_aim='" + dt_CastSeq.Rows[i]["aimMn"].ToString() + "',P_aim='" + dt_CastSeq.Rows[i]["aimP"].ToString() + "',S_aim='" + dt_CastSeq.Rows[i]["aimS"].ToString() + "',Al_aim='" + dt_CastSeq.Rows[i]["aimAl"].ToString() + "',Al_S_aim='" + dt_CastSeq.Rows[i]["aimAlS"].ToString() + "',Cu_aim='" + dt_CastSeq.Rows[i]["aimCu"].ToString() + "',Cr_aim='" + dt_CastSeq.Rows[i]["aimCr"].ToString() + "',Mo_aim='" + dt_CastSeq.Rows[i]["aimMo"].ToString() + "',Ni_aim='" + dt_CastSeq.Rows[i]["aimNi"].ToString() + "',V_aim='" + dt_CastSeq.Rows[i]["aimV"].ToString() + "',Ti_aim='" + dt_CastSeq.Rows[i]["aimTi"].ToString() + "',Nb_aim='" + dt_CastSeq.Rows[i]["aimNb"].ToString() + "',Ca_aim='" + dt_CastSeq.Rows[i]["aimCa"].ToString() + "',W_aim='" + dt_CastSeq.Rows[i]["aimW"].ToString() + "',Sn_aim='" + dt_CastSeq.Rows[i]["aimSn"].ToString() + "',As_aim='" + dt_CastSeq.Rows[i]["aimAs"].ToString() + "',Te_aim='" + dt_CastSeq.Rows[i]["aimTe"].ToString() + "',Bi_aim='" + dt_CastSeq.Rows[i]["aimBi"].ToString() + "',B_aim='" + dt_CastSeq.Rows[i]["aimB"].ToString() + "',Pb_aim='" + dt_CastSeq.Rows[i]["aimPb"].ToString() + "',Mg_aim='" + dt_CastSeq.Rows[i]["aimMg"].ToString() + "',N_aim='" + dt_CastSeq.Rows[i]["aimN"].ToString() + "',Ve_aim='" + dt_CastSeq.Rows[i]["aimVe"].ToString() + "',Co_aim='" + dt_CastSeq.Rows[i]["aimCo"].ToString() + "',Ce_aim='" + dt_CastSeq.Rows[i]["aimCe"].ToString() + "',Sb_aim='" + dt_CastSeq.Rows[i]["aimSb"].ToString() + "',Zr_aim='" + dt_CastSeq.Rows[i]["aimZr"].ToString() + "',O_aim='" + dt_CastSeq.Rows[i]["aimO"].ToString() + "',H_aim='" + dt_CastSeq.Rows[i]["aimH"].ToString() + "' where SteelId='" + StrCastSeq_caster2_SteelID + "' and SeqNo='" + StrCastSeq_caster2_MaxID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatReportLabAnalysis(SeqNo,SteelId,LIQ_min,C_min,Si_min,Mn_min,P_min,S_min,Al_min,Al_S_min,Cu_min,Cr_min,Mo_min,Ni_min,V_min,Ti_min,Nb_min,Ca_min,W_min,Sn_min,As_min,Te_min,Bi_min,B_min,Pb_min,Mg_min,N_min,Ve_min,Co_min,Ce_min,Sb_min,Zr_min,O_min,H_min,C_max,Si_max,Mn_max,P_max,S_max,Al_max,Al_S_max,Cu_max,Cr_max,Mo_max,Ni_max,V_max,Ti_max,Nb_max,Ca_max,W_max,Sn_max,As_max,Te_max,Bi_max,B_max,Pb_max,Mg_max,N_max,Ve_max,Co_max,Ce_max,Sb_max,Zr_max,O_max,H_max,C_aim,Si_aim,Mn_aim,P_aim,S_aim,Al_aim,Al_S_aim,Cu_aim,Cr_aim,Mo_aim,Ni_aim,V_aim,Ti_aim,Nb_aim,Ca_aim,W_aim,Sn_aim,As_aim,Te_aim,Bi_aim,B_aim,Pb_aim,Mg_aim,N_aim,Ve_aim,Co_aim,Ce_aim,Sb_aim,Zr_aim,O_aim,H_aim)values('" + StrCastSeq_caster2_MaxID + "','" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["aimliq"].ToString() + "','" + dt_CastSeq.Rows[i]["minC"].ToString() + "','" + dt_CastSeq.Rows[i]["minSi"].ToString() + "','" + dt_CastSeq.Rows[i]["minMn"].ToString() + "','" + dt_CastSeq.Rows[i]["minP"].ToString() + "','" + dt_CastSeq.Rows[i]["minS"].ToString() + "','" + dt_CastSeq.Rows[i]["minAl"].ToString() + "','" + dt_CastSeq.Rows[i]["minAlS"].ToString() + "','" + dt_CastSeq.Rows[i]["minCu"].ToString() + "','" + dt_CastSeq.Rows[i]["minCr"].ToString() + "','" + dt_CastSeq.Rows[i]["minMo"].ToString() + "','" + dt_CastSeq.Rows[i]["minNi"].ToString() + "','" + dt_CastSeq.Rows[i]["minV"].ToString() + "','" + dt_CastSeq.Rows[i]["minTi"].ToString() + "','" + dt_CastSeq.Rows[i]["minNb"].ToString() + "','" + dt_CastSeq.Rows[i]["minCa"].ToString() + "','" + dt_CastSeq.Rows[i]["minW"].ToString() + "','" + dt_CastSeq.Rows[i]["minsn"].ToString() + "','" + dt_CastSeq.Rows[i]["minAs"].ToString() + "','" + dt_CastSeq.Rows[i]["minTe"].ToString() + "','" + dt_CastSeq.Rows[i]["minBi"].ToString() + "','" + dt_CastSeq.Rows[i]["minB"].ToString() + "','" + dt_CastSeq.Rows[i]["minPb"].ToString() + "','" + dt_CastSeq.Rows[i]["minMg"].ToString() + "','" + dt_CastSeq.Rows[i]["minN"].ToString() + "','" + dt_CastSeq.Rows[i]["minVe"].ToString() + "','" + dt_CastSeq.Rows[i]["minCo"].ToString() + "','" + dt_CastSeq.Rows[i]["minCe"].ToString() + "','" + dt_CastSeq.Rows[i]["minSb"].ToString() + "','" + dt_CastSeq.Rows[i]["minZr"].ToString() + "','" + dt_CastSeq.Rows[i]["minO"].ToString() + "','" + dt_CastSeq.Rows[i]["minH"].ToString() + "','" + dt_CastSeq.Rows[i]["maxC"].ToString() + "','" + dt_CastSeq.Rows[i]["maxSi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxMn"].ToString() + "','" + dt_CastSeq.Rows[i]["maxP"].ToString() + "','" + dt_CastSeq.Rows[i]["maxS"].ToString() + "','" + dt_CastSeq.Rows[i]["maxAl"].ToString() + "','" + dt_CastSeq.Rows[i]["maxAlS"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCu"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCr"].ToString() + "','" + dt_CastSeq.Rows[i]["maxMo"].ToString() + "','" + dt_CastSeq.Rows[i]["maxNi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxV"].ToString() + "','" + dt_CastSeq.Rows[i]["maxTi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxNb"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCa"].ToString() + "','" + dt_CastSeq.Rows[i]["maxW"].ToString() + "','" + dt_CastSeq.Rows[i]["maxSn"].ToString() + "','" + dt_CastSeq.Rows[i]["maxAs"].ToString() + "','" + dt_CastSeq.Rows[i]["maxTe"].ToString() + "','" + dt_CastSeq.Rows[i]["maxBi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxB"].ToString() + "','" + dt_CastSeq.Rows[i]["maxPb"].ToString() + "','" + dt_CastSeq.Rows[i]["maxMg"].ToString() + "','" + dt_CastSeq.Rows[i]["maxN"].ToString() + "','" + dt_CastSeq.Rows[i]["maxVe"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCo"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCe"].ToString() + "','" + dt_CastSeq.Rows[i]["maxSb"].ToString() + "','" + dt_CastSeq.Rows[i]["maxZr"].ToString() + "','" + dt_CastSeq.Rows[i]["maxO"].ToString() + "','" + dt_CastSeq.Rows[i]["maxH"].ToString() + "','" + dt_CastSeq.Rows[i]["aimC"].ToString() + "','" + dt_CastSeq.Rows[i]["aimSi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimMn"].ToString() + "','" + dt_CastSeq.Rows[i]["aimP"].ToString() + "','" + dt_CastSeq.Rows[i]["aimS"].ToString() + "','" + dt_CastSeq.Rows[i]["aimAl"].ToString() + "','" + dt_CastSeq.Rows[i]["aimAlS"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCu"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCr"].ToString() + "','" + dt_CastSeq.Rows[i]["aimMo"].ToString() + "','" + dt_CastSeq.Rows[i]["aimNi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimV"].ToString() + "','" + dt_CastSeq.Rows[i]["aimTi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimNb"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCa"].ToString() + "','" + dt_CastSeq.Rows[i]["aimW"].ToString() + "','" + dt_CastSeq.Rows[i]["aimSn"].ToString() + "','" + dt_CastSeq.Rows[i]["aimAs"].ToString() + "','" + dt_CastSeq.Rows[i]["aimTe"].ToString() + "','" + dt_CastSeq.Rows[i]["aimBi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimB"].ToString() + "','" + dt_CastSeq.Rows[i]["aimPb"].ToString() + "','" + dt_CastSeq.Rows[i]["aimMg"].ToString() + "','" + dt_CastSeq.Rows[i]["aimN"].ToString() + "','" + dt_CastSeq.Rows[i]["aimVe"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCo"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCe"].ToString() + "','" + dt_CastSeq.Rows[i]["aimSb"].ToString() + "','" + dt_CastSeq.Rows[i]["aimZr"].ToString() + "','" + dt_CastSeq.Rows[i]["aimO"].ToString() + "','" + dt_CastSeq.Rows[i]["aimH"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatReportTempReport(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = "SELECT (pdc_tundish_temp.meas_temp - gcc_grade_spec.calc_liq_temp) AS overheat,pdc_tundish_temp.meas_temp as Temp,pdc_tundish_temp.meas_time as Time FROM gcc_grade_spec,pdc_heat,pdc_tundish_temp WHERE pdc_heat.grade_code = gcc_grade_spec.grade_code AND pdc_heat.steel_id     = pdc_tundish_temp.steel_id AND pdc_heat.steel_id     = '" + StrCastSeq_caster2_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatReportTempReport where SteelId='" + StrCastSeq_caster2_SteelID + "' and Time='" + (dt_CastSeq.Rows[i]["Time"].ToString().Replace(",", ".")).ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatReportTempReport set Temp='" + dt_CastSeq.Rows[i]["Temp"].ToString() + "',Temp='" + dt_CastSeq.Rows[i]["OverHeat"].ToString() + "' where SteelId='" + StrCastSeq_caster2_SteelID + "' and Time='" + (dt_CastSeq.Rows[i]["Time"].ToString().Replace(",", ".")).ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatReportTempReport(SeqNo,SteelId,Time,Temp,OverHeat)values('" + StrCastSeq_caster2_MaxID + "','" + StrCastSeq_caster2_SteelID + "','" + (dt_CastSeq.Rows[i]["Time"].ToString().Replace(",", ".")).ToString() + "','" + dt_CastSeq.Rows[i]["Temp"].ToString() + "','" + dt_CastSeq.Rows[i]["OverHeat"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatRpt_PlandataReport(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = "  SELECT pdc_slab.l2_slabid,nvl(pdc_slab_order.length_min,0) as length_min,nvl(pdc_slab_order.length_aim,0) as length_aim, nvl(pdc_slab_order.length_max,0) as length_max, nvl(pdc_slab_order.width_head,0) as width_head, nvl(pdc_slab_order.width_tail,0) as width_tail, nvl(pdc_piece.heat_steel_id,0) as heat_steel_id, nvl(pdc_slab.slab_seq_no,0) as slab_seq_no, nvl(pdc_piece.piece_seq_no,0) as piece_seq_no FROM l2ccs.pdc_piece pdc_piece,l2ccs.pdc_slab_order pdc_slab_order,l2ccs.pdc_slab pdc_slab  WHERE (pdc_slab_order.steel_id(+) = pdc_slab.steel_id) AND (pdc_piece.steel_id           = pdc_slab.steel_id(+)) AND (pdc_piece.heat_steel_id      = '" + StrCastSeq_caster2_SteelID + "')";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_PlandataReport where SteelId='" + StrCastSeq_caster2_SteelID + "' and L2_SlabId='" + dt_CastSeq.Rows[i]["l2_slabid"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_PlandataReport set L2_SlabId='" + dt_CastSeq.Rows[i]["L2_SlabId"].ToString() + "',Length_min='" + dt_CastSeq.Rows[i]["Length_min"].ToString() + "',Length_aim='" + dt_CastSeq.Rows[i]["Length_aim"].ToString() + "',Length_max='" + dt_CastSeq.Rows[i]["Length_max"].ToString() + "',Width_Head='" + dt_CastSeq.Rows[i]["Width_Head"].ToString() + "',Width_Tail='" + dt_CastSeq.Rows[i]["Width_Tail"].ToString() + "',Slab_Seq_No='" + dt_CastSeq.Rows[i]["Slab_Seq_No"].ToString() + "',Piece_Seq_No='" + dt_CastSeq.Rows[i]["Piece_Seq_No"].ToString() + "' where SteelId='" + StrCastSeq_caster2_SteelID + "' and SeqNo='" + StrCastSeq_caster2_MaxID + "'  and L2_SlabId='" + dt_CastSeq.Rows[i]["l2_slabid"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_PlandataReport(SeqNo,SteelId,L2_SlabId,Length_min,Length_aim,Length_max,Width_Head,Width_Tail,Slab_Seq_No,Piece_Seq_No)values('" + StrCastSeq_caster2_MaxID + "','" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["L2_SlabId"].ToString() + "','" + dt_CastSeq.Rows[i]["Length_min"].ToString() + "','" + dt_CastSeq.Rows[i]["Length_aim"].ToString() + "','" + dt_CastSeq.Rows[i]["Length_max"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_Head"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_Tail"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_Seq_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Piece_Seq_No"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatRptAnalysisReport(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = "SELECT a.heatid heatid,  a.sample_no,  a.taken_time,  a.sample_loc,  a.LIQUID_TEMP LIQ,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'C'  AND ae.analysis_id = a.analysis_id  ),0) c,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Si'  AND ae.analysis_id = a.analysis_id  ) ,0)si,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Mn'  AND ae.analysis_id = a.analysis_id  ),0) mn,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'P'  AND ae.analysis_id = a.analysis_id  ),0) p,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'S'  AND ae.analysis_id = a.analysis_id  ),0) s,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Al'  AND ae.analysis_id = a.analysis_id  ),0) al,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Als'  AND ae.analysis_id = a.analysis_id  ),0) als,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Cu'  AND ae.analysis_id = a.analysis_id  ),0) cu,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Cr'  AND ae.analysis_id = a.analysis_id  ),0) cr,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Mo'  AND ae.analysis_id = a.analysis_id  ),0) mo,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ni'  AND ae.analysis_id = a.analysis_id  ),0) ni,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'V'  AND ae.analysis_id = a.analysis_id  ),0) v,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ti'  AND ae.analysis_id = a.analysis_id  ),0) ti,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Nb'  AND ae.analysis_id = a.analysis_id  ),0) nb,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ca'  AND ae.analysis_id = a.analysis_id  ),0) ca,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'W'  AND ae.analysis_id = a.analysis_id  ),0) w,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Sn'  AND ae.analysis_id = a.analysis_id  ),0) sn,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'As'  AND ae.analysis_id = a.analysis_id  ),0) AS a_s,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Te'  AND ae.analysis_id = a.analysis_id  ) ,0)te,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Bi'  AND ae.analysis_id = a.analysis_id  ) ,0)bi,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'B'  AND ae.analysis_id = a.analysis_id  ),0) b,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Pb'  AND ae.analysis_id = a.analysis_id  ),0) Pb,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Mg'  AND ae.analysis_id = a.analysis_id  ),0) Mg,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'N'  AND ae.analysis_id = a.analysis_id  ),0) N,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ve'  AND ae.analysis_id = a.analysis_id  ),0) Ve,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Co'  AND ae.analysis_id = a.analysis_id  ),0) Co,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ce'  AND ae.analysis_id = a.analysis_id  ),0) Ce,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Sb'  AND ae.analysis_id = a.analysis_id  ),0) Sb,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Zr'  AND ae.analysis_id = a.analysis_id  ),0) Zr,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'O'  AND ae.analysis_id = a.analysis_id  ),0) O,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'H'  AND ae.analysis_id = a.analysis_id  ),0) H FROM pd_analysis a,  pdc_heat h WHERE (a.heatid = h.heatid) AND (h.steel_id ='" + StrCastSeq_caster2_SteelID + "')";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRptAnalysisReport where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "' and heatid='" + dt_CastSeq.Rows[i]["heatid"].ToString() + "' and Taken_Time='" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRptAnalysisReport set Sample_No='" + dt_CastSeq.Rows[i]["Sample_No"].ToString() + "',Taken_Time='" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "',Sample_LOC='" + dt_CastSeq.Rows[i]["Sample_LOC"].ToString() + "',LIQ='" + dt_CastSeq.Rows[i]["LIQ"].ToString() + "',C='" + dt_CastSeq.Rows[i]["C"].ToString() + "',Si='" + dt_CastSeq.Rows[i]["Si"].ToString() + "',Mn='" + dt_CastSeq.Rows[i]["Mn"].ToString() + "',P='" + dt_CastSeq.Rows[i]["P"].ToString() + "',S='" + dt_CastSeq.Rows[i]["S"].ToString() + "',Al='" + dt_CastSeq.Rows[i]["Al"].ToString() + "',AlS='" + dt_CastSeq.Rows[i]["AlS"].ToString() + "',Cu='" + dt_CastSeq.Rows[i]["Cu"].ToString() + "',Cr='" + dt_CastSeq.Rows[i]["Cr"].ToString() + "',Mo='" + dt_CastSeq.Rows[i]["Mo"].ToString() + "',Ni='" + dt_CastSeq.Rows[i]["Ni"].ToString() + "',V='" + dt_CastSeq.Rows[i]["V"].ToString() + "',Ti='" + dt_CastSeq.Rows[i]["Ti"].ToString() + "',NB='" + dt_CastSeq.Rows[i]["Nb"].ToString() + "',Ca='" + dt_CastSeq.Rows[i]["Ca"].ToString() + "',W='" + dt_CastSeq.Rows[i]["W"].ToString() + "',Sn='" + dt_CastSeq.Rows[i]["Sn"].ToString() + "',A_S='" + dt_CastSeq.Rows[i]["A_S"].ToString() + "',TE='" + dt_CastSeq.Rows[i]["TE"].ToString() + "',Bi='" + dt_CastSeq.Rows[i]["Bi"].ToString() + "',B='" + dt_CastSeq.Rows[i]["B"].ToString() + "',Pb='" + dt_CastSeq.Rows[i]["Pb"].ToString() + "',Mg='" + dt_CastSeq.Rows[i]["Mg"].ToString() + "',N='" + dt_CastSeq.Rows[i]["N"].ToString() + "',Ve='" + dt_CastSeq.Rows[i]["Ve"].ToString() + "',Co='" + dt_CastSeq.Rows[i]["Co"].ToString() + "',Ce='" + dt_CastSeq.Rows[i]["Ce"].ToString() + "',Sb='" + dt_CastSeq.Rows[i]["Sb"].ToString() + "',Zr='" + dt_CastSeq.Rows[i]["Zr"].ToString() + "',O='" + dt_CastSeq.Rows[i]["O"].ToString() + "',H='" + dt_CastSeq.Rows[i]["H"].ToString() + "' where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "' and heatid='" + dt_CastSeq.Rows[i]["heatid"].ToString() + "' and Taken_Time='" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRptAnalysisReport(SeqNo,SteelId,HeatId,Sample_No,Taken_Time,Sample_LOC,LIQ,C,Si,Mn,P,S,Al,AlS,Cu,Cr,Mo,Ni,V,Ti,NB,Ca,W,Sn,A_S,TE,Bi,B,Pb,Mg,N,Ve,Co,Ce,Sb,Zr,O,H)values('" + StrCastSeq_caster2_MaxID + "','" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "','" + dt_CastSeq.Rows[i]["Sample_No"].ToString() + "','" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Sample_LOC"].ToString() + "','" + dt_CastSeq.Rows[i]["LIQ"].ToString() + "','" + dt_CastSeq.Rows[i]["C"].ToString() + "','" + dt_CastSeq.Rows[i]["Si"].ToString() + "','" + dt_CastSeq.Rows[i]["Mn"].ToString() + "','" + dt_CastSeq.Rows[i]["P"].ToString() + "','" + dt_CastSeq.Rows[i]["S"].ToString() + "','" + dt_CastSeq.Rows[i]["Al"].ToString() + "','" + dt_CastSeq.Rows[i]["AlS"].ToString() + "','" + dt_CastSeq.Rows[i]["Cu"].ToString() + "','" + dt_CastSeq.Rows[i]["Cr"].ToString() + "','" + dt_CastSeq.Rows[i]["Mo"].ToString() + "','" + dt_CastSeq.Rows[i]["Ni"].ToString() + "','" + dt_CastSeq.Rows[i]["V"].ToString() + "','" + dt_CastSeq.Rows[i]["Ti"].ToString() + "','" + dt_CastSeq.Rows[i]["NB"].ToString() + "','" + dt_CastSeq.Rows[i]["Ca"].ToString() + "','" + dt_CastSeq.Rows[i]["W"].ToString() + "','" + dt_CastSeq.Rows[i]["Sn"].ToString() + "','" + dt_CastSeq.Rows[i]["A_S"].ToString() + "','" + dt_CastSeq.Rows[i]["TE"].ToString() + "','" + dt_CastSeq.Rows[i]["Bi"].ToString() + "','" + dt_CastSeq.Rows[i]["B"].ToString() + "','" + dt_CastSeq.Rows[i]["Pb"].ToString() + "','" + dt_CastSeq.Rows[i]["Mg"].ToString() + "','" + dt_CastSeq.Rows[i]["N"].ToString() + "','" + dt_CastSeq.Rows[i]["Ve"].ToString() + "','" + dt_CastSeq.Rows[i]["Co"].ToString() + "','" + dt_CastSeq.Rows[i]["Ce"].ToString() + "','" + dt_CastSeq.Rows[i]["Sb"].ToString() + "','" + dt_CastSeq.Rows[i]["Zr"].ToString() + "','" + dt_CastSeq.Rows[i]["O"].ToString() + "','" + dt_CastSeq.Rows[i]["H"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatRptCutDataReport(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = " SELECT p.strand_no,sl.slab_seq_no   AS slab_no,sl.l2_slabid  AS marking_no,pne.value/1000 AS cold_length,pt.piece_type_desc   AS piece_type,p.head_cast_len   AS cast_length_cut,so.length_min   AS min_length,so.length_aim   AS aim_length,so.length_max  AS max_length,(p.piece_length/1000 - NVL (scrap.scrap_length, 0)) AS good_length,p.piece_length/1000  AS act_length,(DECODE(NVL(piece_weight_meas,0), 0, piece_weight_calc , piece_weight_meas) - NVL (scrap.scrap_mass, 0)) / 1000  AS good_weight,NVL (scrap.scrap_mass, 0) / 1000  AS scrap_weight,p.head_thickness,p.head_width,p.mixzone_begin AS mixzone_begin,p.mixzone_end   AS mixzone_end,p.heat_boundary_pos1,p.sample_weight FROM pdc_piece p,pdc_slab sl,pdc_slab_order so,pdc_heat h,gdc_piece_type pt,pdc_number_entry pne,(SELECT steel_id,SUM (scrap_mass) AS scrap_mass,(SUM (scrap_end) - SUM (scrap_bgn)) AS scrap_length FROM pdc_scrap  GROUP BY steel_id  ) scrap WHERE (p.steel_id    = sl.steel_id(+)) AND (p.steel_id      = so.steel_id(+)) AND (p.steel_id      = scrap.steel_id(+)) AND (p.piece_type    = pt.piece_type(+)) AND (p.heat_steel_id = h.steel_id) AND pne.val_name     = 'ColdLen' AND pne.steel_id     = p.steel_id AND (h.steel_id = '" + StrCastSeq_caster2_SteelID + "') ORDER BY p.head_cast_len ASC";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRptCutDataReport where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "' and Marking_No='" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRptCutDataReport set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Slab_No='" + dt_CastSeq.Rows[i]["Slab_No"].ToString() + "',Marking_No='" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "',Cold_Length='" + dt_CastSeq.Rows[i]["Cold_Length"].ToString() + "',Piece_Type='" + dt_CastSeq.Rows[i]["Piece_Type"].ToString() + "',Cast_Length_Cut='" + dt_CastSeq.Rows[i]["Cast_Length_Cut"].ToString() + "',Min_Length='" + dt_CastSeq.Rows[i]["Min_Length"].ToString() + "',Aim_Length='" + dt_CastSeq.Rows[i]["Aim_Length"].ToString() + "',Max_Length='" + dt_CastSeq.Rows[i]["Max_Length"].ToString() + "',Good_Length='" + dt_CastSeq.Rows[i]["Good_Length"].ToString() + "',Act_Length='" + dt_CastSeq.Rows[i]["Act_Length"].ToString() + "',Good_Weight='" + dt_CastSeq.Rows[i]["Good_Weight"].ToString() + "',scrap_Weight='" + dt_CastSeq.Rows[i]["scrap_Weight"].ToString() + "',Head_Thickness='" + dt_CastSeq.Rows[i]["Head_Thickness"].ToString() + "',Head_Width='" + dt_CastSeq.Rows[i]["Head_Width"].ToString() + "',MixZone_Begin='" + dt_CastSeq.Rows[i]["MixZone_Begin"].ToString() + "',MixZone_End='" + dt_CastSeq.Rows[i]["MixZone_End"].ToString() + "',Heat_Boundary_POS1='" + dt_CastSeq.Rows[i]["Heat_Boundary_POS1"].ToString() + "',Sample_Weight='" + dt_CastSeq.Rows[i]["sample_weight"].ToString() + "' where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "' and Marking_No='" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRptCutDataReport(SeqNo,SteelId,Strand_No,Slab_No,Marking_No,Cold_Length,Piece_Type,Cast_Length_Cut,Min_Length,Aim_Length,Max_Length,Good_Length,Act_Length,Good_Weight,scrap_Weight,Head_Thickness,Head_Width,MixZone_Begin,MixZone_End,Heat_Boundary_POS1,Sample_Weight)values('" + StrCastSeq_caster2_MaxID + "','" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Cold_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Piece_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Length_Cut"].ToString() + "','" + dt_CastSeq.Rows[i]["Min_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Aim_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Max_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Good_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Act_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Good_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["scrap_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Head_Thickness"].ToString() + "','" + dt_CastSeq.Rows[i]["Head_Width"].ToString() + "','" + dt_CastSeq.Rows[i]["MixZone_Begin"].ToString() + "','" + dt_CastSeq.Rows[i]["MixZone_End"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Boundary_POS1"].ToString() + "','" + dt_CastSeq.Rows[i]["sample_weight"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatRptHeader(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = "SELECT  ROUND(nvl(v.slab_weight_net,0),3) AS slab_weight,  nvl(heat.CUT_LOST_WEIGHT,0)   AS cut_lost_weight,  nvl(v.count_slabs,0)          AS slab_count, nvl(heat.avg_slab_width,0) as avg_slab_width, nvl(heat.heatid,0) as heatid,  nvl(heat.heat_seq_no,0) as heat_seq_no,  nvl(heat.grade_code,0) as grade_code, cheat.shift_foreman  AS shift_foreman,  cheat.casting_foreman AS casting_foreman,  heat.prod_orderid, nvl(heat.ladle_no,0) as ladle_no,  nvl(v.ladle_open_weight,0)  AS ladle_open_weight,  nvl(v.ladle_close_weight,0) AS ladle_close_weight,  nvl(v.total_cast_weight,0)  AS cast_weight,  nvl(seq.seq_no,0) as seq_no,  ROUND(nvl(v.yield,2),2) AS yield,  hshift.crew_id   AS crewid,   nvl(heat.treat_counter,0)  AS treat_counter,   nvl(heat.slab_length,0)  AS slab_length,   heat.ladle_close_time  AS ladle_close_time,   nvl(heat.pouring_duration,0)  AS pouring_duration,   heat.ladle_open_time AS ladle_open_time,   nvl(heat.ladle_open_time,0) AS prod_date,   heat.ladle_arrival_time AS ladle_arrival_time,   nvl(heat.tund_powder_type,0)  AS tund_powder_type,   nvl(heat.tund_powder_weight,0) AS tund_powder_weight,  tund_open.tund_open_time AS tund_open_time,   tund_close.tund_close_time,     -1     AS tund_life, nvl(tund_open.tund_open_weight    / 1000 ,0)AS tund_open_weight, nvl( tund_close.tund_close_weight  /1000 ,0) AS tund_close_weight, nvl(DECODE (NVL(tund_open.tund_no,-1), -1, tund_close.tund_no, tund_open.tund_no,0),0) tund_no, nvl(DECODE (NVL(tund_open.car_no, -1), -1, tund_close.car_no, tund_open.car_no,0),0) car_no,  nvl(v.sample_weight,0) AS sample_lost_weight,  nvl(v.scale_loss,0)    AS scale_lost_weight, nvl( heat.ladle_tare_weight,0) as ladle_tare_weight,  nvl((std.cast_len_end     -std.cast_len_bgn) ,0)                           AS total_cast_length,  ROUND(nvl((v.scrap_weight + v.head_crop_weight + v.tail_crop_weight),0),3) AS total_scrap_weight,  param.value                                                         AS casterid FROM pdc_heat heat,  pdc_strand std,  pdc_heat_strand hs,  pdc_heat_shift hshift,  pdc_customer_heat cheat,  pdc_sequence seq,  v_heat_tund_open tund_open,  v_heat_tund_close tund_close,  v_heat_weight v,  gdc_param param WHERE heat.sequence_steel_id    = seq.steel_id AND hs.strand_steel_id          = std.steel_id AND hs.heat_steel_id            = heat.steel_id AND hshift.heat_steel_id(+)     = heat.steel_id AND tund_open.heat_steel_id(+)  = heat.steel_id AND tund_close.heat_steel_id(+) = heat.steel_id AND cheat.steel_id(+)           = heat.steel_id AND v.steel_id(+)               = heat.steel_id AND param.par_context           = 'CASTER' AND param.par_name              = 'CASTERID' AND heat.steel_id= '" + StrCastSeq_caster2_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRptHeader where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "' and HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRptHeader set Slab_Weight ='" + dt_CastSeq.Rows[i]["Slab_Weight"].ToString() + "',Cut_Lost_Weight ='" + dt_CastSeq.Rows[i]["Cut_Lost_Weight"].ToString() + "',Slab_Count='" + dt_CastSeq.Rows[i]["Slab_Count"].ToString() + "',Avg_Slab_Width='" + dt_CastSeq.Rows[i]["Avg_Slab_Width"].ToString() + "',HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "',Heat_Seq_No='" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "',Grade_Code='" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "',Shift_Foreman='" + dt_CastSeq.Rows[i]["Shift_Foreman"].ToString() + "',Casting_Foreman='" + dt_CastSeq.Rows[i]["Casting_Foreman"].ToString() + "',Prod_OrderId='" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "',Ladle_No='" + dt_CastSeq.Rows[i]["Ladle_No"].ToString() + "',Ladle_Open_Weight='" + dt_CastSeq.Rows[i]["Ladle_Open_Weight"].ToString() + "',Ladle_Close_Weight='" + dt_CastSeq.Rows[i]["Ladle_Close_Weight"].ToString() + "',Cast_Weight='" + dt_CastSeq.Rows[i]["Cast_Weight"].ToString() + "',Yield='" + dt_CastSeq.Rows[i]["Yield"].ToString() + "',CrewId='" + dt_CastSeq.Rows[i]["CrewId"].ToString() + "',Treat_Counter='" + dt_CastSeq.Rows[i]["Treat_Counter"].ToString() + "',Slab_Length='" + dt_CastSeq.Rows[i]["Slab_Length"].ToString() + "',Ladle_Close_Time='" + (dt_CastSeq.Rows[i]["Ladle_Close_Time"].ToString()).Replace(",", ".").ToString() + "',Pouring_Duration='" + dt_CastSeq.Rows[i]["Pouring_Duration"].ToString() + "',Ladle_Open_Time='" + (dt_CastSeq.Rows[i]["Ladle_Open_Time"].ToString()).Replace(",", ".").ToString() + "',Prod_date='" + (dt_CastSeq.Rows[i]["Prod_date"].ToString()).Replace(",", ".").ToString() + "',Ladle_Arrival_Time='" + (dt_CastSeq.Rows[i]["Ladle_Arrival_Time"].ToString()).Replace(",", ".").ToString() + "',Tund_Powder_Type='" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "',Tund_Powder_weight='" + dt_CastSeq.Rows[i]["Tund_Powder_weight"].ToString() + "',Tund_Open_Time='" + (dt_CastSeq.Rows[i]["Tund_Open_Time"].ToString()).Replace(",", ".").ToString() + "',Tund_Close_Time='" + (dt_CastSeq.Rows[i]["Tund_Close_Time"].ToString()).Replace(",", ".").ToString() + "',Tund_Life='" + dt_CastSeq.Rows[i]["Tund_Life"].ToString() + "',Tund_Open_Weight='" + dt_CastSeq.Rows[i]["Tund_Open_Weight"].ToString() + "',Tund_Close_Weight='" + dt_CastSeq.Rows[i]["Tund_Close_Weight"].ToString() + "',Tund_No='" + dt_CastSeq.Rows[i]["Tund_No"].ToString() + "',Car_No='" + dt_CastSeq.Rows[i]["Car_No"].ToString() + "',Sample_Lost_Weight='" + dt_CastSeq.Rows[i]["Sample_Lost_Weight"].ToString() + "',Scale_Lost_Weight='" + dt_CastSeq.Rows[i]["Scale_Lost_Weight"].ToString() + "',Ladle_Tare_Weight='" + dt_CastSeq.Rows[i]["Ladle_Tare_Weight"].ToString() + "',Total_Cast_Length='" + dt_CastSeq.Rows[i]["Total_Cast_Length"].ToString() + "',Total_Scrap_Weight='" + dt_CastSeq.Rows[i]["Total_Scrap_Weight"].ToString() + "',CasterId='" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "' where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "' and HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRptHeader(SeqNo,SteelId,Slab_Weight,Cut_Lost_Weight,Slab_Count,Avg_Slab_Width,HeatId,Heat_Seq_No,Grade_Code,Shift_Foreman,Casting_Foreman,Prod_OrderId,Ladle_No,Ladle_Open_Weight,Ladle_Close_Weight,Cast_Weight,Yield,CrewId,Treat_Counter,Slab_Length,Ladle_Close_Time,Pouring_Duration,Ladle_Open_Time,Prod_date,Ladle_Arrival_Time,Tund_Powder_Type,Tund_Powder_weight,Tund_Open_Time,Tund_Close_Time,Tund_Life,Tund_Open_Weight,Tund_Close_Weight,Tund_No,Car_No,Sample_Lost_Weight,Scale_Lost_Weight,Ladle_Tare_Weight,Total_Cast_Length,Total_Scrap_Weight,CasterId)values('" + StrCastSeq_caster2_MaxID + "','" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["slab_weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Cut_Lost_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_Count"].ToString() + "','" + dt_CastSeq.Rows[i]["Avg_Slab_Width"].ToString() + "','" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "','" + dt_CastSeq.Rows[i]["Shift_Foreman"].ToString() + "','" + dt_CastSeq.Rows[i]["Casting_Foreman"].ToString() + "','" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Open_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Close_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Yield"].ToString() + "','" + dt_CastSeq.Rows[i]["CrewId"].ToString() + "','" + dt_CastSeq.Rows[i]["Treat_Counter"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_Length"].ToString() + "','" + (dt_CastSeq.Rows[i]["Ladle_Close_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Pouring_Duration"].ToString() + "','" + (dt_CastSeq.Rows[i]["Ladle_Open_Time"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Prod_date"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Ladle_Arrival_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Powder_weight"].ToString() + "','" + (dt_CastSeq.Rows[i]["Tund_Open_Time"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Tund_Close_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Life"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Open_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Close_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Car_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Sample_Lost_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Scale_Lost_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Tare_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Total_Cast_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Total_Scrap_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatRpt_HeatReportSR(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = "SELECT s.strand_num AS strand_no,  s.avg_cast_speed,  s.temp_speed_tab_num AS copt_temp_tab_num,  s.s_speed_tab_num    AS copt_s_tab_num,  s.smn_speed_tab_num  AS copt_mns_tab_num,  s.tw_speed_tab_num   AS copt_tw_tab_num,  s.hmo_tab_num,  s.softred_tab_num AS hsa_tab_num,  s.gen_tab_num,  s.dsc_tab_num,  s.bops_tab_num AS mms_tab_num,  s.moldadj_tab_num,  s.sgrade_speed_tab_num AS copt_sgrade_tab_num,  s.mlc_tab_num,  s.cast_len_bgn       AS cast_len_begin,  s.cast_len_end       AS cast_len_end,  s.spray_plan_num     AS swplan_tab_num,  clo_tab_num          AS clo_tab_num,  s.clo_sample_tab_num AS clo_smp_tab_num,s.cast_powder_type as Mould_Pow_Type,s.cast_powder_weight as Mould_Pow_Amnt  FROM pdc_strand s,  pdc_heat_strand WHERE pdc_heat_strand.strand_steel_id = s.steel_id AND pdc_heat_strand.heat_steel_id='" + StrCastSeq_caster2_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatReportSR where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatReportSR set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Avg_Cast_Speed='" + dt_CastSeq.Rows[i]["Avg_Cast_Speed"].ToString() + "',Copt_Temp_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_Temp_Tab_Num"].ToString() + "',Copt_S_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_S_Tab_Num"].ToString() + "',Copt_MNS_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_MNS_Tab_Num"].ToString() + "',Copt_TW_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_TW_Tab_Num"].ToString() + "',HMO_Tab_Num='" + dt_CastSeq.Rows[i]["HMO_Tab_Num"].ToString() + "',HSA_Tab_Num='" + dt_CastSeq.Rows[i]["HSA_Tab_Num"].ToString() + "',Gen_Tab_Num='" + dt_CastSeq.Rows[i]["Gen_Tab_Num"].ToString() + "',DSC_Tab_Num='" + dt_CastSeq.Rows[i]["DSC_Tab_Num"].ToString() + "',MMS_Tab_Num='" + dt_CastSeq.Rows[i]["MMS_Tab_Num"].ToString() + "',MOLDADJ_Tab_Num='" + dt_CastSeq.Rows[i]["MOLDADJ_Tab_Num"].ToString() + "',Copt_SGrade_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_SGrade_Tab_Num"].ToString() + "',MLC_Tab_Num='" + dt_CastSeq.Rows[i]["MLC_Tab_Num"].ToString() + "',Cast_Len_Begin='" + dt_CastSeq.Rows[i]["Cast_Len_Begin"].ToString() + "',Cast_Len_End='" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "',SWPlan_Tab_Num='" + dt_CastSeq.Rows[i]["SWPlan_Tab_Num"].ToString() + "',CLO_Tab_Num='" + dt_CastSeq.Rows[i]["CLO_Tab_Num"].ToString() + "',CLO_Smp_Tab_Num='" + dt_CastSeq.Rows[i]["CLO_Smp_Tab_Num"].ToString() + "',Mould_Pow_Type='" + dt_CastSeq.Rows[i]["Mould_Pow_Type"].ToString() + "',Mould_Pow_Amnt='" + dt_CastSeq.Rows[i]["Mould_Pow_Amnt"].ToString() + "' where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatReportSR(SeqNo,SteelId,Strand_No,Avg_Cast_Speed,Copt_Temp_Tab_Num,Copt_S_Tab_Num,Copt_MNS_Tab_Num,Copt_TW_Tab_Num,HMO_Tab_Num,HSA_Tab_Num,Gen_Tab_Num,DSC_Tab_Num,MMS_Tab_Num,MOLDADJ_Tab_Num,Copt_SGrade_Tab_Num,MLC_Tab_Num,Cast_Len_Begin,Cast_Len_End,SWPlan_Tab_Num,CLO_Tab_Num,CLO_Smp_Tab_Num,Mould_Pow_Type,Mould_Pow_Amnt)values('" + StrCastSeq_caster2_MaxID + "','" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Avg_Cast_Speed"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_Temp_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_S_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_MNS_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_TW_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["HMO_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["HSA_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Gen_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["DSC_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["MMS_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["MOLDADJ_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_SGrade_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["MLC_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_Begin"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "','" + dt_CastSeq.Rows[i]["SWPlan_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["CLO_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["CLO_Smp_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Mould_Pow_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Mould_Pow_Amnt"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatRpt_SubMoldStrand(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = "SELECT s.strand_num    AS strand_no,  s.cast_len_bgn       AS cast_start_len,  s.cast_len_end       AS cast_end_len,  s.heat_in_mold_time  AS cast_start_time,  s.heat_out_mold_time AS cast_end_time, nvl(s.mold_no,0) as mold_no FROM pdc_strand s,  pdc_heat_strand hs WHERE s.steel_id     = hs.strand_steel_id AND hs.heat_steel_id='" + StrCastSeq_caster2_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_SubMoldStrand where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_SubMoldStrand set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Cast_Start_Len='" + dt_CastSeq.Rows[i]["Cast_Start_Len"].ToString() + "',Cast_End_Len='" + dt_CastSeq.Rows[i]["Cast_End_Len"].ToString() + "',Cast_Start_Time='" + (dt_CastSeq.Rows[i]["Cast_Start_Time"].ToString()).Replace(",", ".").ToString() + "',Cast_End_Time='" + (dt_CastSeq.Rows[i]["Cast_End_Time"].ToString()).Replace(",", ".").ToString() + "',Mold_No='" + dt_CastSeq.Rows[i]["Mold_No"].ToString() + "' where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_SubMoldStrand(SeqNo,SteelId,Strand_No,Cast_Start_Len,Cast_End_Len,Cast_Start_Time,Cast_End_Time,Mold_No)values('" + StrCastSeq_caster2_MaxID + "','" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Start_Len"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_End_Len"].ToString() + "','" + (dt_CastSeq.Rows[i]["Cast_Start_Time"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Cast_End_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Mold_No"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatRpt_SubPracticeTable(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = "SELECT s.strand_num  AS strand_no,  s.gen_tab_num  AS gen,  s.moldadj_tab_num   AS mold,  s.hmo_tab_num          AS hmo,  s.softred_tab_num      AS hsa,  s.mlc_tab_num          AS mlc,  s.spray_plan_num       AS cooling,  s.bops_tab_num         AS mms,  s.sgrade_speed_tab_num AS copt_speed,  s.temp_speed_tab_num   AS copt_temp,  s.tw_speed_tab_num     AS copt_tw,  s.s_speed_tab_num      AS copt_s,  s.smn_speed_tab_num    AS copt_mns,  s.clo_sample_tab_num   AS clo_smp,  clo_tab_num            AS clo,  s.dsc_tab_num          AS dsc,  0                      AS qual FROM pdc_strand s,  pdc_heat_strand WHERE pdc_heat_strand.strand_steel_id = s.steel_id AND pdc_heat_strand.heat_steel_id='" + StrCastSeq_caster2_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_SubPracticeTable where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_SubPracticeTable set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Gen='" + dt_CastSeq.Rows[i]["Gen"].ToString() + "',Mold='" + dt_CastSeq.Rows[i]["Mold"].ToString() + "',HMO='" + dt_CastSeq.Rows[i]["HMO"].ToString() + "',HSA='" + dt_CastSeq.Rows[i]["HSA"].ToString() + "',MLC='" + dt_CastSeq.Rows[i]["MLC"].ToString() + "',Cooling='" + dt_CastSeq.Rows[i]["Cooling"].ToString() + "',MMS='" + dt_CastSeq.Rows[i]["MMS"].ToString() + "',Copt_Speed='" + dt_CastSeq.Rows[i]["Copt_Speed"].ToString() + "',Copt_Temp='" + dt_CastSeq.Rows[i]["Copt_Temp"].ToString() + "',Copt_TW='" + dt_CastSeq.Rows[i]["Copt_TW"].ToString() + "',Copt_S='" + dt_CastSeq.Rows[i]["Copt_S"].ToString() + "',Copt_MNS='" + dt_CastSeq.Rows[i]["Copt_MNS"].ToString() + "',Clo_SMP='" + dt_CastSeq.Rows[i]["Clo_SMP"].ToString() + "',CLO='" + dt_CastSeq.Rows[i]["CLO"].ToString() + "',DSC='" + dt_CastSeq.Rows[i]["DSC"].ToString() + "',QUAL='" + dt_CastSeq.Rows[i]["QUAL"].ToString() + "' where SeqNo='" + StrCastSeq_caster2_MaxID + "' and SteelId='" + StrCastSeq_caster2_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_SubPracticeTable(SeqNo,SteelId,Strand_No,Gen,Mold,HMO,HSA,MLC,Cooling,MMS,Copt_Speed,Copt_Temp,Copt_TW,Copt_S,Copt_MNS,Clo_SMP,CLO,DSC,QUAL)values('" + StrCastSeq_caster2_MaxID + "','" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Gen"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold"].ToString() + "','" + dt_CastSeq.Rows[i]["HMO"].ToString() + "','" + dt_CastSeq.Rows[i]["HSA"].ToString() + "','" + dt_CastSeq.Rows[i]["MLC"].ToString() + "','" + dt_CastSeq.Rows[i]["Cooling"].ToString() + "','" + dt_CastSeq.Rows[i]["MMS"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_Speed"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_Temp"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_TW"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_S"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_MNS"].ToString() + "','" + dt_CastSeq.Rows[i]["Clo_SMP"].ToString() + "','" + dt_CastSeq.Rows[i]["CLO"].ToString() + "','" + dt_CastSeq.Rows[i]["DSC"].ToString() + "','" + dt_CastSeq.Rows[i]["QUAL"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatRpt_AnalysisElement(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            string StrCastSeqMaxID = "SELECT gcc_element_cat.element_name AS chem_abbr,  gcc_element_cat.display_order     AS display_order,  gcc_element_cat.element_type      AS element_type FROM gcc_element_cat ORDER BY display_order";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select * from HeatRpt_AnalysisElement where DISPLAY_ORDER='" + dt_CastSeq.Rows[i]["DISPLAY_ORDER"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_AnalysisElement set CHEM_ABBR='" + dt_CastSeq.Rows[i]["CHEM_ABBR"].ToString() + "',ELEMENT_TYPE='" + dt_CastSeq.Rows[i]["ELEMENT_TYPE"].ToString() + "' where DISPLAY_ORDER='" + dt_CastSeq.Rows[i]["DISPLAY_ORDER"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_AnalysisElement(CHEM_ABBR,DISPLAY_ORDER,ELEMENT_TYPE)values('" + dt_CastSeq.Rows[i]["CHEM_ABBR"].ToString() + "','" + dt_CastSeq.Rows[i]["DISPLAY_ORDER"].ToString() + "','" + dt_CastSeq.Rows[i]["ELEMENT_TYPE"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster2_HeatRpt_HeatRpt_HeatEquipmentData(string StrCastSeq_caster2_MaxID, string StrCastSeq_caster2_SteelID)
        {
            //Ladle_Shroud_Id,Ladle_Shroud_Life,Ladle_Shroud_Count,Nozzle_Id,Nozzle_Life,Nozzle_Count,Tundish_Id,Tundish_Life,Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxID = "SELECT MAX(eq.customer_id)       AS ladle_shroud_id,  TRUNC(MAX(co.counter_value),2) AS ladle_shroud_life,  COUNT(eq.customer_id)-1        AS ladle_shroud_count FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster2_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'SHROUD' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'HEAT'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Ladle_Shroud_Id='" + dt_CastSeq.Rows[i]["Ladle_Shroud_Id"].ToString() + "',Ladle_Shroud_Life='" + dt_CastSeq.Rows[i]["Ladle_Shroud_Life"].ToString() + "',Ladle_Shroud_Count='" + dt_CastSeq.Rows[i]["Ladle_Shroud_Count"].ToString() + "' where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Ladle_Shroud_Id,Ladle_Shroud_Life,Ladle_Shroud_Count)values('" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["Ladle_Shroud_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Shroud_Life"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Shroud_Count"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }

            //Nozzle_Id,Nozzle_Life,Nozzle_Count,Tundish_Id,Tundish_Life,Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDNozzle = "SELECT MAX(eq.customer_id)       AS nozzle_id,  TRUNC(MAX(co.counter_value),2) AS nozzle_life,  COUNT(eq.customer_id)-1        AS nozzle_count FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster2_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'NOZZLE' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'HEAT'";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxIDNozzle);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Nozzle_Id='" + dt_CastSeq.Rows[i]["Nozzle_Id"].ToString() + "',Nozzle_Life='" + dt_CastSeq.Rows[i]["Nozzle_Life"].ToString() + "',Nozzle_Count='" + dt_CastSeq.Rows[i]["Nozzle_Count"].ToString() + "' where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Nozzle_Id,Nozzle_Life,Nozzle_Count)values('" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["Nozzle_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Nozzle_Life"].ToString() + "','" + dt_CastSeq.Rows[i]["Nozzle_Count"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
            //Tundish_Id,Tundish_Life,Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDTundish = "SELECT eq.customer_id       AS tundish_id,  TRUNC(co.counter_value,2) AS tundish_life FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster2_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'TUNDISH' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'TON' ";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxIDTundish);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set tundish_id='" + dt_CastSeq.Rows[i]["tundish_id"].ToString() + "',Tundish_Life='" + dt_CastSeq.Rows[i]["Tundish_Life"].ToString() + "' where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Tundish_Id,Tundish_Life)values('" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["Tundish_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Tundish_Life"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
            //Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDMold = "SELECT eq.customer_id       AS mold_id,  TRUNC(co.counter_value,2) AS mold_life FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster2_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'MOLD' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'HEAT' ";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxIDMold);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Mold_Id='" + dt_CastSeq.Rows[i]["Mold_Id"].ToString() + "',Mold_Life='" + dt_CastSeq.Rows[i]["Mold_Life"].ToString() + "' where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Mold_Id,Mold_Life)values('" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["Mold_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold_Life"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }

            //Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDStnd = "SELECT strand_no,  cast_length,  event_type,  width_toc,  width_boc,  taper_left,  taper_right FROM pdc_width_change WHERE heat_steel_id = '" + StrCastSeq_caster2_SteelID + "' AND event_type     IN (1,3) ";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERII_Table(StrCastSeqMaxIDStnd);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Cast_Length='" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "',Event_Type='" + dt_CastSeq.Rows[i]["Event_Type"].ToString() + "',Width_TOC='" + dt_CastSeq.Rows[i]["Width_TOC"].ToString() + "',Width_BOC='" + dt_CastSeq.Rows[i]["Width_BOC"].ToString() + "',Taper_Left='" + dt_CastSeq.Rows[i]["Taper_Left"].ToString() + "',Taper_Right='" + dt_CastSeq.Rows[i]["Taper_Right"].ToString() + "' where SteelId='" + StrCastSeq_caster2_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right)values('" + StrCastSeq_caster2_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Event_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_TOC"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_BOC"].ToString() + "','" + dt_CastSeq.Rows[i]["Taper_Left"].ToString() + "','" + dt_CastSeq.Rows[i]["Taper_Right"].ToString() + "')";


                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }

        private void Fn_Caster3_HeatRpt_HeatReportHead(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = "SELECT STEEL_ID AS STEELID,SEQ_NO  AS SeqNo,HEATID  AS HEATID,HEAT_SEQ_NO AS InSeq,SEQUENCE_STEEL_ID AS Heat_Seq_No,SEQUENCE_STEEL_ID AS SEQ_STEEL_ID,PROD_ORDERID      AS PLAN_NO,GRADE_CODE        AS STEEL_GRADE,LADLE_OPEN_TIME   AS CAST_START_TIME,HEAT_STATUS_DESC  AS STATUS,HEAT_STATUS_CODE FROM V_HEAT_DATA WHERE (1=1)AND SEQ_NO='" + StrCastSeq_caster3_MaxID + "' AND HEAT_STATUS_CODE>2 ORDER BY LADLE_OPEN_TIME DESC";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    string StrSTEELID = "";
                    string StrSEQUENCENO = "";
                    string StrCASTSTART = "";
                    StrSTEELID = dt_CastSeq.Rows[i]["STEELID"].ToString();
                    StrSEQUENCENO = dt_CastSeq.Rows[i]["SeqNo"].ToString();
                    if (dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != null && dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString() != string.Empty)
                    {
                        StrCASTSTART = Convert.ToDateTime(dt_CastSeq.Rows[i]["CAST_START_TIME"].ToString().Replace(",", ".")).ToString();
                    }
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SeqNo from HeatReportHead where SeqNo='" + StrSEQUENCENO + "' and SteelId='" + StrSTEELID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatReportHead set HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "',InSeq='" + dt_CastSeq.Rows[i]["InSeq"].ToString() + "',Heat_Seq_No='" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "',Seq_Steel_Id='" + dt_CastSeq.Rows[i]["Seq_Steel_Id"].ToString() + "',Plan_No='" + dt_CastSeq.Rows[i]["Plan_No"].ToString() + "',Steel_Grade='" + dt_CastSeq.Rows[i]["Steel_Grade"].ToString() + "',Cast_Strat_Time='" + StrCASTSTART + "',Status='" + dt_CastSeq.Rows[i]["Status"].ToString() + "',Heat_Status_Code='" + dt_CastSeq.Rows[i]["Heat_Status_Code"].ToString() + "' where SteelId='" + dt_CastSeq.Rows[i]["SteelId"].ToString() + "' and SeqNo='" + dt_CastSeq.Rows[i]["SeqNo"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatReportHead(SteelId,SeqNo,HeatId,InSeq,Heat_Seq_No,Seq_Steel_Id,Plan_No,Steel_Grade,Cast_Strat_Time,Status,Heat_Status_Code)values('" + StrSTEELID + "','" + StrSEQUENCENO + "','" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "','" + dt_CastSeq.Rows[i]["InSeq"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Seq_Steel_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Plan_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Steel_Grade"].ToString() + "','" + StrCASTSTART + "','" + dt_CastSeq.Rows[i]["Status"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Status_Code"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);

                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatReportLabAnalysis(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = "SELECT se.grade_code,gs.grade_code,gs.CALC_LIQ_TEMP aimliq,nvl(SUM (DECODE (se.element_name, 'C', min_conc)),0) minc,  nvl(SUM (DECODE (se.element_name, 'Si', min_conc)),0) minsi, nvl(SUM (DECODE (se.element_name, 'Mn', min_conc)),0) minmn, nvl(SUM (DECODE (se.element_name, 'P', min_conc)),0) minp, nvl(SUM (DECODE (se.element_name, 'S', min_conc)),0) mins, nvl(SUM (DECODE (se.element_name, 'Al', min_conc)),0) minal, nvl(SUM (DECODE (se.element_name, 'Al_S', min_conc)),0) minals, nvl(SUM (DECODE (se.element_name, 'Cu', min_conc)),0) mincu, nvl(SUM (DECODE (se.element_name, 'Cr', min_conc)),0) mincr, nvl(SUM (DECODE (se.element_name, 'Mo', min_conc)),0) minmo, nvl(SUM (DECODE (se.element_name, 'Ni', min_conc)),0) minni, nvl(SUM (DECODE (se.element_name, 'V', min_conc)),0) minv, nvl(SUM (DECODE (se.element_name, 'Ti', min_conc)),0) minti, nvl(SUM (DECODE (se.element_name, 'Nb', min_conc)),0) minnb, nvl(SUM (DECODE (se.element_name, 'Ca', min_conc)),0) minca,nvl(SUM (DECODE (se.element_name, 'W', min_conc)),0) minw, nvl(SUM (DECODE (se.element_name, 'Sn', min_conc)),0) minsn, nvl(SUM (DECODE (se.element_name, 'As', min_conc)),0) minas, nvl(SUM (DECODE (se.element_name, 'Te', min_conc)),0) minte, nvl(SUM (DECODE (se.element_name, 'Bi', min_conc)),0) minbi, nvl(SUM (DECODE (se.element_name, 'B', min_conc)),0) minb, nvl(SUM (DECODE (se.element_name, 'Pb', min_conc)),0) minpb, nvl(SUM (DECODE (se.element_name, 'Mg', min_conc)),0) minmg, nvl(SUM (DECODE (se.element_name, 'N', min_conc)),0) minn,nvl(SUM (DECODE (se.element_name, 'Ve', min_conc)),0) minve, nvl(SUM (DECODE (se.element_name, 'Co', min_conc)),0) minco, nvl(SUM (DECODE (se.element_name, 'Ce', min_conc)),0) mince, nvl(SUM (DECODE (se.element_name, 'Sb', min_conc)),0) minsb, nvl(SUM (DECODE (se.element_name, 'Zr', min_conc)),0) minzr, nvl(SUM (DECODE (se.element_name, 'O', min_conc)),0) mino, nvl(SUM (DECODE (se.element_name, 'H', min_conc)),0) minh, nvl(SUM (DECODE (se.element_name, 'C', max_conc)),0) maxc, nvl(SUM (DECODE (se.element_name, 'Si', max_conc)),0) maxsi, nvl(SUM (DECODE (se.element_name, 'Mn', max_conc)),0) maxmn, nvl(SUM (DECODE (se.element_name, 'P', max_conc)),0) maxp, nvl(SUM (DECODE (se.element_name, 'S', max_conc)),0) maxs, nvl(SUM (DECODE (se.element_name, 'Al', max_conc)),0) maxal, nvl(SUM (DECODE (se.element_name, 'Al_S', max_conc)),0) maxals, nvl(SUM (DECODE (se.element_name, 'Cu', max_conc)),0) maxcu, nvl(SUM (DECODE (se.element_name, 'Cr', max_conc)),0) maxcr, nvl(SUM (DECODE (se.element_name, 'Mo', max_conc)),0) maxmo, nvl(SUM (DECODE (se.element_name, 'Ni', max_conc)),0) maxni, nvl(SUM (DECODE (se.element_name, 'V', max_conc)),0) maxv, nvl(SUM (DECODE (se.element_name, 'Ti', max_conc)),0) maxti, nvl(SUM (DECODE (se.element_name, 'Nb', max_conc)),0) maxnb, nvl(SUM (DECODE (se.element_name, 'Ca', max_conc)),0) maxca, nvl(SUM (DECODE (se.element_name, 'W', max_conc)),0) maxw, nvl(SUM (DECODE (se.element_name, 'Sn', max_conc)),0) maxsn, nvl(SUM (DECODE (se.element_name, 'As', max_conc)),0) maxas, nvl(SUM (DECODE (se.element_name, 'Te', max_conc)),0) maxte, nvl(SUM (DECODE (se.element_name, 'Bi', max_conc)),0) maxbi, nvl(SUM (DECODE (se.element_name, 'B', max_conc)),0) maxb, nvl(SUM (DECODE (se.element_name, 'Pb', max_conc)),0) maxpb, nvl(SUM (DECODE (se.element_name, 'Mg', max_conc)),0) maxmg, nvl(SUM (DECODE (se.element_name, 'N', max_conc)),0) maxn, nvl(SUM (DECODE (se.element_name, 'Ve', max_conc)),0) maxve, nvl(SUM (DECODE (se.element_name, 'Co', max_conc)),0) maxco, nvl(SUM (DECODE (se.element_name, 'Ce', max_conc)),0) maxce, nvl(SUM (DECODE (se.element_name, 'Sb', max_conc)),0) maxsb, nvl(SUM (DECODE (se.element_name, 'Zr', max_conc)),0) maxzr, nvl(SUM (DECODE (se.element_name, 'O', max_conc)),0) maxo, nvl(SUM (DECODE (se.element_name, 'H', max_conc)),0) maxh, nvl(SUM (DECODE (se.element_name, 'C', aim_conc)),0) aimc, nvl(SUM (DECODE (se.element_name, 'Si', aim_conc)),0) aimsi, nvl(SUM (DECODE (se.element_name, 'Mn', aim_conc)),0) aimmn, nvl(SUM (DECODE (se.element_name, 'P', aim_conc)),0) aimp, nvl(SUM (DECODE (se.element_name, 'S', aim_conc)),0) aims, nvl(SUM (DECODE (se.element_name, 'Al', aim_conc)),0) aimal, nvl(SUM (DECODE (se.element_name, 'Al_S', aim_conc)),0) aimals, nvl(SUM (DECODE (se.element_name, 'Cu', aim_conc)),0) aimcu, nvl(SUM (DECODE (se.element_name, 'Cr', aim_conc)),0) aimcr, nvl(SUM (DECODE (se.element_name, 'Mo', aim_conc)),0) aimmo, nvl(SUM (DECODE (se.element_name, 'Ni', aim_conc)),0) aimni, nvl(SUM (DECODE (se.element_name, 'V', aim_conc)),0) aimv, nvl(SUM (DECODE (se.element_name, 'Ti', aim_conc)),0) aimti, nvl(SUM (DECODE (se.element_name, 'Nb', aim_conc)),0) aimnb,nvl(SUM (DECODE (se.element_name, 'Ca', aim_conc)),0) aimca, nvl(SUM (DECODE (se.element_name, 'W', aim_conc)),0) aimw, nvl(SUM (DECODE (se.element_name, 'Sn', aim_conc)),0) aimsn, nvl(SUM (DECODE (se.element_name, 'As', aim_conc)),0) aimas,nvl(SUM (DECODE (se.element_name, 'Te', aim_conc)),0) aimte, nvl(SUM (DECODE (se.element_name, 'Bi', aim_conc)),0) aimbi,nvl(SUM (DECODE (se.element_name, 'B', aim_conc)),0) aimb, nvl(SUM (DECODE (se.element_name, 'Pb', aim_conc)),0) aimpb, nvl(SUM (DECODE (se.element_name, 'Mg', aim_conc)),0) aimmg, nvl(SUM (DECODE (se.element_name, 'N', aim_conc)),0) aimn, nvl(SUM (DECODE (se.element_name, 'Ve', aim_conc)),0) aimve, nvl(SUM (DECODE (se.element_name, 'Co', aim_conc)),0) aimco, nvl(SUM (DECODE (se.element_name, 'Ce', aim_conc)),0) aimce, nvl(SUM (DECODE (se.element_name, 'Sb', aim_conc)),0) aimsb, nvl(SUM (DECODE (se.element_name, 'Zr', aim_conc)),0) aimzr, nvl(SUM (DECODE (se.element_name, 'O', aim_conc)),0) aimo, nvl(SUM (DECODE (se.element_name, 'H', aim_conc)),0) aimh FROM gcc_spec_analysis se, gcc_grade_spec gs WHERE ((gs.grade_code = se.grade_code)AND (se.grade_code    =  (SELECT grade_code FROM pdc_heat WHERE (steel_id = '" + StrCastSeq_caster3_SteelID + "')  ))) GROUP BY (se.grade_code, gs.grade_code,gs.CALC_LIQ_TEMP)";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatReportLabAnalysis where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatReportLabAnalysis set LIQ_min='" + dt_CastSeq.Rows[i]["aimliq"].ToString() + "',C_min='" + dt_CastSeq.Rows[i]["minc"].ToString() + "',Si_min='" + dt_CastSeq.Rows[i]["minSi"].ToString() + "',Mn_min='" + dt_CastSeq.Rows[i]["minMn"].ToString() + "',P_min='" + dt_CastSeq.Rows[i]["minP"].ToString() + "',S_min='" + dt_CastSeq.Rows[i]["minS"].ToString() + "',Al_min='" + dt_CastSeq.Rows[i]["minAl"].ToString() + "',Al_S_min='" + dt_CastSeq.Rows[i]["minAls"].ToString() + "',Cu_min='" + dt_CastSeq.Rows[i]["minCu"].ToString() + "',Cr_min='" + dt_CastSeq.Rows[i]["minCr"].ToString() + "',Mo_min='" + dt_CastSeq.Rows[i]["minMo"].ToString() + "',Ni_min='" + dt_CastSeq.Rows[i]["minNi"].ToString() + "',V_min='" + dt_CastSeq.Rows[i]["minV"].ToString() + "',Ti_min='" + dt_CastSeq.Rows[i]["minTi"].ToString() + "',Nb_min='" + dt_CastSeq.Rows[i]["minNb"].ToString() + "',Ca_min='" + dt_CastSeq.Rows[i]["MinCa"].ToString() + "',W_min='" + dt_CastSeq.Rows[i]["minW"].ToString() + "',Sn_min='" + dt_CastSeq.Rows[i]["minSn"].ToString() + "',As_min='" + dt_CastSeq.Rows[i]["minAs"].ToString() + "',Te_min='" + dt_CastSeq.Rows[i]["minTe"].ToString() + "',Bi_min='" + dt_CastSeq.Rows[i]["minBi"].ToString() + "',B_min='" + dt_CastSeq.Rows[i]["minB"].ToString() + "',Pb_min='" + dt_CastSeq.Rows[i]["minPb"].ToString() + "',Mg_min='" + dt_CastSeq.Rows[i]["minMg"].ToString() + "',N_min='" + dt_CastSeq.Rows[i]["minN"].ToString() + "',Ve_min='" + dt_CastSeq.Rows[i]["minVe"].ToString() + "',Co_min='" + dt_CastSeq.Rows[i]["minCo"].ToString() + "',Ce_min='" + dt_CastSeq.Rows[i]["minCe"].ToString() + "',Sb_min='" + dt_CastSeq.Rows[i]["minSb"].ToString() + "',Zr_min='" + dt_CastSeq.Rows[i]["minZr"].ToString() + "',O_min='" + dt_CastSeq.Rows[i]["minO"].ToString() + "',H_min='" + dt_CastSeq.Rows[i]["minH"].ToString() + "',C_max='" + dt_CastSeq.Rows[i]["maxC"].ToString() + "',Si_max='" + dt_CastSeq.Rows[i]["maxSi"].ToString() + "',Mn_max='" + dt_CastSeq.Rows[i]["maxMn"].ToString() + "',P_max='" + dt_CastSeq.Rows[i]["maxP"].ToString() + "',S_max='" + dt_CastSeq.Rows[i]["maxS"].ToString() + "',Al_max='" + dt_CastSeq.Rows[i]["maxAl"].ToString() + "',Al_S_max='" + dt_CastSeq.Rows[i]["maxAlS"].ToString() + "',Cu_max='" + dt_CastSeq.Rows[i]["maxCu"].ToString() + "',Cr_max='" + dt_CastSeq.Rows[i]["maxCr"].ToString() + "',Mo_max='" + dt_CastSeq.Rows[i]["maxMo"].ToString() + "',Ni_max='" + dt_CastSeq.Rows[i]["maxNi"].ToString() + "',V_max='" + dt_CastSeq.Rows[i]["maxV"].ToString() + "',Ti_max='" + dt_CastSeq.Rows[i]["maxTi"].ToString() + "',Nb_max='" + dt_CastSeq.Rows[i]["maxNb"].ToString() + "',Ca_max='" + dt_CastSeq.Rows[i]["maxCa"].ToString() + "',W_max='" + dt_CastSeq.Rows[i]["maxW"].ToString() + "',Sn_max='" + dt_CastSeq.Rows[i]["maxSn"].ToString() + "',As_max='" + dt_CastSeq.Rows[i]["maxAs"].ToString() + "',Te_max='" + dt_CastSeq.Rows[i]["maxTe"].ToString() + "',Bi_max='" + dt_CastSeq.Rows[i]["maxBi"].ToString() + "',B_max='" + dt_CastSeq.Rows[i]["maxB"].ToString() + "',Pb_max='" + dt_CastSeq.Rows[i]["maxPb"].ToString() + "',Mg_max='" + dt_CastSeq.Rows[i]["maxMg"].ToString() + "',N_max='" + dt_CastSeq.Rows[i]["maxN"].ToString() + "',Ve_max='" + dt_CastSeq.Rows[i]["maxVe"].ToString() + "',Co_max='" + dt_CastSeq.Rows[i]["maxCo"].ToString() + "',Ce_max='" + dt_CastSeq.Rows[i]["maxCe"].ToString() + "',Sb_max='" + dt_CastSeq.Rows[i]["maxSb"].ToString() + "',Zr_max='" + dt_CastSeq.Rows[i]["maxZr"].ToString() + "',O_max='" + dt_CastSeq.Rows[i]["maxO"].ToString() + "',H_max='" + dt_CastSeq.Rows[i]["maxH"].ToString() + "',C_aim='" + dt_CastSeq.Rows[i]["aimC"].ToString() + "',Si_aim='" + dt_CastSeq.Rows[i]["aimSi"].ToString() + "',Mn_aim='" + dt_CastSeq.Rows[i]["aimMn"].ToString() + "',P_aim='" + dt_CastSeq.Rows[i]["aimP"].ToString() + "',S_aim='" + dt_CastSeq.Rows[i]["aimS"].ToString() + "',Al_aim='" + dt_CastSeq.Rows[i]["aimAl"].ToString() + "',Al_S_aim='" + dt_CastSeq.Rows[i]["aimAlS"].ToString() + "',Cu_aim='" + dt_CastSeq.Rows[i]["aimCu"].ToString() + "',Cr_aim='" + dt_CastSeq.Rows[i]["aimCr"].ToString() + "',Mo_aim='" + dt_CastSeq.Rows[i]["aimMo"].ToString() + "',Ni_aim='" + dt_CastSeq.Rows[i]["aimNi"].ToString() + "',V_aim='" + dt_CastSeq.Rows[i]["aimV"].ToString() + "',Ti_aim='" + dt_CastSeq.Rows[i]["aimTi"].ToString() + "',Nb_aim='" + dt_CastSeq.Rows[i]["aimNb"].ToString() + "',Ca_aim='" + dt_CastSeq.Rows[i]["aimCa"].ToString() + "',W_aim='" + dt_CastSeq.Rows[i]["aimW"].ToString() + "',Sn_aim='" + dt_CastSeq.Rows[i]["aimSn"].ToString() + "',As_aim='" + dt_CastSeq.Rows[i]["aimAs"].ToString() + "',Te_aim='" + dt_CastSeq.Rows[i]["aimTe"].ToString() + "',Bi_aim='" + dt_CastSeq.Rows[i]["aimBi"].ToString() + "',B_aim='" + dt_CastSeq.Rows[i]["aimB"].ToString() + "',Pb_aim='" + dt_CastSeq.Rows[i]["aimPb"].ToString() + "',Mg_aim='" + dt_CastSeq.Rows[i]["aimMg"].ToString() + "',N_aim='" + dt_CastSeq.Rows[i]["aimN"].ToString() + "',Ve_aim='" + dt_CastSeq.Rows[i]["aimVe"].ToString() + "',Co_aim='" + dt_CastSeq.Rows[i]["aimCo"].ToString() + "',Ce_aim='" + dt_CastSeq.Rows[i]["aimCe"].ToString() + "',Sb_aim='" + dt_CastSeq.Rows[i]["aimSb"].ToString() + "',Zr_aim='" + dt_CastSeq.Rows[i]["aimZr"].ToString() + "',O_aim='" + dt_CastSeq.Rows[i]["aimO"].ToString() + "',H_aim='" + dt_CastSeq.Rows[i]["aimH"].ToString() + "' where SteelId='" + StrCastSeq_caster3_SteelID + "' and SeqNo='" + StrCastSeq_caster3_MaxID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatReportLabAnalysis(SeqNo,SteelId,LIQ_min,C_min,Si_min,Mn_min,P_min,S_min,Al_min,Al_S_min,Cu_min,Cr_min,Mo_min,Ni_min,V_min,Ti_min,Nb_min,Ca_min,W_min,Sn_min,As_min,Te_min,Bi_min,B_min,Pb_min,Mg_min,N_min,Ve_min,Co_min,Ce_min,Sb_min,Zr_min,O_min,H_min,C_max,Si_max,Mn_max,P_max,S_max,Al_max,Al_S_max,Cu_max,Cr_max,Mo_max,Ni_max,V_max,Ti_max,Nb_max,Ca_max,W_max,Sn_max,As_max,Te_max,Bi_max,B_max,Pb_max,Mg_max,N_max,Ve_max,Co_max,Ce_max,Sb_max,Zr_max,O_max,H_max,C_aim,Si_aim,Mn_aim,P_aim,S_aim,Al_aim,Al_S_aim,Cu_aim,Cr_aim,Mo_aim,Ni_aim,V_aim,Ti_aim,Nb_aim,Ca_aim,W_aim,Sn_aim,As_aim,Te_aim,Bi_aim,B_aim,Pb_aim,Mg_aim,N_aim,Ve_aim,Co_aim,Ce_aim,Sb_aim,Zr_aim,O_aim,H_aim)values('" + StrCastSeq_caster3_MaxID + "','" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["aimliq"].ToString() + "','" + dt_CastSeq.Rows[i]["minC"].ToString() + "','" + dt_CastSeq.Rows[i]["minSi"].ToString() + "','" + dt_CastSeq.Rows[i]["minMn"].ToString() + "','" + dt_CastSeq.Rows[i]["minP"].ToString() + "','" + dt_CastSeq.Rows[i]["minS"].ToString() + "','" + dt_CastSeq.Rows[i]["minAl"].ToString() + "','" + dt_CastSeq.Rows[i]["minAlS"].ToString() + "','" + dt_CastSeq.Rows[i]["minCu"].ToString() + "','" + dt_CastSeq.Rows[i]["minCr"].ToString() + "','" + dt_CastSeq.Rows[i]["minMo"].ToString() + "','" + dt_CastSeq.Rows[i]["minNi"].ToString() + "','" + dt_CastSeq.Rows[i]["minV"].ToString() + "','" + dt_CastSeq.Rows[i]["minTi"].ToString() + "','" + dt_CastSeq.Rows[i]["minNb"].ToString() + "','" + dt_CastSeq.Rows[i]["minCa"].ToString() + "','" + dt_CastSeq.Rows[i]["minW"].ToString() + "','" + dt_CastSeq.Rows[i]["minsn"].ToString() + "','" + dt_CastSeq.Rows[i]["minAs"].ToString() + "','" + dt_CastSeq.Rows[i]["minTe"].ToString() + "','" + dt_CastSeq.Rows[i]["minBi"].ToString() + "','" + dt_CastSeq.Rows[i]["minB"].ToString() + "','" + dt_CastSeq.Rows[i]["minPb"].ToString() + "','" + dt_CastSeq.Rows[i]["minMg"].ToString() + "','" + dt_CastSeq.Rows[i]["minN"].ToString() + "','" + dt_CastSeq.Rows[i]["minVe"].ToString() + "','" + dt_CastSeq.Rows[i]["minCo"].ToString() + "','" + dt_CastSeq.Rows[i]["minCe"].ToString() + "','" + dt_CastSeq.Rows[i]["minSb"].ToString() + "','" + dt_CastSeq.Rows[i]["minZr"].ToString() + "','" + dt_CastSeq.Rows[i]["minO"].ToString() + "','" + dt_CastSeq.Rows[i]["minH"].ToString() + "','" + dt_CastSeq.Rows[i]["maxC"].ToString() + "','" + dt_CastSeq.Rows[i]["maxSi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxMn"].ToString() + "','" + dt_CastSeq.Rows[i]["maxP"].ToString() + "','" + dt_CastSeq.Rows[i]["maxS"].ToString() + "','" + dt_CastSeq.Rows[i]["maxAl"].ToString() + "','" + dt_CastSeq.Rows[i]["maxAlS"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCu"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCr"].ToString() + "','" + dt_CastSeq.Rows[i]["maxMo"].ToString() + "','" + dt_CastSeq.Rows[i]["maxNi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxV"].ToString() + "','" + dt_CastSeq.Rows[i]["maxTi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxNb"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCa"].ToString() + "','" + dt_CastSeq.Rows[i]["maxW"].ToString() + "','" + dt_CastSeq.Rows[i]["maxSn"].ToString() + "','" + dt_CastSeq.Rows[i]["maxAs"].ToString() + "','" + dt_CastSeq.Rows[i]["maxTe"].ToString() + "','" + dt_CastSeq.Rows[i]["maxBi"].ToString() + "','" + dt_CastSeq.Rows[i]["maxB"].ToString() + "','" + dt_CastSeq.Rows[i]["maxPb"].ToString() + "','" + dt_CastSeq.Rows[i]["maxMg"].ToString() + "','" + dt_CastSeq.Rows[i]["maxN"].ToString() + "','" + dt_CastSeq.Rows[i]["maxVe"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCo"].ToString() + "','" + dt_CastSeq.Rows[i]["maxCe"].ToString() + "','" + dt_CastSeq.Rows[i]["maxSb"].ToString() + "','" + dt_CastSeq.Rows[i]["maxZr"].ToString() + "','" + dt_CastSeq.Rows[i]["maxO"].ToString() + "','" + dt_CastSeq.Rows[i]["maxH"].ToString() + "','" + dt_CastSeq.Rows[i]["aimC"].ToString() + "','" + dt_CastSeq.Rows[i]["aimSi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimMn"].ToString() + "','" + dt_CastSeq.Rows[i]["aimP"].ToString() + "','" + dt_CastSeq.Rows[i]["aimS"].ToString() + "','" + dt_CastSeq.Rows[i]["aimAl"].ToString() + "','" + dt_CastSeq.Rows[i]["aimAlS"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCu"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCr"].ToString() + "','" + dt_CastSeq.Rows[i]["aimMo"].ToString() + "','" + dt_CastSeq.Rows[i]["aimNi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimV"].ToString() + "','" + dt_CastSeq.Rows[i]["aimTi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimNb"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCa"].ToString() + "','" + dt_CastSeq.Rows[i]["aimW"].ToString() + "','" + dt_CastSeq.Rows[i]["aimSn"].ToString() + "','" + dt_CastSeq.Rows[i]["aimAs"].ToString() + "','" + dt_CastSeq.Rows[i]["aimTe"].ToString() + "','" + dt_CastSeq.Rows[i]["aimBi"].ToString() + "','" + dt_CastSeq.Rows[i]["aimB"].ToString() + "','" + dt_CastSeq.Rows[i]["aimPb"].ToString() + "','" + dt_CastSeq.Rows[i]["aimMg"].ToString() + "','" + dt_CastSeq.Rows[i]["aimN"].ToString() + "','" + dt_CastSeq.Rows[i]["aimVe"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCo"].ToString() + "','" + dt_CastSeq.Rows[i]["aimCe"].ToString() + "','" + dt_CastSeq.Rows[i]["aimSb"].ToString() + "','" + dt_CastSeq.Rows[i]["aimZr"].ToString() + "','" + dt_CastSeq.Rows[i]["aimO"].ToString() + "','" + dt_CastSeq.Rows[i]["aimH"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatReportTempReport(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = "SELECT (pdc_tundish_temp.meas_temp - gcc_grade_spec.calc_liq_temp) AS overheat,pdc_tundish_temp.meas_temp as Temp,pdc_tundish_temp.meas_time as Time FROM gcc_grade_spec,pdc_heat,pdc_tundish_temp WHERE pdc_heat.grade_code = gcc_grade_spec.grade_code AND pdc_heat.steel_id     = pdc_tundish_temp.steel_id AND pdc_heat.steel_id     = '" + StrCastSeq_caster3_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatReportTempReport where SteelId='" + StrCastSeq_caster3_SteelID + "' and Time='" + (dt_CastSeq.Rows[i]["Time"].ToString().Replace(",", ".")).ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatReportTempReport set Temp='" + dt_CastSeq.Rows[i]["Temp"].ToString() + "',Temp='" + dt_CastSeq.Rows[i]["OverHeat"].ToString() + "' where SteelId='" + StrCastSeq_caster3_SteelID + "' and Time='" + (dt_CastSeq.Rows[i]["Time"].ToString().Replace(",", ".")).ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatReportTempReport(SeqNo,SteelId,Time,Temp,OverHeat)values('" + StrCastSeq_caster3_MaxID + "','" + StrCastSeq_caster3_SteelID + "','" + (dt_CastSeq.Rows[i]["Time"].ToString().Replace(",", ".")).ToString() + "','" + dt_CastSeq.Rows[i]["Temp"].ToString() + "','" + dt_CastSeq.Rows[i]["OverHeat"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatRpt_PlandataReport(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = "  SELECT pdc_slab.l2_slabid,nvl(pdc_slab_order.length_min,0) as length_min,nvl(pdc_slab_order.length_aim,0) as length_aim, nvl(pdc_slab_order.length_max,0) as length_max, nvl(pdc_slab_order.width_head,0) as width_head, nvl(pdc_slab_order.width_tail,0) as width_tail, nvl(pdc_piece.heat_steel_id,0) as heat_steel_id, nvl(pdc_slab.slab_seq_no,0) as slab_seq_no, nvl(pdc_piece.piece_seq_no,0) as piece_seq_no FROM l2ccs.pdc_piece pdc_piece,l2ccs.pdc_slab_order pdc_slab_order,l2ccs.pdc_slab pdc_slab  WHERE (pdc_slab_order.steel_id(+) = pdc_slab.steel_id) AND (pdc_piece.steel_id           = pdc_slab.steel_id(+)) AND (pdc_piece.heat_steel_id      = '" + StrCastSeq_caster3_SteelID + "')";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_PlandataReport where SteelId='" + StrCastSeq_caster3_SteelID + "' and L2_SlabId='" + dt_CastSeq.Rows[i]["l2_slabid"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_PlandataReport set L2_SlabId='" + dt_CastSeq.Rows[i]["L2_SlabId"].ToString() + "',Length_min='" + dt_CastSeq.Rows[i]["Length_min"].ToString() + "',Length_aim='" + dt_CastSeq.Rows[i]["Length_aim"].ToString() + "',Length_max='" + dt_CastSeq.Rows[i]["Length_max"].ToString() + "',Width_Head='" + dt_CastSeq.Rows[i]["Width_Head"].ToString() + "',Width_Tail='" + dt_CastSeq.Rows[i]["Width_Tail"].ToString() + "',Slab_Seq_No='" + dt_CastSeq.Rows[i]["Slab_Seq_No"].ToString() + "',Piece_Seq_No='" + dt_CastSeq.Rows[i]["Piece_Seq_No"].ToString() + "' where SteelId='" + StrCastSeq_caster3_SteelID + "' and SeqNo='" + StrCastSeq_caster3_MaxID + "'  and L2_SlabId='" + dt_CastSeq.Rows[i]["l2_slabid"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_PlandataReport(SeqNo,SteelId,L2_SlabId,Length_min,Length_aim,Length_max,Width_Head,Width_Tail,Slab_Seq_No,Piece_Seq_No)values('" + StrCastSeq_caster3_MaxID + "','" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["L2_SlabId"].ToString() + "','" + dt_CastSeq.Rows[i]["Length_min"].ToString() + "','" + dt_CastSeq.Rows[i]["Length_aim"].ToString() + "','" + dt_CastSeq.Rows[i]["Length_max"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_Head"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_Tail"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_Seq_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Piece_Seq_No"].ToString() + "')";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatRptAnalysisReport(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = "SELECT a.heatid heatid,  a.sample_no,  a.taken_time,  a.sample_loc,  a.LIQUID_TEMP LIQ,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'C'  AND ae.analysis_id = a.analysis_id  ),0) c,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Si'  AND ae.analysis_id = a.analysis_id  ) ,0)si,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Mn'  AND ae.analysis_id = a.analysis_id  ),0) mn,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'P'  AND ae.analysis_id = a.analysis_id  ),0) p,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'S'  AND ae.analysis_id = a.analysis_id  ),0) s,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Al'  AND ae.analysis_id = a.analysis_id  ),0) al,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Als'  AND ae.analysis_id = a.analysis_id  ),0) als,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Cu'  AND ae.analysis_id = a.analysis_id  ),0) cu,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Cr'  AND ae.analysis_id = a.analysis_id  ),0) cr,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Mo'  AND ae.analysis_id = a.analysis_id  ),0) mo,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ni'  AND ae.analysis_id = a.analysis_id  ),0) ni,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'V'  AND ae.analysis_id = a.analysis_id  ),0) v,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ti'  AND ae.analysis_id = a.analysis_id  ),0) ti,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Nb'  AND ae.analysis_id = a.analysis_id  ),0) nb,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ca'  AND ae.analysis_id = a.analysis_id  ),0) ca,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'W'  AND ae.analysis_id = a.analysis_id  ),0) w,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Sn'  AND ae.analysis_id = a.analysis_id  ),0) sn,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'As'  AND ae.analysis_id = a.analysis_id  ),0) AS a_s,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Te'  AND ae.analysis_id = a.analysis_id  ) ,0)te,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Bi'  AND ae.analysis_id = a.analysis_id  ) ,0)bi,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'B'  AND ae.analysis_id = a.analysis_id  ),0) b,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Pb'  AND ae.analysis_id = a.analysis_id  ),0) Pb,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Mg'  AND ae.analysis_id = a.analysis_id  ),0) Mg,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'N'  AND ae.analysis_id = a.analysis_id  ),0) N,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ve'  AND ae.analysis_id = a.analysis_id  ),0) Ve,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Co'  AND ae.analysis_id = a.analysis_id  ),0) Co,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Ce'  AND ae.analysis_id = a.analysis_id  ),0) Ce,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Sb'  AND ae.analysis_id = a.analysis_id  ),0) Sb,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'Zr'  AND ae.analysis_id = a.analysis_id  ),0) Zr,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'O'  AND ae.analysis_id = a.analysis_id  ),0) O,  nvl((SELECT MAX (actual_value)  FROM pd_analysis_entry ae  WHERE element_name = 'H'  AND ae.analysis_id = a.analysis_id  ),0) H FROM pd_analysis a,  pdc_heat h WHERE (a.heatid = h.heatid) AND (h.steel_id ='" + StrCastSeq_caster3_SteelID + "')";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRptAnalysisReport where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "' and heatid='" + dt_CastSeq.Rows[i]["heatid"].ToString() + "' and Taken_Time='" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRptAnalysisReport set Sample_No='" + dt_CastSeq.Rows[i]["Sample_No"].ToString() + "',Taken_Time='" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "',Sample_LOC='" + dt_CastSeq.Rows[i]["Sample_LOC"].ToString() + "',LIQ='" + dt_CastSeq.Rows[i]["LIQ"].ToString() + "',C='" + dt_CastSeq.Rows[i]["C"].ToString() + "',Si='" + dt_CastSeq.Rows[i]["Si"].ToString() + "',Mn='" + dt_CastSeq.Rows[i]["Mn"].ToString() + "',P='" + dt_CastSeq.Rows[i]["P"].ToString() + "',S='" + dt_CastSeq.Rows[i]["S"].ToString() + "',Al='" + dt_CastSeq.Rows[i]["Al"].ToString() + "',AlS='" + dt_CastSeq.Rows[i]["AlS"].ToString() + "',Cu='" + dt_CastSeq.Rows[i]["Cu"].ToString() + "',Cr='" + dt_CastSeq.Rows[i]["Cr"].ToString() + "',Mo='" + dt_CastSeq.Rows[i]["Mo"].ToString() + "',Ni='" + dt_CastSeq.Rows[i]["Ni"].ToString() + "',V='" + dt_CastSeq.Rows[i]["V"].ToString() + "',Ti='" + dt_CastSeq.Rows[i]["Ti"].ToString() + "',NB='" + dt_CastSeq.Rows[i]["Nb"].ToString() + "',Ca='" + dt_CastSeq.Rows[i]["Ca"].ToString() + "',W='" + dt_CastSeq.Rows[i]["W"].ToString() + "',Sn='" + dt_CastSeq.Rows[i]["Sn"].ToString() + "',A_S='" + dt_CastSeq.Rows[i]["A_S"].ToString() + "',TE='" + dt_CastSeq.Rows[i]["TE"].ToString() + "',Bi='" + dt_CastSeq.Rows[i]["Bi"].ToString() + "',B='" + dt_CastSeq.Rows[i]["B"].ToString() + "',Pb='" + dt_CastSeq.Rows[i]["Pb"].ToString() + "',Mg='" + dt_CastSeq.Rows[i]["Mg"].ToString() + "',N='" + dt_CastSeq.Rows[i]["N"].ToString() + "',Ve='" + dt_CastSeq.Rows[i]["Ve"].ToString() + "',Co='" + dt_CastSeq.Rows[i]["Co"].ToString() + "',Ce='" + dt_CastSeq.Rows[i]["Ce"].ToString() + "',Sb='" + dt_CastSeq.Rows[i]["Sb"].ToString() + "',Zr='" + dt_CastSeq.Rows[i]["Zr"].ToString() + "',O='" + dt_CastSeq.Rows[i]["O"].ToString() + "',H='" + dt_CastSeq.Rows[i]["H"].ToString() + "' where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "' and heatid='" + dt_CastSeq.Rows[i]["heatid"].ToString() + "' and Taken_Time='" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRptAnalysisReport(SeqNo,SteelId,HeatId,Sample_No,Taken_Time,Sample_LOC,LIQ,C,Si,Mn,P,S,Al,AlS,Cu,Cr,Mo,Ni,V,Ti,NB,Ca,W,Sn,A_S,TE,Bi,B,Pb,Mg,N,Ve,Co,Ce,Sb,Zr,O,H)values('" + StrCastSeq_caster3_MaxID + "','" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "','" + dt_CastSeq.Rows[i]["Sample_No"].ToString() + "','" + (dt_CastSeq.Rows[i]["Taken_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Sample_LOC"].ToString() + "','" + dt_CastSeq.Rows[i]["LIQ"].ToString() + "','" + dt_CastSeq.Rows[i]["C"].ToString() + "','" + dt_CastSeq.Rows[i]["Si"].ToString() + "','" + dt_CastSeq.Rows[i]["Mn"].ToString() + "','" + dt_CastSeq.Rows[i]["P"].ToString() + "','" + dt_CastSeq.Rows[i]["S"].ToString() + "','" + dt_CastSeq.Rows[i]["Al"].ToString() + "','" + dt_CastSeq.Rows[i]["AlS"].ToString() + "','" + dt_CastSeq.Rows[i]["Cu"].ToString() + "','" + dt_CastSeq.Rows[i]["Cr"].ToString() + "','" + dt_CastSeq.Rows[i]["Mo"].ToString() + "','" + dt_CastSeq.Rows[i]["Ni"].ToString() + "','" + dt_CastSeq.Rows[i]["V"].ToString() + "','" + dt_CastSeq.Rows[i]["Ti"].ToString() + "','" + dt_CastSeq.Rows[i]["NB"].ToString() + "','" + dt_CastSeq.Rows[i]["Ca"].ToString() + "','" + dt_CastSeq.Rows[i]["W"].ToString() + "','" + dt_CastSeq.Rows[i]["Sn"].ToString() + "','" + dt_CastSeq.Rows[i]["A_S"].ToString() + "','" + dt_CastSeq.Rows[i]["TE"].ToString() + "','" + dt_CastSeq.Rows[i]["Bi"].ToString() + "','" + dt_CastSeq.Rows[i]["B"].ToString() + "','" + dt_CastSeq.Rows[i]["Pb"].ToString() + "','" + dt_CastSeq.Rows[i]["Mg"].ToString() + "','" + dt_CastSeq.Rows[i]["N"].ToString() + "','" + dt_CastSeq.Rows[i]["Ve"].ToString() + "','" + dt_CastSeq.Rows[i]["Co"].ToString() + "','" + dt_CastSeq.Rows[i]["Ce"].ToString() + "','" + dt_CastSeq.Rows[i]["Sb"].ToString() + "','" + dt_CastSeq.Rows[i]["Zr"].ToString() + "','" + dt_CastSeq.Rows[i]["O"].ToString() + "','" + dt_CastSeq.Rows[i]["H"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatRptCutDataReport(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = " SELECT p.strand_no,sl.slab_seq_no   AS slab_no,sl.l2_slabid  AS marking_no,pne.value/1000 AS cold_length,pt.piece_type_desc   AS piece_type,p.head_cast_len   AS cast_length_cut,so.length_min   AS min_length,so.length_aim   AS aim_length,so.length_max  AS max_length,(p.piece_length/1000 - NVL (scrap.scrap_length, 0)) AS good_length,p.piece_length/1000  AS act_length,(DECODE(NVL(piece_weight_meas,0), 0, piece_weight_calc , piece_weight_meas) - NVL (scrap.scrap_mass, 0)) / 1000  AS good_weight,NVL (scrap.scrap_mass, 0) / 1000  AS scrap_weight,p.head_thickness,p.head_width,p.mixzone_begin AS mixzone_begin,p.mixzone_end   AS mixzone_end,p.heat_boundary_pos1,p.sample_weight FROM pdc_piece p,pdc_slab sl,pdc_slab_order so,pdc_heat h,gdc_piece_type pt,pdc_number_entry pne,(SELECT steel_id,SUM (scrap_mass) AS scrap_mass,(SUM (scrap_end) - SUM (scrap_bgn)) AS scrap_length FROM pdc_scrap  GROUP BY steel_id  ) scrap WHERE (p.steel_id    = sl.steel_id(+)) AND (p.steel_id      = so.steel_id(+)) AND (p.steel_id      = scrap.steel_id(+)) AND (p.piece_type    = pt.piece_type(+)) AND (p.heat_steel_id = h.steel_id) AND pne.val_name     = 'ColdLen' AND pne.steel_id     = p.steel_id AND (h.steel_id = '" + StrCastSeq_caster3_SteelID + "') ORDER BY p.head_cast_len ASC";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRptCutDataReport where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "' and Marking_No='" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRptCutDataReport set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Slab_No='" + dt_CastSeq.Rows[i]["Slab_No"].ToString() + "',Marking_No='" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "',Cold_Length='" + dt_CastSeq.Rows[i]["Cold_Length"].ToString() + "',Piece_Type='" + dt_CastSeq.Rows[i]["Piece_Type"].ToString() + "',Cast_Length_Cut='" + dt_CastSeq.Rows[i]["Cast_Length_Cut"].ToString() + "',Min_Length='" + dt_CastSeq.Rows[i]["Min_Length"].ToString() + "',Aim_Length='" + dt_CastSeq.Rows[i]["Aim_Length"].ToString() + "',Max_Length='" + dt_CastSeq.Rows[i]["Max_Length"].ToString() + "',Good_Length='" + dt_CastSeq.Rows[i]["Good_Length"].ToString() + "',Act_Length='" + dt_CastSeq.Rows[i]["Act_Length"].ToString() + "',Good_Weight='" + dt_CastSeq.Rows[i]["Good_Weight"].ToString() + "',scrap_Weight='" + dt_CastSeq.Rows[i]["scrap_Weight"].ToString() + "',Head_Thickness='" + dt_CastSeq.Rows[i]["Head_Thickness"].ToString() + "',Head_Width='" + dt_CastSeq.Rows[i]["Head_Width"].ToString() + "',MixZone_Begin='" + dt_CastSeq.Rows[i]["MixZone_Begin"].ToString() + "',MixZone_End='" + dt_CastSeq.Rows[i]["MixZone_End"].ToString() + "',Heat_Boundary_POS1='" + dt_CastSeq.Rows[i]["Heat_Boundary_POS1"].ToString() + "',Sample_Weight='" + dt_CastSeq.Rows[i]["sample_weight"].ToString() + "' where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "' and Marking_No='" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRptCutDataReport(SeqNo,SteelId,Strand_No,Slab_No,Marking_No,Cold_Length,Piece_Type,Cast_Length_Cut,Min_Length,Aim_Length,Max_Length,Good_Length,Act_Length,Good_Weight,scrap_Weight,Head_Thickness,Head_Width,MixZone_Begin,MixZone_End,Heat_Boundary_POS1,Sample_Weight)values('" + StrCastSeq_caster3_MaxID + "','" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Marking_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Cold_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Piece_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Length_Cut"].ToString() + "','" + dt_CastSeq.Rows[i]["Min_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Aim_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Max_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Good_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Act_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Good_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["scrap_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Head_Thickness"].ToString() + "','" + dt_CastSeq.Rows[i]["Head_Width"].ToString() + "','" + dt_CastSeq.Rows[i]["MixZone_Begin"].ToString() + "','" + dt_CastSeq.Rows[i]["MixZone_End"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Boundary_POS1"].ToString() + "','" + dt_CastSeq.Rows[i]["sample_weight"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatRptHeader(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = "SELECT  ROUND(nvl(v.slab_weight_net,0),3) AS slab_weight,  nvl(heat.CUT_LOST_WEIGHT,0)   AS cut_lost_weight,  nvl(v.count_slabs,0)          AS slab_count, nvl(heat.avg_slab_width,0) as avg_slab_width, nvl(heat.heatid,0) as heatid,  nvl(heat.heat_seq_no,0) as heat_seq_no,  nvl(heat.grade_code,0) as grade_code, cheat.shift_foreman  AS shift_foreman,  cheat.casting_foreman AS casting_foreman,  heat.prod_orderid, nvl(heat.ladle_no,0) as ladle_no,  nvl(v.ladle_open_weight,0)  AS ladle_open_weight,  nvl(v.ladle_close_weight,0) AS ladle_close_weight,  nvl(v.total_cast_weight,0)  AS cast_weight,  nvl(seq.seq_no,0) as seq_no,  ROUND(nvl(v.yield,2),2) AS yield,  hshift.crew_id   AS crewid,   nvl(heat.treat_counter,0)  AS treat_counter,   nvl(heat.slab_length,0)  AS slab_length,   heat.ladle_close_time  AS ladle_close_time,   nvl(heat.pouring_duration,0)  AS pouring_duration,   heat.ladle_open_time AS ladle_open_time,   nvl(heat.ladle_open_time,0) AS prod_date,   heat.ladle_arrival_time AS ladle_arrival_time,   nvl(heat.tund_powder_type,0)  AS tund_powder_type,   nvl(heat.tund_powder_weight,0) AS tund_powder_weight,  tund_open.tund_open_time AS tund_open_time,   tund_close.tund_close_time,     -1     AS tund_life, nvl(tund_open.tund_open_weight    / 1000 ,0)AS tund_open_weight, nvl( tund_close.tund_close_weight  /1000 ,0) AS tund_close_weight, nvl(DECODE (NVL(tund_open.tund_no,-1), -1, tund_close.tund_no, tund_open.tund_no,0),0) tund_no, nvl(DECODE (NVL(tund_open.car_no, -1), -1, tund_close.car_no, tund_open.car_no,0),0) car_no,  nvl(v.sample_weight,0) AS sample_lost_weight,  nvl(v.scale_loss,0)    AS scale_lost_weight, nvl( heat.ladle_tare_weight,0) as ladle_tare_weight,  nvl((std.cast_len_end     -std.cast_len_bgn) ,0)                           AS total_cast_length,  ROUND(nvl((v.scrap_weight + v.head_crop_weight + v.tail_crop_weight),0),3) AS total_scrap_weight,  param.value                                                         AS casterid FROM pdc_heat heat,  pdc_strand std,  pdc_heat_strand hs,  pdc_heat_shift hshift,  pdc_customer_heat cheat,  pdc_sequence seq,  v_heat_tund_open tund_open,  v_heat_tund_close tund_close,  v_heat_weight v,  gdc_param param WHERE heat.sequence_steel_id    = seq.steel_id AND hs.strand_steel_id          = std.steel_id AND hs.heat_steel_id            = heat.steel_id AND hshift.heat_steel_id(+)     = heat.steel_id AND tund_open.heat_steel_id(+)  = heat.steel_id AND tund_close.heat_steel_id(+) = heat.steel_id AND cheat.steel_id(+)           = heat.steel_id AND v.steel_id(+)               = heat.steel_id AND param.par_context           = 'CASTER' AND param.par_name              = 'CASTERID' AND heat.steel_id= '" + StrCastSeq_caster3_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRptHeader where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "' and HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRptHeader set Slab_Weight ='" + dt_CastSeq.Rows[i]["Slab_Weight"].ToString() + "',Cut_Lost_Weight ='" + dt_CastSeq.Rows[i]["Cut_Lost_Weight"].ToString() + "',Slab_Count='" + dt_CastSeq.Rows[i]["Slab_Count"].ToString() + "',Avg_Slab_Width='" + dt_CastSeq.Rows[i]["Avg_Slab_Width"].ToString() + "',HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "',Heat_Seq_No='" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "',Grade_Code='" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "',Shift_Foreman='" + dt_CastSeq.Rows[i]["Shift_Foreman"].ToString() + "',Casting_Foreman='" + dt_CastSeq.Rows[i]["Casting_Foreman"].ToString() + "',Prod_OrderId='" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "',Ladle_No='" + dt_CastSeq.Rows[i]["Ladle_No"].ToString() + "',Ladle_Open_Weight='" + dt_CastSeq.Rows[i]["Ladle_Open_Weight"].ToString() + "',Ladle_Close_Weight='" + dt_CastSeq.Rows[i]["Ladle_Close_Weight"].ToString() + "',Cast_Weight='" + dt_CastSeq.Rows[i]["Cast_Weight"].ToString() + "',Yield='" + dt_CastSeq.Rows[i]["Yield"].ToString() + "',CrewId='" + dt_CastSeq.Rows[i]["CrewId"].ToString() + "',Treat_Counter='" + dt_CastSeq.Rows[i]["Treat_Counter"].ToString() + "',Slab_Length='" + dt_CastSeq.Rows[i]["Slab_Length"].ToString() + "',Ladle_Close_Time='" + (dt_CastSeq.Rows[i]["Ladle_Close_Time"].ToString()).Replace(",", ".").ToString() + "',Pouring_Duration='" + dt_CastSeq.Rows[i]["Pouring_Duration"].ToString() + "',Ladle_Open_Time='" + (dt_CastSeq.Rows[i]["Ladle_Open_Time"].ToString()).Replace(",", ".").ToString() + "',Prod_date='" + (dt_CastSeq.Rows[i]["Prod_date"].ToString()).Replace(",", ".").ToString() + "',Ladle_Arrival_Time='" + (dt_CastSeq.Rows[i]["Ladle_Arrival_Time"].ToString()).Replace(",", ".").ToString() + "',Tund_Powder_Type='" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "',Tund_Powder_weight='" + dt_CastSeq.Rows[i]["Tund_Powder_weight"].ToString() + "',Tund_Open_Time='" + (dt_CastSeq.Rows[i]["Tund_Open_Time"].ToString()).Replace(",", ".").ToString() + "',Tund_Close_Time='" + (dt_CastSeq.Rows[i]["Tund_Close_Time"].ToString()).Replace(",", ".").ToString() + "',Tund_Life='" + dt_CastSeq.Rows[i]["Tund_Life"].ToString() + "',Tund_Open_Weight='" + dt_CastSeq.Rows[i]["Tund_Open_Weight"].ToString() + "',Tund_Close_Weight='" + dt_CastSeq.Rows[i]["Tund_Close_Weight"].ToString() + "',Tund_No='" + dt_CastSeq.Rows[i]["Tund_No"].ToString() + "',Car_No='" + dt_CastSeq.Rows[i]["Car_No"].ToString() + "',Sample_Lost_Weight='" + dt_CastSeq.Rows[i]["Sample_Lost_Weight"].ToString() + "',Scale_Lost_Weight='" + dt_CastSeq.Rows[i]["Scale_Lost_Weight"].ToString() + "',Ladle_Tare_Weight='" + dt_CastSeq.Rows[i]["Ladle_Tare_Weight"].ToString() + "',Total_Cast_Length='" + dt_CastSeq.Rows[i]["Total_Cast_Length"].ToString() + "',Total_Scrap_Weight='" + dt_CastSeq.Rows[i]["Total_Scrap_Weight"].ToString() + "',CasterId='" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "' where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "' and HeatId='" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRptHeader(SeqNo,SteelId,Slab_Weight,Cut_Lost_Weight,Slab_Count,Avg_Slab_Width,HeatId,Heat_Seq_No,Grade_Code,Shift_Foreman,Casting_Foreman,Prod_OrderId,Ladle_No,Ladle_Open_Weight,Ladle_Close_Weight,Cast_Weight,Yield,CrewId,Treat_Counter,Slab_Length,Ladle_Close_Time,Pouring_Duration,Ladle_Open_Time,Prod_date,Ladle_Arrival_Time,Tund_Powder_Type,Tund_Powder_weight,Tund_Open_Time,Tund_Close_Time,Tund_Life,Tund_Open_Weight,Tund_Close_Weight,Tund_No,Car_No,Sample_Lost_Weight,Scale_Lost_Weight,Ladle_Tare_Weight,Total_Cast_Length,Total_Scrap_Weight,CasterId)values('" + StrCastSeq_caster3_MaxID + "','" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["slab_weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Cut_Lost_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_Count"].ToString() + "','" + dt_CastSeq.Rows[i]["Avg_Slab_Width"].ToString() + "','" + dt_CastSeq.Rows[i]["HeatId"].ToString() + "','" + dt_CastSeq.Rows[i]["Heat_Seq_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Grade_Code"].ToString() + "','" + dt_CastSeq.Rows[i]["Shift_Foreman"].ToString() + "','" + dt_CastSeq.Rows[i]["Casting_Foreman"].ToString() + "','" + dt_CastSeq.Rows[i]["Prod_OrderId"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Open_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Close_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Yield"].ToString() + "','" + dt_CastSeq.Rows[i]["CrewId"].ToString() + "','" + dt_CastSeq.Rows[i]["Treat_Counter"].ToString() + "','" + dt_CastSeq.Rows[i]["Slab_Length"].ToString() + "','" + (dt_CastSeq.Rows[i]["Ladle_Close_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Pouring_Duration"].ToString() + "','" + (dt_CastSeq.Rows[i]["Ladle_Open_Time"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Prod_date"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Ladle_Arrival_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Powder_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Powder_weight"].ToString() + "','" + (dt_CastSeq.Rows[i]["Tund_Open_Time"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Tund_Close_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Life"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Open_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_Close_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Tund_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Car_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Sample_Lost_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Scale_Lost_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Tare_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["Total_Cast_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Total_Scrap_Weight"].ToString() + "','" + dt_CastSeq.Rows[i]["CasterId"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatRpt_HeatReportSR(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = "SELECT s.strand_num AS strand_no,  s.avg_cast_speed,  s.temp_speed_tab_num AS copt_temp_tab_num,  s.s_speed_tab_num    AS copt_s_tab_num,  s.smn_speed_tab_num  AS copt_mns_tab_num,  s.tw_speed_tab_num   AS copt_tw_tab_num,  s.hmo_tab_num,  s.softred_tab_num AS hsa_tab_num,  s.gen_tab_num,  s.dsc_tab_num,  s.bops_tab_num AS mms_tab_num,  s.moldadj_tab_num,  s.sgrade_speed_tab_num AS copt_sgrade_tab_num,  s.mlc_tab_num,  s.cast_len_bgn       AS cast_len_begin,  s.cast_len_end       AS cast_len_end,  s.spray_plan_num     AS swplan_tab_num,  clo_tab_num          AS clo_tab_num,  s.clo_sample_tab_num AS clo_smp_tab_num,s.cast_powder_type as Mould_Pow_Type,s.cast_powder_weight as Mould_Pow_Amnt  FROM pdc_strand s,  pdc_heat_strand WHERE pdc_heat_strand.strand_steel_id = s.steel_id AND pdc_heat_strand.heat_steel_id='" + StrCastSeq_caster3_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatReportSR where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatReportSR set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Avg_Cast_Speed='" + dt_CastSeq.Rows[i]["Avg_Cast_Speed"].ToString() + "',Copt_Temp_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_Temp_Tab_Num"].ToString() + "',Copt_S_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_S_Tab_Num"].ToString() + "',Copt_MNS_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_MNS_Tab_Num"].ToString() + "',Copt_TW_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_TW_Tab_Num"].ToString() + "',HMO_Tab_Num='" + dt_CastSeq.Rows[i]["HMO_Tab_Num"].ToString() + "',HSA_Tab_Num='" + dt_CastSeq.Rows[i]["HSA_Tab_Num"].ToString() + "',Gen_Tab_Num='" + dt_CastSeq.Rows[i]["Gen_Tab_Num"].ToString() + "',DSC_Tab_Num='" + dt_CastSeq.Rows[i]["DSC_Tab_Num"].ToString() + "',MMS_Tab_Num='" + dt_CastSeq.Rows[i]["MMS_Tab_Num"].ToString() + "',MOLDADJ_Tab_Num='" + dt_CastSeq.Rows[i]["MOLDADJ_Tab_Num"].ToString() + "',Copt_SGrade_Tab_Num='" + dt_CastSeq.Rows[i]["Copt_SGrade_Tab_Num"].ToString() + "',MLC_Tab_Num='" + dt_CastSeq.Rows[i]["MLC_Tab_Num"].ToString() + "',Cast_Len_Begin='" + dt_CastSeq.Rows[i]["Cast_Len_Begin"].ToString() + "',Cast_Len_End='" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "',SWPlan_Tab_Num='" + dt_CastSeq.Rows[i]["SWPlan_Tab_Num"].ToString() + "',CLO_Tab_Num='" + dt_CastSeq.Rows[i]["CLO_Tab_Num"].ToString() + "',CLO_Smp_Tab_Num='" + dt_CastSeq.Rows[i]["CLO_Smp_Tab_Num"].ToString() + "',Mould_Pow_Type='" + dt_CastSeq.Rows[i]["Mould_Pow_Type"].ToString() + "',Mould_Pow_Amnt='" + dt_CastSeq.Rows[i]["Mould_Pow_Amnt"].ToString() + "' where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatReportSR(SeqNo,SteelId,Strand_No,Avg_Cast_Speed,Copt_Temp_Tab_Num,Copt_S_Tab_Num,Copt_MNS_Tab_Num,Copt_TW_Tab_Num,HMO_Tab_Num,HSA_Tab_Num,Gen_Tab_Num,DSC_Tab_Num,MMS_Tab_Num,MOLDADJ_Tab_Num,Copt_SGrade_Tab_Num,MLC_Tab_Num,Cast_Len_Begin,Cast_Len_End,SWPlan_Tab_Num,CLO_Tab_Num,CLO_Smp_Tab_Num,Mould_Pow_Type,Mould_Pow_Amnt)values('" + StrCastSeq_caster3_MaxID + "','" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Avg_Cast_Speed"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_Temp_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_S_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_MNS_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_TW_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["HMO_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["HSA_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Gen_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["DSC_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["MMS_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["MOLDADJ_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_SGrade_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["MLC_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_Begin"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Len_End"].ToString() + "','" + dt_CastSeq.Rows[i]["SWPlan_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["CLO_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["CLO_Smp_Tab_Num"].ToString() + "','" + dt_CastSeq.Rows[i]["Mould_Pow_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Mould_Pow_Amnt"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatRpt_SubMoldStrand(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = "SELECT s.strand_num    AS strand_no,  s.cast_len_bgn       AS cast_start_len,  s.cast_len_end       AS cast_end_len,  s.heat_in_mold_time  AS cast_start_time,  s.heat_out_mold_time AS cast_end_time, nvl(s.mold_no,0) as mold_no FROM pdc_strand s,  pdc_heat_strand hs WHERE s.steel_id     = hs.strand_steel_id AND hs.heat_steel_id='" + StrCastSeq_caster3_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_SubMoldStrand where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_SubMoldStrand set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Cast_Start_Len='" + dt_CastSeq.Rows[i]["Cast_Start_Len"].ToString() + "',Cast_End_Len='" + dt_CastSeq.Rows[i]["Cast_End_Len"].ToString() + "',Cast_Start_Time='" + (dt_CastSeq.Rows[i]["Cast_Start_Time"].ToString()).Replace(",", ".").ToString() + "',Cast_End_Time='" + (dt_CastSeq.Rows[i]["Cast_End_Time"].ToString()).Replace(",", ".").ToString() + "',Mold_No='" + dt_CastSeq.Rows[i]["Mold_No"].ToString() + "' where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_SubMoldStrand(SeqNo,SteelId,Strand_No,Cast_Start_Len,Cast_End_Len,Cast_Start_Time,Cast_End_Time,Mold_No)values('" + StrCastSeq_caster3_MaxID + "','" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Start_Len"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_End_Len"].ToString() + "','" + (dt_CastSeq.Rows[i]["Cast_Start_Time"].ToString()).Replace(",", ".").ToString() + "','" + (dt_CastSeq.Rows[i]["Cast_End_Time"].ToString()).Replace(",", ".").ToString() + "','" + dt_CastSeq.Rows[i]["Mold_No"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatRpt_SubPracticeTable(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = "SELECT s.strand_num  AS strand_no,  s.gen_tab_num  AS gen,  s.moldadj_tab_num   AS mold,  s.hmo_tab_num          AS hmo,  s.softred_tab_num      AS hsa,  s.mlc_tab_num          AS mlc,  s.spray_plan_num       AS cooling,  s.bops_tab_num         AS mms,  s.sgrade_speed_tab_num AS copt_speed,  s.temp_speed_tab_num   AS copt_temp,  s.tw_speed_tab_num     AS copt_tw,  s.s_speed_tab_num      AS copt_s,  s.smn_speed_tab_num    AS copt_mns,  s.clo_sample_tab_num   AS clo_smp,  clo_tab_num            AS clo,  s.dsc_tab_num          AS dsc,  0                      AS qual FROM pdc_strand s,  pdc_heat_strand WHERE pdc_heat_strand.strand_steel_id = s.steel_id AND pdc_heat_strand.heat_steel_id='" + StrCastSeq_caster3_SteelID + "'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_SubPracticeTable where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_SubPracticeTable set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Gen='" + dt_CastSeq.Rows[i]["Gen"].ToString() + "',Mold='" + dt_CastSeq.Rows[i]["Mold"].ToString() + "',HMO='" + dt_CastSeq.Rows[i]["HMO"].ToString() + "',HSA='" + dt_CastSeq.Rows[i]["HSA"].ToString() + "',MLC='" + dt_CastSeq.Rows[i]["MLC"].ToString() + "',Cooling='" + dt_CastSeq.Rows[i]["Cooling"].ToString() + "',MMS='" + dt_CastSeq.Rows[i]["MMS"].ToString() + "',Copt_Speed='" + dt_CastSeq.Rows[i]["Copt_Speed"].ToString() + "',Copt_Temp='" + dt_CastSeq.Rows[i]["Copt_Temp"].ToString() + "',Copt_TW='" + dt_CastSeq.Rows[i]["Copt_TW"].ToString() + "',Copt_S='" + dt_CastSeq.Rows[i]["Copt_S"].ToString() + "',Copt_MNS='" + dt_CastSeq.Rows[i]["Copt_MNS"].ToString() + "',Clo_SMP='" + dt_CastSeq.Rows[i]["Clo_SMP"].ToString() + "',CLO='" + dt_CastSeq.Rows[i]["CLO"].ToString() + "',DSC='" + dt_CastSeq.Rows[i]["DSC"].ToString() + "',QUAL='" + dt_CastSeq.Rows[i]["QUAL"].ToString() + "' where SeqNo='" + StrCastSeq_caster3_MaxID + "' and SteelId='" + StrCastSeq_caster3_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_SubPracticeTable(SeqNo,SteelId,Strand_No,Gen,Mold,HMO,HSA,MLC,Cooling,MMS,Copt_Speed,Copt_Temp,Copt_TW,Copt_S,Copt_MNS,Clo_SMP,CLO,DSC,QUAL)values('" + StrCastSeq_caster3_MaxID + "','" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Gen"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold"].ToString() + "','" + dt_CastSeq.Rows[i]["HMO"].ToString() + "','" + dt_CastSeq.Rows[i]["HSA"].ToString() + "','" + dt_CastSeq.Rows[i]["MLC"].ToString() + "','" + dt_CastSeq.Rows[i]["Cooling"].ToString() + "','" + dt_CastSeq.Rows[i]["MMS"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_Speed"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_Temp"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_TW"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_S"].ToString() + "','" + dt_CastSeq.Rows[i]["Copt_MNS"].ToString() + "','" + dt_CastSeq.Rows[i]["Clo_SMP"].ToString() + "','" + dt_CastSeq.Rows[i]["CLO"].ToString() + "','" + dt_CastSeq.Rows[i]["DSC"].ToString() + "','" + dt_CastSeq.Rows[i]["QUAL"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatRpt_AnalysisElement(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            string StrCastSeqMaxID = "SELECT gcc_element_cat.element_name AS chem_abbr,  gcc_element_cat.display_order     AS display_order,  gcc_element_cat.element_type      AS element_type FROM gcc_element_cat ORDER BY display_order";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select * from HeatRpt_AnalysisElement where DISPLAY_ORDER='" + dt_CastSeq.Rows[i]["DISPLAY_ORDER"].ToString() + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_AnalysisElement set CHEM_ABBR='" + dt_CastSeq.Rows[i]["CHEM_ABBR"].ToString() + "',ELEMENT_TYPE='" + dt_CastSeq.Rows[i]["ELEMENT_TYPE"].ToString() + "' where DISPLAY_ORDER='" + dt_CastSeq.Rows[i]["DISPLAY_ORDER"].ToString() + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_AnalysisElement(CHEM_ABBR,DISPLAY_ORDER,ELEMENT_TYPE)values('" + dt_CastSeq.Rows[i]["CHEM_ABBR"].ToString() + "','" + dt_CastSeq.Rows[i]["DISPLAY_ORDER"].ToString() + "','" + dt_CastSeq.Rows[i]["ELEMENT_TYPE"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }
        private void Fn_Caster3_HeatRpt_HeatRpt_HeatEquipmentData(string StrCastSeq_caster3_MaxID, string StrCastSeq_caster3_SteelID)
        {
            //Ladle_Shroud_Id,Ladle_Shroud_Life,Ladle_Shroud_Count,Nozzle_Id,Nozzle_Life,Nozzle_Count,Tundish_Id,Tundish_Life,Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxID = "SELECT MAX(eq.customer_id)       AS ladle_shroud_id,  TRUNC(MAX(co.counter_value),2) AS ladle_shroud_life,  COUNT(eq.customer_id)-1        AS ladle_shroud_count FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster3_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'SHROUD' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'HEAT'";
            DBConnections clsObj_CastSeq = new DBConnections();
            DataTable dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxID);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Ladle_Shroud_Id='" + dt_CastSeq.Rows[i]["Ladle_Shroud_Id"].ToString() + "',Ladle_Shroud_Life='" + dt_CastSeq.Rows[i]["Ladle_Shroud_Life"].ToString() + "',Ladle_Shroud_Count='" + dt_CastSeq.Rows[i]["Ladle_Shroud_Count"].ToString() + "' where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Ladle_Shroud_Id,Ladle_Shroud_Life,Ladle_Shroud_Count)values('" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["Ladle_Shroud_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Shroud_Life"].ToString() + "','" + dt_CastSeq.Rows[i]["Ladle_Shroud_Count"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }

            //Nozzle_Id,Nozzle_Life,Nozzle_Count,Tundish_Id,Tundish_Life,Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDNozzle = "SELECT MAX(eq.customer_id)       AS nozzle_id,  TRUNC(MAX(co.counter_value),2) AS nozzle_life,  COUNT(eq.customer_id)-1        AS nozzle_count FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster3_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'NOZZLE' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'HEAT'";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxIDNozzle);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Nozzle_Id='" + dt_CastSeq.Rows[i]["Nozzle_Id"].ToString() + "',Nozzle_Life='" + dt_CastSeq.Rows[i]["Nozzle_Life"].ToString() + "',Nozzle_Count='" + dt_CastSeq.Rows[i]["Nozzle_Count"].ToString() + "' where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Nozzle_Id,Nozzle_Life,Nozzle_Count)values('" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["Nozzle_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Nozzle_Life"].ToString() + "','" + dt_CastSeq.Rows[i]["Nozzle_Count"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
            //Tundish_Id,Tundish_Life,Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDTundish = "SELECT eq.customer_id       AS tundish_id,  TRUNC(co.counter_value,2) AS tundish_life FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster3_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'TUNDISH' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'TON' ";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxIDTundish);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set tundish_id='" + dt_CastSeq.Rows[i]["tundish_id"].ToString() + "',Tundish_Life='" + dt_CastSeq.Rows[i]["Tundish_Life"].ToString() + "' where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Tundish_Id,Tundish_Life)values('" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["Tundish_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Tundish_Life"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
            //Mold_Id,Mold_Life,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDMold = "SELECT eq.customer_id       AS mold_id,  TRUNC(co.counter_value,2) AS mold_life FROM pdc_heat_equip eq,  pdc_heat_equip_counter co WHERE eq.steel_id      = '" + StrCastSeq_caster3_SteelID + "' AND eq.steel_id        = co.steel_id(+) AND eq.equip_type      = 'MOLD' AND eq.equip_id        = co.equip_id(+) AND co.counter_type(+) = 'HEAT' ";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxIDMold);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Mold_Id='" + dt_CastSeq.Rows[i]["Mold_Id"].ToString() + "',Mold_Life='" + dt_CastSeq.Rows[i]["Mold_Life"].ToString() + "' where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Mold_Id,Mold_Life)values('" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["Mold_Id"].ToString() + "','" + dt_CastSeq.Rows[i]["Mold_Life"].ToString() + "')";

                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }

            //Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right
            string StrCastSeqMaxIDStnd = "SELECT strand_no,  cast_length,  event_type,  width_toc,  width_boc,  taper_left,  taper_right FROM pdc_width_change WHERE heat_steel_id = '" + StrCastSeq_caster3_SteelID + "' AND event_type     IN (1,3) ";
            clsObj_CastSeq = new DBConnections();
            dt_CastSeq = new DataTable();
            dt_CastSeq = clsObj.DBSelectQueryCASTERIII_Table(StrCastSeqMaxIDStnd);
            if (dt_CastSeq.Rows.Count > 0)
            {
                for (int i = 0; i < dt_CastSeq.Rows.Count; i++)
                {
                    DataTable dt_MIS_CASTER = new DataTable();
                    string StrQueryCASTER = "select SteelId from HeatRpt_HeatEquipmentData where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                    dt_MIS_CASTER = clsObj.DBSelectQueryMIS_Table(StrQueryCASTER);
                    if (dt_MIS_CASTER.Rows.Count > 0)
                    {
                        string StrUpdateQuery = "Update HeatRpt_HeatEquipmentData set Strand_No='" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "',Cast_Length='" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "',Event_Type='" + dt_CastSeq.Rows[i]["Event_Type"].ToString() + "',Width_TOC='" + dt_CastSeq.Rows[i]["Width_TOC"].ToString() + "',Width_BOC='" + dt_CastSeq.Rows[i]["Width_BOC"].ToString() + "',Taper_Left='" + dt_CastSeq.Rows[i]["Taper_Left"].ToString() + "',Taper_Right='" + dt_CastSeq.Rows[i]["Taper_Right"].ToString() + "' where SteelId='" + StrCastSeq_caster3_SteelID + "'";
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                    }
                    else
                    {
                        string StrInsertQuery = "insert HeatRpt_HeatEquipmentData(SteelId,Strand_No,Cast_Length,Event_Type,Width_TOC,Width_BOC,Taper_Left,Taper_Right)values('" + StrCastSeq_caster3_SteelID + "','" + dt_CastSeq.Rows[i]["Strand_No"].ToString() + "','" + dt_CastSeq.Rows[i]["Cast_Length"].ToString() + "','" + dt_CastSeq.Rows[i]["Event_Type"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_TOC"].ToString() + "','" + dt_CastSeq.Rows[i]["Width_BOC"].ToString() + "','" + dt_CastSeq.Rows[i]["Taper_Left"].ToString() + "','" + dt_CastSeq.Rows[i]["Taper_Right"].ToString() + "')";


                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrInsertQuery);
                    }
                }
            }
        }



        # endregion

        private void timer_FinalChemToMis_Tick(object sender, EventArgs e)
        {

            //MessageBox.Show(System.DateTime.Now.ToString());
            FnFinAimChem();
        }
        //For the manual Update of missing data  from LIMS to MIS
        public void FnLIMStoMISManualUpd()
        {
            DataTable dt_LiveHeatIdfrmMIS = new DataTable();
            string StrLiveHeatIdfrmMIS = "SELECT TOP 1 CONARC_A,CONARC_B FROM ANNOUNCED_HEAT";
            dt_LiveHeatIdfrmMIS = clsObj.DBSelectQueryMIS_Table(StrLiveHeatIdfrmMIS);
            foreach (DataRow dtRow in dt_LiveHeatIdfrmMIS.Rows)
            {
                foreach (DataColumn dc in dt_LiveHeatIdfrmMIS.Columns)
                {
                    string HeatId = dtRow[dc].ToString();

                    if (dt_LiveHeatIdfrmMIS.Rows.Count > 0)
                    {
                        string StrQry = "";
                        DataTable dt_LIMSChemToMIS = new DataTable();
                        //string StrQuery = "SELECT * FROM TestResultSpectro2 where HeatID='"+StrHEATID+"' UNION ALL select  *from TestResultSpectro1  where HeatID='"+StrHEATID+"'";
                        //string StrQuery = "SELECT * FROM TestResultSpectro2 where HeatID='" + HeatId + "' UNION select  *from TestResultSpectro1  where HeatID='" + HeatId + "'";
                        string StrQuery = "SELECT * FROM TestResultSpectro2 where HeatID='14B1121' UNION select  *from TestResultSpectro1  where HeatID='14B1121'";
                        //dt_LIMSChemToMIS = DBSelectQueryLMIS_Table(StrQuery);
                        dt_LIMSChemToMIS = clsObj.DBSelectQueryLMIS_Table(StrQuery);
                        if (dt_LIMSChemToMIS.Rows.Count > 0)
                        {
                            DataTable dt_Mis_HeatChem = new DataTable();
                            for (int i = 0; i < dt_LIMSChemToMIS.Rows.Count; i++)
                            {

                                string StrQueryMISHeatChem = "select *from SMS_Live_Chemistry_Status where HeatID='" + dt_LIMSChemToMIS.Rows[i]["HeatID"].ToString() + "' and SampleNo='" + dt_LIMSChemToMIS.Rows[i]["SampleNo"].ToString() + "' and TreatID='" + dt_LIMSChemToMIS.Rows[i]["TreatID"].ToString() + "'";
                                //dt_Mis_HeatChem = DBSelectQueryMIS_Table(StrQueryMISHeatChem);
                                dt_Mis_HeatChem = clsObj.DBSelectQueryMIS_Table(StrQueryMISHeatChem);
                                if (dt_Mis_HeatChem.Rows.Count == 0)
                                {
                                    //insert
                                    StrQry = "INSERT INTO SMS_Live_Chemistry_Status(HeatID,SampleNo,L2Status,TreatID,C,Mn,S,P,Si,Al,N,Cu,Ni,Cr,Re,Nb,V,Ti,Mo,B," +
                                            "Sn,Ca,Al_S,O,Co,Pb,W,Mg,Ce,A_s,Bi,Sb,Zr,H,Te,Al_l,Zn,Fe,Mn_Si,B_S,B_Oxy,SP1,SP2,SP3) VALUES('" + dt_LIMSChemToMIS.Rows[i]["HeatID"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["SampleNo"] + "','" + dt_LIMSChemToMIS.Rows[i]["L2Status"] + "','" + dt_LIMSChemToMIS.Rows[i]["TreatID"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["C"] + "','" + dt_LIMSChemToMIS.Rows[i]["Mn"] + "','" + dt_LIMSChemToMIS.Rows[i]["S"] + "','" + dt_LIMSChemToMIS.Rows[i]["P"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["Si"] + "','" + dt_LIMSChemToMIS.Rows[i]["Al"] + "','" + dt_LIMSChemToMIS.Rows[i]["N"] + "','" + dt_LIMSChemToMIS.Rows[i]["Cu"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["Ni"] + "','" + dt_LIMSChemToMIS.Rows[i]["Cr"] + "','" + dt_LIMSChemToMIS.Rows[i]["Re"] + "','" + dt_LIMSChemToMIS.Rows[i]["Nb"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["V"] + "','" + dt_LIMSChemToMIS.Rows[i]["Ti"] + "','" + dt_LIMSChemToMIS.Rows[i]["Mo"] + "','" + dt_LIMSChemToMIS.Rows[i]["B"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["Sn"] + "','" + dt_LIMSChemToMIS.Rows[i]["Ca"] + "','" + dt_LIMSChemToMIS.Rows[i]["Al_S"] + "','" + dt_LIMSChemToMIS.Rows[i]["O"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["Co"] + "','" + dt_LIMSChemToMIS.Rows[i]["Pb"] + "','" + dt_LIMSChemToMIS.Rows[i]["W"] + "','" + dt_LIMSChemToMIS.Rows[i]["Mg"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["Ce"] + "','" + dt_LIMSChemToMIS.Rows[i]["A_s"] + "','" + dt_LIMSChemToMIS.Rows[i]["Bi"] + "','" + dt_LIMSChemToMIS.Rows[i]["Sb"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["Zr"] + "','" + dt_LIMSChemToMIS.Rows[i]["H"] + "','" + dt_LIMSChemToMIS.Rows[i]["Te"] + "','" + dt_LIMSChemToMIS.Rows[i]["Al_l"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["Zn"] + "','" + dt_LIMSChemToMIS.Rows[i]["Fe"] + "','" + dt_LIMSChemToMIS.Rows[i]["Mn_Si"] + "','" + dt_LIMSChemToMIS.Rows[i]["B_S"] + "','" +
                                            dt_LIMSChemToMIS.Rows[i]["B_Oxy"] + "','" + dt_LIMSChemToMIS.Rows[i]["SP1"] + "','" + dt_LIMSChemToMIS.Rows[i]["SP2"] + "','" + dt_LIMSChemToMIS.Rows[i]["SP3"] + "')";
                                    //bool StrQStatus = DBInsertUpdateDeleteMIS(StrQry);
                                    bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrQry);
                                }
                                else
                                {
                                    // update
                                    StrQry = "UPDATE SMS_Live_Chemistry_Status SET [HeatID] = '" + dt_LIMSChemToMIS.Rows[i]["HeatID"] + "',[SampleNo] ='" + dt_LIMSChemToMIS.Rows[i]["SampleNo"] +
                                        "',[L2Status] ='" + dt_LIMSChemToMIS.Rows[i]["L2Status"] + "',[TreatID] ='" + dt_LIMSChemToMIS.Rows[i]["TreatID"] + "',[C] ='" + dt_LIMSChemToMIS.Rows[i]["C"] +
                                        "',[Mn] ='" + dt_LIMSChemToMIS.Rows[i]["Mn"] + "',[S] ='" + dt_LIMSChemToMIS.Rows[i]["S"] + "' ,[P] ='" + dt_LIMSChemToMIS.Rows[i]["P"] + "',[Si] = '" +
                                        dt_LIMSChemToMIS.Rows[i]["Si"] + "',[Al] ='" + dt_LIMSChemToMIS.Rows[i]["Al"] + "',[N] ='" + dt_LIMSChemToMIS.Rows[i]["N"] + "',[Cu] ='" + dt_LIMSChemToMIS.Rows[i]["Cu"] +
                                        "',[Ni] ='" + dt_LIMSChemToMIS.Rows[i]["Ni"] + "',[Cr]='" + dt_LIMSChemToMIS.Rows[i]["Cr"] + "',[Re] ='" + dt_LIMSChemToMIS.Rows[i]["Re"] + "',[Nb]='" + dt_LIMSChemToMIS.Rows[i]["Nb"] +
                                        "' ,[V] ='" + dt_LIMSChemToMIS.Rows[i]["V"] + "',[Ti] ='" + dt_LIMSChemToMIS.Rows[i]["Ti"] + "',[Mo] ='" + dt_LIMSChemToMIS.Rows[i]["Mo"] + "',[B] ='" + dt_LIMSChemToMIS.Rows[i]["B"] +
                                        "',[Sn] ='" + dt_LIMSChemToMIS.Rows[i]["Sn"] + "',[Ca] ='" + dt_LIMSChemToMIS.Rows[i]["Ca"] + "',[Al_S] ='" + dt_LIMSChemToMIS.Rows[i]["Al_S"] + "',[O] = '" + dt_LIMSChemToMIS.Rows[i]["O"] +
                                        "',[Co] ='" + dt_LIMSChemToMIS.Rows[i]["Co"] + "',[Pb] ='" + dt_LIMSChemToMIS.Rows[i]["Pb"] + "',[W] ='" + dt_LIMSChemToMIS.Rows[i]["W"] + "' ,[Mg] ='" + dt_LIMSChemToMIS.Rows[i]["Mg"] +
                                        "',[Ce] ='" + dt_LIMSChemToMIS.Rows[i]["Ce"] + "',[A_s] ='" + dt_LIMSChemToMIS.Rows[i]["A_s"] + "',[Bi] ='" + dt_LIMSChemToMIS.Rows[i]["Bi"] + "',[Sb] ='" + dt_LIMSChemToMIS.Rows[i]["Sb"] +
                                        "',[Zr] ='" + dt_LIMSChemToMIS.Rows[i]["Zr"] + "',[H] ='" + dt_LIMSChemToMIS.Rows[i]["H"] + "',[Te] ='" + dt_LIMSChemToMIS.Rows[i]["Te"] + "',[Al_l] ='" + dt_LIMSChemToMIS.Rows[i]["Al_l"] +
                                        "',[Zn] ='" + dt_LIMSChemToMIS.Rows[i]["Zn"] + "',[Fe] ='" + dt_LIMSChemToMIS.Rows[i]["Fe"] + "',[Mn_Si] ='" + dt_LIMSChemToMIS.Rows[i]["Mn_Si"] + "',[B_S]='" + dt_LIMSChemToMIS.Rows[i]["B_S"] +
                                        "',[B_Oxy] ='" + dt_LIMSChemToMIS.Rows[i]["B_Oxy"] + "',[SP1]='" + dt_LIMSChemToMIS.Rows[i]["SP1"] + "',[SP2] ='" + dt_LIMSChemToMIS.Rows[i]["SP2"] + "' ,[SP3] ='" + dt_LIMSChemToMIS.Rows[i]["SP3"] +
                                        "' WHERE HeatID='" + dt_LIMSChemToMIS.Rows[i]["HeatID"].ToString() + "' and SampleNo='" + dt_LIMSChemToMIS.Rows[i]["SampleNo"].ToString() + "' and TreatID='" + dt_LIMSChemToMIS.Rows[i]["TreatID"].ToString() + "'";
                                    //bool StrQStatus2 = DBInsertUpdateDeleteMIS(StrQry);
                                    bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrQry);
                                }
                            }
                        }
                    }
                }
            }
        }
        public void FnLIMStoMIS()
        {
            //@Author- MOHAN SINGH ,LastEdited:-19052014(Added Spectro3 Data)
            string StrQry = "";
            DataTable dt_LIMSChemToMIS = new DataTable();
            //string StrQuery = "SELECT * FROM TestResultSpectro2 where HeatID='"+StrHEATID+"' UNION ALL select  *from TestResultSpectro1  where HeatID='"+StrHEATID+"'";
            //string StrQuery = "select TRS1.*,HTRDt.TimeOfAnalysis,HTRDt.MachineName from TestResultSpectro1 TRS1 inner join  HeatTestResultDataTable HTRDt on TRS1.HeatID=HTRDt.HeatID and  TRS1.TreatID=HTRDt.TreatID and TRS1.SampleNo=HTRDt.SampleNo  where  HTRDt.TimeOfAnalysis=(select MAX(TimeOfAnalysis) from HeatTestResultDataTable where MachineName='Spectro1') UNION ALL select TRS2.*,HTRDt.TimeOfAnalysis,HTRDt.MachineName from TestResultSpectro2 TRS2 inner join  HeatTestResultDataTable HTRDt on TRS2.HeatID=HTRDt.HeatID and TRS2.TreatID=HTRDt.TreatID and TRS2.SampleNo=HTRDt.SampleNo  where  HTRDt.TimeOfAnalysis=(select MAX(TimeOfAnalysis) from HeatTestResultDataTable where MachineName='Spectro2')";
            string StrQuery = "select TRS1.*,HTRDt.TimeOfAnalysis,HTRDt.MachineName from TestResultSpectro1 TRS1 inner join HeatTestResultDataTable HTRDt on TRS1.HeatID=HTRDt.HeatID and  TRS1.TreatID=HTRDt.TreatID" +
                                " and TRS1.SampleNo=HTRDt.SampleNo  where  HTRDt.TimeOfAnalysis=(select MAX(TimeOfAnalysis) from HeatTestResultDataTable where MachineName='Spectro1') " +
                                 " UNION ALL" +
                                " select TRS2.*,HTRDt.TimeOfAnalysis,HTRDt.MachineName from TestResultSpectro2 TRS2 inner join HeatTestResultDataTable HTRDt on TRS2.HeatID=HTRDt.HeatID and TRS2.TreatID=HTRDt.TreatID" +
                                " and TRS2.SampleNo=HTRDt.SampleNo  where  HTRDt.TimeOfAnalysis=(select MAX(TimeOfAnalysis) from HeatTestResultDataTable where MachineName='Spectro2') " +
                                " UNION ALL " +
                                " select TRS3.*,HTRDt.TimeOfAnalysis,HTRDt.MachineName from TestResultSpectro3 TRS3 inner join HeatTestResultDataTable HTRDt on TRS3.HeatID=HTRDt.HeatID and TRS3.TreatID=HTRDt.TreatID" +
                                " and TRS3.SampleNo=HTRDt.SampleNo  where  HTRDt.TimeOfAnalysis=(select MAX(TimeOfAnalysis) from HeatTestResultDataTable where MachineName='Spectro3') "+
                                " UNION ALL " +
                                "select TRS4.*,HTRDt.TimeOfAnalysis,HTRDt.MachineName from TestResultBofCLab TRS4 inner join HeatTestResultDataTable HTRDt on TRS4.HeatID=HTRDt.HeatID and TRS4.TreatID=HTRDt.TreatID "+
                                " and TRS4.SampleNo=HTRDt.SampleNo  where  HTRDt.TimeOfAnalysis=(select MAX(TimeOfAnalysis) from HeatTestResultDataTable where MachineName='BofCLab')";
            //dt_LIMSChemToMIS = DBSelectQueryLMIS_Table(StrQuery);
            dt_LIMSChemToMIS = clsObj.DBSelectQueryLMIS_Table(StrQuery);
            if (dt_LIMSChemToMIS.Rows.Count > 0)
            {
                DataTable dt_Mis_HeatChem = new DataTable();
                for (int i = 0; i < dt_LIMSChemToMIS.Rows.Count; i++)
                {
                    string StrQueryMISHeatChem = "select *from SMS_Live_Chemistry_Status where HeatID='" + dt_LIMSChemToMIS.Rows[i]["HeatID"].ToString() + "' and SampleNo='" + dt_LIMSChemToMIS.Rows[i]["SampleNo"].ToString() + "' and TreatID='" + dt_LIMSChemToMIS.Rows[i]["TreatID"].ToString() + "'";
                    //dt_Mis_HeatChem = DBSelectQueryMIS_Table(StrQueryMISHeatChem);
                    dt_Mis_HeatChem = clsObj.DBSelectQueryMIS_Table(StrQueryMISHeatChem);
                    if (dt_Mis_HeatChem.Rows.Count == 0)
                    {
                        //insert
                        StrQry = "INSERT INTO SMS_Live_Chemistry_Status(HeatID,SampleNo,L2Status,TreatID,C,Mn,S,P,Si,Al,N,Cu,Ni,Cr,Re,Nb,V,Ti,Mo,B," +
                                "Sn,Ca,Al_S,O,Co,Pb,W,Mg,Ce,A_s,Bi,Sb,Zr,H,Te,Al_l,Zn,Fe,Mn_Si,B_S,B_Oxy,SP1,SP2,SP3,TimeOfAnalysis,MachineName) VALUES('" + dt_LIMSChemToMIS.Rows[i]["HeatID"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["SampleNo"] + "','" + dt_LIMSChemToMIS.Rows[i]["L2Status"] + "','" + dt_LIMSChemToMIS.Rows[i]["TreatID"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["C"] + "','" + dt_LIMSChemToMIS.Rows[i]["Mn"] + "','" + dt_LIMSChemToMIS.Rows[i]["S"] + "','" + dt_LIMSChemToMIS.Rows[i]["P"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["Si"] + "','" + dt_LIMSChemToMIS.Rows[i]["Al"] + "','" + dt_LIMSChemToMIS.Rows[i]["N"] + "','" + dt_LIMSChemToMIS.Rows[i]["Cu"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["Ni"] + "','" + dt_LIMSChemToMIS.Rows[i]["Cr"] + "','" + dt_LIMSChemToMIS.Rows[i]["Re"] + "','" + dt_LIMSChemToMIS.Rows[i]["Nb"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["V"] + "','" + dt_LIMSChemToMIS.Rows[i]["Ti"] + "','" + dt_LIMSChemToMIS.Rows[i]["Mo"] + "','" + dt_LIMSChemToMIS.Rows[i]["B"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["Sn"] + "','" + dt_LIMSChemToMIS.Rows[i]["Ca"] + "','" + dt_LIMSChemToMIS.Rows[i]["Al_S"] + "','" + dt_LIMSChemToMIS.Rows[i]["O"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["Co"] + "','" + dt_LIMSChemToMIS.Rows[i]["Pb"] + "','" + dt_LIMSChemToMIS.Rows[i]["W"] + "','" + dt_LIMSChemToMIS.Rows[i]["Mg"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["Ce"] + "','" + dt_LIMSChemToMIS.Rows[i]["A_s"] + "','" + dt_LIMSChemToMIS.Rows[i]["Bi"] + "','" + dt_LIMSChemToMIS.Rows[i]["Sb"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["Zr"] + "','" + dt_LIMSChemToMIS.Rows[i]["H"] + "','" + dt_LIMSChemToMIS.Rows[i]["Te"] + "','" + dt_LIMSChemToMIS.Rows[i]["Al_l"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["Zn"] + "','" + dt_LIMSChemToMIS.Rows[i]["Fe"] + "','" + dt_LIMSChemToMIS.Rows[i]["Mn_Si"] + "','" + dt_LIMSChemToMIS.Rows[i]["B_S"] + "','" +
                                dt_LIMSChemToMIS.Rows[i]["B_Oxy"] + "','" + dt_LIMSChemToMIS.Rows[i]["SP1"] + "','" + dt_LIMSChemToMIS.Rows[i]["SP2"] + "','" + dt_LIMSChemToMIS.Rows[i]["SP3"] + "','" + dt_LIMSChemToMIS.Rows[i]["TimeOfAnalysis"] + "','" + dt_LIMSChemToMIS.Rows[i]["MachineName"] + "')";
                        //bool StrQStatus = DBInsertUpdateDeleteMIS(StrQry);
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrQry);
                    }
                    else
                    {
                        // update
                        StrQry = "UPDATE SMS_Live_Chemistry_Status SET [HeatID] = '" + dt_LIMSChemToMIS.Rows[i]["HeatID"] + "',[SampleNo] ='" + dt_LIMSChemToMIS.Rows[i]["SampleNo"] +
                            "',[L2Status] ='" + dt_LIMSChemToMIS.Rows[i]["L2Status"] + "',[TreatID] ='" + dt_LIMSChemToMIS.Rows[i]["TreatID"] + "',[C] ='" + dt_LIMSChemToMIS.Rows[i]["C"] +
                            "',[Mn] ='" + dt_LIMSChemToMIS.Rows[i]["Mn"] + "',[S] ='" + dt_LIMSChemToMIS.Rows[i]["S"] + "' ,[P] ='" + dt_LIMSChemToMIS.Rows[i]["P"] + "',[Si] = '" +
                            dt_LIMSChemToMIS.Rows[i]["Si"] + "',[Al] ='" + dt_LIMSChemToMIS.Rows[i]["Al"] + "',[N] ='" + dt_LIMSChemToMIS.Rows[i]["N"] + "',[Cu] ='" + dt_LIMSChemToMIS.Rows[i]["Cu"] +
                            "',[Ni] ='" + dt_LIMSChemToMIS.Rows[i]["Ni"] + "',[Cr]='" + dt_LIMSChemToMIS.Rows[i]["Cr"] + "',[Re] ='" + dt_LIMSChemToMIS.Rows[i]["Re"] + "',[Nb]='" + dt_LIMSChemToMIS.Rows[i]["Nb"] +
                            "' ,[V] ='" + dt_LIMSChemToMIS.Rows[i]["V"] + "',[Ti] ='" + dt_LIMSChemToMIS.Rows[i]["Ti"] + "',[Mo] ='" + dt_LIMSChemToMIS.Rows[i]["Mo"] + "',[B] ='" + dt_LIMSChemToMIS.Rows[i]["B"] +
                            "',[Sn] ='" + dt_LIMSChemToMIS.Rows[i]["Sn"] + "',[Ca] ='" + dt_LIMSChemToMIS.Rows[i]["Ca"] + "',[Al_S] ='" + dt_LIMSChemToMIS.Rows[i]["Al_S"] + "',[O] = '" + dt_LIMSChemToMIS.Rows[i]["O"] +
                            "',[Co] ='" + dt_LIMSChemToMIS.Rows[i]["Co"] + "',[Pb] ='" + dt_LIMSChemToMIS.Rows[i]["Pb"] + "',[W] ='" + dt_LIMSChemToMIS.Rows[i]["W"] + "' ,[Mg] ='" + dt_LIMSChemToMIS.Rows[i]["Mg"] +
                            "',[Ce] ='" + dt_LIMSChemToMIS.Rows[i]["Ce"] + "',[A_s] ='" + dt_LIMSChemToMIS.Rows[i]["A_s"] + "',[Bi] ='" + dt_LIMSChemToMIS.Rows[i]["Bi"] + "',[Sb] ='" + dt_LIMSChemToMIS.Rows[i]["Sb"] +
                            "',[Zr] ='" + dt_LIMSChemToMIS.Rows[i]["Zr"] + "',[H] ='" + dt_LIMSChemToMIS.Rows[i]["H"] + "',[Te] ='" + dt_LIMSChemToMIS.Rows[i]["Te"] + "',[Al_l] ='" + dt_LIMSChemToMIS.Rows[i]["Al_l"] +
                            "',[Zn] ='" + dt_LIMSChemToMIS.Rows[i]["Zn"] + "',[Fe] ='" + dt_LIMSChemToMIS.Rows[i]["Fe"] + "',[Mn_Si] ='" + dt_LIMSChemToMIS.Rows[i]["Mn_Si"] + "',[B_S]='" + dt_LIMSChemToMIS.Rows[i]["B_S"] +
                            "',[B_Oxy] ='" + dt_LIMSChemToMIS.Rows[i]["B_Oxy"] + "',[SP1]='" + dt_LIMSChemToMIS.Rows[i]["SP1"] + "',[SP2] ='" + dt_LIMSChemToMIS.Rows[i]["SP2"] + "' ,[SP3] ='" + dt_LIMSChemToMIS.Rows[i]["SP3"] + "',[TimeOfAnalysis]='" + dt_LIMSChemToMIS.Rows[i]["TimeOfAnalysis"] + "',[MachineName]='" + dt_LIMSChemToMIS.Rows[i]["MachineName"] +
                            "' WHERE HeatID='" + dt_LIMSChemToMIS.Rows[i]["HeatID"].ToString() + "' and SampleNo='" + dt_LIMSChemToMIS.Rows[i]["SampleNo"].ToString() + "' and TreatID='" + dt_LIMSChemToMIS.Rows[i]["TreatID"].ToString() + "'";
                        //bool StrQStatus2 = DBInsertUpdateDeleteMIS(StrQry);
                        bool StrQStatus = clsObj.DBInsertUpdateDeleteMIS(StrQry);
                    }
                }
            }

        }
        private void timer_LIMSToMISHeatChem_Tick(object sender, EventArgs e)
        {
            FnLIMStoMIS();
            //FnLIMStoMISManualUpd();
        }

        private void txtError_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {

            DateTime start = DateTime.Parse(dtpDateFrom.Text);
            DateTime end = DateTime.Parse(dtpDateTo.Text);

            for (DateTime counter = start; counter <= end; counter = counter.AddDays(1))
            {
                lblmsg.Text = "";
                string StrPLANT = "";
                string StrPLANTNO = "";
                string StrMIS_PLANTNO = "";
                //string StrDate = dtpDateFrom.Text;
                string StrDate = Convert.ToDateTime(counter).ToString("yyyy-MM-dd");
                string StrDeleteQuery = "delete from smsHeatTracker where DateStamp='" + StrDate + "'";
                bool StrDStatus = clsObj.DBInsertUpdateDeleteMIS(StrDeleteQuery);
                StrPLANT = "CON"; StrPLANTNO = "1"; StrMIS_PLANTNO = "3";
                FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                //FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                StrPLANT = "CON"; StrPLANTNO = "2"; StrMIS_PLANTNO = "4";
                FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                //FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                StrPLANT = "LF"; StrPLANTNO = "1"; StrMIS_PLANTNO = "9";
                FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                //FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                StrPLANT = "LF"; StrPLANTNO = "2"; StrMIS_PLANTNO = "10";
                FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                //FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                StrPLANT = "HMD"; StrPLANTNO = "1"; StrMIS_PLANTNO = "1";
                FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                //FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                StrPLANT = "HMD"; StrPLANTNO = "2"; StrMIS_PLANTNO = "2";
                FnSMS_SMCDB_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                //FnSMS_SMCDB_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                StrPLANT = "CASTER-I"; StrPLANTNO = "1"; StrMIS_PLANTNO = "14";
                FnSMS_CasterI_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                //FnSMS_CasterI_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                StrPLANT = "CASTER-II"; StrPLANTNO = "2"; StrMIS_PLANTNO = "15";
                FnSMS_CasterII_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                //FnSMS_CasterII_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                StrPLANT = "CASTER-III"; StrPLANTNO = "3"; StrMIS_PLANTNO = "16";
                FnSMS_CasterIII_HeatTracking(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                //FnSMS_CasterIII_UpdateHeatEndTime(StrPLANT, StrPLANTNO, StrMIS_PLANTNO, StrDate);
                ////string StrUpdateQuery = "update  smsHeatTracker set actEnd='24' where actStart > actEnd and DateStamp<'" + StrDate + "'";
                ////bool StrUStatus = clsObj.DBInsertUpdateDeleteMIS(StrUpdateQuery);
                lblmsg.Text = "Completed";
            }
        }

        private void btnBlowArcData_Click(object sender, EventArgs e)
        {
            FnSMS_SMCDB_HeatBlow_Arc("2014-04-20");
        }
    }
}
