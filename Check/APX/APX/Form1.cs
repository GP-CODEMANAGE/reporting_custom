using Systems;TESTIN_G
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using System.Threading;
using System.Security.Principal;
//using Spire.Xls;
using System.Xml;
using System.Data.Common;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using System.Net;
using System.Configuration;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using OfficeOpenXml.Style;
using System.Net.Mail;
using OfficeOpenXml;
using Spire.Xls;
//using Spire.Xls;

namespace APXTESTIN_G
{
    public partial class Form1 : Form
    {
        public string execType = string.Empty;
        Logs LG = new Logs();
        public static string LogFileName = "LogFile " + DateTime.Now;
        public static string Logname = LogFileName.Replace(":", "-").Replace("/", "-");

        public static string vLogFile = AppDomain.CurrentDomain.BaseDirectory.ToString() + @"\LogFile" + @"\" + Logname + ".txt";

        bool bProceed = true;
        string strFilePath = ConfigurationManager.AppSettings["SharedirvePath"].ToString();
        public StreamWriter sw = null;
        int Pload = 0;
        int Tload = 0;
        DataTable dtExcel = null;
        Microsoft.Xrm.Sdk.IOrganizationService service = null;

        public Form1()
        {
            InitializeComponent();
            LG.CreateLogFile(vLogFile);
            lblMessage.Text = "";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //PositionErrorFileExcel(service, this.sw);


            dateTimePicker1.Value = DateTime.Today;
            txtHide.Text = dateTimePicker1.Value.ToString("MM/dd/yyyy");
            txtHide.Enabled = false;

            BindSecurityType();
            //   ReconExcel();
            //    generatesExcelsheets();


          //  generatesExcelsheets();

           //DataSet ds=  GetReportData();
           // GenerateExcel(ds);




        }

        public void BindSecurityType()
        {
            try
            {
                string Gresham_String = ConfigurationManager.AppSettings["Gresham_String_db"].ToString();
                string sqlstr = "SP_S_ACCOUNT_SOURCE";
                DataSet ds = new DataSet();

                SqlConnection conn = new SqlConnection(Gresham_String);
                SqlCommand cmd = new SqlCommand(sqlstr, conn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(ds);


                cbAccountSource.DataSource = ds.Tables[0].DefaultView;
                cbAccountSource.DisplayMember = "Name";
                cbAccountSource.ValueMember = "Id";


            }
            catch (Exception exc)
            {
                MessageBox.Show("Error occured: " + exc.Message + "  Inner Exeption: " + exc.InnerException);
            }
        }
        private void cbLoadType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void Loads()
        {
            lblSecurityMsg.Visible = false;
            dGVSecurity.Visible = false;
            //if (txtHide.Text != "" && cbUsedDate.Checked == true)
            //{
            execType = "B";

            btnLoad.Enabled = false;

            lblMessage.Text = "Loading...";
            lblMessage.Refresh();

            if (startTruncate())
            {

                #region CRMConnection
                try
                {

                    string Gresham_String = ConfigurationManager.AppSettings["Gresham_String_db"].ToString();
                    string strOrganizationUri = ConfigurationManager.AppSettings["strOrganizationUri"].ToString();
                    ClientCredentials Credentials = new ClientCredentials();
                    Credentials.Windows.ClientCredential = (NetworkCredential)CredentialCache.DefaultCredentials;

                    //Credentials.UserName.UserName = @"corp\gbhagia";
                    //Credentials.UserName.Password = "51ngl3malt";

                    Uri OrganizationUri = new Uri(strOrganizationUri);
                    Uri HomeRealmUri = null;

                    using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null))
                    {
                        // This statement is required to enable early-bound type support.
                        serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());
                        serviceProxy.Timeout = new TimeSpan(0, 20, 0);
                        service = (IOrganizationService)serviceProxy;
                    }
                }

                catch (System.Web.Services.Protocols.SoapException exc)
                {
                    bProceed = false;
                    string strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
                    LogMessage(sw, service, strDescription, 62, "GeneralError");
                    // LG.AddinLogFile(Form1.vLogFile, strDescription);
                }
                catch (Exception exc)
                {
                    bProceed = false;
                    string strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                    LogMessage(sw, service, strDescription, 62, "GeneralError");
                    // LG.AddinLogFile(Form1.vLogFile, strDescription);
                }
                #endregion


                bool result = Runjobs(service);
                //Pload = 1;
                //Tload = 1;
                // lblMessage.Text = "Job Run Completed..";
                // lblMessage.Refresh();

                if (result)
                {
                    if (cbUsedDate.Checked)
                    {
                        DataSet dsData = getDataSet("EXEC SP_S_NEW_SECURITY_LIST @AsOfDate='" + txtHide.Text + "'");

                        if (dsData.Tables[0].Rows.Count > 0)
                        {
                            dGVSecurity.Visible = true;
                            dGVSecurity.DataSource = dsData.Tables[0];
                            btnLoad.Visible = false;

                            btnContinue.Visible = true;
                            btnCancel.Visible = true;
                        }
                        else
                        {
                            lblMessage.Text = "Loading...";
                            lblMessage.Refresh();
                            dGVSecurity.Visible = false;
                            potfolioCodeUpdate(service);
                            saveToCrm(service);
                            if (result)
                                lblMessage.Text = "Load Completed...";
                            else
                                lblMessage.Text = "Job Failed";
                            lblMessage.Refresh();
                        }

                    }
                    else
                    {

                        DataSet dsData = getDataSet("EXEC SP_S_NEW_SECURITY_LIST");

                        lblMessage.Text = "Loading...";
                        lblMessage.Refresh();
                        dGVSecurity.Visible = false;
                        potfolioCodeUpdate(service);
                        saveToCrm(service);

                        if (result)
                            lblMessage.Text = "Load Completed...";
                        else
                            lblMessage.Text = "Job Failed";
                        lblMessage.Refresh();

                        if (dtExcel != null)
                        {
                            if (dtExcel.Rows.Count > 0)
                            {
                                lblSecurityMsg.Visible = true;
                                lblSecurityMsg.Text = @"New securities and\or security types were created. Please check your email and process them";
                                dGVSecurity.Visible = true;
                                dGVSecurity.DataSource = dtExcel;
                            }

                        }

                    }
                }
                else
                {
                    lblMessage.Text = "Job Failed";
                    lblMessage.Refresh();
                }
            }
            else
            {
                lblMessage.Text = "Load Fail. Please Conatct Administrator";
                lblMessage.Refresh();
            }

            btnLoad.Enabled = true;
            //}
            //else
            //    lblMessage.Text = "Please Select Date";
        }

        public bool Runjobs(IOrganizationService service)
        {

            string Gresham_String = null;
            string strDescription = "";


            Gresham_String = ConfigurationManager.AppSettings["Gresham_String_db"].ToString();

            SqlConnection Gresham_con = new SqlConnection(Gresham_String);
            DataTable dt = CreateDataTable();

            SqlConnection CRM_con;
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter dagersham = new SqlDataAdapter();
            SqlDataAdapter da_CRM;
            DataSet ds_gresham = new DataSet();
            DataSet ds = new DataSet();

            string greshamquery, vContain;
            int totalCount = 0;
            int successCount = 0;
            int failiureCount = 0;

            // strDescription = "Before Job Run";
            //LogMessage(sw, service, strDescription, 62, "GeneralError");


            #region JobLoad
            if (execType == "P")
            {
                try
                {

                    #region JOB For Position (6_29_2017 - sasmit)
                    Gresham_con = new SqlConnection(Gresham_String);
                    Gresham_con.Open();
                    greshamquery = "EXEC SP_S_ExecuteJobs @TypeId = 1";
                    cmd = new SqlCommand();
                    cmd.Connection = Gresham_con;
                    cmd.CommandText = greshamquery;
                    // cmd.CommandTimeout = 600;
                    cmd.ExecuteNonQuery();
                    Gresham_con.Close();
                    strDescription = "Job Load Completed.. EXEC SP_S_ExecuteJobs @TypeId = 1 ";
                    LG.AddinLogFile(Form1.vLogFile, strDescription);
                    Pload = 1;
                    Tload = 1;
                    #endregion

                }
                catch (Exception ex)
                {
                    //throw;
                    bProceed = false;
                    strDescription = "Load Job Failed to Execute for Positions" + ex.Message;
                    LogMessage(sw, service, strDescription, 64, "JobLoad");
                    // LG.AddinLogFile(Form1.vLogFile, strDescription);
                    Pload = 0;
                    Tload = 0;
                    return false;
                }
            }
            else if (execType == "T")
            {
                try
                {
                    if (mergeFiles()) //to add new line    
                    {

                        #region JOB For Position (6_29_2017 - sasmit)
                        Gresham_con = new SqlConnection(Gresham_String);
                        Gresham_con.Open();
                        greshamquery = "EXEC SP_S_ExecuteJobs @TypeId = 2";
                        cmd = new SqlCommand();
                        cmd.Connection = Gresham_con;
                        cmd.CommandText = greshamquery;
                        // cmd.CommandTimeout0 = 600;
                        cmd.ExecuteNonQuery();
                        Gresham_con.Close();
                        strDescription = "Job Load Completed.. EXEC SP_S_ExecuteJobs @TypeId = 2 ";
                        LG.AddinLogFile(Form1.vLogFile, strDescription);
                        Pload = 1;
                        Tload = 1;

                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    //throw;
                    bProceed = false;
                    strDescription = "Load Job Failed to Execute for Transactions-" + ex.Message;
                    LogMessage(sw, service, strDescription, 64, "JobLoad");
                    // LG.AddinLogFile(Form1.vLogFile, strDescription);
                    Pload = 0;
                    Tload = 0;
                    return false;
                }
            }
            else if (execType == "B")
            {
                try
                {
                    if (mergeFiles())   //to add new line 
                    // if (true)
                    {

                        vContain = "------- Before Exceuting JOB EXEC SP_S_ExecuteJobs @TypeId = 2  " + DateTime.Now.ToString() + "-------";
                        LG.AddinLogFile(Form1.vLogFile, vContain);

                        #region JOB For Position (6_29_2017 - sasmit)
                        Gresham_con = new SqlConnection(Gresham_String);
                        Gresham_con.Open();
                        greshamquery = "EXEC SP_S_ExecuteJobs @TypeId = 2";
                        cmd = new SqlCommand();
                        cmd.Connection = Gresham_con;
                        cmd.CommandText = greshamquery;
                        // cmd.CommandTimeout = 600;
                        cmd.ExecuteNonQuery();
                        Gresham_con.Close();
                        vContain = "------- After Exceuting JOB EXEC SP_S_ExecuteJobs @TypeId = 2" + DateTime.Now.ToString() + "-------";
                        LG.AddinLogFile(Form1.vLogFile, vContain);
                        // LG.AddinLogFile(Form1.vLogFile, strDescription);
                        Pload = 1;
                        Tload = 1;
                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    bProceed = false;
                    strDescription = "Load Job Failed to Execute for Transactions- EXEC SP_S_ExecuteJobs @TypeId = 2 " + ex.Message;
                    LogMessage(sw, service, strDescription, 64, "JobLoad");
                    // LG.AddinLogFile(Form1.vLogFile, strDescription);
                    // LG.AddinLogFile(Form1.vLogFile, strDescription);
                    Pload = 0;
                    Tload = 0;
                    return false;
                }

                try
                {

                    #region JOB For Position (6_29_2017 - sasmit)
                    Gresham_con = new SqlConnection(Gresham_String);
                    Gresham_con.Open();
                    greshamquery = "EXEC SP_S_ExecuteJobs @TypeId = 1";
                    cmd = new SqlCommand();
                    cmd.Connection = Gresham_con;
                    cmd.CommandText = greshamquery;
                    // cmd.CommandTimeout = 600;
                    cmd.ExecuteNonQuery();
                    Gresham_con.Close();
                    strDescription = "Load Job Execute for Transactions- EXEC SP_S_ExecuteJobs @TypeId = 1 ";
                    LG.AddinLogFile(Form1.vLogFile, strDescription);

                    Pload = 1;
                    Tload = 1;
                    #endregion
                }
                catch (Exception ex)
                {
                    bProceed = false;
                    strDescription = "Load Job Failed to Execute for Positions" + ex.Message;
                    LogMessage(sw, service, strDescription, 64, "JobLoad");
                    //LG.AddinLogFile(Form1.vLogFile, strDescription);
                    Pload = 0;
                    Tload = 0;
                    return false;
                }
            }

            return true;
            #endregion



        }

        public void potfolioCodeUpdate(IOrganizationService service)
        {
            if (chkHistoricalTrxnPstn.Checked)
            {
                #region Historical Transaction And Position
                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter dagersham = new SqlDataAdapter();
                try
                {
                    string Gresham_String = ConfigurationManager.AppSettings["Gresham_String_db"].ToString();

                    SqlConnection Gresham_con = new SqlConnection(Gresham_String);
                    DataTable dt = CreateDataTable();
                    SqlConnection CRM_con;
                    SqlDataAdapter da_CRM;
                    DataSet ds_gresham = new DataSet();
                    DataSet ds = new DataSet();

                    Gresham_con = new SqlConnection(Gresham_String);
                    Gresham_con.Open();
                    string greshamquery = "SP_U_TRASACTION_PORTCODE_WITH9_NEW_GA @HistoricalAccountFlg=1";
                    cmd = new SqlCommand();
                    cmd.Connection = Gresham_con;
                    cmd.CommandText = greshamquery;
                    cmd.CommandTimeout = 600;
                    cmd.ExecuteNonQuery();
                    Gresham_con.Close();

                    Gresham_con = new SqlConnection(Gresham_String);
                    Gresham_con.Open();
                    greshamquery = "SP_U_POSITION_PORTCODE_WITH9_NEW_GA @HistoricalAccountFlg=1";
                    cmd = new SqlCommand();
                    cmd.Connection = Gresham_con;
                    cmd.CommandText = greshamquery;
                    cmd.CommandTimeout = 600;
                    cmd.ExecuteNonQuery();
                    Gresham_con.Close();
                }
                catch (System.Web.Services.Protocols.SoapException exc)
                {
                    bProceed = false;
                    string strDescription = "Trxn and Position Portfolio Code Update failed. Error Detail:" + exc.Detail.InnerText;
                    LogMessage(sw, service, strDescription, 62, "Trxn_Position_PortCode_Update");
                }
                catch (Exception exc)
                {
                    bProceed = false;
                    string strDescription = "Trxn and Position Portfolio Code Update failed. Error Detail:" + exc.Message;
                    LogMessage(sw, service, strDescription, 62, "Trxn_Position_PortCode_Update");
                }
                finally
                {
                    cmd.Dispose();
                    dagersham.Dispose();
                }

                #endregion
            }
        }

        public void saveToCrm(IOrganizationService service)
        {
            // DataTable dtExcel = null;
            DataTable dtExcel1 = null;
            string greshamquery = "";
            string Gresham_String = ConfigurationManager.AppSettings["Gresham_String_db"].ToString();

            SqlConnection Gresham_con = new SqlConnection(Gresham_String);
            DataTable dt = CreateDataTable();
            SqlConnection CRM_con;
            SqlDataAdapter da_CRM;
            DataSet ds_gresham = new DataSet();
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter dagersham = new SqlDataAdapter();
            string strDescription = "";

            string vContain = "";
            int successCount = 0;
            int failiureCount = 0;
            int totalCount = 0;

            if (Pload == 1 && Tload == 1)
            {
                if (bProceed == true)
                {
                    #region Position AS OF DATE Update
                    try
                    {
                        greshamquery = "SP_U_POSITIONS_ASOFDATE_NEW_GA";
                        cmd.Connection = Gresham_con;
                        cmd.CommandText = greshamquery;
                        cmd.CommandTimeout = 600;
                        dagersham = new SqlDataAdapter(cmd);
                        ds = new DataSet();
                        dagersham.Fill(ds);

                        int result = Convert.ToInt32(ds.Tables[0].Rows[0]["UpdateCount"]);

                        vContain = "-------SP_U_POSITIONS_ASOFDATE_NEW_GA " + DateTime.Now.ToString() + "-------";
                        LG.AddinLogFile(Form1.vLogFile, vContain);
                    }
                    catch (System.Web.Services.Protocols.SoapException exc)
                    {
                        bProceed = false;
                        strDescription = "Position AS OF DATE Update failed. Error Detail:" + exc.Detail.InnerText;
                        LogMessage(sw, service, strDescription, 62, "Positions");
                        //vContain = "-------SP_U_POSITIONS_ASOFDATE_NEW_GA fail1" + DateTime.Now.ToString() + "-------" + exc.Message; ;
                        //LG.AddinLogFile(Form1.vLogFile, vContain);
                    }
                    catch (Exception exc)
                    {
                        bProceed = false;
                        strDescription = "Position AS OF DATE Update failed. Error Detail:" + exc.Message;
                        LogMessage(sw, service, strDescription, 62, "Positions");

                        //vContain = "-------SP_U_POSITIONS_ASOFDATE_NEW_GA fail1" + DateTime.Now.ToString() + "-------" + exc.Message; ;
                        //LG.AddinLogFile(Form1.vLogFile, vContain);
                    }
                    finally
                    {
                        cmd.Dispose();
                        dagersham.Dispose();
                        ds.Dispose();
                    }
                    #endregion
                }

                //label3.Text = "Custodian Load Starts";
                if (bProceed == true)
                {
                    #region Custodian

                    ////////////////////////////////// Update Starts //////////////////////////////////////
                    #region Update CUSTODIAN  //Commented
                    //try
                    //{
                    //    greshamquery = "SP_S_CUSTODIAN_UPDATE '" + execType + "'";
                    //    dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                    //    ds_gresham = new DataSet();
                    //    dagersham.Fill(ds_gresham);
                    //    totalCount = ds_gresham.Tables[0].Rows.Count;
                    //    sw.WriteLine("---------------------------- CUSTODIAN Update Starts -------------------");
                    //    sw.WriteLine("Update Starts for Custodian on: " + DateTime.Now.ToString());
                    //}
                    //catch (System.Web.Services.Protocols.SoapException exc)
                    //{
                    //    totalCount = 0;
                    //    strDescription = "Custodian Update failed, please contact administrator. Error Detail: " + exc.Detail.InnerText;
                    //    LogMessage(sw, service, strDescription, 62, "Custodian");
                    //}
                    //catch (Exception exc)
                    //{
                    //    totalCount = 0;
                    //    strDescription = "Custodian Update failed, please contact administrator. Error Detail: " + exc.Message;
                    //    LogMessage(sw, service, strDescription, 62, "Custodian");
                    //}

                    //for (int i = 0; i < totalCount; i++)
                    //{
                    //    try
                    //    {
                    //        Guid CustodianID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["accountid"]));
                    //        account objCustodian = new account();

                    //        objCustodian.accountid = new Key();
                    //        objCustodian.accountid.Value = CustodianID;

                    //        //CrmNumber crmInt = new CrmNumber();
                    //        //crmInt.Value = Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["SOURCE CODE"]);

                    //        objCustodian.name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SOURCE CODE TEXT"]);
                    //        //objCustodian.ssi_code = crmInt;
                    //        objCustodian.accountclassificationcode = new Picklist();
                    //        objCustodian.accountclassificationcode.Value = 200000;

                    //        service.Update(objCustodian);
                    //        ////Thread.Sleep(sleepTime);
                    //        successCount = successCount + 1;
                    //    }
                    //    catch (System.Web.Services.Protocols.SoapException exc)
                    //    {
                    //        failiureCount = failiureCount + 1;
                    //        strDescription = "Update failed for Custodian : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SOURCE CODE TEXT"]) + " Error Detail:" + exc.Detail.InnerText;
                    //        LogMessage(sw, service, strDescription, 25, "Custodian");
                    //    }
                    //    catch (Exception exc)
                    //    {
                    //        failiureCount = failiureCount + 1;
                    //        strDescription = "Update failed for Custodian : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SOURCE CODE TEXT"]) + " Error Detail:" + exc.Message;
                    //        LogMessage(sw, service, strDescription, 25, "Custodian");
                    //    }

                    //}

                    //sw.WriteLine("Update Ends for Custodian on: " + DateTime.Now.ToString());

                    //strDescription = "Total Custodian Update Failed: " + failiureCount;
                    //LogMessage(sw, service, strDescription, 39, "Custodian");

                    //strDescription = "Total Custodian Updated: " + successCount;
                    //LogMessage(sw, service, strDescription, 24, "Custodian");
                    //sw.WriteLine("---------------------------- Custodian Update Ends  -------------------");

                    #endregion
                    //////////////////////////////// Update Ends ////////////////////////////////////////
                    successCount = 0;
                    failiureCount = 0;
                    //////////////////////////////// Insert Starts ///////////////////////////////////////
                    #region Insert Household Custodian
                    try
                    {
                        greshamquery = "SP_S_CUSTODIAN_INSERT_NEW_GA '" + execType + "'";
                        dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                        ds_gresham = new DataSet();
                        dagersham.SelectCommand.CommandTimeout = 600;
                        dagersham.Fill(ds_gresham);
                        totalCount = ds_gresham.Tables[0].Rows.Count;

                        vContain = "---------------------------- New Custodian Insert Starts -------------------";
                        LG.AddinLogFile(Form1.vLogFile, vContain);
                        vContain = "Insert Starts for New Custodian on: " + DateTime.Now.ToString();
                        LG.AddinLogFile(Form1.vLogFile, vContain);

                    }
                    //catch (System.Web.Services.Protocols.SoapException exc)
                    //{
                    //    bProceed = false;
                    //    totalCount = 0;
                    //    strDescription = "Custodian Insert failed, please contact administrator. Error Detail: " + exc.Detail.InnerText;
                    //    LogMessage(sw, service, strDescription, 62, "Custodian");
                    //}
                    catch (Exception exc)
                    {
                        bProceed = false;
                        totalCount = 0;
                        strDescription = "Custodian Insert failed, please contact administrator. Error Detail: " + exc.Message;
                        LogMessage(sw, service, strDescription, 62, "Custodian");
                    }

                    if (bProceed == true)
                        for (int i = 0; i < totalCount; i++)
                        {
                            try
                            {
                                if (bProceed == true)
                                {
                                    //  account objAccount = new account();

                                    Microsoft.Xrm.Sdk.Entity objAccount = new Microsoft.Xrm.Sdk.Entity("account");


                                    //objAccount.accountid = new Key();
                                    //objAccount.accountid.Value = Guid.NewGuid();

                                    //CrmNumber crmInt = new CrmNumber();
                                    //crmInt.Value = Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["SOURCE CODE"]);

                                    // objAccount.name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SOURCE CODE TEXT"]);
                                    objAccount["name"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SOURCE CODE TEXT"]);

                                    //objAccount.ssi_code = crmInt;
                                    //objAccount.accountclassificationcode = new Picklist();
                                    //objAccount.accountclassificationcode.Value = 200000;
                                    objAccount["accountclassificationcode"] = new Microsoft.Xrm.Sdk.OptionSetValue(200000);

                                    // objAccount["accountclassificationcode"] = 200000;


                                    //  service.Create(objAccount);
                                    Guid newAccountId = service.Create(objAccount);

                                    //Thread.Sleep(sleepTime);
                                    successCount = successCount + 1;
                                    strDescription = "New Custodian inserted: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SOURCE CODE TEXT"]);
                                    LogMessage(sw, service, strDescription, 21, "Custodian");
                                }
                                else
                                    break;
                            }
                            catch (System.Web.Services.Protocols.SoapException exc)
                            {
                                bProceed = true;
                                failiureCount = failiureCount + 1;
                                strDescription = "Insert failed for Custodian : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SOURCE CODE TEXT"]) + "Error Detail: " + exc.Detail.InnerText;
                                LogMessage(sw, service, strDescription, 22, "Custodian");
                                AddException("Insert failed for Custodian", "SOURCE CODE TEXT : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SOURCE CODE TEXT"]), "", "", "", "", "", "", "", exc.Message, dt);
                            }
                            catch (Exception exc)
                            {
                                bProceed = true;
                                failiureCount = failiureCount + 1;
                                strDescription = "Insert failed for Custodian : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SOURCE CODE TEXT"]) + exc.Message;
                                LogMessage(sw, service, strDescription, 22, "Custodian");
                                AddException("Insert failed for Custodian", "SOURCE CODE TEXT : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SOURCE CODE TEXT"]), "", "", "", "", "", "", "", exc.Message, dt);
                            }
                        }

                    LG.AddinLogFile(Form1.vLogFile, "Insert Ends for Custodian on: " + DateTime.Now.ToString());

                    strDescription = "Total Custodian failed to insert: " + failiureCount;
                    LogMessage(sw, service, strDescription, 38, "Custodian");

                    strDescription = "Total Custodian inserted: " + successCount;
                    LogMessage(sw, service, strDescription, 23, "Custodian");
                    LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Custodian Insert End -------------------");
                    LG.AddinLogFile(Form1.vLogFile, " ");
                    ds_gresham.Dispose();
                    dagersham.Dispose();

                    #endregion
                    ////////////////////////////////// Insert Ends ////////////////////////////////////////

                    #endregion
                }
                successCount = 0;
                failiureCount = 0;

                if (bProceed == true)
                {
                    //label3.Text = "Account Load Starts";
                    #region Accounts
                    //////////////////////////////// Insert Starts ///////////////////////////////////////
                    #region Insert Accounts
                    try
                    {
                        greshamquery = "SP_S_ACCOUNT_INSERT_NEW_GA '" + execType + "'";
                        dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                        ds_gresham = new DataSet();
                        dagersham.SelectCommand.CommandTimeout = 600;
                        dagersham.Fill(ds_gresham);
                        totalCount = ds_gresham.Tables[0].Rows.Count;
                        //  sw.WriteLine("---------------------------- New Accounts Insert Starts -------------------");
                        LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Accounts Insert Starts -------------------");
                        LG.AddinLogFile(Form1.vLogFile, "Insert Starts for New Acoounts on: " + DateTime.Now.ToString());


                    }
                    catch (System.Web.Services.Protocols.SoapException exc)
                    {
                        bProceed = false;
                        totalCount = 0;
                        strDescription = "Account Insert failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                        LogMessage(sw, service, strDescription, 62, "Account");

                    }
                    catch (Exception exc)
                    {
                        bProceed = false;
                        totalCount = 0;
                        strDescription = "Account Insert failed, please contact administrator. Error Detail:" + exc.Message;
                        LogMessage(sw, service, strDescription, 62, "Account");
                    }
                    finally
                    {
                        ds_gresham.Dispose();
                        dagersham.Dispose();
                    }

                    if (bProceed == true)
                        for (int i = 0; i < totalCount; i++)
                        {
                            try
                            {
                                if (bProceed == true)
                                {

                                    Entity objAccount = new Entity("ssi_account");
                                    //  ssi_account objAccount = new ssi_account();

                                    //objAccount.ssi_accountid = new Key();
                                    //objAccount.ssi_accountid.Value = Guid.NewGuid();
                                    objAccount["ssi_source"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(cbAccountSource.SelectedValue));
                                    //objAccount.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]);
                                    objAccount["ssi_name"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]);

                                    //objAccount.ssi_name1 = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME1"]);
                                    objAccount["ssi_name1"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME1"]);

                                    //objAccount.ssi_name2 = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME2"]);
                                    objAccount["ssi_name2"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME2"]);

                                    // objAccount.ssi_name3 = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME3"]);
                                    objAccount["ssi_name3"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME3"]);

                                    //Added by Dhaval --24-July-2015
                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_custodianid"]) != "")
                                    {
                                        //objAccount.ssi_custodianid = new Lookup();
                                        //objAccount.ssi_custodianid.type = EntityName.account.ToString();
                                        //objAccount.ssi_custodianid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_custodianid"]));
                                        objAccount["ssi_custodianid"] = new EntityReference("account", new Guid(Convert.ToString(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_custodianid"]))));

                                    }

                                    // objAccount.ssi_custodianaccountnumber = Convert.ToString(ds_gresham.Tables[0].Rows[i]["CUSTODIANACCOUNTNMB"]);
                                    //  objAccount.ssi_custodianaccountnumber = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME3"]);
                                    objAccount["ssi_custodianaccountnumber"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME3"]);

                                    if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["AccountType"]) != "")
                                    {

                                        objAccount["ssi_accounttype"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["AccountType"]));

                                    }
                                    //  service.Create(objAccount);
                                    Guid newAccountId = service.Create(objAccount);

                                    //Thread.Sleep(sleepTime);
                                    successCount = successCount + 1;
                                    strDescription = "New Account inserted: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]);
                                    LogMessage(sw, service, strDescription, 1, "Account");
                                }
                                else
                                    break;
                            }
                            catch (System.Web.Services.Protocols.SoapException exc)
                            {
                                bProceed = true;
                                failiureCount = failiureCount + 1;
                                strDescription = " Insert failed for Account : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]) + " Error Detail: " + exc.Detail.InnerText;
                                LogMessage(sw, service, strDescription, 2, "Account");
                                AddException("Insert failed for Account", "PORT CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]), "", "", "", "", "", "", "", exc.Message, dt);

                            }
                            catch (Exception exc)
                            {
                                bProceed = true;
                                failiureCount = failiureCount + 1;
                                LG.AddinLogFile(Form1.vLogFile, "Insert failed for Account : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]));
                                strDescription = "Insert failed for Account : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]) + " Error Detail: " + exc.Message;
                                LogMessage(sw, service, strDescription, 2, "Account");
                                AddException("Insert failed for Account", "PORT CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]), "", "", "", "", "", "", "", exc.Message, dt);
                            }
                            finally
                            {
                                ds_gresham.Dispose();
                                dagersham.Dispose();
                            }
                        }

                    LG.AddinLogFile(Form1.vLogFile, "Insert Ends for Acoount on: " + DateTime.Now.ToString());

                    strDescription = "Total Accounts inserted: " + successCount;
                    LogMessage(sw, service, strDescription, 3, "Account");

                    /////// Total Insert Failed
                    strDescription = "Accounts failed to insert: " + failiureCount;
                    LogMessage(sw, service, strDescription, 30, "Account");

                    LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Account Insert End -------------------");
                    LG.AddinLogFile(Form1.vLogFile, " ");
                    #endregion
                    //////////////////////////////// Insert Ends /////////////////////////////////////////

                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    if (bProceed == true)
                    {
                        //////////////////////////////// Update Starts ///////////////////////////////////////
                        #region Update Accounts
                        try
                        {
                            greshamquery = "SP_S_ACCOUNT_UPDATE_NEW_GA '" + execType + "'";
                            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                            ds_gresham = new DataSet();
                            dagersham.SelectCommand.CommandTimeout = 600;
                            dagersham.Fill(ds_gresham);
                            totalCount = ds_gresham.Tables[0].Rows.Count;
                            // sw.WriteLine("---------------------------- Account Update Starts -------------------");
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- Account Update Starts -------------------");
                            // sw.WriteLine("Update Starts for Accounts on: " + DateTime.Now.ToString());
                            LG.AddinLogFile(Form1.vLogFile, "Update Starts for Accounts on: " + DateTime.Now.ToString());
                        }
                        catch (System.Web.Services.Protocols.SoapException exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Account Update failed, please contact administrator." + exc.Detail.InnerText;
                            LogMessage(sw, service, strDescription, 62, "Account");
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Account Update failed, please contact administrator." + exc.Message;
                            LogMessage(sw, service, strDescription, 62, "Account");
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {


                                        // ssi_account objAccount = new ssi_account();
                                        Entity objAccount = new Entity("ssi_account");

                                        Guid AccountID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]));
                                        objAccount["ssi_accountid"] = AccountID;

                                        // objAccount.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]);
                                        objAccount["ssi_name"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]); ;

                                        //objAccount.ssi_name1 = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME1"]);
                                        objAccount["ssi_name1"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME1"]);

                                        //objAccount.ssi_name2 = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME2"]);
                                        objAccount["ssi_name2"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME2"]);

                                        //objAccount.ssi_name3 = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME3"]);
                                        objAccount["ssi_name3"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME3"]);

                                        //objAccount.ssi_custodianaccountnumber = Convert.ToString(ds_gresham.Tables[0].Rows[i]["CUSTODIANACCOUNTNMB"]);

                                        //objAccount.ssi_accountnumber = Convert.ToString(ds_gresham.Tables[0].Rows[i]["CUSTODIANACCOUNTNMB"]);
                                        objAccount["ssi_accountnumber"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["CUSTODIANACCOUNTNMB"]);

                                        // AS OF DATE
                                        //objAccount.ssi_asofdate = new CrmDateTime();
                                        //objAccount.ssi_asofdate.Value = Convert.ToString(DateTime.Today.ToShortDateString());
                                        objAccount["ssi_asofdate"] = DateTime.Now;   // Convert.ToDateTime(Convert.ToString(DateTime.Today.ToShortDateString()));

                                        //if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_custodianid"]) != "")
                                        //{
                                        //    objAccount.ssi_custodianid = new Lookup();
                                        //    objAccount.ssi_custodianid.type = EntityName.account.ToString();
                                        //    objAccount.ssi_custodianid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_custodianid"]));
                                        //}

                                        service.Update(objAccount);
                                        //Thread.Sleep(sleepTime);

                                        successCount = successCount + 1;
                                    }
                                    else
                                        break;
                                }
                                catch (System.Web.Services.Protocols.SoapException exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Update failed for Account " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]) + " Error Detail: " + exc.Message + " " + exc.Detail.InnerText;
                                    LogMessage(sw, service, strDescription, 5, "Account");
                                    AddException("Update failed for Account", "PORT CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]), "", "", "", "", "", "", "", exc.Message, dt);

                                }
                                catch (Exception exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Update failed for Account : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]) + " Error Detail: " + exc.Message;
                                    LogMessage(sw, service, strDescription, 5, "Account");
                                    AddException("Update failed for Account", "PORT CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]), "", "", "", "", "", "", "", exc.Message, dt);
                                }
                            }

                        ds_gresham.Dispose();
                        dagersham.Dispose();

                        //sw.WriteLine("Update Ends for Acoount on: " + DateTime.Now.ToString());
                        LG.AddinLogFile(Form1.vLogFile, "Update Ends for Acoount on: " + DateTime.Now.ToString());
                        strDescription = "Total Accounts Failed to Update: " + failiureCount;
                        LogMessage(sw, service, strDescription, 31, "Account");

                        strDescription = "Total Accounts Updated: " + successCount;
                        LogMessage(sw, service, strDescription, 4, "Account");
                        // sw.WriteLine("---------------------------- Account Update Ends  -------------------");
                        LG.AddinLogFile(Form1.vLogFile, "---------------------------- Account Update Ends  -------------------");

                        #endregion
                    }

                    if (execType == "B" || execType == "P")
                    {
                        successCount = 0;
                        failiureCount = 0;
                        totalCount = 0;

                        if (bProceed == true)
                        {
                            #region Update Account for Report Code
                            try
                            {
                                greshamquery = "SP_S_ACCOUNT_UPDATE_ReportCode_NEW_GA '" + execType + "'";
                                dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                ds_gresham = new DataSet();
                                dagersham.SelectCommand.CommandTimeout = 600;
                                dagersham.Fill(ds_gresham);
                                totalCount = ds_gresham.Tables[0].Rows.Count;
                                //  sw.WriteLine("---------------------------- Account Update for ReportCode Starts -------------------");
                                LG.AddinLogFile(Form1.vLogFile, "---------------------------- Account Update for ReportCode Starts -------------------");
                                // sw.WriteLine("Update Starts for Accounts ReportCode on: " + DateTime.Now.ToString());
                                LG.AddinLogFile(Form1.vLogFile, "Update Starts for Accounts ReportCode on: " + DateTime.Now.ToString());
                            }
                            catch (System.Web.Services.Protocols.SoapException exc)
                            {
                                bProceed = false;
                                totalCount = 0;
                                strDescription = "Account Update for ReportCode failed, please contact administrator. Error Detail: " + exc.Detail.InnerText;
                                LG.AddinLogFile(Form1.vLogFile, strDescription);
                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                totalCount = 0;
                                strDescription = "Account Update for ReportCode failed, please contact administrator. Error Detail: " + exc.Message;
                                LG.AddinLogFile(Form1.vLogFile, strDescription);
                            }

                            if (bProceed == true)
                                for (int i = 0; i < totalCount; i++)
                                {
                                    try
                                    {
                                        if (bProceed == true)
                                        {
                                            Guid AccountID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]));

                                            //ssi_account objAccount = new ssi_account();
                                            Entity objAccount = new Entity("ssi_account");

                                            //objAccount.ssi_accountid = new Key();
                                            //objAccount.ssi_accountid.Value = AccountID;
                                            objAccount["ssi_accountid"] = AccountID;

                                            //objAccount.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]);
                                            //objAccount.ssi_adventreportcode = Convert.ToString(ds_gresham.Tables[0].Rows[i]["report code"]);
                                            objAccount["ssi_adventreportcode"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["report code"]);

                                            // Managed flag
                                            //objAccount.ssi_managedcode = new CrmBoolean();
                                            //objAccount.ssi_managedcode.Value = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["MANAGED"]).ToLower());
                                            objAccount["ssi_managedcode"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["MANAGED"]).ToLower());

                                            // Start date
                                            //objAccount.ssi_startdate = new CrmDateTime();
                                            //objAccount.ssi_startdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["START DATE"]);
                                            objAccount["ssi_startdate"] = Convert.ToDateTime(Convert.ToString(ds_gresham.Tables[0].Rows[i]["START DATE"]));

                                            service.Update(objAccount);
                                            //Thread.Sleep(sleepTime);

                                            successCount = successCount + 1;
                                        }
                                        else
                                            break;
                                    }
                                    catch (System.Web.Services.Protocols.SoapException exc)
                                    {
                                        bProceed = true;
                                        failiureCount = failiureCount + 1;
                                        strDescription = "Update failed for Account ReportCode " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]) + " Error Detail: " + exc.Detail.InnerText;
                                        LogMessage(sw, service, strDescription, 5, "Account");
                                        AddException("Update failed for Account ReportCode", "PORT CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]), "", "", "", "", "", "", "", exc.Message, dt);

                                    }
                                    catch (Exception exc)
                                    {
                                        bProceed = true;
                                        failiureCount = failiureCount + 1;
                                        strDescription = "Update failed for Account ReportCode : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]) + " Error Detail: " + exc.Message;
                                        LogMessage(sw, service, strDescription, 5, "Account");
                                        AddException("Update failed for Account ReportCode", "PORT CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]), "", "", "", "", "", "", "", exc.Message, dt);
                                    }
                                }

                            ds_gresham.Dispose();
                            dagersham.Dispose();

                            //    sw.WriteLine("Update Ends for Account ReportCode on: " + DateTime.Now.ToString());
                            LG.AddinLogFile(Form1.vLogFile, "Update Ends for Account ReportCode on: " + DateTime.Now.ToString());

                            strDescription = "Total Account ReportCode Failed to Update: " + failiureCount;
                            LogMessage(sw, service, strDescription, 31, "Account");

                            strDescription = "Total Account ReportCode Updated: " + successCount;
                            LogMessage(sw, service, strDescription, 4, "Account");
                            //  sw.WriteLine("---------------------------- Account ReportCode Update Ends  -------------------");
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- Account ReportCode Update Ends  -------------------");

                            #endregion
                        }
                    }

                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    if (bProceed == true)
                    {
                        #region Update Account for Closed Date and Closed Flag
                        try
                        {
                            greshamquery = "SP_S_CLOSED_ACCOUNT_UPDATE_NEW_GA";
                            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                            ds_gresham = new DataSet();
                            dagersham.SelectCommand.CommandTimeout = 600;
                            dagersham.Fill(ds_gresham);
                            totalCount = ds_gresham.Tables[0].Rows.Count;
                            //sw.WriteLine("---------------------------- Account Update for Closed Date Starts -------------------");
                            //sw.WriteLine("Update Starts for Accounts Closed Date on: " + DateTime.Now.ToString());

                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- Account Update for Closed Date Starts -------------------");
                            LG.AddinLogFile(Form1.vLogFile, "Update Starts for Accounts Closed Date on: " + DateTime.Now.ToString());
                        }
                        catch (System.Web.Services.Protocols.SoapException exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Account Update for Closed Date failed, please contact administrator. Error Detail: " + exc.Detail.InnerText;
                            LG.AddinLogFile(Form1.vLogFile, "Account Update for Closed Date failed, please contact administrator.");
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Account Update for Closed Date failed, please contact administrator. Error Detail: " + exc.Message;
                            LG.AddinLogFile(Form1.vLogFile, "Account Update for Closed Date failed, please contact administrator.");
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        Guid AccountID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]));

                                        //  ssi_account objAccount = new ssi_account();
                                        Entity objAccount = new Entity("ssi_account");

                                        //objAccount.ssi_accountid = new Key();
                                        //objAccount.ssi_accountid.Value = AccountID;
                                        objAccount["ssi_accountid"] = AccountID;

                                        // Closed date
                                        //objAccount.ssi_terminationdate = new CrmDateTime();
                                        //objAccount.ssi_terminationdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Closed Date"]);
                                        objAccount["ssi_terminationdate"] = Convert.ToDateTime(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Closed Date"]));

                                        // Closed flag
                                        //objAccount.ssi_accountclosed = new CrmBoolean();
                                        //objAccount.ssi_accountclosed.Value = true;
                                        objAccount["ssi_terminationdate"] = true;

                                        service.Update(objAccount);
                                        //Thread.Sleep(sleepTime);

                                        successCount = successCount + 1;
                                    }
                                    else
                                        break;
                                }
                                catch (System.Web.Services.Protocols.SoapException exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Update failed for Account Closed Date " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]) + " Error Detail: " + exc.Detail.InnerText;
                                    LogMessage(sw, service, strDescription, 5, "Account");
                                    AddException("Update failed for Account Closed Date", "PORT CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]), "", "", "", "", "", "", "", exc.Message, dt);

                                }
                                catch (Exception exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Update failed for Account Closed Date : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]) + " Error Detail: " + exc.Message;
                                    LogMessage(sw, service, strDescription, 5, "Account");
                                    AddException("Update failed for Account Closed Date", "PORT CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["PORT CODE"]), "", "", "", "", "", "", "", exc.Message, dt);
                                }
                            }

                        ds_gresham.Dispose();
                        dagersham.Dispose();

                        LG.AddinLogFile(Form1.vLogFile, "Update Ends for Account Closed Date on: " + DateTime.Now.ToString());
                        strDescription = "Total Account Failed to Update for Closed Date: " + failiureCount;
                        LogMessage(sw, service, strDescription, 31, "Account");

                        strDescription = "Total Account Updated for Closed Date: " + successCount;
                        LogMessage(sw, service, strDescription, 4, "Account");
                        LG.AddinLogFile(Form1.vLogFile, "---------------------------- Account Closed Date Update Ends  -------------------");

                        #endregion
                    }

                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    if (bProceed == true)
                    {
                        #region LookThroughAccount Group Update for EndDate
                        try
                        {
                            greshamquery = @"SP_S_LOOKTHROUGH_CLOSEDDDATE_UPDATE_NEW_GA";

                            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                            ds_gresham = new DataSet();
                            dagersham.SelectCommand.CommandTimeout = 600;
                            dagersham.Fill(ds_gresham);
                            totalCount = ds_gresham.Tables[0].Rows.Count;
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- LookThroughAccount Group Update Starts -------------------");
                            LG.AddinLogFile(Form1.vLogFile, "Update Starts for LookThroughAccount Group on: " + DateTime.Now.ToString());
                            ////Console.WriteLine("gresham Dataset Built for Account Update ");
                        }
                        catch (System.Web.Services.Protocols.SoapException exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "LookThroughAccount Group Update failed, please contact administrator." + exc.Detail.InnerText;
                            LogMessage(sw, service, strDescription, 62, "LookThroughAccountGroup");
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "LookThroughAccount Group Update failed, please contact administrator." + exc.Message;
                            LogMessage(sw, service, strDescription, 62, "LookThroughAccountGroup");
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        Guid LookThroughID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lookthroughaccountgroupid"]));

                                        //ssi_lookthroughaccountgroup objLookThrough = new ssi_lookthroughaccountgroup();
                                        Entity objLookThrough = new Entity("ssi_lookthroughaccountgroup");


                                        //objLookThrough.ssi_lookthroughaccountgroupid = new Key();
                                        //objLookThrough.ssi_lookthroughaccountgroupid.Value = LookThroughID;

                                        objLookThrough["ssi_lookthroughaccountgroupid"] = LookThroughID;

                                        //objLookThrough.ssi_enddate = new CrmDateTime();
                                        //objLookThrough.ssi_enddate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Closed Date"]);
                                        objLookThrough["ssi_enddate"] = Convert.ToDateTime(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Closed Date"]));


                                        service.Update(objLookThrough);
                                        //Thread.Sleep(sleepTime);

                                        successCount = successCount + 1;
                                        //Console.WriteLine(i.ToString() + " Updated");
                                    }
                                    else
                                        break;
                                }
                                catch (System.Web.Services.Protocols.SoapException exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Update failed for LookThroughAccount Group " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lookthroughaccountgroupid"]) + " Error Detail: " + exc.Message + " " + exc.Detail.InnerText;
                                    LogMessage(sw, service, strDescription, 65, "Account");
                                    AddException("Update failed for LookThroughAccount Group", "LookThrough Account Group Id : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lookthroughaccountgroupid"]), "Closed Date : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Closed Date"]), "", "", "", "", "", "", exc.Message, dt);

                                }
                                catch (Exception exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Update failed for LookThroughAccount Group : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lookthroughaccountgroupid"]) + " Error Detail: " + exc.Message;
                                    LogMessage(sw, service, strDescription, 65, "LookThroughAccountGroup");
                                    AddException("Update failed for LookThroughAccount Group", "LookThrough Account Group Id : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_lookthroughaccountgroupid"]), "Closed Date : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Closed Date"]), "", "", "", "", "", "", exc.Message, dt);
                                }
                            }

                        ds_gresham.Dispose();
                        dagersham.Dispose();

                        LG.AddinLogFile(Form1.vLogFile, "Update Ends for LookThroughAccount Group on: " + DateTime.Now.ToString());
                        strDescription = "Total LookThroughAccount Group Failed to Update: " + failiureCount;
                        LogMessage(sw, service, strDescription, 62, "LookThroughAccountGroup");

                        strDescription = "Total LookThroughAccount Group Updated: " + successCount;
                        LogMessage(sw, service, strDescription, 65, "LookThroughAccountGroup");
                        LG.AddinLogFile(Form1.vLogFile, "---------------------------- LookThroughAccount Group Update Ends  -------------------");

                        #endregion
                    }
                    //////////////////////////////// Update Ends /////////////////////////////////////////
                    #endregion
                }
                successCount = 0;
                failiureCount = 0;
                totalCount = 0;

                if (bProceed == true)
                {
                    //label3.Text = "Asset Load Starts";
                    #region Asset Class

                    if (bProceed == true)
                    {
                        //////////////////////////////// Update Starts ///////////////////////////////////
                        #region Update Asset Class Commented
                        //try
                        //{
                        //    greshamquery = "SP_S_ASSET_UPDATE '" + execType + "'";
                        //    dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                        //    ds_gresham = new DataSet();
                        //    dagersham.Fill(ds_gresham);
                        //    totalCount = ds_gresham.Tables[0].Rows.Count;
                        //    sw.WriteLine("---------------------------- Asset Class Update Starts -------------------");
                        //    sw.WriteLine("Update Starts for Asset Class on: " + DateTime.Now.ToString());

                        //    //lblMessage.Text = "gresham Dataset Built for Asset Class Update ";
                        //}
                        //catch (System.Web.Services.Protocols.SoapException exc)
                        //{
                        //    totalCount = 0;
                        //    strDescription = "Asset Class Update failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                        //    LogMessage(sw, service, strDescription, 62, "Asset Class");
                        //}
                        //catch (Exception exc)
                        //{
                        //    totalCount = 0;
                        //    strDescription = "Asset Class Update failed, please contact administrator. Error Detail:" + exc.Message;
                        //    LogMessage(sw, service, strDescription, 62, "Asset Class");
                        //}

                        //for (int i = 0; i < totalCount; i++)
                        //{
                        //    try
                        //    {
                        //        sas_assetclass objAssetClass = new sas_assetclass();

                        //        Guid AssetClassID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]));

                        //        objAssetClass.sas_assetclassid = new Key();
                        //        objAssetClass.sas_assetclassid.Value = AssetClassID;

                        //        objAssetClass.sas_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]);
                        //        objAssetClass.ssi_code = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS CODE"]);

                        //        service.Update(objAssetClass);
                        //        //Thread.Sleep(sleepTime);
                        //        successCount = successCount + 1;
                        //        ////Console.WriteLine(i.ToString() + " Updated");
                        //    }
                        //    catch (System.Web.Services.Protocols.SoapException exc)
                        //    {
                        //        failiureCount = failiureCount + 1;
                        //        strDescription = "Update failed for Asset Class : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]) + " Error Detail: " + exc.Detail.InnerText;
                        //        LogMessage(sw, service, strDescription, 20, "Asset Class");
                        //    }
                        //    catch (Exception exc)
                        //    {
                        //        ////Console.WriteLine("Update failed for Asset Class : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]));
                        //        failiureCount = failiureCount + 1;
                        //        strDescription = "Update failed for Asset Class : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]) + " Error Detail: " + exc.Message;
                        //        LogMessage(sw, service, strDescription, 20, "Asset Class");
                        //    }
                        //}

                        //sw.WriteLine("Update Ends for Asset Class on: " + DateTime.Now.ToString());

                        //strDescription="Total Asset Class Failed to Update: " + failiureCount;
                        //LogMessage(sw,service,strDescription,37,"Asset Class");
                        //strDescription="Total Asset Class Updated: " + successCount;
                        //LogMessage(sw,service,strDescription,19,"Asset Class");

                        //ds_gresham.Dispose();
                        //dagersham.Dispose();
                        //sw.WriteLine("---------------------------- Asset Class Update Ends  -------------------");

                        #endregion
                        //////////////////////////////// Update Ends ///////////////////////////////////
                    }
                    successCount = 0;
                    failiureCount = 0;
                    if (bProceed == true)
                    {
                        //////////////////////////////// Insert Starts ///////////////////////////////////////
                        #region Insert AssetClass
                        try
                        {
                            greshamquery = "SP_S_ASSET_INSERT_NEW_GA '" + execType + "'";
                            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                            ds_gresham = new DataSet();
                            dagersham.SelectCommand.CommandTimeout = 600;
                            dagersham.Fill(ds_gresham);
                            totalCount = ds_gresham.Tables[0].Rows.Count;
                            ////Console.WriteLine("gresham Dataset Built for Asset Class Insert ");
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Asset Class Insert Starts -------------------");
                            LG.AddinLogFile(Form1.vLogFile, "Insert Starts for New Asset Class on: " + DateTime.Now.ToString());
                        }
                        catch (System.Web.Services.Protocols.SoapException exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Asset Class Insert failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                            LogMessage(sw, service, strDescription, 62, "Asset Class");
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Asset Class Insert failed, please contact administrator. Error Detail:" + exc.Message;
                            LogMessage(sw, service, strDescription, 62, "Asset Class");
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        // sas_assetclass objAssetClass = new sas_assetclass();
                                        Entity objAssetClass = new Entity("sas_assetclass");

                                        //objAssetClass.sas_assetclassid = new Key();
                                        //objAssetClass.sas_assetclassid.Value = Guid.NewGuid();
                                        // objAssetClass["sas_assetclassid"] 

                                        //   objAssetClass.sas_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]);
                                        objAssetClass["sas_name"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]);
                                        //  objAssetClass.ssi_code = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS CODE"]);
                                        objAssetClass["ssi_code"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS CODE"]);


                                        //service.Create(objAssetClass);
                                        Guid newAssetClassId = service.Create(objAssetClass);

                                        //Thread.Sleep(sleepTime);
                                        ////Console.WriteLine(i.ToString() + " Inserted");
                                        successCount = successCount + 1;
                                        strDescription = "New Asset Class inserted: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]);
                                        LogMessage(sw, service, strDescription, 16, "Asset Class");
                                    }
                                    else
                                        break;
                                }
                                catch (System.Web.Services.Protocols.SoapException exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Insert failed for Asset Class : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]) + " Error Detail: " + exc.Detail.InnerText;
                                    LogMessage(sw, service, strDescription, 17, "Asset Class");
                                    AddException("Insert failed for Asset Class", "ASSET CLASS NAME : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]), "ASSET CLASS CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS CODE"]), "", "", "", "", "", "", exc.Message, dt);

                                }
                                catch (Exception exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Insert failed for Asset Class : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]) + " Error Detail: " + exc.Message;
                                    LogMessage(sw, service, strDescription, 17, "Asset Class");
                                    AddException("Insert failed for Asset Class", "ASSET CLASS NAME : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS NAME"]), "ASSET CLASS CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ASSET CLASS CODE"]), "", "", "", "", "", "", exc.Message, dt);

                                }
                            }

                        LG.AddinLogFile(Form1.vLogFile, "Insert Ends for Asset Class on: " + DateTime.Now.ToString());

                        strDescription = "Total New Asset Class failed to insert: " + failiureCount;
                        LogMessage(sw, service, strDescription, 36, "Asset Class");
                        strDescription = "Total New Asset Class inserted: " + successCount;
                        LogMessage(sw, service, strDescription, 18, "Asset Class");

                        ds_gresham.Dispose();
                        dagersham.Dispose();
                        LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Asset Class Insert End -------------------");
                        LG.AddinLogFile(Form1.vLogFile, " ");

                        #endregion
                        //////////////////////////////// Insert Ends ///////////////////////////////////////
                    }
                    #endregion
                }
                successCount = 0;
                failiureCount = 0;

                if (bProceed == true)
                {
                    //label3.Text = "Security Type Load Starts";
                    #region Security Type
                    if (bProceed == true)
                    {
                        //////////////////////////////// Update Starts /////////////////////////////////////
                        #region Update Security Type Commented
                        //try
                        //{
                        //    greshamquery = "SP_S_SECURITY_TYPE_UPDATE '" + execType + "'";
                        //    dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                        //    ds_gresham = new DataSet();
                        //    dagersham.Fill(ds_gresham);
                        //    totalCount = ds_gresham.Tables[0].Rows.Count;
                        //    sw.WriteLine("---------------------------- Security Type Update Starts -------------------");
                        //    sw.WriteLine("Update Starts for Security Type on: " + DateTime.Now.ToString());
                        //    //Console.WriteLine("gresham Dataset Built for Security Type Update ");
                        //}
                        //catch (System.Web.Services.Protocols.SoapException exc)
                        //{
                        //    totalCount = 0;
                        //    strDescription = "Security Type Update failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                        //    LogMessage(sw, service, strDescription, 62, "Security Type");
                        //}
                        //catch (Exception exc)
                        //{
                        //    totalCount = 0;
                        //    strDescription = "Security Type Update failed, please contact administrator. Error Detail:" + exc.Message;
                        //    LogMessage(sw, service, strDescription, 62, "Security Type");
                        //}

                        //for (int i = 0; i < totalCount; i++)
                        //{
                        //    try
                        //    {
                        //        Guid SecurityTypeID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_SECURITYTYPEID"]));

                        //        ssi_securitytype objSecurityType = new ssi_securitytype();

                        //        objSecurityType.ssi_securitytypeid = new Key();
                        //        objSecurityType.ssi_securitytypeid.Value = SecurityTypeID;

                        //        objSecurityType.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC TYPE"]);
                        //        objSecurityType.ssi_symbol = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC Type"]);

                        //        service.Update(objSecurityType);
                        //        //Thread.Sleep(sleepTime);
                        //        successCount = successCount + 1;
                        //        //Console.WriteLine(i.ToString() + " Updated");
                        //    }
                        //    catch (System.Web.Services.Protocols.SoapException exc)
                        //    {
                        //        failiureCount = failiureCount + 1;
                        //        strDescription = "Update failed for Security Type : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC TYPE"]) + " Error Detail: " + exc.Detail.InnerText;
                        //        LogMessage(sw, service, strDescription, 15, "Security Type");
                        //    }
                        //    catch (Exception exc)
                        //    {
                        //        failiureCount = failiureCount + 1;
                        //        strDescription = "Update failed for Security Type : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC TYPE"]) + " Error Detail: " + exc.Message;
                        //        LogMessage(sw, service, strDescription, 15, "Security Type");
                        //    }
                        //}

                        //sw.WriteLine("Update Ends for Security Type on: " + DateTime.Now.ToString());

                        //strDescription= "Total Security Type failed to Update: " + failiureCount;
                        //LogMessage(sw,service,strDescription,35,"Security Type");
                        //strDescription="Total Security Type Updated: " + successCount;
                        //LogMessage(sw, service, strDescription, 14, "Security Type");

                        //sw.WriteLine("---------------------------- Security Type Update Ends  -------------------");

                        #endregion
                        //////////////////////////////// Update Ends ///////////////////////////////////
                    }
                    successCount = 0;
                    failiureCount = 0;

                    if (bProceed == true)
                    {
                        //////////////////////////////// Insert Starts ///////////////////////////////////////
                        #region Insert Security Type
                        try
                        {
                            greshamquery = "SP_S_SECURITY_TYPE_INSERT_NEW_GA '" + execType + "'";
                            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                            ds_gresham = new DataSet();
                            dagersham.SelectCommand.CommandTimeout = 600;
                            dagersham.Fill(ds_gresham);
                            totalCount = ds_gresham.Tables[0].Rows.Count;

                            ////////////////////////////////Security Type Excel ///////////////////////////////////////
                            dtExcel1 = ds_gresham.Tables[0];
                            ////////////////////////////////Security Type Excel ///////////////////////////////////////

                            //Console.WriteLine("gresham Dataset Built for Security Type Insert ");
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Security Type Starts -------------------");
                            LG.AddinLogFile(Form1.vLogFile, "Insert Starts for New Security Type on: " + DateTime.Now.ToString());
                        }
                        catch (System.Web.Services.Protocols.SoapException exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Security Type Insert failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                            LogMessage(sw, service, strDescription, 62, "Security Type");
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Security Type Insert failed, please contact administrator. Error Detail:" + exc.Message;
                            LogMessage(sw, service, strDescription, 62, "Security Type");
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        // ssi_securitytype objSecurityType = new ssi_securitytype();
                                        Entity objSecurityType = new Entity("ssi_securitytype");


                                        //objSecurityType.ssi_securitytypeid = new Key();
                                        //objSecurityType.ssi_securitytypeid.Value = Guid.NewGuid();

                                        //  objSecurityType.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC Type"]);
                                        objSecurityType["ssi_name"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC Type"]);

                                        //  objSecurityType.ssi_symbol = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC Type"]);
                                        objSecurityType["ssi_symbol"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC Type"]);

                                        //service.Create(objSecurityType);
                                        Guid newSecurityTypeID = service.Create(objSecurityType);

                                        //Thread.Sleep(sleepTime);
                                        //Console.WriteLine(i.ToString() + " Inserted");
                                        successCount = successCount + 1;
                                        strDescription = "New Security Type inserted: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC Type"]);
                                        LogMessage(sw, service, strDescription, 11, "Security Type");
                                    }
                                    else
                                        break;
                                }
                                catch (System.Web.Services.Protocols.SoapException exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Insert failed for Security Type : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC Type"]) + exc.Detail.InnerText;
                                    LogMessage(sw, service, strDescription, 12, "Security Type");
                                    AddException("Insert failed for Security Type", "SEC Type : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC Type"]), "", "", "", "", "", "", "", exc.Message, dt);

                                }
                                catch (Exception exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Insert failed for Security Type : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC Type"]) + exc.Message;
                                    LogMessage(sw, service, strDescription, 12, "Security Type");
                                    AddException("Insert failed for Security Type", "SEC Type : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC Type"]), "", "", "", "", "", "", "", exc.Message, dt);
                                }
                            }

                        LG.AddinLogFile(Form1.vLogFile, "Insert Ends for Security Type on: " + DateTime.Now.ToString());

                        strDescription = "Total Security Type failed to insert: " + failiureCount;
                        LogMessage(sw, service, strDescription, 34, "Security Type");

                        strDescription = "Total Security Type inserted: " + successCount;
                        LogMessage(sw, service, strDescription, 13, "Security Type");
                        LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Security Type Insert End -------------------");
                        LG.AddinLogFile(Form1.vLogFile, " ");
                        ds_gresham.Dispose();
                        dagersham.Dispose();

                        #endregion
                    }
                    //////////////////////////////// Insert Ends ///////////////////////////////////////
                    #endregion
                }
                successCount = 0;
                failiureCount = 0;

                if (bProceed == true)
                {
                    //label3.Text = "Security Load Starts";
                    #region Security

                    if (bProceed == true)
                    {
                        //////////////////////////////// Insert Starts ///////////////////////////////////////
                        #region Insert Security
                        try
                        {
                            if (cbUsedDate.Checked)
                                greshamquery = "SP_S_SECURITY_INSERT_NEW_GA @LoadTypeTxt='" + execType + "',@AsOfDate='" + txtHide.Text + "'";
                            else
                                greshamquery = "SP_S_SECURITY_INSERT_NEW_GA @LoadTypeTxt='" + execType + "'";

                            cmd.Connection = Gresham_con;
                            cmd.CommandText = greshamquery;
                            cmd.CommandTimeout = 600;
                            dagersham = new SqlDataAdapter(cmd);
                            ds_gresham = new DataSet();
                            dagersham.Fill(ds_gresham);

                            totalCount = ds_gresham.Tables[0].Rows.Count;
                            if (ds_gresham.Tables.Count == 2)
                            {
                                dtExcel = ds_gresham.Tables[1];
                            }

                            //Console.WriteLine("gresham Dataset Built for Security Insert ");
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Security Insert Starts -------------------");
                            LG.AddinLogFile(Form1.vLogFile, "Insert Starts for New Security on: " + DateTime.Now.ToString());
                        }
                        catch (System.Web.Services.Protocols.SoapException exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Security Insert failed, please contact administrator." + "Error Detail: " + exc.Detail.InnerText;
                            LogMessage(sw, service, strDescription, 51, "Security");
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Security Insert failed, please contact administrator." + "Error Detail: " + exc.Message;
                            LogMessage(sw, service, strDescription, 51, "Security");
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        //ssi_security objSecurity = new ssi_security();
                                        Entity objSecurity = new Entity("ssi_security");

                                        //objSecurity.ssi_securityid = new Key();
                                        //objSecurity.ssi_securityid.Value = Guid.NewGuid();

                                        //objSecurity.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC NAME"]);
                                        objSecurity["ssi_name"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC NAME"]);

                                        //    objSecurity.ssi_symbol = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC SYMBOL"]);
                                        objSecurity["ssi_symbol"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC SYMBOL"]);

                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["SECURITYTYPEID"]) != "")
                                        {
                                            //objSecurity.ssi_securitytypeid = new Lookup();
                                            //objSecurity.ssi_securitytypeid.type = EntityName.ssi_securitytype.ToString();
                                            //objSecurity.ssi_securitytypeid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["SECURITYTYPEID"]));

                                            objSecurity["ssi_securitytypeid"] = new EntityReference("ssi_securitytype", new Guid(Convert.ToString(Convert.ToString(ds_gresham.Tables[0].Rows[i]["SECURITYTYPEID"]))));
                                        }

                                        //greshamadvised (sectorflg)  (Commented at 06_07_2018 by abhi)
                                        //if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower() == "true")
                                        //{
                                        //    //objSecurity.ssi_greshamadvised = new CrmBoolean();
                                        //    //objSecurity.ssi_greshamadvised.Value = true; //Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());
                                        //    objSecurity["ssi_greshamadvised"] = true;
                                        //}

                                        //if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower() == "false")
                                        //{
                                        //    //objSecurity.ssi_greshamadvised = new CrmBoolean();
                                        //    //objSecurity.ssi_greshamadvised.Value = false; //Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());
                                        //    objSecurity["ssi_greshamadvised"] = false;

                                        //}


                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_assetclassid"]) != "")    // new added abhi 11/03/2017     change  assetclassid to ssi_assetclassid on 11_27_2017
                                        {
                                            //objSecurity.ssi_assetclassid = new Lookup();
                                            //objSecurity.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                                            //objSecurity.ssi_assetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["assetclassid"]));
                                            objSecurity["ssi_assetclassid"] = new EntityReference("sas_assetclass", new Guid(Convert.ToString(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_assetclassid"]))));

                                        }

                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["asset_code"]) != "")
                                        {

                                            objSecurity["ssi_apxassetclass"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["asset_code"]);  // new added abhi 11/02/2017

                                        }

                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_sectypeassetclass"]) != "")
                                        {

                                            objSecurity["ssi_sectypeassetclass"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_sectypeassetclass"]);  // new added abhi 11/27/2017

                                        }

                                        //  service.Create(objSecurity);
                                        Guid newSecurityID = service.Create(objSecurity);

                                        successCount = successCount + 1;
                                        //Thread.Sleep(sleepTime);
                                        strDescription = "New Security inserted: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC SYMBOL"]);
                                        LogMessage(sw, service, strDescription, 46, "Security");
                                    }
                                    else
                                        break;
                                }
                                catch (System.Web.Services.Protocols.SoapException exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Insert failed for Security : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC SYMBOL"]) + " Error Detail: " + exc.Detail.InnerText;
                                    LogMessage(sw, service, strDescription, 47, "Security");
                                    AddException("Insert failed for Security", "SEC SYMBOL : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC SYMBOL"]), "", "", "", "", "", "", "", exc.Message, dt);

                                }
                                catch (Exception exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Insert failed for Security : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC SYMBOL"]) + " Error Detail: " + exc.Message;
                                    LogMessage(sw, service, strDescription, 47, "Security");
                                    AddException("Insert failed for Security", "SEC SYMBOL : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC SYMBOL"]), "", "", "", "", "", "", "", exc.Message, dt);
                                }
                            }

                        LG.AddinLogFile(Form1.vLogFile, "Insert Ends for Security on: " + DateTime.Now.ToString());

                        strDescription = "Total Security failed to insert: " + failiureCount;
                        LogMessage(sw, service, strDescription, 51, "Security");

                        strDescription = "Total Security inserted: " + successCount;
                        LogMessage(sw, service, strDescription, 48, "Security");
                        LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Security Insert End -------------------");
                        LG.AddinLogFile(Form1.vLogFile, " ");
                        ds_gresham.Dispose();
                        dagersham.Dispose();

                        #endregion
                        //////////////////////////////// Insert Ends /////////////////////////////////////////
                    }
                    successCount = 0;
                    failiureCount = 0;

                    if (bProceed == true)
                    {
                        //////////////////////////////// Update Starts //////////////////////////////////////
                        #region Update Security
                        try
                        {
                            greshamquery = "SP_S_SECURITY_UPDATE_NEW_GA '" + execType + "'";
                            cmd.Connection = Gresham_con;
                            cmd.CommandText = greshamquery;
                            cmd.CommandTimeout = 600;
                            dagersham = new SqlDataAdapter(cmd);
                            ds_gresham = new DataSet();
                            dagersham.Fill(ds_gresham);
                            totalCount = ds_gresham.Tables[0].Rows.Count;
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- Security Update Starts -------------------");
                            LG.AddinLogFile(Form1.vLogFile, "Update Starts for Security on: " + DateTime.Now.ToString());
                        }
                        catch (System.Web.Services.Protocols.SoapException exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Security Update failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                            LogMessage(sw, service, strDescription, 62, "Security");
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Security Update failed, please contact administrator. Error Detail:" + exc.Message;
                            LogMessage(sw, service, strDescription, 62, "Security");
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        Guid SecurityID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]));

                                        //   ssi_security objSecurity = new ssi_security();
                                        Entity objSecurity = new Entity("ssi_security");

                                        //objSecurity.ssi_securityid = new Key();
                                        //objSecurity.ssi_securityid.Value = SecurityID;
                                        objSecurity["ssi_securityid"] = SecurityID;

                                        // objSecurity.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC NAME"]);
                                        objSecurity["ssi_name"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC NAME"]);

                                        //objSecurity.ssi_symbol = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC SYMBOL"]);
                                        objSecurity["ssi_symbol"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC SYMBOL"]);

                                        //objSecurity.ssi_securitytypeid = new Lookup();
                                        //objSecurity.ssi_securitytypeid.type = EntityName.ssi_securitytype.ToString();
                                        //objSecurity.ssi_securitytypeid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SecurityTypeId"]));
                                        objSecurity["ssi_securitytypeid"] = new EntityReference("ssi_securitytype", new Guid(Convert.ToString(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_SecurityTypeId"]))));

                                        //objSecurity.ssi_assetclassid = new Lookup();
                                        //objSecurity.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                                        //objSecurity.ssi_assetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AssetClassId"]));
                                        //  objSecurity["ssi_assetclassid"] = new EntityReference("sas_assetclass", new Guid(Convert.ToString(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_AssetClassId"]))));    // new added abhi 11/03/2017


                                        objSecurity["ssi_apxassetclass"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["asset_code"]); // new added abhi 11/02/2017




                                        //commneted on 29 jan 2014 by dharamendra.
                                        /*
                                        //greshamadvised (sectorflg)
                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower() == "true")
                                        {
                                            objSecurity.ssi_greshamadvised = new CrmBoolean();
                                            objSecurity.ssi_greshamadvised.Value = true; //Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());
                                        }

                                        if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower() == "false")
                                        {
                                            objSecurity.ssi_greshamadvised = new CrmBoolean();
                                            objSecurity.ssi_greshamadvised.Value = false; //Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());
                                        }
                                        */
                                        service.Update(objSecurity);
                                        //Thread.Sleep(sleepTime);
                                        successCount = successCount + 1;
                                        //Console.WriteLine(i.ToString() + " Updated");
                                    }
                                    else
                                        break;
                                }
                                catch (System.Web.Services.Protocols.SoapException exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Update failed for Security : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC NAME"]) + " Error Detail:" + exc.Detail.InnerText;
                                    LogMessage(sw, service, strDescription, 50, "Security");
                                    AddException("Update failed for Security", "SEC NAME : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC NAME"]), "", "", "", "", "", "", "", exc.Message, dt);

                                }
                                catch (Exception exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Update failed for Security : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC NAME"]) + " Error Detail:" + exc.Message;
                                    LogMessage(sw, service, strDescription, 50, "Security");
                                    AddException("Update failed for Security", "SEC NAME : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SEC NAME"]), "", "", "", "", "", "", "", exc.Message, dt);
                                }
                            }

                        LG.AddinLogFile(Form1.vLogFile, "Update Ends for Security on: " + DateTime.Now.ToString());
                        strDescription = "Total Security failed to Update: " + failiureCount;
                        LogMessage(sw, service, strDescription, 52, "Security");

                        strDescription = "Total Security Updated: " + successCount;
                        LogMessage(sw, service, strDescription, 49, "Security");

                        LG.AddinLogFile(Form1.vLogFile, "---------------------------- Security Update Ends  -------------------");

                        #endregion
                        //////////////////////////////// Update Ends ////////////////////////////////////////
                    }
                    #endregion
                }
                successCount = 0;
                failiureCount = 0;

                #region Group Commented
                //////////////////////////////// Insert Starts ///////////////////////////////////////
                #region Insert Group // Commetned
                //try
                //{
                //    greshamquery = "SP_S_GROUP_INSERT '" + execType + "'";
                //    dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                //    ds_gresham = new DataSet();
                //    dagersham.Fill(ds_gresham);
                //    totalCount = ds_gresham.Tables[0].Rows.Count;
                //    //Console.WriteLine("gresham Dataset Built for Group Insert ");
                //    sw.WriteLine("---------------------------- New Group Insert Starts -------------------");
                //    sw.WriteLine("Insert Starts for New Group on: " + DateTime.Now.ToString());
                //}
                //catch (System.Web.Services.Protocols.SoapException exc)
                //{
                //    totalCount = 0;
                //    strDescription = "Group Insert failed, please contact administrator." + "Error Detail:" + exc.Detail.InnerText;
                //    LogMessage(sw, service, strDescription, 62, "Group");
                //}
                //catch (Exception exc)
                //{
                //    totalCount = 0;
                //    strDescription = "Group Insert failed, please contact administrator." + "Error Detail:" + exc.Message;
                //    LogMessage(sw, service, strDescription, 62, "Group");
                //}

                //for (int i = 0; i < totalCount; i++)
                //{
                //    try
                //    {
                //        sas_reportrollupgroup objGroup = new sas_reportrollupgroup();

                //        objGroup.sas_reportrollupgroupid = new Key();
                //        objGroup.sas_reportrollupgroupid.Value = Guid.NewGuid();

                //        objGroup.sas_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME1"]);
                //        objGroup.ssi_adventgroup = new CrmBoolean();
                //        objGroup.ssi_adventgroup.Value = true;

                //        service.Create(objGroup);
                //        //Console.WriteLine(i.ToString() + " Inserted");
                //        successCount = successCount + 1;
                //        //Thread.Sleep(sleepTime);
                //        strDescription = "New Group inserted: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME1"]);
                //        LogMessage(sw, service, strDescription, 53, "Group");
                //    }
                //    catch (System.Web.Services.Protocols.SoapException exc)
                //    {
                //        failiureCount = failiureCount + 1;
                //        strDescription = "Insert failed for Group : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME1"]) + " Error Detail:" + exc.Detail.InnerText;
                //        LogMessage(sw, service, strDescription, 54, "Group");
                //    }
                //    catch (Exception exc)
                //    {
                //        failiureCount = failiureCount + 1;
                //        strDescription = "Insert failed for Group : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["NAME1"]) + " Error Detail:" + exc.Message;
                //        LogMessage(sw, service, strDescription, 54, "Group");
                //    }
                //}

                //sw.WriteLine("Insert Ends for Group on: " + DateTime.Now.ToString());

                //strDescription = "Total Group failed to insert: " + failiureCount;
                //LogMessage(sw, service, strDescription, 56, "Group");
                //strDescription = "Total Group inserted: " + successCount;
                //LogMessage(sw, service, strDescription, 55, "Group");
                //sw.WriteLine("---------------------------- New Group Insert End -------------------");
                //sw.WriteLine();
                //ds_gresham.Dispose();
                //dagersham.Dispose();

                #endregion
                //////////////////////////////// Insert Ends /////////////////////////////////////////
                #endregion

                successCount = 0;
                failiureCount = 0;
                //label3.Text = "Group Account M-M Load Starts";
                if (bProceed == true)
                {
                    //////////////////////////////// Insert Starts ///////////////////////////////////////
                    #region Insert Look Through Account Group  // Commented
                    //try
                    //{
                    //    greshamquery = "SP_S_GROUP_RELSHIP_INSERT '" + execType + "'";
                    //    dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                    //    ds_gresham = new DataSet();
                    //    dagersham.Fill(ds_gresham);
                    //    totalCount = ds_gresham.Tables[0].Rows.Count;
                    //    sw.WriteLine("---------------------------- New Look Through Account Group Starts -------------------");
                    //    sw.WriteLine("Insert Starts for Look Through Account Group on: " + DateTime.Now.ToString());
                    //}
                    //catch (System.Web.Services.Protocols.SoapException exc)
                    //{
                    //    totalCount = 0;
                    //    strDescription = "Look Through Account Group, please contact administrator. Error Detail: " + exc.Detail.InnerText;
                    //    LogMessage(sw, service, strDescription, 62, "LookThroughAccountGroup");
                    //}
                    //catch (Exception exc)
                    //{
                    //    totalCount = 0;
                    //    strDescription = "Look Through Account Group, please contact administrator. Error Detail: " + exc.Message;
                    //    LogMessage(sw, service, strDescription, 62, "LookThroughAccountGroup");
                    //}

                    //for (int i = 0; i < totalCount; i++)
                    //{
                    //    try
                    //    {
                    //        ssi_lookthroughaccountgroup objAccountGroup = new ssi_lookthroughaccountgroup();

                    //        objAccountGroup.ssi_lookthroughaccountgroupid = new Key();
                    //        objAccountGroup.ssi_lookthroughaccountgroupid.Value = Guid.NewGuid();

                    //        objAccountGroup.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["LookTroughName"]);

                    //        // objAccountGroup.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["]);
                    //        objAccountGroup.ssi_reportgroupid = new Lookup();
                    //        objAccountGroup.ssi_reportgroupid.type = EntityName.sas_reportrollupgroup.ToString();
                    //        objAccountGroup.ssi_reportgroupid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_reportrollupgroupid"]));

                    //        objAccountGroup.ssi_accountid = new Lookup();
                    //        objAccountGroup.ssi_accountid.type = EntityName.ssi_account.ToString();
                    //        objAccountGroup.ssi_accountid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]));

                    //        //ownership
                    //        //if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_ownership"]) != "")
                    //        //{
                    //        objAccountGroup.ssi_ownership = new CrmDecimal();
                    //        objAccountGroup.ssi_ownership.Value = 100;
                    //        //}
                    //        //else
                    //        //{
                    //        //    objAccountGroup.ssi_ownership = new CrmDecimal();
                    //        //    objAccountGroup.ssi_ownership.IsNull = true;
                    //        //    objAccountGroup.ssi_ownership.IsNullSpecified = true;
                    //        //}

                    //        service.Create(objAccountGroup);
                    //        successCount = successCount + 1;
                    //        strDescription = "New Look Through Account Group inserted: Reportrollupgroupid:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_reportrollupgroupid"]) + " Accountid:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]);
                    //        //sw.WriteLine(strDescription);
                    //        LogMessage(sw, service, strDescription, 58, "LookThroughAccountGroup");
                    //    }
                    //    catch (System.Web.Services.Protocols.SoapException exc)
                    //    {
                    //        failiureCount = failiureCount + 1;
                    //        strDescription = "Insert failed for Look Through Account Group : Reportrollupgroupid:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_reportrollupgroupid"]) + " Accountid:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + " Error Detail:" + exc.Detail.InnerText;
                    //        LogMessage(sw, service, strDescription, 57, "LookThroughAccountGroup");

                    //    }
                    //    catch (Exception exc)
                    //    {
                    //        failiureCount = failiureCount + 1;
                    //        strDescription = "Insert failed for Look Through Account Group : Reportrollupgroupid:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_reportrollupgroupid"]) + " Accountid:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + " Error Detail:" + exc.Message;
                    //        LogMessage(sw, service, strDescription, 57, "LookThroughAccountGroup");
                    //    }
                    //}

                    //sw.WriteLine("Insert Ends for Look Through Account Group on: " + DateTime.Now.ToString());

                    //strDescription = "Total Look Through Account Group failed to insert: " + failiureCount;
                    //LogMessage(sw, service, strDescription, 60, "LookThroughAccountGroup");
                    //strDescription = "Total Look Through Account Group inserted: " + successCount;
                    //LogMessage(sw, service, strDescription, 59, "LookThroughAccountGroup");
                    //sw.WriteLine("---------------------------- New Look Through Account Group Insert End -------------------");
                    //sw.WriteLine();
                    //ds_gresham.Dispose();
                    //dagersham.Dispose();

                    #endregion
                }

                //////////////////////////////// Insert Ends /////////////////////////////////////////

                successCount = 0;
                failiureCount = 0;
                //label3.Text = "Transaction Code Load Starts";
                if (bProceed == true)
                {
                    #region Transaction Code
                    if (bProceed == true)
                    {
                        //////////////////////////////// Update Starts ///////////////////////////////////
                        #region Update Transaction Code //Commented
                        //try
                        //{
                        //    greshamquery = "SP_S_TRXN_CODE_UPDATE '" + execType + "'";
                        //    dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                        //    ds_gresham = new DataSet();
                        //    dagersham.Fill(ds_gresham);
                        //    totalCount = ds_gresham.Tables[0].Rows.Count;
                        //    sw.WriteLine("---------------------------- Transaction Code Update Starts -------------------");
                        //    sw.WriteLine("Update Starts for Transaction Code on: " + DateTime.Now.ToString());
                        //}
                        //catch (System.Web.Services.Protocols.SoapException exc)
                        //{
                        //    totalCount = 0;
                        //    strDescription = "Transaction Code Update failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                        //    LogMessage(sw, service, strDescription, 62, "Transaction Code");
                        //}
                        //catch (Exception exc)
                        //{
                        //    totalCount = 0;
                        //    strDescription = "Transaction Code Update failed, please contact administrator. Error Detail:" + exc.Message;
                        //    LogMessage(sw, service, strDescription, 62, "Transaction Code");
                        //}

                        //for (int i = 0; i < totalCount; i++)
                        //{
                        //    try
                        //    {
                        //        ssi_transactioncode objTransactionCode = new ssi_transactioncode();
                        //        Guid TransactionCodeID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_transactioncodeId"]));

                        //        objTransactionCode.ssi_transactioncodeid = new Key();
                        //        objTransactionCode.ssi_transactioncodeid.Value = TransactionCodeID;

                        //        objTransactionCode.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]);
                        //        //objTransactionCode.ssi_code = Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]);

                        //        service.Update(objTransactionCode);
                        //        //Thread.Sleep(sleepTime);
                        //        successCount = successCount + 1;
                        //    }
                        //    catch (System.Web.Services.Protocols.SoapException exc)
                        //    {
                        //        failiureCount = failiureCount + 1;
                        //        strDescription = "Update failed for Transaction Code : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]) + " Error Detail:" + exc.Detail.InnerText;
                        //        LogMessage(sw, service, strDescription, 10, "Transaction Code");
                        //    }
                        //    catch (Exception exc)
                        //    {
                        //        failiureCount = failiureCount + 1;
                        //        strDescription = "Update failed for Transaction Code : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]) + " Error Detail:" + exc.Message;
                        //        LogMessage(sw, service, strDescription, 10, "Transaction Code");
                        //    }
                        //}

                        //sw.WriteLine("Update Ends for Transaction Code on: " + DateTime.Now.ToString());

                        //strDescription= "Total Transaction Code Update Failed: " + failiureCount;
                        //LogMessage(sw,service,strDescription,33,"Transaction Code");

                        //strDescription = "Total Transaction Code Updated: " + successCount;
                        //LogMessage(sw,service,strDescription,9,"Transaction Code");
                        //sw.WriteLine("---------------------------- Transaction Code Update Ends  -------------------");

                        #endregion
                        //////////////////////////////// Update Ends ////////////////////////////////////
                    }
                    successCount = 0;
                    failiureCount = 0;
                    totalCount = 0;

                    if (bProceed == true)
                    {
                        //////////////////////////////// Insert Starts ///////////////////////////////////////
                        #region Insert Transaction Code
                        try
                        {
                            greshamquery = "SP_S_TRXN_CODE_INSERT_NEW_GA '" + execType + "'";
                            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                            ds_gresham = new DataSet();
                            dagersham.SelectCommand.CommandTimeout = 600;
                            dagersham.Fill(ds_gresham);
                            totalCount = ds_gresham.Tables[0].Rows.Count;
                            //Console.WriteLine("gresham Dataset Built for Transaction Code Insert ");
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Transaction Code Insert Starts -------------------");
                            LG.AddinLogFile(Form1.vLogFile, "Insert Starts for New Transaction Code on: " + DateTime.Now.ToString());
                        }
                        catch (System.Web.Services.Protocols.SoapException exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Transaction Code Insert failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                            LogMessage(sw, service, strDescription, 62, "Transaction Code");
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            totalCount = 0;
                            strDescription = "Transaction Code Insert failed, please contact administrator. Error Detail:" + exc.Message;
                            LogMessage(sw, service, strDescription, 62, "Transaction Code");
                        }

                        if (bProceed == true)
                            for (int i = 0; i < totalCount; i++)
                            {
                                try
                                {
                                    if (bProceed == true)
                                    {
                                        //ssi_transactioncode objTransactionCode = new ssi_transactioncode();
                                        Entity objTransactionCode = new Entity("ssi_transactioncode");

                                        //objTransactionCode.ssi_transactioncodeid = new Key();
                                        //objTransactionCode.ssi_transactioncodeid.Value = Guid.NewGuid();

                                        //  objTransactionCode.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]);
                                        objTransactionCode["ssi_name"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]);
                                        //objTransactionCode.ssi_code = Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]);

                                        //     service.Create(objTransactionCode);
                                        Guid newTransactionCodeId = service.Create(objTransactionCode);

                                        //Thread.Sleep(sleepTime);
                                        //Console.WriteLine(i.ToString() + " Inserted");
                                        successCount = successCount + 1;
                                        strDescription = "New Transaction Code inserted: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]);
                                        LogMessage(sw, service, strDescription, 6, "Transaction Code");
                                    }
                                    else
                                        break;
                                }
                                catch (System.Web.Services.Protocols.SoapException exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Insert failed for Transaction Code : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]) + " Error Detail: " + exc.Detail.InnerText;
                                    LogMessage(sw, service, strDescription, 26, "Transaction Code");
                                    AddException("Insert failed for Transaction Code", "TRAN CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]), "", "", "", "", "", "", "", exc.Message, dt);

                                }
                                catch (Exception exc)
                                {
                                    bProceed = true;
                                    failiureCount = failiureCount + 1;
                                    strDescription = "Insert failed for Transaction Code : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]) + " Error Detail: " + exc.Message;
                                    LogMessage(sw, service, strDescription, 26, "Transaction Code");
                                    AddException("Insert failed for Transaction Code", "TRAN CODE : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRAN CODE"]), "", "", "", "", "", "", "", exc.Message, dt);
                                }
                            }

                        LG.AddinLogFile(Form1.vLogFile, "Insert Ends for Transaction Code on: " + DateTime.Now.ToString());

                        strDescription = "Total Transaction Code failed to insert: " + failiureCount;
                        LogMessage(sw, service, strDescription, 32, "Transaction Code");

                        strDescription = "Total Transaction Code inserted: " + successCount;
                        LogMessage(sw, service, strDescription, 8, "Transaction Code");
                        LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Transaction Code Insert End -------------------");
                        LG.AddinLogFile(Form1.vLogFile, " ");
                        ds_gresham.Dispose();
                        dagersham.Dispose();

                        #endregion
                        //////////////////////////////// Insert Ends /////////////////////////////////////////
                    }
                    #endregion
                }

                if (bProceed == true)
                {
                    if (execType == "T" || execType == "B")
                    {
                        successCount = 0;
                        failiureCount = 0;
                        totalCount = 0;
                        //label3.Text = "Transaction Load Starts";
                        bool transactionSuccess = true;
                        #region Transaction
                        if (bProceed == true)
                        {
                            ///////////////////////////////// Delete Starts /////////////////////////////////////
                            #region Transaction Delete
                            //Guid TransactionID = new Guid();
                            try
                            {
                                LG.AddinLogFile(Form1.vLogFile, "---------------------------- Transaction Delete Starts -------------------");
                                LG.AddinLogFile(Form1.vLogFile, "Delete Starts for Transaction on: " + DateTime.Now.ToString());
                                greshamquery = "SP_S_TRXN_DELETE_NEW_GA";

                                //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                cmd.Connection = Gresham_con;
                                cmd.CommandText = greshamquery;
                                cmd.CommandTimeout = 4500;    //900 to 4500 by abhi
                                dagersham = new SqlDataAdapter(cmd);
                                ds_gresham = new DataSet();
                                dagersham.Fill(ds_gresham);

                                //dagersham.Fill(ds_gresham);
                                //successCount = Convert.ToInt32(ds_gresham.Tables[0].Rows[0]["DeleteCount"]);

                                for (int j = 0; j < ds_gresham.Tables[0].Rows.Count; j++)
                                {
                                    Guid UUID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[j]["ssi_transactionlogId"]));
                                    service.Delete("ssi_transactionlog", UUID);
                                    successCount = successCount + 1;
                                }

                            }
                            catch (System.Web.Services.Protocols.SoapException exc)
                            {
                                bProceed = false;
                                transactionSuccess = false;
                                strDescription = "Transaction Delete failed, please contact administrator." + " Error Detail: " + exc.Detail.InnerText + exc.StackTrace.ToString();
                                LogMessage(sw, service, strDescription, 62, "Transaction");
                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                transactionSuccess = false;
                                strDescription = "Transaction Delte failed, please contact administrator." + " Error Detail: " + exc.Message + exc.StackTrace.ToString();
                                LogMessage(sw, service, strDescription, 62, "Transaction");
                            }

                            LG.AddinLogFile(Form1.vLogFile, "Delete Ends for Transaction on: " + DateTime.Now.ToString());

                            strDescription = "Total Transaction Deleted: " + successCount;
                            LogMessage(sw, service, strDescription, 41, "Transaction");
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- Transaction Delete Ends  -------------------");

                            #endregion
                            /////////////////////////////// Delete Ends //////////////////////////////////////
                        }
                        successCount = 0;
                        failiureCount = 0;

                        if (bProceed == true)
                        {
                            ////////////////////////////// Insert Starts ///////////////////////////////////////
                            if (transactionSuccess)
                            {
                                #region Insert Transaction
                                try
                                {
                                    LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Transaction Insert Starts -------------------");
                                    greshamquery = "SP_S_TRXN_INSERT_NEW_GA";
                                    dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                    ds_gresham = new DataSet();
                                    dagersham.SelectCommand.CommandTimeout = 4500;   //600 to 4500 by abhi
                                    dagersham.Fill(ds_gresham);
                                    totalCount = ds_gresham.Tables[0].Rows.Count;
                                    //Console.WriteLine("gresham Dataset Built for Transaction Insert ");                
                                    LG.AddinLogFile(Form1.vLogFile, "Insert Starts for New Transaction on: " + DateTime.Now.ToString());
                                }
                                catch (System.Web.Services.Protocols.SoapException exc)
                                {
                                    bProceed = true;
                                    totalCount = 0;
                                    transactionSuccess = false;
                                    strDescription = "Transaction Insert failed, please contact administrator. Error Detail: " + exc.Detail.InnerText;
                                    LogMessage(sw, service, strDescription, 62, "Transaction");
                                }
                                catch (Exception exc)
                                {
                                    bProceed = true;
                                    totalCount = 0;
                                    transactionSuccess = false;
                                    strDescription = "Transaction Insert failed, please contact administrator. Error Detail: " + exc.Message;
                                    LogMessage(sw, service, strDescription, 62, "Transaction");
                                }

                                if (bProceed == true)
                                    for (int i = 0; i < totalCount; i++)
                                    {
                                        try
                                        {
                                            if (bProceed == true)
                                            {
                                                //ssi_transactionlog objTransaction = new ssi_transactionlog();
                                                Entity objTransaction = new Entity("ssi_transactionlog");


                                                //objTransaction.ssi_transactionlogid = new Key();
                                                //objTransaction.ssi_transactionlogid.Value = Guid.NewGuid();

                                                //account
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) != "")
                                                {
                                                    //objTransaction.ssi_accountid = new Lookup();
                                                    //objTransaction.ssi_accountid.type = EntityName.ssi_account.ToString();
                                                    //objTransaction.ssi_accountid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]));
                                                    objTransaction["ssi_accountid"] = new EntityReference("ssi_account", new Guid(Convert.ToString(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]))));
                                                }

                                                //quantity
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["QUANTITY"]) != "")
                                                {
                                                    //objTransaction.ssi_quantity = new CrmDecimal();
                                                    //objTransaction.ssi_quantity.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["QUANTITY"]);
                                                    objTransaction["ssi_quantity"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["QUANTITY"]);
                                                }
                                                else
                                                {
                                                    //objTransaction.ssi_quantity = new CrmDecimal();
                                                    //objTransaction.ssi_quantity.IsNull = true;
                                                    //objTransaction.ssi_quantity.IsNullSpecified = true;
                                                    objTransaction["ssi_quantity"] = null;
                                                }

                                                //greshamadvised (sectorflg)
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]) != "")
                                                {
                                                    //objTransaction.ssi_grehamadvised = new CrmBoolean();
                                                    //objTransaction.ssi_grehamadvised.Value = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());
                                                    objTransaction["ssi_grehamadvised"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());
                                                }

                                                //Trade date
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE DATE"]) != "")
                                                {
                                                    //objTransaction.ssi_tradedate = new CrmDateTime();
                                                    //objTransaction.ssi_tradedate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE DATE"]);
                                                    objTransaction["ssi_tradedate"] = Convert.ToDateTime(Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE DATE"]));
                                                }

                                                //Settle date
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["SETTLE DATE"]) != "")
                                                {
                                                    //objTransaction.ssi_settledate = new CrmDateTime();
                                                    //objTransaction.ssi_settledate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["SETTLE DATE"]);
                                                    objTransaction["ssi_settledate"] = Convert.ToDateTime(Convert.ToString(ds_gresham.Tables[0].Rows[i]["SETTLE DATE"]));
                                                }

                                                //Trade Amt
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE AMT"]) != "")
                                                {
                                                    //objTransaction.ssi_value = new CrmMoney();
                                                    //objTransaction.ssi_value.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["TRADE AMT"]);
                                                    objTransaction["ssi_value"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["TRADE AMT"]));
                                                }
                                                else
                                                {
                                                    //objTransaction.ssi_value = new CrmMoney();
                                                    //objTransaction.ssi_value.IsNull = true;
                                                    //objTransaction.ssi_value.IsNullSpecified = true;
                                                    objTransaction["ssi_value"] = null;
                                                }

                                                //Security
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]) != "")
                                                {
                                                    //objTransaction.ssi_securityid = new Lookup();
                                                    //objTransaction.ssi_securityid.type = EntityName.ssi_security.ToString();
                                                    //objTransaction.ssi_securityid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]));

                                                    objTransaction["ssi_securityid"] = new EntityReference("ssi_security", new Guid(Convert.ToString(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]))));
                                                }

                                                //Assetclassid 
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]) != "")
                                                {
                                                    //objTransaction.ssi_assetclassid = new Lookup();
                                                    //objTransaction.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                                                    //objTransaction.ssi_assetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]));
                                                    objTransaction["ssi_assetclassid"] = new EntityReference("sas_assetclass", new Guid(Convert.ToString(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]))));
                                                }

                                                //transactioncodeid
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_transactioncodeid"]) != "")
                                                {
                                                    //objTransaction.ssi_transactioncodeid = new Lookup();
                                                    //objTransaction.ssi_transactioncodeid.type = EntityName.ssi_transactionlog.ToString();
                                                    //objTransaction.ssi_transactioncodeid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_transactioncodeid"]));
                                                    objTransaction["ssi_transactioncodeid"] = new EntityReference("ssi_transactionlog", new Guid(Convert.ToString(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_transactioncodeid"]))));
                                                }

                                                //Currency (default to USD)
                                                //objTransaction.transactioncurrencyid = new Lookup();
                                                //objTransaction.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                                                //objTransaction.transactioncurrencyid.Value = new Guid("215A7268-A2E1-DD11-A826-001D09665E8F");
                                                objTransaction["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("215A7268-A2E1-DD11-A826-001D09665E8F"));

                                                //below 2 cols added on 19 dec 2013
                                                //Assetclassid 
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]) != "")
                                                {
                                                    //objTransaction.ssi_subassetclassid = new Lookup();
                                                    //objTransaction.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                                                    //objTransaction.ssi_subassetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]));
                                                    objTransaction["ssi_subassetclassid"] = new EntityReference("ssi_subassetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"])));
                                                }

                                                //greshamadvised (sectorflg)
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"]) != "")
                                                {
                                                    //objTransaction.ssi_benchmarkid = new Lookup();
                                                    //objTransaction.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                                                    //objTransaction.ssi_benchmarkid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"])); ;
                                                    objTransaction["ssi_benchmarkid"] = new EntityReference("sas_benchmark", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"])));
                                                }

                                                //below 3 cols added on 11th Feb 2016

                                                //Ssi_TransactionCodeIdName
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TransactionCodeIdName"]) != "")
                                                {
                                                    //objTransaction.ssi_name = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TransactionCodeIdName"]);
                                                    objTransaction["ssi_name"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_TransactionCodeIdName"]);
                                                }
                                                //Comments
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Comments"]) != "")
                                                {
                                                    //objTransaction.ssi_comment = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Comments"]);
                                                    objTransaction["ssi_comment"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["Comments"]);
                                                }
                                                //Transaction_Type
                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Transaction_Type"]) != "")
                                                {
                                                    int TType = Convert.ToInt32(ds_gresham.Tables[0].Rows[i]["Transaction_Type"]);
                                                    //objTransaction.ssi_transactiontype = new Picklist();
                                                    //objTransaction.ssi_transactiontype.Value = TType;
                                                    objTransaction["ssi_transactiontype"] = new Microsoft.Xrm.Sdk.OptionSetValue(TType);
                                                }

                                                if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["OriginalSecurityId"]) != "")//added 12/14/2017-sasmit
                                                {

                                                    objTransaction["ssi_originalsecurityid"] = new EntityReference("ssi_transactionlog", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["OriginalSecurityId"])));
                                                }
                                                objTransaction["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(100000001);   // change by AbhiS 12/12/2017

                                                Guid NewTransactionId = service.Create(objTransaction);
                                                //Thread.Sleep(sleepTime);

                                                successCount = successCount + 1;
                                            }
                                            else
                                                break;
                                        }
                                        catch (System.Web.Services.Protocols.SoapException exc)
                                        {
                                            bProceed = true;
                                            failiureCount = failiureCount + 1;
                                            transactionSuccess = false;
                                            string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", Cusip: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Cusip"]) + ", Trade Date: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE DATE"]) + ", Trade Amt:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE AMT"]);
                                            strDescription = failiureText + "    IdNmb:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IdNmb"]) + " Error Detail:" + exc.Detail.InnerText;
                                            LogMessage(sw, service, strDescription, 26, "Transaction");
                                            AddException("Insert failed for Transaction", "Account : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AccountName"]), "Security Name : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecurityName"]), "SecurityType :" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecurityType"]), "Security Symbol: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecuritySymbol"]), "Trade Date: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE DATE"]), "Trade Amt:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE AMT"]), "AssetClassName : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AssetClassName"]), "SubAssetClassName : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SubAssetClassName"]), exc.Message, dt);

                                        }
                                        catch (Exception exc)
                                        {
                                            bProceed = true;
                                            transactionSuccess = false;
                                            failiureCount = failiureCount + 1;
                                            string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", Cusip: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["Cusip"]) + ", Trade Date: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE DATE"]) + ", Trade Amt:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE AMT"]);
                                            strDescription = failiureText + "    IdNmb:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IdNmb"]) + " Error Detail:" + exc.Message;
                                            LogMessage(sw, service, strDescription, 26, "Transaction");
                                            AddException("Insert failed for Transaction", "Account : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AccountName"]), "Security Name : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecurityName"]), "SecurityType :" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecurityType"]), "Security Symbol " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecuritySymbol"]), "Trade Date: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE DATE"]), "Trade Amt:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["TRADE AMT"]), "AssetClassName : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AssetClassName"]), "SubAssetClassName : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SubAssetClassName"]), exc.Message, dt);
                                        }
                                    }

                                LG.AddinLogFile(Form1.vLogFile, "Insert Ends for Transaction on: " + DateTime.Now.ToString());

                                strDescription = "Total Transaction failed to insert: " + failiureCount;
                                LogMessage(sw, service, strDescription, 40, "Transaction");

                                strDescription = "Total Transaction inserted: " + successCount;
                                LogMessage(sw, service, strDescription, 29, "Transaction");
                                LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Transaction Insert Ends -------------------");
                                LG.AddinLogFile(Form1.vLogFile, " ");

                                ds_gresham.Dispose();
                                dagersham.Dispose();

                                #endregion
                            }
                        }
                        ////////////////////////////////// Insert Ends ////////////////////////////////////////

                        #endregion
                    }
                }

                if (bProceed == true)
                {
                    if (execType == "P" || execType == "B")
                    {
                        successCount = 0;
                        failiureCount = 0;
                        totalCount = 0;
                        bool positionSuccess = true;

                        #region Position
                        if (bProceed == true)
                        {
                            ///////////////////////////////// Delete Starts /////////////////////////////////////
                            #region Delete Position
                            try
                            {
                                LG.AddinLogFile(Form1.vLogFile, "---------------------------- Position Delete Starts -------------------");
                                LG.AddinLogFile(Form1.vLogFile, "Delete Starts for Position on: " + DateTime.Now.ToString());

                                //greshamquery = "SP_S_POSITIONS_DELETE ";
                                //dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                //ds_gresham = new DataSet();
                                //dagersham.Fill(ds_gresham);

                                greshamquery = "SP_S_POSITIONS_DELETE_NEW_GA";
                                cmd.Connection = Gresham_con;
                                cmd.CommandText = greshamquery;
                                cmd.CommandTimeout = 4500; //300 to 4500 by abhi
                                dagersham = new SqlDataAdapter(cmd);
                                ds_gresham = new DataSet();
                                dagersham.Fill(ds_gresham);

                                //successCount = Convert.ToInt32(ds_gresham.Tables[0].Rows[0]["DeleteCount"]);

                                for (int j = 0; j < ds_gresham.Tables[0].Rows.Count; j++)
                                {
                                    Guid UUID = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[j]["ssi_PositionId"]));
                                    //service.Delete(EntityName.ssi_position.ToString(), UUID);
                                    service.Delete("ssi_position", UUID);
                                    successCount = successCount + 1;
                                }
                            }
                            catch (System.Web.Services.Protocols.SoapException exc)
                            {
                                bProceed = false;
                                positionSuccess = false;
                                strDescription = "Position Delete failed, please contact administrator. Error Detail" + exc.Detail.InnerText;
                                LogMessage(sw, service, strDescription, 62, "Position");
                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                positionSuccess = false;
                                strDescription = "Position Delete failed, please contact administrator. Error Detail" + exc.Message;
                                LogMessage(sw, service, strDescription, 62, "Position");
                            }

                            strDescription = "Total Position Deleted: " + successCount;
                            LogMessage(sw, service, strDescription, 44, "Position");
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- Position Delete Ends  -------------------");

                            #endregion
                            /////////////////////////////// Delete Ends ///////////////////////////////////////
                        }
                        successCount = 0;
                        failiureCount = 0;
                        totalCount = 0;

                        if (bProceed == true)
                        {
                            #region Position Join Load
                            try
                            {
                                cmd.CommandText = "SP_S_POSITION_JOIN_LOAD_NEW_GA";
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Connection = Gresham_con;
                                cmd.CommandTimeout = 600;    //600 to 4500 by abhi
                                Gresham_con.Open();
                                cmd.ExecuteNonQuery();
                                LG.AddinLogFile(Form1.vLogFile, "SP_S_POSITION_JOIN_LOAD_NEW_GA Excecute");
                            }
                            catch (System.Web.Services.Protocols.SoapException exc)
                            {
                                bProceed = false;
                                positionSuccess = false;
                                strDescription = "Position Join Load failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                                LogMessage(sw, service, strDescription, 62, "Position");

                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                positionSuccess = false;
                                strDescription = "Position Join Load failed, please contact administrator. Error Detail:" + exc.Message;
                                LogMessage(sw, service, strDescription, 62, "Position");
                            }

                            #endregion
                        }

                        successCount = 0;
                        failiureCount = 0;
                        totalCount = 0;

                        if (positionSuccess && bProceed == true)
                        {
                            ////////////////////////////// Insert Starts ////////////////////////////////////////////
                            #region Insert Position
                            try
                            {
                                greshamquery = "SP_S_POSITIONS_INSERT_NEW_GA";
                                dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
                                ds_gresham = new DataSet();
                                dagersham.SelectCommand.CommandTimeout = 4500; //600 to 4500 by abhi
                                dagersham.Fill(ds_gresham);

                                totalCount = ds_gresham.Tables[0].Rows.Count;
                                LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Position Insert Starts -------------------");
                                LG.AddinLogFile(Form1.vLogFile, "Insert Starts for New Position on: " + DateTime.Now.ToString());
                            }
                            catch (System.Web.Services.Protocols.SoapException exc)
                            {
                                bProceed = false;
                                totalCount = 0;
                                positionSuccess = false;
                                strDescription = "Position Insert failed, please contact administrator. Error Detail:" + exc.Detail.InnerText;
                                LogMessage(sw, service, strDescription, 62, "Position");
                            }
                            catch (Exception exc)
                            {
                                bProceed = false;
                                totalCount = 0;
                                positionSuccess = false;
                                strDescription = "Position Insert failed, please contact administrator. Error Detail:" + exc.Message;
                                LogMessage(sw, service, strDescription, 62, "Position");
                            }

                            if (bProceed == true)
                                for (int i = 0; i < totalCount; i++)
                                {
                                    try
                                    {
                                        if (bProceed == true)
                                        {
                                            //ssi_position objPosition = new ssi_position();
                                            Entity objPosition = new Entity("ssi_position");

                                            //primary key ssi_positionid
                                            //objPosition.ssi_positionid = new Key();
                                            //objPosition.ssi_positionid.Value = Guid.NewGuid();

                                            //accountid
                                            //if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) != "")
                                            //{
                                            //objPosition.ssi_accountid = new Lookup();
                                            //objPosition.ssi_accountid.type = EntityName.ssi_account.ToString();
                                            //objPosition.ssi_accountid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]));
                                            objPosition["ssi_accountid"] = new EntityReference("ssi_account", new Guid(Convert.ToString(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]))));
                                            //}

                                            //Quantity
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["QUANTITY"]) != "")
                                            {
                                                //objPosition.ssi_quantity = new CrmDecimal();
                                                //objPosition.ssi_quantity.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["QUANTITY"]);
                                                objPosition["ssi_quantity"] = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["QUANTITY"]);
                                            }
                                            else
                                            {
                                                //objPosition.ssi_quantity = new CrmDecimal();
                                                //objPosition.ssi_quantity.IsNull = true;
                                                //objPosition.ssi_quantity.IsNullSpecified = true;
                                                objPosition["ssi_quantity"] = null;
                                            }

                                            //GreshamAdvised flag
                                            //objPosition.ssi_greshamadvised = new CrmBoolean();
                                            //objPosition.ssi_greshamadvised.Value = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());


                                            //Unit Adj Cost
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["UNIT ADJ COST"]) != "")
                                            {
                                                //objPosition.ssi_adjustedunitcost = new CrmMoney();
                                                //objPosition.ssi_adjustedunitcost.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["UNIT ADJ COST"]);
                                                objPosition["ssi_adjustedunitcost"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["UNIT ADJ COST"]));
                                            }
                                            else
                                            {
                                                //objPosition.ssi_adjustedunitcost = new CrmMoney();
                                                //objPosition.ssi_adjustedunitcost.IsNull = true;
                                                //objPosition.ssi_adjustedunitcost.IsNullSpecified = true;
                                                objPosition["ssi_adjustedunitcost"] = null;
                                            }

                                            //Adjusted Cost
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ADJUSTED COST"]) != "")
                                            {
                                                //objPosition.ssi_totaladjustedcost = new CrmMoney();
                                                //objPosition.ssi_totaladjustedcost.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["ADJUSTED COST"]);
                                                objPosition["ssi_totaladjustedcost"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["ADJUSTED COST"]));
                                            }
                                            else
                                            {
                                                //objPosition.ssi_totaladjustedcost = new CrmMoney();
                                                //objPosition.ssi_totaladjustedcost.IsNull = true;
                                                //objPosition.ssi_totaladjustedcost.IsNullSpecified = true;
                                                objPosition["ssi_totaladjustedcost"] = null;
                                            }

                                            //price
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["PRICE"]) != "")
                                            {
                                                //objPosition.ssi_price = new CrmMoney();
                                                //objPosition.ssi_price.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["PRICE"]);
                                                objPosition["ssi_price"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["PRICE"]));
                                            }
                                            else
                                            {
                                                //objPosition.ssi_price = new CrmMoney();
                                                //objPosition.ssi_price.IsNull = true;
                                                //objPosition.ssi_price.IsNullSpecified = true;
                                                objPosition["ssi_price"] = null;
                                            }

                                            //Market Value
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["MKT VALUE"]) != "")
                                            {
                                                //objPosition.ssi_marketvalue = new CrmMoney();
                                                //objPosition.ssi_marketvalue.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["MKT VALUE"]);
                                                objPosition["ssi_marketvalue"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["MKT VALUE"]));
                                            }
                                            else
                                            {
                                                //objPosition.ssi_marketvalue = new CrmMoney();
                                                //objPosition.ssi_marketvalue.IsNull = true;
                                                //objPosition.ssi_marketvalue.IsNullSpecified = true;
                                                objPosition["ssi_marketvalue"] = null;
                                            }

                                            //Unreal G L
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["UNREAL G L"]) != "")
                                            {
                                                //objPosition.ssi_unrealizedgl = new CrmMoney();
                                                //objPosition.ssi_unrealizedgl.Value = Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["UNREAL G L"]);
                                                objPosition["ssi_unrealizedgl"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(ds_gresham.Tables[0].Rows[i]["UNREAL G L"]));
                                            }
                                            else
                                            {
                                                //objPosition.ssi_unrealizedgl = new CrmMoney();
                                                //objPosition.ssi_unrealizedgl.IsNull = true;
                                                //objPosition.ssi_unrealizedgl.IsNullSpecified = true;
                                                objPosition["ssi_unrealizedgl"] = null;
                                            }

                                            // Managed flag
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["MANAGED"]) != "")
                                            {
                                                //objPosition.ssi_managedcode = new CrmBoolean();
                                                //objPosition.ssi_managedcode.Value = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["MANAGED"]).ToLower());
                                                objPosition["ssi_managedcode"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["MANAGED"]).ToLower().ToLower());
                                            }

                                            // Saving flag
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["SAVINGS"]) != "")
                                            {
                                                //objPosition.ssi_saving = new CrmBoolean();
                                                //objPosition.ssi_saving.Value = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["SAVINGS"]).ToLower());
                                                objPosition["ssi_saving"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["SAVINGS"]).ToLower().ToLower());
                                            }
                                            // Start date
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["START DATE"]) != "")
                                            {
                                                //objPosition.ssi_startdate = new CrmDateTime();
                                                //objPosition.ssi_startdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["START DATE"]);
                                                objPosition["ssi_startdate"] = Convert.ToDateTime(Convert.ToString(ds_gresham.Tables[0].Rows[i]["START DATE"]));
                                            }
                                            // AS OF DATE
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]) != "")
                                            {
                                                //objPosition.ssi_asofdate = new CrmDateTime();
                                                //objPosition.ssi_asofdate.Value = Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]);
                                                objPosition["ssi_asofdate"] = Convert.ToDateTime(Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]));
                                            }
                                            //SecurityId
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]) != "")
                                            {
                                                //objPosition.ssi_securityid = new Lookup();
                                                //objPosition.ssi_securityid.type = EntityName.ssi_security.ToString();
                                                //objPosition.ssi_securityid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"]));
                                                objPosition["ssi_securityid"] = new EntityReference("ssi_security", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_securityid"])));
                                            }
                                            //sas_assetclassid
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]) != "")
                                            {
                                                //objPosition.ssi_assetclassid = new Lookup();
                                                //objPosition.ssi_assetclassid.type = EntityName.sas_assetclass.ToString();
                                                //objPosition.ssi_assetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"]));
                                                objPosition["ssi_assetclassid"] = new EntityReference("sas_assetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sas_assetclassid"])));
                                            }
                                            //objPosition.ssi_transactioncodeid = new Lookup();
                                            //objPosition.ssi_transactioncodeid.type = EntityName.ssi_transactionlog.ToString();
                                            //objPosition.ssi_transactioncodeid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_transactioncodeid"]));

                                            //CustodianId
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["custodianid"]) != "")
                                            {
                                                //objPosition.ssi_custodianid = new Lookup();
                                                //objPosition.ssi_custodianid.type = EntityName.account.ToString();
                                                //objPosition.ssi_custodianid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["custodianid"]));
                                                objPosition["ssi_custodianid"] = new EntityReference("account", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["custodianid"])));
                                            }

                                            // objPosition.ssi_adventreportcode = Convert.ToString(ds_gresham.Tables[0].Rows[i]["REPORT CODE"]);
                                            objPosition["ssi_adventreportcode"] = Convert.ToString(ds_gresham.Tables[0].Rows[i]["REPORT CODE"]);

                                            // Currency ( default to USD )
                                            //objPosition.transactioncurrencyid = new Lookup();
                                            //objPosition.transactioncurrencyid.type = EntityName.transactioncurrency.ToString();
                                            //objPosition.transactioncurrencyid.Value = new Guid("215A7268-A2E1-DD11-A826-001D09665E8F");
                                            objPosition["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("215A7268-A2E1-DD11-A826-001D09665E8F"));


                                            //below 3 col added on 19 dec 2013
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]) != "")
                                            {
                                                //objPosition.ssi_subassetclassid = new Lookup();
                                                //objPosition.ssi_subassetclassid.type = EntityName.ssi_subassetclass.ToString();
                                                //objPosition.ssi_subassetclassid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"]));
                                                objPosition["ssi_subassetclassid"] = new EntityReference("ssi_subassetclass", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_subassetclassId"])));
                                            }

                                            //SubAssetclassid 
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"]) != "")
                                            {
                                                //objPosition.ssi_benchmarkid = new Lookup();
                                                //objPosition.ssi_benchmarkid.type = EntityName.sas_benchmark.ToString();
                                                //objPosition.ssi_benchmarkid.Value = new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"]));
                                                objPosition["ssi_benchmarkid"] = new EntityReference("sas_benchmark", new Guid(Convert.ToString(ds_gresham.Tables[0].Rows[i]["Ssi_BenchmarkSubAssetClassId"])));
                                            }

                                            //greshamadvised (sectorflg)
                                            if (Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]) != "")
                                            {
                                                //objPosition.ssi_greshamadvised = new CrmBoolean();
                                                //objPosition.ssi_greshamadvised.Value = Convert.ToBoolean(ds_gresham.Tables[0].Rows[i]["sectorflg"]);
                                                objPosition["ssi_greshamadvised"] = Convert.ToBoolean(Convert.ToString(ds_gresham.Tables[0].Rows[i]["sectorflg"]).ToLower());
                                            }

                                            objPosition["ssi_datasource"] = new Microsoft.Xrm.Sdk.OptionSetValue(100000000);   // change by AbhiS 12/12/2017


                                            Guid newPositionID = service.Create(objPosition);
                                            //Thread.Sleep(sleepTime);
                                            successCount = successCount + 1;
                                        }
                                        else
                                            break;
                                    }
                                    catch (System.Web.Services.Protocols.SoapException exc)
                                    {
                                        bProceed = true;
                                        positionSuccess = false;
                                        failiureCount = failiureCount + 1;
                                        string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]) + ", MKT VALUE:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["MKT VALUE"]);
                                        strDescription = "Insert failed for Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position " + failiureText + " Error Detail: " + exc.Detail.InnerText;
                                        LogMessage(sw, service, strDescription, 27, "Position");
                                        //AddException("Insert failed for Position", "Account : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]), " AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]), " MKT VALUE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["MKT VALUE"]), "IdNmb:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IdNmb"]),"", "", exc.Message, dt);
                                        AddException("Insert failed for Position", "Account : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AccountName"]), "Security Name : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecurityName"]), "SecurityType :" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecurityType"]), "Security Symbol " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecuritySymbol"]), "Trade Date: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]), "MKT VALUE " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["MKT VALUE"]), "AssetClassName : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AssetClassName"]), "SubAssetClassName : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SubAssetClassName"]), exc.Message, dt);

                                    }
                                    catch (Exception exc)
                                    {
                                        bProceed = true;
                                        positionSuccess = false;
                                        failiureCount = failiureCount + 1;
                                        string failiureText = "Account:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["ssi_accountid"]) + ", AS OF DATE: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]) + ", MKT VALUE:" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["MKT VALUE"]);
                                        strDescription = "Insert failed for Position (IDNMB) : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["IDNMB"]) + " for Position" + failiureText + " Error Detail: " + exc.Message;
                                        LogMessage(sw, service, strDescription, 27, "Position");
                                        AddException("Insert failed for Position", "Account : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AccountName"]), "Security Name : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecurityName"]), "SecurityType :" + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecurityType"]), "Security Symbol " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SecuritySymbol"]), "Trade Date: " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AS OF DATE"]), "MKT VALUE " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["MKT VALUE"]), "AssetClassName : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["AssetClassName"]), "SubAssetClassName : " + Convert.ToString(ds_gresham.Tables[0].Rows[i]["SubAssetClassName"]), exc.Message, dt);
                                    }
                                }

                            LG.AddinLogFile(Form1.vLogFile, "Insert Ends for Position on: " + DateTime.Now.ToString());

                            strDescription = "Total Position failed to insert: " + failiureCount;
                            LogMessage(sw, service, strDescription, 43, "Position");

                            strDescription = "Total Position inserted: " + successCount;
                            LogMessage(sw, service, strDescription, 28, "Position");
                            LG.AddinLogFile(Form1.vLogFile, "---------------------------- New Position Insert Ends -------------------");
                            LG.AddinLogFile(Form1.vLogFile, " ");
                            ds_gresham.Dispose();
                            dagersham.Dispose();

                            #endregion
                            ////////////////////////////////// Insert Ends //////////////////////////////////////////
                        }
                        #endregion
                    }
                }

                try
                {
                    string vContain1 = "-------Before Excel  -" + DateTime.Now.ToString() + "-------";
                    LG.AddinLogFile(Form1.vLogFile, vContain1);
                    DataSet data = GetReportData(service, this.sw);
                    if (data.Tables[1].Rows.Count > 0)
                    {
                        GenerateExcel(data);
                    }
                    string vContain2 = "-------GenerateExcel(data) Count--" + data.Tables[1].Rows.Count;
                    LG.AddinLogFile(Form1.vLogFile, vContain2);

                    vContain2 = "-------After Excel  -" + DateTime.Now.ToString() + "-------";
                    LG.AddinLogFile(Form1.vLogFile, vContain2);

                    vContain2 = "-------GenerateExcel(data) After Excel  Count--" + data.Tables[0].Rows.Count;
                    LG.AddinLogFile(Form1.vLogFile, vContain2);

                      this.generatesExcelsheets(service, this.sw);



                    //    PositionErrorFileExcel(service, this.sw);

                    vContain2 = "Row count  -" + dt.Rows.Count + DateTime.Now.ToString();
                    LG.AddinLogFile(Form1.vLogFile, vContain2);
                    if (dt.Rows.Count > 0)
                    {
                        this.ExportErrorExcel(dt);
                        vContain2 = "Export done  -" + dt.Rows.Count + DateTime.Now.ToString();
                        LG.AddinLogFile(Form1.vLogFile, vContain2);
                    }

                }
                catch (System.Web.Services.Protocols.SoapException exception16)
                {
                    string str5 = "Failed to generate error file. Error Detail:" + exception16.Detail.InnerText;
                    LogMessage(this.sw, service, str5, 62, "AxysLoadErrorFile");
                    //  LG.AddinLogFile(Form1.vLogFile, str5);
                }
                catch (Exception exception17)
                {
                    string str5 = "Failed to generate error file. Error Detail:" + exception17.Message;
                    LogMessage(this.sw, service, str5, 62, "AxysLoadErrorFile");
                    //   LG.AddinLogFile(Form1.vLogFile, str5);
                }

                //sw.Flush();
                //sw.Close();

                //DataSet dsData = getDataSet("EXEC SP_S_NEW_SECURITY_LIST");    //@AsOfDate='" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "'


                string excelFile = "";
                //if (dtExcel != null||dtExcel1!=null)
                //{
                //    if (dtExcel.Rows.Count > 0 )
                //    {
                //        excelFile = ExcelGenrate(dtExcel);
                //    }
                //    else
                //        excelFile = "";
                //}



                excelFile = ExcelGenrate(dtExcel, dtExcel1);
                if (excelFile != "")
                    SendEmail("Please Find Attached List of Newly Created Securities....", "New Securities Created", "", excelFile);

                //  lblMessage.Text = "Load Complete..";
            }

        }

        private void AddException(string Column1, string Column2, string Column3, string Column4, string Column5, string Column6, string Column7, string Column8, string Column9, string Column10, DataTable dt)
        {
            DataRow dr = dt.NewRow();
            dr["Column1"] = Column1;
            dr["Column2"] = Column2;
            dr["Column3"] = Column3;
            dr["Column4"] = Column4;
            dr["Column5"] = Column5;
            dr["Column6"] = Column6;
            dr["Column7"] = Column7;
            dr["Column8"] = Column8;
            dr["Column9"] = Column9;
            dr["Column10"] = Column10;
            dt.Rows.Add(dr);
        }


        private DataSet GetReportData(IOrganizationService service, StreamWriter sw)
        {
            string Gresham_String = ConfigurationManager.AppSettings["Gresham_String_db"].ToString();

            string connectionString = Gresham_String;// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=TransactionLoad_DB;Data Source=GRPAO1-VWSQL02";
            SqlConnection selectConnection = new SqlConnection(connectionString);
            SqlCommand selectCommand = new SqlCommand();
            SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet dataSet = new DataSet();
            DataSet set2 = new DataSet();
            string selectCommandText = string.Empty;
            string strDescription = string.Empty;
            try
            {
                selectCommandText = "exec [SP_S_POSITION_TRANSACTION_EXCEL_NEW_GA] @LoadTypeFlg='" + cbAccountSource.SelectedValue + "'";
                adapter = new SqlDataAdapter(selectCommandText, selectConnection);
                dataSet = new DataSet();
                selectCommand.Connection = selectConnection;
                selectCommand.CommandText = selectCommandText;
                selectCommand.CommandTimeout = 600;
                new SqlDataAdapter(selectCommand).Fill(dataSet);
            }
            catch (System.Web.Services.Protocols.SoapException exception)
            {
                strDescription = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exception.Detail.InnerText;
                //  LogMessage(sw, service, strDescription, 62, "GeneralError");
                LG.AddinLogFile(Form1.vLogFile, strDescription);
            }
            catch (Exception exception2)
            {
                strDescription = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exception2.Message;
                //  LogMessage(sw, service, strDescription, 62, "GeneralError");
                LG.AddinLogFile(Form1.vLogFile, strDescription);
            }
            return dataSet;
        }

        private DataTable CreateDataTable()
        {
            DataTable dt = new DataTable();
            dt.Clear();
            dt.Columns.Add("Column1");
            dt.Columns.Add("Column2");
            dt.Columns.Add("Column3");
            dt.Columns.Add("Column4");
            dt.Columns.Add("Column5");
            dt.Columns.Add("Column6");
            dt.Columns.Add("Column7");
            dt.Columns.Add("Column8");
            dt.Columns.Add("Column9");
            dt.Columns.Add("Column10");

            return dt;
        }
        private void LogMessage(StreamWriter sw, IOrganizationService service, string strDescription, int intIssueType, string strFileLoading)
        {
            try
            {
                //sw.WriteLine(strDescription);
                LG.AddinLogFile(Form1.vLogFile, strDescription);


                //   ssi_loadlog objLoadLog = new ssi_loadlog();
                Entity objLoadLog = new Entity("ssi_loadlog");
                //objLoadLog.ssi_name =Convert.ToString(DateTime.Today);

                //objLoadLog.ssi_date = new CrmDateTime();
                //objLoadLog.ssi_date.Value = DateTime.Now.ToString();
                objLoadLog["ssi_date"] = DateTime.Now;

                // objLoadLog.ssi_fileloading = strFileLoading;
                objLoadLog["ssi_fileloading"] = strFileLoading;

                // objLoadLog.ssi_descriptionofissue = strDescription;
                objLoadLog["ssi_descriptionofissue"] = strDescription;

                //objLoadLog.ssi_typeofissue = new Picklist();
                //objLoadLog.ssi_typeofissue.Value = intIssueType;
                objLoadLog["ssi_typeofissue"] = new Microsoft.Xrm.Sdk.OptionSetValue(intIssueType);


                service.Create(objLoadLog);
            }
            catch (Exception e)
            {
                string str5 = "Error in LogMessage" + e.ToString();
                //Logs LG1 = new Logs();
                LG.AddinLogFile(Form1.vLogFile, str5);
            }

        }
        protected bool mergeFiles()
        {
            try
            {
                StreamReader rdr = new StreamReader(strFilePath + "transGA.txt");
                string master = rdr.ReadToEnd();
                rdr.Close();

                rdr = new StreamReader(strFilePath + "AxysLoadSampleData.txt");
                string newdata = rdr.ReadToEnd();
                rdr.Close();
                rdr.Dispose();
                //The New Data .csv file will have headers. Need to remove those.
                newdata = newdata.Substring(newdata.IndexOf('\n') + 1);
                StreamWriter wtr = new StreamWriter(strFilePath + "transGA.txt");
                wtr.Write(master + "\r\n" /*That \r\n may or may not be necessary*/ + newdata);
                wtr.Close();
                wtr.Dispose();
                string vContain = "-------MergeFiles Complete ------";
                LG.AddinLogFile(Form1.vLogFile, vContain);
                return true;
            }
            catch (Exception e)
            {
                string vContain = "-------MergeFiles Error -" + e.ToString();
                LG.AddinLogFile(Form1.vLogFile, vContain);
                return false;
            }
        }

        public DataSet getDataSet(string sqlQuery)
        {
            try
            {
                string Gresham_String = ConfigurationManager.AppSettings["Gresham_String_db"].ToString();
                SqlConnection Gresham_con = new SqlConnection(Gresham_String);
                DataSet ds = new DataSet();
                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter dagersham = new SqlDataAdapter();


                cmd.Connection = Gresham_con;
                cmd.CommandText = sqlQuery;
                cmd.CommandTimeout = 600;
                dagersham = new SqlDataAdapter(cmd);
                ds = new DataSet();
                dagersham.Fill(ds);

                return ds;
            }
            catch (Exception e)
            {
                string strDescription = e.ToString() + e.StackTrace;
                LG.AddinLogFile(Form1.vLogFile, strDescription);
                return null;
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            txtHide.Text = "";
            txtHide.Text = dateTimePicker1.Value.ToString("MM/dd/yyyy");
            // txtHide.Visible = true;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnContinue_Click(object sender, EventArgs e)
        {
            lblMessage.Text = "Working...";
            lblMessage.Refresh();

            potfolioCodeUpdate(service);
            saveToCrm(service);
            btnLoad.Visible = true;
            btnLoad.Enabled = true;
            btnContinue.Visible = false;
            btnCancel.Visible = false;

            lblMessage.Text = "Load Completed...";
            lblMessage.Refresh();
            dGVSecurity.Visible = false;

        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            btnLoad.Enabled = false;

           
            Loads();
            btnLoad.Enabled = true;
            //  dGVSecurity.Visible = false;
            //  string dt = ;
            //DataSet dsData = getDataSet("EXEC SP_S_NEW_SECURITY_LIST @AsOfDate='" + dateTimePicker1.Value.ToString("MM/dd/yyyy") + "'");
            //dGVSecurity.Visible = true;
            //dGVSecurity.DataSource = dsData.Tables[0];
            //btnLoad.Visible = false;
            //btnContinue.Visible = true;
            //btnCancel.Visible = true;
        }

        public void SendEmail(string mailmessage, string subject, string mailTo, string Attachment1)
        {
            try
            {

                MailMessage myMessage = new MailMessage();
                // SmtpClient SMTPSERVER = new SmtpClient();

                string EmailID = ConfigurationManager.AppSettings["FromEmailId"].ToString(); // AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
                string Password = ConfigurationManager.AppSettings["Password"].ToString(); // AppLogic.GetParam(AppLogic.ConfigParam.Password);
                string SMTPHost = ConfigurationManager.AppSettings["SMTPHost"].ToString(); // AppLogic.GetParam(AppLogic.ConfigParam.SMTPHost);
                string ToEmailIDs1 = ConfigurationManager.AppSettings["ToEmailIDs"].ToString(); //AppLogic.GetParam(AppLogic.ConfigParam.ToEmailIDs1);
                int Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"].ToString()); //Convert.ToInt32(AppLogic.GetParam(AppLogic.ConfigParam.Port));

                //string EmailID = "gbhagia@greshampartners.com";
                //string Password = "";
                //string SMTPHost = "10.0.0.2";
                //int Port = 25;

                //int Port = = ConfigurationSettings.AppSettings["EmailId"].ToString();
                //string Password = ConfigurationSettings.AppSettings["Password"].ToString();
                //string SMTPHost = ConfigurationSettings.AppSettings["SMTPHost"].ToString();
                //string ToEmailIDs = ConfigurationSettings.AppSettings["ToEmailIDs"].ToString();
                //int Port = Convert.ToInt32(ConfigurationSettings.AppSettings["Port"]);

                myMessage.From = new MailAddress(EmailID, "APX Load");
                string[] strTo = ToEmailIDs1.Split('|');


                for (int i = 0; i < strTo.Length; i++)
                {
                    if (strTo[i] != "")
                    {
                        myMessage.To.Add(new MailAddress(strTo[i]));
                    }
                }

                // myMessage.Bcc.Add("skane@infograte.com");
                myMessage.Bcc.Add(new MailAddress("auto-emails@infograte.com"));

                myMessage.Subject = subject;

                if (Attachment1 != "")
                    myMessage.Attachments.Add(new Attachment(Attachment1));

                myMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                myMessage.Body = mailmessage;

                myMessage.IsBodyHtml = true;

                SmtpClient SMTPSERVER = new SmtpClient(SMTPHost, Port);
                SMTPSERVER.DeliveryMethod = SmtpDeliveryMethod.Network;
                //SMTPSERVER.Host = SMTPHost;
                //SMTPSERVER.Port = Port;


                SMTPSERVER.EnableSsl = true;
                // smtp.EnableSsl = true;
                SMTPSERVER.UseDefaultCredentials = true;
                System.Net.NetworkCredential basicAuthenticationInfo = new System.Net.NetworkCredential(EmailID, Password);
                SMTPSERVER.Credentials = basicAuthenticationInfo;
                SMTPSERVER.Send(myMessage);

                myMessage.Dispose();
                myMessage = null;
                SMTPSERVER = null;
                LG.AddinLogFile(Form1.vLogFile, "Email Send..");
            }
            catch (Exception ex)
            {

                LG.AddinLogFile(Form1.vLogFile, "Error In Sending Mail :" + ex.ToString()+ex.StackTrace);
            }
        }

        #region Excel Import
        public string ExcelGenrate(DataTable dtData, DataTable dtSecurityType)
        {
            if (dtData.Rows.Count == 0 && dtSecurityType.Rows.Count == 0)
            {
                return "";
            }
            string ExcelFile = AppDomain.CurrentDomain.BaseDirectory.ToString() + @"\ReportOutput\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            string ExcelfilePath = AppDomain.CurrentDomain.BaseDirectory.ToString() + @"\ReportOutput\";
            LG.AddinLogFile(Form1.vLogFile, dtData.Rows.Count.ToString());

            //if (dtData.Rows.Count > 0 )
            //{
            try
            {
                string ExcelFilePath = ExcelFile;// ExcelfilePath + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                if (!Directory.Exists(ExcelfilePath))
                {
                    Directory.CreateDirectory(ExcelfilePath);
                }
                FileInfo newFile = new FileInfo(ExcelFilePath);

                using (OfficeOpenXml.ExcelPackage pck = new OfficeOpenXml.ExcelPackage(newFile))
                {
                    OfficeOpenXml.ExcelWorksheet ws = null;
                    if (dtData.Rows.Count > 0)
                    {
                        ws = pck.Workbook.Worksheets.Add("Security");
                        ws.Cells["A1"].LoadFromDataTable(dtData, true);
                    }
                    if (dtSecurityType.Rows.Count > 0)
                    {
                        ws = pck.Workbook.Worksheets.Add("Security Type");
                        ws.Cells["A1"].LoadFromDataTable(dtSecurityType, true);
                    }
                    pck.Save();
                }
                return ExcelFilePath;
            }
            catch (Exception e)
            {
                string vContain = "Excel Report Genration Fail,  Error " + e.ToString();
                LG.AddinLogFile(Form1.vLogFile, vContain);

                return "";
            }
            //}
            //else
            //{
            //    string vContain = "Excel Report Genration Fail,  Error, Data Not Found";
            //    LG.AddinLogFile(Form1.vLogFile, vContain);

            //    return "";
            //}
        }
        //   public void generatesExcelsheets(IOrganizationService service, StreamWriter sw)
        //   public void generatesExcelsheets(IOrganizationService service, StreamWriter sw)
        ///  
        ///  

        ///  public void generatesExcelsheets()
        public void generatesExcelsheets(IOrganizationService service, StreamWriter sw)
        {
            #region Spire License Code
            string License = ConfigurationSettings.AppSettings["SpireLicense"].ToString();
            Spire.License.LicenseProvider.SetLicenseKey(License);
            Spire.License.LicenseProvider.LoadLicense();
            #endregion

            DataSet reportData = this.GetReportData();
            // DataSet reportData = this.GetReportData(service, sw);
            // DataSet reportData = GetReportData();
            DataSet lodataset = reportData.Copy();
            DataSet set3 = reportData.Copy();
            if (lodataset.Tables[0].Rows.Count >= 1)
            {
                //string strDescription = "Total Transaction and Position not inserted : " + Convert.ToString(lodataset.Tables[0].Rows.Count) + ". Please check excel file for details.";
                //LogMessage(sw, service, strDescription, 43, "Transaction and Position");

                int num3;
                int num4;
                reportData.AcceptChanges();
                lodataset.AcceptChanges();
                DateTime time2 = new DateTime().AddDays(1.0);
                string str = Convert.ToDateTime(lodataset.Tables[0].Rows[0]["As Of Date"]).ToShortDateString();

                int rowcount = lodataset.Tables[0].Rows.Count;
                string str2 = time2.ToShortDateString();
                //reportData = this.RemoveColumns("_", reportData);
                //lodataset = this.RemoveColumns("_", lodataset);
                //lodataset.AcceptChanges();
                string str3 = "AxysLoadErrorFile_" + DateTime.Now.ToString("MMddyyhhmmss") + ".xlsx";
                string fileName = Application.StartupPath + @"\Axys_Transaction_Position_ErrorFile_Template.xlsx";
                string destFileName = Application.StartupPath + @"\ReportOutput\" + str3;
                string str6 = Application.StartupPath + @"\ReportOutput\" + str3.Replace("xlsx", "xml");
                //   string str7 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + destFileName + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
                string str7 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + destFileName + "';Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";    // change by abhi 09/11/2017
                DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");
                FileInfo info = new FileInfo(fileName);
                info.CopyTo(destFileName, true);

                #region not used
                //using (DbConnection connection = factory.CreateConnection())
                //{
                //    connection.ConnectionString = str7;
                //    connection.Open();
                //    string str8 = string.Empty;
                //    for (int j = 0; j < 1; j++)      // lodataset.Tables.Count     only for 1st dataset
                //    {
                //        str8 = Convert.ToString((int)(j + 1));
                //        string str9 = "Insert into [Sheet" + str8 + "$] (";
                //        string str10 = "";
                //        string str11 = "";
                //        num3 = 0;
                //        while (num3 < lodataset.Tables[j].Columns.Count)
                //        {
                //            str11 = str11 + "'" + lodataset.Tables[j].Columns[num3].ColumnName.Replace("'", "''") + "'";
                //            str10 = str10 + "id" + (num3 + 1);
                //            if (num3 < (lodataset.Tables[j].Columns.Count - 1))
                //            {
                //                str11 = str11 + ",";
                //                str10 = str10 + ",";
                //            }
                //            num3++;
                //        }
                //        str9 = str9 + str10 + ") Values (" + str11 + ")";
                //        DbCommand command = connection.CreateCommand();
                //        try
                //        {
                //            command.CommandText = str9;
                //            command.ExecuteNonQuery();
                //        }
                //        catch
                //        {
                //        }
                //        finally
                //        {
                //            if (command != null)
                //            {
                //                command.Dispose();
                //            }
                //        }
                //        num4 = 0;
                //        while (num4 < lodataset.Tables[j].Rows.Count)
                //        {
                //            str9 = "Insert into [Sheet" + str8 + "$] (";
                //            str11 = "";
                //            num3 = 0;
                //            while (num3 < lodataset.Tables[j].Columns.Count)
                //            {
                //                str11 = str11 + "'" + lodataset.Tables[j].Rows[num4][num3].ToString().Replace("'", "''") + "'";
                //                if (num3 < (lodataset.Tables[j].Columns.Count - 1))
                //                {
                //                    str11 = str11 + ",";
                //                }
                //                num3++;
                //            }
                //            str9 = str9 + str10 + ") Values (" + str11 + ")";
                //            command = connection.CreateCommand();
                //            try
                //            {
                //                command.CommandText = str9;
                //                command.ExecuteNonQuery();
                //            }
                //            catch
                //            {
                //            }
                //            finally
                //            {
                //                if (command != null)
                //                {
                //                    command.Dispose();
                //                }
                //            }
                //            num4++;
                //        }
                //    }
                //    connection.Close();
                //}
                #endregion

                // Code changes (remove oledb)
                Workbook workbooknew = new Workbook();
                workbooknew.LoadFromFile(fileName);

                Worksheet sheetnew = workbooknew.Worksheets[0];
                sheetnew.InsertDataTable(lodataset.Tables[0], true, 6, 1);
                workbooknew.SaveToFile(destFileName);

                Workbook workbook = new Workbook();
                workbook.LoadFromFile(destFileName);
                for (int i = 0; i < 1; i++)
                {
                    int num8;
                    Worksheet worksheet = workbook.Worksheets[i];
                    worksheet.PageSetup.TopMargin = 0.25;
                    for (int k = 1; k < 23; k++)
                    {
                        worksheet.Range[1, k].Text = "";
                    }
                    worksheet.Range[5, 1].Text = "Transaction and Position";
                    //worksheet.Range[3, 1].Text = str;
                    worksheet.Range[3, 1].Style.Font.IsItalic = true;
                    worksheet.Range[4, 1].Style.Font.IsItalic = true;
                    num4 = 0;
                    while (num4 < reportData.Tables[i].Rows.Count)
                    {
                        num8 = num4 + 7;
                        num3 = 1;
                        while (num3 <= lodataset.Tables[i].Columns.Count)
                        {
                            if (!string.IsNullOrEmpty(worksheet.Range[num8, num3].Text))
                            {
                                try
                                {
                                    if (!worksheet.Range[num8, num3].Text.Contains("E"))
                                    {
                                        //worksheet.Range[num8, num3].Text = Convert.ToString(Math.Round(Convert.ToDecimal(worksheet.Range[num8, num3].Text), 2));
                                    }
                                    else
                                    {
                                        //worksheet.Range[num8, num3].Text = Convert.ToString(Math.Round(Convert.ToDecimal(Convert.ToDouble(worksheet.Range[num8, num3].Text))));
                                    }
                                }
                                catch
                                {
                                }
                            }
                            if (num4 == 0)
                            {
                                worksheet.Range[6, num3].Style.Font.FontName = "Frutiger 55 Roman";
                                worksheet.Range[6, num3].Style.Font.Size = 9.0;
                                worksheet.Range[6, num3].RowHeight = 12.0;
                                worksheet.Range[6, num3].VerticalAlignment = VerticalAlignType.Bottom;
                                worksheet.Range[6, num3].Style.Font.IsBold = true;
                                worksheet.Range[6, num3].Style.HorizontalAlignment = HorizontalAlignType.Right;
                            }
                            num3++;
                        }
                        num4++;
                    }
                    num4 = 0;
                    while (num4 < reportData.Tables[i].Rows.Count)
                    {
                        num8 = num4 + 7;
                        num3 = 2;
                        while (num3 < lodataset.Tables[i].Columns.Count)
                        {
                            try
                            {
                                if (!string.IsNullOrEmpty(worksheet.Range[num8, num3].Text) && !worksheet.Range[num8, num3].Text.Contains("%"))
                                {
                                    if (worksheet.Range[num8, num3].Text.Contains("("))
                                    {
                                        //worksheet.Range[num8, num3].Text = Convert.ToDouble((double)(-1.0 * Convert.ToDouble(worksheet.Range[num8, num3].Text.Replace("(", "").Replace(")", "")))).ToString();
                                    }
                                    //worksheet.Range[num8, num3].NumberValue = Convert.ToDouble(worksheet.Range[num8, num3].Text);
                                    worksheet.Range[num8, num3].NumberFormat = @"#,##0_);[Black]\(#,##0\)";
                                }
                            }
                            catch
                            {
                            }
                            num3++;
                        }
                        num4++;
                    }
                    for (int m = 0; m < set3.Tables[i].Rows.Count; m++)
                    {
                        int num11 = m + 7;
                        for (int n = 0; n < set3.Tables[i].Columns.Count; n++)
                        {
                            if (Convert.ToString(set3.Tables[i].Rows[m][n]).ToUpper() == "TRUE")
                            {
                                for (int num13 = 0; num13 < reportData.Tables[i].Columns.Count; num13++)
                                {
                                    if (i == 0)
                                    {
                                        if (worksheet.Range[6, num13 + 1].Text.Contains(set3.Tables[i].Columns[n].ColumnName.Replace("_", "")))
                                        {
                                            worksheet.Range[num11, num13 + 1].Style.Color = Color.Yellow;
                                        }
                                    }
                                    else
                                    {
                                        worksheet.Range[num11, num13 + 1].Style.Color = Color.Yellow;
                                    }
                                }
                            }
                        }
                    }
                    for (num4 = 0; num4 < reportData.Tables[i].Rows.Count; num4++)
                    {
                        num8 = num4 + 7;
                        for (num3 = 1; num3 <= lodataset.Tables[i].Columns.Count; num3++)
                        {
                            if (num4 == 0)
                            {
                                worksheet.Range[6, num3].Style.Font.FontName = "Frutiger 55 Roman";
                                worksheet.Range[6, num3].Style.Font.Size = 12.0;
                                worksheet.Range[6, num3].RowHeight = 42.0;
                                worksheet.Range[6, num3].VerticalAlignment = VerticalAlignType.Center;
                                worksheet.Range[6, num3].Style.Font.IsBold = true;
                                worksheet.Range[6, num3].Style.HorizontalAlignment = HorizontalAlignType.Center;
                                worksheet.Range[6, num3].IsWrapText = true;
                                worksheet.Range[6, num3].Style.Color = Color.FromArgb(0xd8, 0xd8, 0xd8);
                            }
                        }
                    }
                    worksheet.Range[6, 1, 6, 11].HorizontalAlignment = HorizontalAlignType.Center;
                    worksheet.Range[7, 1, 5000, 11].HorizontalAlignment = HorizontalAlignType.Right;
                    worksheet.Range[6, 11, 6, 11].RowHeight = 51.0;
                    worksheet.Range[7, 11, 5000, 11].RowHeight = 16.5;
                    worksheet.Range[6, 1, 5000, 1].ColumnWidth = 17.0;
                    worksheet.Range[6, 2, 5000, 2].ColumnWidth = 20.0;
                    worksheet.Range[6, 3, 5000, 3].ColumnWidth = 21.0;
                    worksheet.Range[6, 4, 5000, 4].ColumnWidth = 18.0;
                    worksheet.Range[6, 5, 5000, 5].ColumnWidth = 20.0;
                    worksheet.Range[6, 6, 5000, 6].ColumnWidth = 20.0;
                    worksheet.Range[6, 7, 5000, 7].ColumnWidth = 20.0;
                    worksheet.Range[6, 8, 5000, 8].ColumnWidth = 40.0;
                    worksheet.Range[7, 1, 5000, 1].HorizontalAlignment = HorizontalAlignType.Center;
                    worksheet.Range[7, 5, 5000, 5].HorizontalAlignment = HorizontalAlignType.Right;
                    worksheet.Range[7, 6, 5000, 6].HorizontalAlignment = HorizontalAlignType.Right;
                    worksheet.Range[7, 8, 5000, 8].HorizontalAlignment = HorizontalAlignType.Right;
                    worksheet.Range[7, 9, 5000, 9].HorizontalAlignment = HorizontalAlignType.Right;
                    worksheet.Range[7, 10, 5000, 10].HorizontalAlignment = HorizontalAlignType.Right;
                    worksheet.Range[7, 11, 5000, 11].HorizontalAlignment = HorizontalAlignType.Right;
                    worksheet.Range[7, 12, 5000, 12].HorizontalAlignment = HorizontalAlignType.Right;
                    worksheet.Range[7, 13, 5000, 13].HorizontalAlignment = HorizontalAlignType.Right;
                    worksheet.Range[5, 1, 5, 1].Style.Color = Color.FromArgb(112, 128, 144);
                    worksheet.Range[5, 1, 5, 1].Style.Font.IsBold = true;


                    worksheet.Range[1, 1, rowcount + 7, 9].AutoFitColumns();
                }

                workbook.SaveToFile(destFileName, ExcelVersion.Version2016);

                #region remove
                //workbook.SaveAsXml(str6);
                //workbook = null;
                //XmlDocument document = new XmlDocument();
                //document.Load(str6);
                //XmlElement documentElement = document.DocumentElement;
                //XmlNode lastChild = documentElement.LastChild;
                //XmlNode firstChild = documentElement.FirstChild;
                //// documentElement.RemoveChild(lastChild);
                //foreach (XmlNode node3 in documentElement)
                //{
                //    if (node3.Name == "ss:Worksheet")
                //    {
                //        foreach (XmlNode node4 in node3.ChildNodes)
                //        {
                //            if (node4.Name == "x:WorksheetOptions")
                //            {
                //                foreach (XmlNode node5 in node4.ChildNodes)
                //                {
                //                    if (node5.Name == "x:PageSetup")
                //                    {
                //                        try
                //                        {
                //                            if (!node3.Attributes[0].InnerText.ToLower().Contains("cover"))
                //                            {
                //                                node5.ChildNodes[0].Attributes[1].InnerText = "&C&\"Frutiger 55 Roman,Italic\"&10Page &P of &N&R&\"Frutiger 55 Roman,Italic\"&10&KD8D8D8&D,&T";
                //                            }
                //                            else
                //                            {
                //                                node5.ChildNodes[0].Attributes[1].InnerText = "&R&\"Frutiger 55 Roman,Italic\"&10&KD8D8D8&D,&T";
                //                            }
                //                        }
                //                        catch
                //                        {
                //                        }
                //                    }
                //                }
                //            }
                //        }
                //    }
                //    if (node3.Name == "ss:Styles")
                //    {
                //        foreach (XmlNode node6 in node3.ChildNodes)
                //        {
                //            try
                //            {
                //                foreach (XmlNode node7 in node6.ChildNodes)
                //                {
                //                    if (node7.Name == "ss:Interior")
                //                    {
                //                        if (node7.Attributes[0].InnerText == "#C0C0C0")
                //                        {
                //                            node7.Attributes[0].InnerText = "#B7DDE8";
                //                        }
                //                        if (node7.Attributes[0].InnerText == "#808080")
                //                        {
                //                            node7.Attributes[0].InnerText = "#d3d3d3";
                //                        }
                //                    }
                //                }
                //                foreach (XmlNode node7 in node6.ChildNodes)
                //                {
                //                    if (node7.Name == "ss:Borders")
                //                    {
                //                        foreach (XmlNode node8 in node7.ChildNodes)
                //                        {
                //                            if (node8.Attributes["ss:Color"].InnerText == "#C0C0C0")
                //                            {
                //                                node8.Attributes["ss:Color"].InnerText = "#B7DDE8";
                //                            }
                //                            if (node8.Attributes["ss:Color"].InnerText == "#808080")
                //                            {
                //                                node8.Attributes["ss:Color"].InnerText = "#d3d3d3";
                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //            catch
                //            {
                //            }
                //        }
                //    }
                //}
                //document.Save(str6);
                //document = null;

                //info = null;
                //info = new FileInfo(destFileName);
                //info.Delete();
                //  info = new FileInfo(str6);
                // info.CopyTo(destFileName, true);
                //  info = null;
                #endregion



                #region xml to xlsx
                //Workbook workbook1 = new Workbook();
                //workbook1.LoadFromXml(str6);
                // workbook1.SaveToFile(destFileName, ExcelVersion.Version2016);
                #endregion

                //  new FileInfo(str6).Delete();



                string sourceFileName = Application.StartupPath + @"\ReportOutput\" + str3;
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.InitialDirectory = @"C:\";
                dialog.Filter = "Excel 2003 (*.xls)|*.xls|Excel 2007 file (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                dialog.FileName = str3;
                dialog.AddExtension = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string str13 = dialog.FileName;
                    System.IO.File.Copy(sourceFileName, str13);
                }
            }
        }

        public void PositionErrorFileExcel(IOrganizationService service, StreamWriter sw)
        {
            try
            {

                DataSet reportData = GetReportData(service, sw);
                //   DataSet reportData = GetReportData();

                //DataSet reportData = GetData("exec [sp_r_anziano_position_recon]", service, sw);
                DataSet lodataset = reportData.Copy();
                DataSet set3 = reportData.Copy();

                DataTable dtData = lodataset.Tables[0];

                //dtData.Columns.Remove("AsofDate");
                //dtData.Columns.Remove("_CHIP Underlying Mgr");
                //dtData.Columns.Remove("_CHIP Client Position");

                // string Date = Convert.ToDateTime(reportData.Tables[0].Rows[0][3]).ToString("MM/dd/yyyy");

                string Filename = @"\Axys_Transaction_Position_ErrorFile" + DateTime.Now.ToString("MMddyyhhmmss") + ".xlsx";
                string ExcelFilePath = Application.StartupPath + @"\ReportOutput" + Filename;
                string TemplateFilePath = Application.StartupPath + @"\Axys_Transaction_Position_ErrorFile_Template.xlsx";
                File.Copy(TemplateFilePath, ExcelFilePath);

                if (dtData.Rows.Count >= 1)
                {
                    FileInfo newFile = new FileInfo(ExcelFilePath);

                    using (ExcelPackage pck = new ExcelPackage(newFile))
                    {
                        //  ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Data");
                        //ExcelWorkbook excelWorkBook = pck.Workbook;
                        //ExcelWorksheet ws = excelWorkBook.Worksheets.();
                        //  pck.Workbook.Worksheets.First();
                        //  OfficeOpenXml.ExcelWorksheet ws = pck.Workbook.Worksheets

                        ExcelWorksheet ws = pck.Workbook.Worksheets[1];

                        ws.Cells["A6"].LoadFromDataTable(dtData, true);

                        string RangeCell = "A2:C2";
                        //ws.Cells[1, 2].Value = "Client Position Recon";
                        //ws.Cells[2, 1, 2, 3].Merge = true;
                        //ws.Cells[2, 1, 2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //ws.Cells[2, 1, 2, 3].Style.Font.Size = 14;    //  Frutiger 55 Roman
                        //ws.Cells[2, 1, 2, 3].Style.Font.Name = "Frutiger 55 Roman";
                        //ws.Cells[2, 1, 2, 3].Style.Font.Bold = true;

                        //ws.Cells["A2"].Value = "Client Position Recon";
                        //ws.Cells[2, 1, 2, 3].Merge = true;
                        //ws.Cells[2, 1, 2, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //ws.Cells[2, 1, 2, 3].Style.Font.Size = 14;    //  Frutiger 55 Roman
                        //ws.Cells[2, 1, 2, 3].Style.Font.Name = "Frutiger 55 Roman";
                        //ws.Cells[2, 1, 2, 3].Style.Font.Bold = true;


                        //ws.Cells["A3"].Value = Date;// "12/31/2015";
                        //ws.Cells[3, 1, 3, 3].Merge = true;
                        //ws.Cells[3, 1, 3, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //ws.Cells[3, 1, 3, 3].Style.Font.Size = 12;    //  Frutiger 55 Roman
                        //ws.Cells[3, 1, 3, 3].Style.Font.Name = "Frutiger 55 Roman";
                        //ws.Cells[3, 1, 3, 3].Style.Font.Italic = true;

                        ws.Cells["A5"].Value = "Transaction and Position";
                        ws.Cells[5, 1, 5, 4].Merge = true;
                        ws.Cells[5, 1, 5, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[5, 1, 5, 4].Style.Font.Size = 11;    //  Frutiger 55 Roman
                        ws.Cells[5, 1, 5, 4].Style.Font.Name = "Calibri";
                        ws.Cells[5, 1, 5, 4].Style.Font.Bold = true;

                        ws.Column(1).Width = 17.00;
                        ws.Column(2).Width = 20.00;
                        ws.Column(3).Width = 21.00;
                        ws.Column(4).Width = 21.00;
                        ws.Column(5).Width = 20.00;
                        ws.Column(6).Width = 20.00;
                        ws.Column(7).Width = 21.00;
                        ws.Column(8).Width = 40.00;
                        ws.Column(9).Width = 40.00;

                        ws.Column(1).Style.WrapText = true;
                        ws.Column(2).Style.WrapText = true;
                        ws.Column(3).Style.WrapText = true;
                        ws.Column(4).Style.WrapText = true;
                        ws.Column(5).Style.WrapText = true;
                        ws.Column(6).Style.WrapText = true;
                        ws.Column(7).Style.WrapText = true;
                        ws.Column(8).Style.WrapText = true;
                        ws.Column(9).Style.WrapText = true;
                        ws.Row(6).Height = 51.00;

                        Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#D3D3D3");
                        ws.Cells["A5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells["A5"].Style.Fill.BackgroundColor.SetColor(colFromHex);

                        colFromHex = System.Drawing.ColorTranslator.FromHtml("#B7DDE8");
                        ws.Cells["A6:I6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells["A6:I6"].Style.Fill.BackgroundColor.SetColor(colFromHex);

                        ws.Cells["A6:I6"].Style.Font.Size = 12;
                        ws.Cells["A6:I6"].Style.Font.Name = "Frutiger 55 Roman";
                        ws.Cells["A6:I6"].Style.Font.Bold = true;

                        ws.Cells["A6:I6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells["A6:I6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        int istartRow = 7, iEndRow = 0;

                        iEndRow = dtData.Rows.Count + 7;

                        ws.Cells[istartRow, 1, iEndRow, 9].Style.Font.Size = 11;
                        ws.Cells[istartRow, 1, iEndRow, 9].Style.Font.Name = "Calibri";

                        //for column 1 
                        ws.Cells[istartRow, 1, iEndRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        //for column 2
                        ws.Cells[istartRow, 2, iEndRow, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        //for column 3
                        ws.Cells[istartRow, 3, iEndRow, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        //for column 4
                        ws.Cells[istartRow, 4, iEndRow, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        //for column 5
                        ws.Cells[istartRow, 5, iEndRow, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        //for column 6
                        ws.Cells[istartRow, 6, iEndRow, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        //for column 7
                        ws.Cells[istartRow, 7, iEndRow, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        //for column 8
                        ws.Cells[istartRow, 8, iEndRow, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        //for column 9
                        ws.Cells[istartRow, 9, iEndRow, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        ws.Cells[7, 1, 500, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[7, 5, 500, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[7, 6, 500, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[7, 8, 500, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[7, 9, 500, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[7, 10, 500, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[7, 11, 500, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[7, 12, 500, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        ws.Cells[7, 13, 500, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        ws.Cells[5, 1, 5, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(112, 128, 144));
                        ws.Cells[5, 1, 5, 1].Style.Font.Bold = true;



                        int i = 0;
                        while (i < reportData.Tables[0].Rows.Count)
                        {
                            int no = i + 7;
                            try
                            {
                                //if (!ws.Cells[no, 2].Text.Contains("E"))
                                //{
                                //    ws.Cells[no, 2].Value = Math.Round(Convert.ToDecimal(ws.Cells[no, 2].Text), 2);
                                //}
                                //else
                                //{
                                //    ws.Cells[no, 2].Value = Math.Round(Convert.ToDecimal(Convert.ToDouble(ws.Cells[no, 2].Text)), 2);
                                //}
                            }
                            catch
                            {
                            }

                            ws.Cells[no, 2].Style.Numberformat.Format = "#,###.#0_);[Black](#,###.#0)";   //#,##0.00

                            //    try
                            //    {
                            //        if (!ws.Cells[no, 3].Text.Contains("E"))
                            //        {
                            //            ws.Cells[no, 3].Value = Math.Round(Convert.ToDecimal(ws.Cells[no, 3].Text), 2);
                            //        }
                            //        else
                            //        {
                            //            ws.Cells[no, 3].Value = Math.Round(Convert.ToDecimal(Convert.ToDouble(ws.Cells[no, 3].Text)), 2);
                            //        }
                            //    }
                            //    catch
                            //    {
                            //    }


                            //    ws.Cells[no, 3].Style.Numberformat.Format = "#,###.#0_);[Black](#,###.#0)";   //#,##0.00

                            i++;
                        }

                        int num4 = 0;
                        while (num4 < reportData.Tables[0].Rows.Count)
                        {
                            int num8 = num4 + 7;
                            int num3 = 2;
                            while (num3 < lodataset.Tables[0].Columns.Count)
                            {
                                try
                                {
                                    if (!string.IsNullOrEmpty(ws.Cells[num8, num3].Text) && !ws.Cells[num8, num3].Text.Contains("%"))
                                    {
                                        if (ws.Cells[num8, num3].Text.Contains("("))
                                        {
                                            //   ws.Cells[num8, num3].Value = Convert.ToDouble((double)(-1.0 * Convert.ToDouble(ws.Cells[num8, num3].Text.Replace("(", "").Replace(")", "")))).ToString();
                                        }
                                        //ws.Cells[num8, num3].Value = Convert.ToDouble(ws.Cells[num8, num3].Text);
                                        if (num3 != 7 || num3 != 8)
                                        {
                                            ws.Cells[num8, num3].Style.Numberformat.Format = @"#,##0_);[Black]\(#,##0\)";
                                        }
                                    }
                                }
                                catch
                                {
                                }
                                num3++;
                            }
                            num4++;
                        }
                        num4 = 0;
                        i = 0;

                        try
                        {
                            for (int m = 0; m < set3.Tables[0].Rows.Count; m++)
                            {
                                int num11 = m + 7;
                                for (int n = 0; n < set3.Tables[0].Columns.Count; n++)
                                {
                                    if (Convert.ToString(set3.Tables[0].Rows[m][n]).ToUpper() == "TRUE")
                                    {
                                        for (int num13 = 0; num13 < reportData.Tables[0].Columns.Count; num13++)
                                        {
                                            if (i == 0)
                                            {
                                                try
                                                {
                                                    if (ws.Cells[6, num13 + 1].Value.ToString().Contains(set3.Tables[0].Columns[n].ColumnName.Replace("_", "")))
                                                    {

                                                        ws.Cells[num11, num13 + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                        ws.Cells[num11, num13 + 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                                    }

                                                }
                                                catch { }
                                            }
                                            else
                                            {
                                                ws.Cells[num11, num13 + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                ws.Cells[num11, num13 + 1].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                            }
                                        }
                                    }
                                }
                            }

                        }
                        catch { }

                        pck.Save();
                    }
                    SaveFileDialog dialog = new SaveFileDialog();
                    dialog.InitialDirectory = @"C:\";
                    dialog.Filter = "All files (*.*)|*.*";
                    dialog.FileName = Filename;
                    dialog.AddExtension = true;
                    dialog.ShowHelp = true;
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        string str13 = dialog.FileName;
                        System.IO.File.Copy(ExcelFilePath, str13);
                    }

                }
            }
            catch (Exception exception2)
            {
                string strDescription = "There was an error occured in Excel File, Please contact administrator. <br/>Error Detail:" + exception2.Message;
                // LogMessage(sw, service, strDescription, 62, "GeneralError");
            }
        }

        private void ExportErrorExcel(DataTable dt)
        {
            // Bind table data to Stream Writer to export data to respective folder
            string str3 = "AxysLoadErrorFile1_" + DateTime.Now.ToString("MMddyyhhmmss") + ".xls";
            //string fileName = Application.StartupPath + @"\Axys_Transaction_Position_ErrorFile_Template.xls";

            string sourceFileName = Application.StartupPath + @"\ReportOutput\" + str3;


            string destFileName = Application.StartupPath + @"\ReportOutput\" + str3;

            StreamWriter wr = new StreamWriter(destFileName);
            //// Write Columns to excel file
            //for (int i = 0; i < dt.Columns.Count; i++)
            //{
            //    wr.Write(dt.Columns[i].ToString().ToUpper() + "\t");
            //}
            //wr.WriteLine();
            //write rows to excel file
            for (int i = 0; i < (dt.Rows.Count); i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Rows[i][j] != null)
                    {
                        wr.Write(Convert.ToString(dt.Rows[i][j]) + "\t");
                    }
                    else
                    {
                        wr.Write("\t");
                    }
                }
                wr.WriteLine();
            }
            wr.Close();

            //FileInfo fi = new FileInfo(destFileName);
            //if (fi.Exists)
            //{
            //    System.Diagnostics.Process.Start(destFileName);
            //}

            SaveFileDialog dialog = new SaveFileDialog();
            dialog.InitialDirectory = @"C:\";
            dialog.Filter = "Excel 2003 (*.xls)|*.xls|Excel 2007 file (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            dialog.FileName = str3;
            dialog.AddExtension = true;
            dialog.ShowHelp = true;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string str13 = dialog.FileName;
                System.IO.File.Copy(sourceFileName, str13);
            }
        }

        public string GenerateExcel(DataSet ds)
        {
            #region Spire License Code
            string License = ConfigurationSettings.AppSettings["SpireLicense"].ToString();
            Spire.License.LicenseProvider.SetLicenseKey(License);
            Spire.License.LicenseProvider.LoadLicense();
            #endregion
            try
            {
                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\ReportOutput"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\ReportOutput");
                }

                String lsFileNamforFinalXls = "APX to CRM Recon " + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                string ExcelFilePath = System.Windows.Forms.Application.StartupPath + "\\ReportOutput\\" + lsFileNamforFinalXls;

                if (System.IO.File.Exists(ExcelFilePath))
                {
                    System.IO.File.Delete(ExcelFilePath);
                }
                #region EPP code
                // FileInfo newFile = new FileInfo(ExcelFilePath);

                //using (OfficeOpenXml.ExcelPackage pck = new OfficeOpenXml.ExcelPackage(newFile))
                //{

                //    for (int i = 1; i <= ds.Tables.Count; i++)
                //    {
                //        OfficeOpenXml.ExcelWorksheet ws = null;
                //        //string vContain11 = " i =" + i + "    " + ds.Tables.Count;
                //        LG.AddinLogFile(Form1.vLogFile, " i =" + i + "    " + ds.Tables.Count);

                //        if (i == 1)
                //        {
                //            ws = pck.Workbook.Worksheets.Add("Pivot");
                //        }
                //        else if (i == 2)
                //        {
                //            ws = pck.Workbook.Worksheets.Add("Posit File Data");
                //        }
                //        else if (i == 3)
                //        {
                //            ws = pck.Workbook.Worksheets.Add("CRM Positions Data");
                //        }
                //        else if (i == 4)
                //        {
                //            ws = pck.Workbook.Worksheets.Add("Excluded");
                //        }

                //        if (i != 1 && ds.Tables[i - 1].Rows.Count > 0)
                //        {
                //            ws.Cells["A1"].LoadFromDataTable(ds.Tables[i - 1], true);
                //            WorksheetFormatting(ws);
                //        }
                //        else if (i != 1)
                //        {
                //            ws.Cells["J9:L10"].Merge = true;
                //            ws.Cells["J9:L10"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                //            ws.Cells["J9:L10"].Value = "No Data Found";
                //            ws.Cells["J9:L10"].Style.Font.Size = 16;
                //            //ws.Cells["J9:L10"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                //        }

                //    }
                //    pck.Save();

                //}
                #endregion

                Workbook book = new Workbook();
                book.CreateEmptySheets(ds.Tables.Count);
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    LG.AddinLogFile(Form1.vLogFile, " i =" + i + "    " + ds.Tables[i].Rows.Count);
                    Worksheet sheet = book.Worksheets[i];
                    if (i == 0)
                    {
                        sheet.Name = "Pivot";
                    }
                    else if (i == 1)
                    {
                        sheet.Name = "Posit File Data";
                    }
                    else if (i == 2)
                    {
                        sheet.Name = "CRM Positions Data";
                    }
                    else if (i == 3)
                    {
                        sheet.Name = "Excluded";
                    }

                    if (i != 0 && ds.Tables[i].Rows.Count > 0)
                    {
                        sheet.Range[1, 1, 1, ds.Tables[i].Columns.Count].Style.Font.IsBold = true;
                        sheet.InsertDataTable(ds.Tables[i], true, 1, 1);
                        sheet.Range[1, 1, ds.Tables[i].Rows.Count + 1, ds.Tables[i].Columns.Count].AutoFitColumns();
                        sheet.Range[1, 1, ds.Tables[i].Rows.Count + 1, ds.Tables[i].Columns.Count].Style.HorizontalAlignment = HorizontalAlignType.Center;

                    }
                    else if (i != 0)
                    {
                        sheet.Range[10, 10, 12, 12].Merge();
                        sheet.Range[10, 10, 12, 12].Value = "No Data Found";
                        sheet.Range[10, 10, 12, 12].Style.Font.Size = 16;
                        sheet.Range[10, 10, 12, 12].Style.HorizontalAlignment = HorizontalAlignType.Center;
                        // ws.Cells["J9:L10"].Merge = true;
                        //  ws.Cells["J9:L10"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        // ws.Cells["J9:L10"].Value = "No Data Found";
                        // ws.Cells["J9:L10"].Style.Font.Size = 16;
                    }
                }

                // formating chnages for sheet no 2 (posit file Data)
                Worksheet sheet1 = book.Worksheets[1];
                for (int liCounter = 0; liCounter < ds.Tables[1].Rows.Count; liCounter++)
                {
                    int lisrc = liCounter + 2;
                    string val1 = ds.Tables[1].Rows[liCounter]["QUANTITY"].ToString();
                    try
                    {

                        if (val1 != "")
                        {
                            sheet1.Range[lisrc, 9].Text = "";
                            sheet1.Range[lisrc, 9].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet1.Range[lisrc, 9].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }

                    try
                    {
                        val1 = ds.Tables[1].Rows[liCounter]["UNIT ADJ COST"].ToString();
                        if (val1 != "")
                        {
                            sheet1.Range[lisrc, 12].Text = "";
                            sheet1.Range[lisrc, 12].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet1.Range[lisrc, 12].NumberFormat = "#,##0_);[Black](#,##0)";
                        }

                    }
                    catch { }

                    try
                    {

                        val1 = ds.Tables[1].Rows[liCounter]["ADJUSTED COST"].ToString();
                        if (val1 != "")
                        {
                            sheet1.Range[lisrc, 13].Text = "";
                            sheet1.Range[lisrc, 13].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet1.Range[lisrc, 13].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }

                    try
                    {

                        val1 = ds.Tables[1].Rows[liCounter]["PRICE DATE"].ToString();
                        if (val1 != "")
                        {
                            sheet1.Range[lisrc, 14].Text = "";
                            sheet1.Range[lisrc, 14].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet1.Range[lisrc, 14].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }

                    try
                    {

                        val1 = ds.Tables[1].Rows[liCounter]["MKT VALUE"].ToString();
                        if (val1 != "")
                        {
                            sheet1.Range[lisrc, 15].Text = "";
                            sheet1.Range[lisrc, 15].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet1.Range[lisrc, 15].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }

                    try
                    {
                        val1 = ds.Tables[1].Rows[liCounter]["UNREAL G/L"].ToString();
                        if (val1 != "")
                        {
                            sheet1.Range[lisrc, 16].Text = "";
                            sheet1.Range[lisrc, 16].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet1.Range[lisrc, 16].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }

                }


                // formating chnages for sheet no 2 (CRM position data)
                Worksheet sheet2 = book.Worksheets[2];
                for (int liCounter = 0; liCounter < ds.Tables[2].Rows.Count; liCounter++)
                {
                    int lisrc = liCounter + 2;

                    string val1 = ds.Tables[2].Rows[liCounter]["Market Value"].ToString();

                    try
                    {
                        if (val1 != "")
                        {
                            sheet2.Range[lisrc, 8].Text = "";
                            sheet2.Range[lisrc, 8].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet2.Range[lisrc, 8].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }

                    try
                    {
                        val1 = ds.Tables[2].Rows[liCounter]["Quantity"].ToString();
                        if (val1 != "")
                        {
                            sheet2.Range[lisrc, 9].Text = "";
                            sheet2.Range[lisrc, 9].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet2.Range[lisrc, 9].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }

                    try
                    {
                        val1 = ds.Tables[2].Rows[liCounter]["Price"].ToString();
                        if (val1 != "")
                        {
                            sheet2.Range[lisrc, 10].Text = "";
                            sheet2.Range[lisrc, 10].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet2.Range[lisrc, 10].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }


                }

                // formating chnages for sheet no 2 (excluded)
                Worksheet sheet3 = book.Worksheets[2];
                for (int liCounter = 0; liCounter < ds.Tables[3].Rows.Count; liCounter++)
                {
                    int lisrc = liCounter + 2;
                    string val1 = ds.Tables[3].Rows[liCounter]["QUANTITY"].ToString();
                    try
                    {
                        if (val1 != "")
                        {
                            sheet3.Range[lisrc, 9].Text = "";
                            sheet3.Range[lisrc, 9].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet3.Range[lisrc, 9].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }

                    try
                    {

                        val1 = ds.Tables[3].Rows[liCounter]["UNIT ADJ COST"].ToString();
                        if (val1 != "")
                        {
                            sheet3.Range[lisrc, 12].Text = "";
                            sheet3.Range[lisrc, 12].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet3.Range[lisrc, 12].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }

                    try
                    {
                        val1 = ds.Tables[3].Rows[liCounter]["ADJUSTED COST"].ToString();
                        if (val1 != "")
                        {
                            sheet3.Range[lisrc, 13].Text = "";
                            sheet3.Range[lisrc, 13].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet3.Range[lisrc, 13].NumberFormat = "#,##0_);[Black](#,##0)";
                        }

                    }
                    catch { }

                    try
                    {
                        val1 = ds.Tables[3].Rows[liCounter]["PRICE DATE"].ToString();
                        if (val1 != "")
                        {
                            sheet3.Range[lisrc, 14].Text = "";
                            sheet3.Range[lisrc, 14].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet3.Range[lisrc, 14].NumberFormat = "#,##0_);[Black](#,##0)";
                        }

                    }
                    catch { }

                    try
                    {
                        val1 = ds.Tables[3].Rows[liCounter]["MKT VALUE"].ToString();
                        if (val1 != "")
                        {
                            sheet3.Range[lisrc, 15].Text = "";
                            sheet3.Range[lisrc, 15].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet3.Range[lisrc, 15].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }

                    try
                    {
                        val1 = ds.Tables[3].Rows[liCounter]["UNREAL G/L"].ToString();
                        if (val1 != "")
                        {
                            sheet3.Range[lisrc, 16].Text = "";
                            sheet3.Range[lisrc, 16].NumberValue = Convert.ToDouble(val1);  //Convert.ToString(Convert.ToDouble(val));
                            sheet3.Range[lisrc, 16].NumberFormat = "#,##0_);[Black](#,##0)";
                        }
                    }
                    catch { }


                }

                book.SaveToFile(ExcelFilePath, ExcelVersion.Version2010);


                LG.AddinLogFile(Form1.vLogFile, "Excel Report Generated Succesfully ");

                #region Throw EXCEL

                SaveFileDialog dialog = new SaveFileDialog();
                dialog.InitialDirectory = @"C:\";
                dialog.Filter = "Excel 2003 (*.xls)|*.xls|Excel 2007 file (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                dialog.FileName = lsFileNamforFinalXls;
                dialog.AddExtension = true;
                dialog.ShowHelp = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string str13 = dialog.FileName;
                    System.IO.File.Copy(ExcelFilePath, str13);
                }
                #endregion



                return ExcelFilePath;
            }
            catch (Exception e)
            {
                string vContain = "-------Error Generating excel APX to CRM Recon.xlsx  -" + e.ToString() + " " + DateTime.Now.ToString() + "-------";
                LG.AddinLogFile(Form1.vLogFile, vContain);
                return "";
            }
        }
        public void WorksheetFormatting(OfficeOpenXml.ExcelWorksheet ws)
        {
            int totalCols = ws.Dimension.End.Column;
            var headerCells = ws.Cells[1, 1, 1, totalCols];
            var headerFont = headerCells.Style.Font;
            headerFont.Bold = true;

            int totalRows = ws.Dimension.End.Row;
            var allCells = ws.Cells[1, 1, totalRows, totalCols];
            allCells.AutoFitColumns();
            allCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        private DataSet GetReportData()
        {
            string Gresham_String = ConfigurationManager.AppSettings["Gresham_String_db"].ToString();

            string connectionString = Gresham_String;// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=TransactionLoad_DB;Data Source=GRPAO1-VWSQL02";
            SqlConnection selectConnection = new SqlConnection(connectionString);
            SqlCommand selectCommand = new SqlCommand();
            SqlDataAdapter adapter = new SqlDataAdapter();
            DataSet dataSet = new DataSet();
            DataSet set2 = new DataSet();
            string selectCommandText = string.Empty;
            string strDescription = string.Empty;
            try
            {
                //selectCommandText = "exec [SP_S_POSITION_TRANSACTION_EXCEL_NEW_GA]";
                selectCommandText = "exec [SP_S_POSITION_TRANSACTION_EXCEL_NEW_GA] @LoadTypeFlg='" + cbAccountSource.SelectedValue + "'";
                adapter = new SqlDataAdapter(selectCommandText, selectConnection);
                dataSet = new DataSet();
                selectCommand.Connection = selectConnection;
                selectCommand.CommandText = selectCommandText;
                selectCommand.CommandTimeout = 600;
                new SqlDataAdapter(selectCommand).Fill(dataSet);
            }
            catch (System.Web.Services.Protocols.SoapException exception)
            {
                strDescription = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exception.Detail.InnerText;
                // LogMessage(sw, service, strDescription, 62, "GeneralError");
                LG.AddinLogFile(Form1.vLogFile, strDescription);
            }
            catch (Exception exception2)
            {
                strDescription = "There was an error occured, Please contact administrator. <br/>Error Detail:" + exception2.Message;
                //LogMessage(sw, service, strDescription, 62, "GeneralError");
                LG.AddinLogFile(Form1.vLogFile, strDescription);
            }
            return dataSet;
        }

        #endregion

        private void cbUsedDate_CheckedChanged(object sender, EventArgs e)
        {
            if (cbUsedDate.Checked)
            {
                txtHide.Enabled = true;
            }
            else
                txtHide.Enabled = false;
        }

        public bool startTruncate()
        {
            string strDescription = "-----Truncate Start -------";
            LG.AddinLogFile(Form1.vLogFile, strDescription);

            try
            {
                string Gresham_String = ConfigurationManager.AppSettings["Gresham_String_db"].ToString();

                SqlConnection Gresham_con = new SqlConnection(Gresham_String);
                DataTable dt = CreateDataTable();

                SqlConnection CRM_con;
                SqlCommand cmd = new SqlCommand();
                SqlDataAdapter dagersham = new SqlDataAdapter();
                SqlDataAdapter da_CRM;
                DataSet ds_gresham = new DataSet();
                DataSet ds = new DataSet();

                Gresham_con = new SqlConnection(Gresham_String);
                Gresham_con.Open();
                string greshamquery = "SP_S_ExecuteJobs @TypeId = 0";
                cmd = new SqlCommand();
                cmd.Connection = Gresham_con;
                cmd.CommandText = greshamquery;
                // cmd.CommandTimeout = 600;
                cmd.ExecuteNonQuery();
                Gresham_con.Close();
                 strDescription = "Truncate Completed EXEC SP_S_ExecuteJobs @TypeId = 0 ";
                LG.AddinLogFile(Form1.vLogFile, strDescription);

                return true;
            }
            catch (Exception Ex)
            {
                string vContain = "-------Truncate Error ----" + Ex.ToString();
                LG.AddinLogFile(Form1.vLogFile, vContain);
                return false;
            }
        
        }
    }
}
