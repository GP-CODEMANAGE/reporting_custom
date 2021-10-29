using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
using System.IO;
using RKLib.ExportData;
//using CrmSdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.Configuration;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using System.Net;

using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Smo.Agent;
using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class BatchReport_BenchMarkUpload : System.Web.UI.Page
{
    DB clsDB = null;
    SqlConnection cn = null;
    bool bProceed = true;

    string strDescription;
    GeneralMethods clsGM = new GeneralMethods();
    public String _dbErrorMsg;


    protected void Page_Load(object sender, EventArgs e)
    {

    }
    #region old Code
    //public static CrmService GetCrmService(string crmServerUrl, string organizationName, string CallerId)
    //{
    //    // Get the CRM Users appointments
    //    // Setup the Authentication Token
    //    CrmAuthenticationToken token = new CrmAuthenticationToken();
    //    token.AuthenticationType = 0; // Use Active Directory authentication.
    //    token.OrganizationName = organizationName;
    //    // string username = WindowsIdentity.GetCurrent().Name;

    //    if (CallerId != "")
    //        token.CallerId = new Guid(CallerId);

    //    CrmService service = new CrmService();

    //    if (crmServerUrl != null &&
    //        crmServerUrl.Length > 0)
    //    {
    //        UriBuilder builder = new UriBuilder(crmServerUrl);
    //        builder.Path = "//MSCRMServices//2007//CrmService.asmx";
    //        service.Url = builder.Uri.ToString();
    //    }

    //    service.CrmAuthenticationTokenValue = token;
    //    service.Credentials = System.Net.CredentialCache.DefaultCredentials;

    //    //////////////////////////// impersonate service to crm user /////////////////////////////

    //    // WhoAmIRequest userRequest = new WhoAmIRequest();
    //    // Execute the request.
    //    // WhoAmIResponse user = (WhoAmIResponse)service.Execute(userRequest);
    //    // string currentuser = user.UserId.ToString();


    //    //string currentuser = "62DE1F95-8203-DE11-A38C-001D09665E8F";
    //    //token.CallerId = new Guid(currentuser);

    //    return service;
    //}
    #endregion
    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        // string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
        //Changed Windows to - ADFS Claims Login 8_9_2019
        IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
        string strName = claimsIdentity.Name;

        //////////
        //"select top 1 internalemailaddress,systemuserid from systemuser where domainname= 'Signature\\" + strName + "'";
        sqlstr = "select top 1 internalemailaddress,systemuserid from systemuser where domainname= '" + strName + "'";
        DB clsDB = new DB();
        DataSet lodataset = clsDB.getDataSet(sqlstr);
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);
            //return UserID = "DFCE21B1-B81E-E211-A2B7-0002A5443D86";
        }
        else
        {
            return UserID = "";
        }
    }

    protected void Button1_Click1(object sender, EventArgs e)
    {
        int intResult = 0;

        //    sas_publicperformance obj = null;
        Microsoft.Xrm.Sdk.Entity obj = null;


        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        // string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;


        try
        {
            string UserId = GetcurrentUser();

            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblError.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
        }


        if (FileUpload1.HasFile == true)
        {

            if (System.IO.Path.GetExtension(FileUpload1.FileName) == ".xls" || System.IO.Path.GetExtension(FileUpload1.FileName) == ".xlsx")
            {
                if (Request.Url.AbsoluteUri.Contains("localhost"))
                {
                    FileUpload1.PostedFile.SaveAs(@"C:\\Reports\\" + FileUpload1.FileName);

                    if (File.Exists(@"C:\\Reports\\" + FileUpload1.FileName))
                    {
                        File.Delete(@"C:\\Reports\\Public_performance_Data.xlsx");
                        FileUpload1.PostedFile.SaveAs(@"C:\\Reports\\" + FileUpload1.FileName);
                        File.Move(@"C:\\Reports\\" + FileUpload1.FileName, @"C:\\Reports\\Public_performance_Data.xlsx");
                    }
                }
                else
                {
                    try
                    {
                        string extension = System.IO.Path.GetExtension(FileUpload1.FileName);
                        //Response.Write("FileUpload1.FileName:" + FileUpload1.FileName + "<br/><br/><br/>");
                        string strFileName = "Public_performance_Data" + extension;
                        //Response.Write("New FileName:" + strFileName + "<br/><br/><br/>");
                        FileUpload1.PostedFile.SaveAs(AppLogic.GetParam(AppLogic.ConfigParam.PerformanceFilePath) + strFileName);

                        if (File.Exists(AppLogic.GetParam(AppLogic.ConfigParam.PerformanceFilePath) + strFileName))
                        {
                            File.Delete(AppLogic.GetParam(AppLogic.ConfigParam.PerformanceFilePath) + "Public_performance_Data.xlsx");
                            FileUpload1.PostedFile.SaveAs(AppLogic.GetParam(AppLogic.ConfigParam.PerformanceFilePath) + strFileName);
                            File.Move(AppLogic.GetParam(AppLogic.ConfigParam.PerformanceFilePath) + strFileName, AppLogic.GetParam(AppLogic.ConfigParam.PerformanceFilePath) + "Public_performance_Data.xlsx");
                        }
                    }
                    catch (Exception exc)
                    {
                        Response.Write(exc.Message + exc.StackTrace);
                    }
                }
            }


            try
            {
                int retVal = 1;
                string con = AppLogic.GetParam(AppLogic.ConfigParam.DBTransactions);
                try
                {
                    using (SqlConnection connection3 = new SqlConnection(con))
                    {
                        DateTime time;
                        ServerConnection serverConnection = new ServerConnection(connection3);
                        Server server = new Server(serverConnection);
                        Job job = server.JobServer.Jobs["PublicPerformanceData"];
                        JobHistoryFilter filter = new JobHistoryFilter();
                        filter.JobName = "PublicPerformanceData";
                        time = time = job.LastRunDate;
                        job.Start();
                        while (time == job.LastRunDate)
                        {
                            job.Refresh();
                        }
                        if (job.LastRunOutcome == CompletionResult.Succeeded)
                        {
                            retVal = 0;
                        }
                        else
                        {
                            retVal = 1;
                        }
                    }
                }
                catch (Exception exception3)
                {
                    lblError.Text = "Load Job Failed to Execute." + exception3.Message;
                }
                /*
                cn = new SqlConnection(con);

                string strsql = "SP_I_PUBLIC_PERFORMANCE";
                SqlCommand cmd = new SqlCommand();

                SqlParameter returncode = cmd.Parameters.Add("@ReturnIdNmb", SqlDbType.Int);
                returncode.Direction = ParameterDirection.Output;

                //cmd.Parameters["@returncode"].Direction = ParameterDirection.Output;
                cmd.CommandText = strsql;
                cmd.Connection = cn;
                cmd.CommandType = CommandType.StoredProcedure;

                cn.Open();
                //Response.Write(sqlconn.State + "<br/><br/><br/>");
                //Response.Write(sqlconn.Database + "<br/><br/><br/>"); 
                int result = cmd.ExecuteNonQuery();
                //System.Threading.Thread.Sleep(1000);

                int retVal = (int)cmd.Parameters["@ReturnIdNmb"].Value;

                //Response.Write(Convert.ToString(retVal)); 
                */
                if (retVal == 0)
                {
                    clsDB = new DB();

                    #region Update Bench Marks

                    DataSet loDatasetUpdate = new DataSet();
                    loDatasetUpdate = LoadDataSet("EXEC DBO.SP_U_PUBLIC_PERFORMANCE");

                    if (loDatasetUpdate.Tables.Count > 0)
                    {
                        if (loDatasetUpdate.Tables[0].Rows.Count > 0)
                        {
                            for (int j = 0; j < loDatasetUpdate.Tables[0].Rows.Count; j++)
                            {
                                obj = new Microsoft.Xrm.Sdk.Entity("sas_publicperformance");



                                Microsoft.Xrm.Sdk.Entity objPosition = new Microsoft.Xrm.Sdk.Entity("ssi_position");

                                if (Convert.ToString(loDatasetUpdate.Tables[0].Rows[j]["Sas_PublicPerformanceID"]) != "")
                                {
                                    // obj.sas_publicperformanceid = new Key();
                                    // obj.sas_publicperformanceid.Value = new Guid(Convert.ToString(loDatasetUpdate.Tables[0].Rows[j]["Sas_PublicPerformanceID"]));

                                    Guid sas_publicperformanceid = new Guid(Convert.ToString(loDatasetUpdate.Tables[0].Rows[j]["Sas_PublicPerformanceID"]));

                                    obj["sas_publicperformanceid"] = sas_publicperformanceid;

                                    if (Convert.ToString(loDatasetUpdate.Tables[0].Rows[j]["Sas_Performance"]) != "")
                                    {
                                        // obj.sas_performance = new CrmDecimal();
                                        // obj.sas_performance.Value = Convert.ToDecimal(loDatasetUpdate.Tables[0].Rows[j]["Sas_Performance"]);

                                        obj["sas_performance"] = Convert.ToDecimal(loDatasetUpdate.Tables[0].Rows[j]["Sas_Performance"]);
                                    }

                                    service.Update(obj);
                                    intResult++;
                                }
                            }
                        }
                    }





                    #endregion

                    #region Insert Bench Marks

                    DataSet loDataset = new DataSet();
                    loDataset = LoadDataSet("EXEC DBO.SP_S_PUBLIC_PERFORMANCE");

                    if (loDataset.Tables.Count > 0)
                    {
                        if (loDataset.Tables[0].Rows.Count > 0)
                        {
                            for (int i = 0; i < loDataset.Tables[0].Rows.Count; i++)
                            {
                                obj = new Microsoft.Xrm.Sdk.Entity("sas_publicperformance");

                                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_StartDate"]) != "")
                                {
                                    //  obj.sas_startdate = new CrmDateTime();
                                    //  obj.sas_startdate.Value = Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_StartDate"]);

                                    obj["sas_startdate"] = Convert.ToDateTime(Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_StartDate"]));
                                }


                                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_EndDate"]) != "")
                                {
                                    // obj.sas_enddate = new CrmDateTime();
                                    // obj.sas_enddate.Value = Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_EndDate"]);

                                    obj["sas_enddate"] = Convert.ToDateTime(Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_EndDate"]));
                                }

                                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_Performance"]) != "")
                                {
                                    // obj.sas_performance = new CrmDecimal();
                                    // obj.sas_performance.Value = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Sas_Performance"]);

                                    obj["sas_performance"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Sas_Performance"]);
                                }


                                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_PerformanceId"]) != "")//display name benchmark 
                                {
                                    // obj.sas_performanceid = new Lookup();
                                    // obj.sas_performanceid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_PerformanceId"]));

                                    obj["sas_performanceid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["Sas_PerformanceId"])));
                                }

                                if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_ClientAccountID"]) != "")
                                {
                                    // obj.ssi_clientaccountid = new Lookup();
                                    // obj.ssi_clientaccountid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_ClientAccountID"]));

                                    obj["ssi_clientaccountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_ClientAccountID"])));


                                }


                                if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_fundID"]) != "")
                                {
                                    // obj.ssi_fundid = new Lookup();
                                    // obj.ssi_fundid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_fundID"]));

                                    obj["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_fundID"])));


                                }
                                //added 1_8_2019 Dynamo Changes
                                obj["ssi_source"] = new Microsoft.Xrm.Sdk.OptionSetValue(100000001); // Fundadmin

                                service.Create(obj);
                                intResult++;
                            }
                        }
                    }

                    #endregion


                    if (intResult > 0)
                    {
                        #region Excel for data issues
                        string strsql11 = "SP_S_PUBLIC_PERF_NOT_INSERTED_RECORDS";
                        DataSet ds = LoadDataSet(strsql11);

                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            lnkIssues.Visible = true;
                            System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();
                            grid.HeaderStyle.Font.Bold = true;
                            grid.DataSource = ds;

                            grid.DataBind();

                            // render the DataGrid control to a file
                            using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Data_Issues.xls")))
                            {
                                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                                {
                                    grid.RenderControl(hw);
                                }
                            }

                        }
                        else
                        {
                            lnkIssues.Visible = false;
                        }

                        if (ds.Tables[1].Rows.Count > 0)
                        {
                            lnkDuplicate.Visible = true;
                            System.Web.UI.WebControls.DataGrid gridduplicate = new System.Web.UI.WebControls.DataGrid();
                            gridduplicate.HeaderStyle.Font.Bold = true;
                            gridduplicate.DataSource = ds.Tables[1];

                            gridduplicate.DataBind();

                            // render the DataGrid control to a file
                            using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Duplicate_Data.xls")))
                            {
                                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                                {
                                    gridduplicate.RenderControl(hw);
                                }
                            }

                        }
                        else
                        {
                            lnkDuplicate.Visible = false;
                        }

                        if (ds.Tables[2].Rows.Count > 0)
                        {
                            lnkMissing.Visible = true;
                            System.Web.UI.WebControls.DataGrid gridMissing = new System.Web.UI.WebControls.DataGrid();
                            gridMissing.HeaderStyle.Font.Bold = true;
                            gridMissing.DataSource = ds.Tables[2];

                            gridMissing.DataBind();

                            // render the DataGrid control to a file
                            using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Missing_Data.xls")))
                            {
                                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                                {
                                    gridMissing.RenderControl(hw);
                                }
                            }

                        }
                        else
                        {
                            lnkMissing.Visible = false;
                        }

                        if (ds.Tables[4].Rows.Count > 0)
                        {
                            lnkNotLoaded.Visible = true;
                            System.Web.UI.WebControls.DataGrid gridnotloaded = new System.Web.UI.WebControls.DataGrid();
                            gridnotloaded.HeaderStyle.Font.Bold = true;
                            gridnotloaded.DataSource = ds.Tables[4];

                            gridnotloaded.DataBind();

                            // render the DataGrid control to a file
                            using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Data_not_Loaded.xls")))
                            {
                                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                                {
                                    gridnotloaded.RenderControl(hw);
                                }
                            }

                        }
                        else
                        {
                            lnkNotLoaded.Visible = false;
                        }

                        #endregion
                        lblError.Text = Convert.ToString(loDataset.Tables[0].Rows.Count) + " records created <br/>" + Convert.ToString(loDatasetUpdate.Tables[0].Rows.Count) + " records Updated.";
                    }
                    else
                    {
                        #region Excel for data issues
                        string strsql11 = "SP_S_PUBLIC_PERF_NOT_INSERTED_RECORDS";
                        DataSet ds = LoadDataSet(strsql11);

                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            lnkIssues.Visible = true;
                            System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();
                            grid.HeaderStyle.Font.Bold = true;
                            grid.DataSource = ds;

                            grid.DataBind();

                            // render the DataGrid control to a file
                            using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Data_Issues.xls")))
                            {
                                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                                {
                                    grid.RenderControl(hw);
                                }
                            }

                        }
                        else
                        {
                            lnkIssues.Visible = false;
                        }

                        if (ds.Tables[1].Rows.Count > 0)
                        {
                            lnkDuplicate.Visible = true;
                            System.Web.UI.WebControls.DataGrid gridduplicate = new System.Web.UI.WebControls.DataGrid();
                            gridduplicate.HeaderStyle.Font.Bold = true;
                            gridduplicate.DataSource = ds.Tables[1];

                            gridduplicate.DataBind();

                            // render the DataGrid control to a file
                            using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Duplicate_Data.xls")))
                            {
                                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                                {
                                    gridduplicate.RenderControl(hw);
                                }
                            }

                        }
                        else
                        {
                            lnkDuplicate.Visible = false;
                        }

                        if (ds.Tables[2].Rows.Count > 0)
                        {
                            lnkMissing.Visible = true;
                            System.Web.UI.WebControls.DataGrid gridMissing = new System.Web.UI.WebControls.DataGrid();
                            gridMissing.HeaderStyle.Font.Bold = true;
                            gridMissing.DataSource = ds.Tables[2];

                            gridMissing.DataBind();

                            // render the DataGrid control to a file
                            using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Missing_Data.xls")))
                            {
                                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                                {
                                    gridMissing.RenderControl(hw);
                                }
                            }

                        }
                        else
                        {
                            lnkMissing.Visible = false;
                        }

                        if (ds.Tables[4].Rows.Count > 0)
                        {
                            lnkNotLoaded.Visible = true;
                            System.Web.UI.WebControls.DataGrid gridnotloaded = new System.Web.UI.WebControls.DataGrid();
                            gridnotloaded.HeaderStyle.Font.Bold = true;
                            gridnotloaded.DataSource = ds.Tables[4];

                            gridnotloaded.DataBind();

                            // render the DataGrid control to a file
                            using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Data_not_Loaded.xls")))
                            {
                                using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                                {
                                    gridnotloaded.RenderControl(hw);
                                }
                            }

                        }
                        else
                        {
                            lnkNotLoaded.Visible = false;
                        }

                        #endregion
                        lblError.Text = Convert.ToString(loDataset.Tables[0].Rows.Count) + " records created <br/>" + Convert.ToString(loDatasetUpdate.Tables[0].Rows.Count) + " records Updated.";
                    }
                }
                else
                {
                    #region Excel for data issues
                    string strsql11 = "SP_S_PUBLIC_PERF_NOT_INSERTED_RECORDS";
                    DataSet ds = LoadDataSet(strsql11);

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        lnkIssues.Visible = true;
                        System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();
                        grid.HeaderStyle.Font.Bold = true;
                        grid.DataSource = ds;

                        grid.DataBind();

                        // render the DataGrid control to a file
                        using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Data_Issues.xls")))
                        {
                            using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                            {
                                grid.RenderControl(hw);
                            }
                        }

                    }
                    else
                    {
                        lnkIssues.Visible = false;
                    }

                    if (ds.Tables[1].Rows.Count > 0)
                    {
                        lnkDuplicate.Visible = true;
                        System.Web.UI.WebControls.DataGrid gridduplicate = new System.Web.UI.WebControls.DataGrid();
                        gridduplicate.HeaderStyle.Font.Bold = true;
                        gridduplicate.DataSource = ds.Tables[1];

                        gridduplicate.DataBind();

                        // render the DataGrid control to a file
                        using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Duplicate_Data.xls")))
                        {
                            using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                            {
                                gridduplicate.RenderControl(hw);
                            }
                        }

                    }
                    else
                    {
                        lnkDuplicate.Visible = false;
                    }

                    if (ds.Tables[2].Rows.Count > 0)
                    {
                        lnkMissing.Visible = true;
                        System.Web.UI.WebControls.DataGrid gridMissing = new System.Web.UI.WebControls.DataGrid();
                        gridMissing.HeaderStyle.Font.Bold = true;
                        gridMissing.DataSource = ds.Tables[2];

                        gridMissing.DataBind();

                        // render the DataGrid control to a file
                        using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Missing_Data.xls")))
                        {
                            using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                            {
                                gridMissing.RenderControl(hw);
                            }
                        }

                    }
                    else
                    {
                        lnkMissing.Visible = false;
                    }

                    if (ds.Tables[4].Rows.Count > 0)
                    {
                        lnkNotLoaded.Visible = true;
                        System.Web.UI.WebControls.DataGrid gridnotloaded = new System.Web.UI.WebControls.DataGrid();
                        gridnotloaded.HeaderStyle.Font.Bold = true;
                        gridnotloaded.DataSource = ds.Tables[4];

                        gridnotloaded.DataBind();

                        // render the DataGrid control to a file
                        using (StreamWriter sw = new StreamWriter(Server.MapPath("./ExcelTemplate/Public_performance_Data_not_Loaded.xls")))
                        {
                            using (HtmlTextWriter hw = new HtmlTextWriter(sw))
                            {
                                gridnotloaded.RenderControl(hw);
                            }
                        }

                    }
                    else
                    {
                        lnkNotLoaded.Visible = false;
                    }

                    #endregion
                    lblError.Text = "File upload failed.";
                    lblError.Visible = true;
                }

            }
            catch (System.Web.Services.Protocols.SoapException exc1)
            {

                Response.Write("<br/>Exception: " + exc1.Detail.InnerText);

            }
            catch (Exception exc)
            {
                Response.Write(exc.Message + exc.StackTrace);
            }
            finally
            {
                if (cn != null)
                    if (cn.State != System.Data.ConnectionState.Open)
                        cn.Close();
            }



        }
    }

    private SqlConnection OpenConnection()
    {
        try
        {
            //"Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
            //"Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TransactionLoad_DB;Data Source=SQL01";
            string ConnString = AppLogic.GetParam(AppLogic.ConfigParam.DBTransactions);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=TransactionLoad_DB;Data Source=sql01\\crmtest";
            cn = new SqlConnection(ConnString);
            cn.Open();
            return cn;
        }
        catch (Exception ex)
        {
            _dbErrorMsg = ex.Message;
            return cn;
        }
    }

    // To close Database Connection
    private void CloseConnection()
    {
        cn.Close();
        cn.Dispose();
    }

    private DataSet LoadDataSet(String sqlstr)
    {
        cn = OpenConnection();
        SqlDataAdapter da = new SqlDataAdapter(sqlstr, cn);
        da.SelectCommand.CommandTimeout = 1800;//prv - 300
        DataSet ds = new DataSet();
        da.Fill(ds);
        da.Dispose();
        CloseConnection();
        return (ds);
    }
    protected void lnkSamplefile_Click(object sender, EventArgs e)
    {
        string filePath = Server.MapPath("./ExcelTemplate/Public_performance_Data.xls");
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment;filename=Public_performance_Data.xls");
        Response.TransmitFile(filePath);
        Response.End();
    }
    protected void lnkIssues_Click(object sender, EventArgs e)
    {
        string filePath = Server.MapPath("./ExcelTemplate/Public_performance_Data_Issues.xls");
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment;filename=Public_performance_Data_Issues.xls");
        Response.TransmitFile(filePath);
        Response.End();
    }

    protected void lnkDuplicate_Click(object sender, EventArgs e)
    {
        string filePath = Server.MapPath("./ExcelTemplate/Public_performance_Duplicate_Data.xls");
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment;filename=Public_performance_Duplicate_Data.xls");
        Response.TransmitFile(filePath);
        Response.End();
    }

    protected void lnkMissing_Click(object sender, EventArgs e)
    {
        string filePath = Server.MapPath("./ExcelTemplate/Public_performance_Missing_Data.xls");
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment;filename=Public_performance_Missing_Data.xls");
        Response.TransmitFile(filePath);
        Response.End();
    }

    protected void lnkNotLoaded_Click(object sender, EventArgs e)
    {
        string filePath = Server.MapPath("./ExcelTemplate/Public_performance_Data_not_Loaded.xls");
        Response.ContentType = "application/octet-stream";
        Response.AddHeader("Content-Disposition", "attachment;filename=Public_performance_Data_not_Loaded.xls");
        Response.TransmitFile(filePath);
        Response.End();
    }
}
