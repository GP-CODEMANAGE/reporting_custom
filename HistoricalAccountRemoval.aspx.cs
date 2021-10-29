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
using System.Security.Principal;
using System.IO;
using System.Text;
using Spire.Xls;
using System.Xml;
//using CrmSdk;

using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class HistoricalAccountRemoval : System.Web.UI.Page
{
    string sqlstr = string.Empty;
    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = new DB();
    public StreamWriter sw = null;
    public string execType = string.Empty;
    string strDescription = string.Empty;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            fillHousehold();
        }
    }

    public void fillHousehold()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("SP_S_GET_HOUSEHOLDNAME_HISTORICAL");
        ddlHouseHold.Items.Clear();
        ddlHouseHold.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHouseHold.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        #region Declaration
        string LogFileName = "LogFile " + DateTime.Now;
        LogFileName = LogFileName.Replace(":", "-");
        LogFileName = LogFileName.Replace("/", "-");
        sw = new StreamWriter(Request.PhysicalApplicationPath + "\\Log\\" + LogFileName + ".txt", true);
        //sw = new StreamWriter(LogFileName + ".txt", true);

        // string Gresham_String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TransactionLoad_DB;Data Source=SQL01";
        //string Gresham_String = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GreshamPartners_MSCRM;Data Source=SQL01";
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);

        // string CRM_constring = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GreshamPartners_MSCRM;Data Source=SQL01";

        ////string crmServerUrl = "http://Crm01/";
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);
        ////string crmServerURL = "http://server:5555/";

        //string orgName = "GreshamPartners";
        //string orgName = "Webdev";

        string Userid = GetcurrentUser();
        bool bProceed = true;
        string strDescription;
       // CrmService service = null;
        IOrganizationService service = null;
        try
        {
            //service = GetCrmService(crmServerUrl, orgName, Userid);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
            LogMessage(sw, service, strDescription, 62, "GeneralError");
            sw.WriteLine("step 1 ");
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            LogMessage(sw, service, strDescription, 62, "GeneralError");
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            LogMessage(sw, service, strDescription, 62, "GeneralError");
        }

        //service.PreAuthenticate = true;

        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        SqlConnection Gresham_con = new SqlConnection(Gresham_String);

        SqlConnection CRM_con;
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();
        string greshamquery;
        int totalCount = 0;
        int successCount = 0;
        int failiureCount = 0;


        #endregion

        try
        {
            sw.WriteLine("---------------------------- Historical Account Removal Starts -------------------");
            greshamquery = "SP_S_HISTORICAL_DELETE @HHUUID='" + ddlHouseHold.SelectedValue + "'";
            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.SelectCommand.CommandTimeout = 1800;
            dagersham.Fill(ds_gresham);
            //totalCount = ds_gresham.Tables[0].Rows.Count;

            //successCount = Convert.ToInt32(ds_gresham.Tables[0].Rows[0]["DeleteCount"]);

            //strDescription = "Total Historical Account Removed: " + successCount;

            for (int i = 0; i < ds_gresham.Tables.Count; i++)
            {
                string EntityName = ds_gresham.Tables[i].Columns[0].ColumnName.Remove(ds_gresham.Tables[i].Columns[0].ColumnName.Length - 2, 2).ToLower();
                DeleteData(ds_gresham.Tables[i], EntityName, ds_gresham.Tables[i].Columns[0].ColumnName, Userid, sw);
            }

    
            strDescription = "Historical Account Removed successfully for " + ddlHouseHold.SelectedItem.Text;
            LogMessage(sw, service, strDescription, 62, "HistoricalAccountRemoval");
            sw.WriteLine("---------------------------- Historical Account Removal Ends  -------------------");
            lblError.Text = "Historical Account Removed successfully for " + ddlHouseHold.SelectedItem.Text + ".";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            totalCount = 0;
            strDescription = "Historical Account Removal failed, please contact administrator. Error Detail: " + exc.Detail.InnerText;
            LogMessage(sw, service, strDescription, 62, "HistoricalAccountRemoval");
        }
        catch (Exception exc)
        {
            totalCount = 0;
            strDescription = "Historical Account Removal failed, please contact administrator. Error Detail: " + exc.Message;
            LogMessage(sw, service, strDescription, 62, "HistoricalAccountRemoval");
        }
    }

    private void DeleteData(DataTable dt, string entityName, string ColumnName, string UserId, StreamWriter sw)
    {
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);
        //string orgName = "GreshamPartners";
       // CrmService service = null;
        IOrganizationService service = null;
        int successcount = 0;
        try
        {
           // service = GetCrmService(crmServerUrl, orgName, UserId);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            //bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblError.Text = strDescription;
            sw.WriteLine(strDescription);
        }
        catch (Exception exc)
        {
            //bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
            sw.WriteLine(strDescription);
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                Guid UUID = new Guid(Convert.ToString(dt.Rows[j][ColumnName]));
                service.Delete(entityName.ToString(), UUID);
                successcount = successcount + 1;
            }
        }
        catch (Exception ex)
        {
            Response.Write("<br/>" + ex.Message);
            sw.WriteLine(strDescription);
            strDescription = "Falied to Delete " + entityName + " : " + ex.Message;
            LogMessage(sw, service, strDescription, 77, "" + entityName + " Load");
        }
        //return successcount;
    }


   // private static void LogMessage(StreamWriter sw, CrmService service, string strDescription, int intIssueType, string strFileLoading)
    private static void LogMessage(StreamWriter sw, IOrganizationService service, string strDescription, int intIssueType, string strFileLoading)
    {
        try
        {
            sw.WriteLine(strDescription);

            //ssi_loadlog objLoadLog = new ssi_loadlog();
            Entity objLoadLog = new Entity("ssi_loadlog");
            //objLoadLog.ssi_name =Convert.ToString(DateTime.Today);

            //objLoadLog.ssi_date = new CrmDateTime();
            //objLoadLog.ssi_date.Value = DateTime.Now.ToString();
            objLoadLog["ssi_date"] = DateTime.Now;


            //objLoadLog.ssi_fileloading = strFileLoading;
            objLoadLog["ssi_fileloading"] = strFileLoading;

            //objLoadLog.ssi_descriptionofissue = strDescription;
            objLoadLog["ssi_descriptionofissue"] = strDescription;

            //objLoadLog.ssi_typeofissue = new Picklist();
            //objLoadLog.ssi_typeofissue.Value = intIssueType;
            objLoadLog["ssi_typeofissue"] = new Microsoft.Xrm.Sdk.OptionSetValue(intIssueType);


            service.Create(objLoadLog);
        }
        catch (Exception exc)
        {
            //HttpContext.Current.Response.Write(exc.Message.ToString());
            sw.WriteLine(exc.Message);
            sw.Flush();
            sw.Close();
            throw;
        }
    }

    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        //   string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
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

    #region OLDCODE Commented CRM2016 Upgrade
    ///// <summary>
    ///// Set up the CRM Service.
    ///// </summary>
    ///// <param name="organizationName">My Organization</param>
    ///// <returns>CrmService configured with AD Authentication</returns>
    //public static CrmService GetCrmService(string crmServerUrl, string organizationName, string CallerId)
    //{
    //    // Get the CRM Users appointments
    //    // Setup the Authentication Token
    //    CrmAuthenticationToken token = new CrmAuthenticationToken();
    //    token.AuthenticationType = 0; // Use Active Directory authentication.
    //    token.OrganizationName = organizationName;
    //    //string username = WindowsIdentity.GetCurrent().Name;

    //    //if (username == "CORP\\gbhagia")
    //    //{
    //    //    // Use the global user ID of the system user that is to be impersonated.
    //    //    token.CallerId = new Guid("EE8E3A77-59E2-DD11-831F-001D09665E8F");//deb
    //    //    //token.CallerId = new Guid("C42C7E05-8303-DE11-A38C-001D09665E8F");//gary                
    //    //}
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

    //    return service;
    //}
    #endregion

}
