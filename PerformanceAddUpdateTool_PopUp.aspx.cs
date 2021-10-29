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
//using CrmSdk;

using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System.Threading;
using Microsoft.IdentityModel.Claims;
using System.IO;
using System.Net;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;

public partial class PerformanceAddUpdateTool_PopUp : System.Web.UI.Page
{
    Logs lg = new Logs();
    public StreamWriter sw = null;
    public string Filename = "";
    bool bProceed = true;
    string strDescription;
    int totalCount = 0;
    int successcount = 0;
    public string _UUID = string.Empty;
	GeneralMethods clsGM = new GeneralMethods();
	
    DB clsdb = new DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        Response.Cache.SetCacheability(HttpCacheability.NoCache);

        if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "" && Convert.ToString(Request.QueryString["uuid"]) != null)
        {
            _UUID = "'" + Convert.ToString(Request.QueryString["uuid"]) + "'";

        }
        if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["type"]) != "" && Convert.ToString(Request.QueryString["type"]) != null)
        {
            lblPerfType.Text = Convert.ToString(Request.QueryString["type"]);
        }
        if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["name"]) != "" && Convert.ToString(Request.QueryString["name"]) != null)
        {
			lblName.Text = HttpUtility.UrlDecode(Convert.ToString(Request.QueryString["name"]));
        }
        if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["asofdate"]) != "" && Convert.ToString(Request.QueryString["asofdate"]) != null)
        {
            lblDate.Text = Convert.ToDateTime(Request.QueryString["asofdate"]).ToString("MM/dd/yyyy");
        }
    }
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        Response.Write("HERE");
        DateTime dtmain = DateTime.Now;
        string LogFileName = string.Empty;
        LogFileName = "Log-" + DateTime.Now;
        LogFileName = LogFileName.Replace(":", "-");
        LogFileName = LogFileName.Replace("/", "-");
        LogFileName = Server.MapPath("") + @"\Logs" + "/" + LogFileName + ".txt";
        sw = new StreamWriter(LogFileName);
        sw.Close();
        HttpContext.Current.Session["Filename"] = LogFileName;
        ViewState["Filename"] = LogFileName;

        Response.Write("LOGCREATED" + LogFileName);

        LogFileName = (string)ViewState["Filename"];

        Session["Filename"] = LogFileName;

        lg.AddinLogFile(Session["Filename"].ToString(), "Start Page Load " + dtmain);




        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        Type tp = this.GetType();

        // string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        // string orgName = "GreshamPartners";
        //CrmService service = null;
		
		IOrganizationService service = null;
		
        lblMessage.Text = "";

        string UserId = GetcurrentUser();
        lg.AddinLogFile(Session["Filename"].ToString(), "Current UserId " + UserId);
        try
        {
            //service = GetCrmService(crmServerUrl, orgName, UserId);
			service = GetCrmService();
            lg.AddinLogFile(Session["Filename"].ToString(), "Curronnection Succesfull ");
            strDescription = "Crm Service starts successfully";
        }
        // catch (System.Web.Services.Protocols.SoapException exc)
		catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblMessage.Text = strDescription;
            lg.AddinLogFile(Session["Filename"].ToString(), "Curronnection Error : " + strDescription);
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
            lg.AddinLogFile(Session["Filename"].ToString(), "Curronnection Error1 : " + strDescription);
        }

        // service.PreAuthenticate = true;
        // service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        lg.AddinLogFile(Session["Filename"].ToString(), "START : ");
        if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "" && Convert.ToString(Request.QueryString["uuid"]) != null)
        {

            if (txtPerformance.Text == "")
            {
                sb.Append("\n<script type=text/javascript>\n");
                sb.Append("\n alert('Please enter value in performance.');");
                sb.Append("</script>");
                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                return;
            }

            if (!IsDuplicateExists())
            {
                //sas_publicperformance objPerformance = new sas_publicperformance();
				Entity objPerformance = new Entity("sas_publicperformance");

                // objPerformance.sas_performance = new CrmDecimal();
                // objPerformance.sas_performance.Value = Convert.ToDecimal(txtPerformance.Text.Trim());
				
				objPerformance["sas_performance"] = Convert.ToDecimal(txtPerformance.Text.Trim());

                if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "" && Convert.ToString(Request.QueryString["type"]) == "FUND")
                {
                    // objPerformance.ssi_fundid = new Lookup();
                    // objPerformance.ssi_fundid.type = EntityName.ssi_fund.ToString();
                    // objPerformance.ssi_fundid.Value = new Guid(Convert.ToString(Request.QueryString["uuid"]));
					
					objPerformance["ssi_fundid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_fund", new Guid(Convert.ToString(Request.QueryString["uuid"])));
                }

                if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "" && Convert.ToString(Request.QueryString["type"]) == "ACCOUNT")
                {
                    // objPerformance.ssi_clientaccountid = new Lookup();
                    // objPerformance.ssi_clientaccountid.type = EntityName.ssi_account.ToString();
                    // objPerformance.ssi_clientaccountid.Value = new Guid(Convert.ToString(Request.QueryString["uuid"]));
					
					objPerformance["ssi_clientaccountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(Request.QueryString["uuid"])));
                }

                if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "" && Convert.ToString(Request.QueryString["type"]) == "BENCHMARK")
                {
                    // objPerformance.sas_performanceid = new Lookup();
                    // objPerformance.sas_performanceid.type = EntityName.sas_benchmark.ToString();
                    // objPerformance.sas_performanceid.Value = new Guid(Convert.ToString(Request.QueryString["uuid"]));
					
					objPerformance["sas_performanceid"] = new Microsoft.Xrm.Sdk.EntityReference("sas_benchmark", new Guid(Convert.ToString(Request.QueryString["uuid"])));
                }

                // objPerformance.sas_enddate = new CrmDateTime();
                // objPerformance.sas_enddate.Value = lblDate.Text.Trim();
				
				objPerformance["sas_enddate"] = Convert.ToDateTime(lblDate.Text.Trim());

                // objPerformance.sas_startdate = new CrmDateTime();
                // objPerformance.sas_startdate.Value = lblDate.Text.Split('/')[0] + "/01/" + lblDate.Text.Split('/')[2];

				objPerformance["sas_startdate"] = Convert.ToDateTime(lblDate.Text.Split('/')[0] + "/01/" + lblDate.Text.Split('/')[2]);

                service.Create(objPerformance);
                successcount++;

                Response.Write(successcount.ToString());


                if (successcount > 0)
                {
                    Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "close", "<script type='text/javascript'>ReturnToParent('true');</script>");
                }
            }
            else
            {
                lblMessage.Visible = true;
                if (lblMessage.Text == "")
                    lblMessage.Text = "Record already exists.";
            }
        }
    }
    public IOrganizationService GetCrmService()
    {
        Microsoft.Xrm.Sdk.IOrganizationService service;
        try
        {


            ClientCredentials Credentials = new ClientCredentials();
            // Credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;

            //  Credentials.Windows.ClientCredential = (NetworkCredential)CredentialCache.DefaultCredentials;

            //string str = ((WindowsIdentity)HttpContext.Current.User.Identity).Name;
            string str = string.Empty;
            if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
            {
                str = "corp\\gbhagia";
            }
            else
            {
                IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
                str = claimsIdentity.Name;

            }
            lg.AddinLogFile(Session["Filename"].ToString(), "User " + str);
            string UserID = string.Empty;
            string sqlstr = "select top 1 internalemailaddress,systemuserid from systemuser where domainname= '" + str + "'";
            DB clsDB = new DB();
            DataSet lodataset = clsDB.getDataSet(sqlstr);
            if (lodataset.Tables[0].Rows.Count > 0)
            {
                UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);

                // Response.Write("UserID:" + UserID);
                //return UserID = "DFCE21B1-B81E-E211-A2B7-0002A5443D86";
            }
            lg.AddinLogFile(Session["Filename"].ToString(), "UserID from DB  " + UserID);

            // *** for specific user credential *********  //
            Credentials.UserName.UserName = "corp\\crmadmin";
            Credentials.UserName.Password = "51ngl3malt_51ngl3malt";




            // Credentials.UserName.UserName = "corp\\crmadmin";
            // Credentials.UserName.Password = "W!gmxF26ggw]";

            //This URL needs to be updated to match the servername and Organization for the environment.
            //  Uri OrganizationUri = new Uri("http://gp-crm2016/GreshamPartners/XRMServices/2011/Organization.svc");
            Uri OrganizationUri = new Uri(AppLogic.GetParam(AppLogic.ConfigParam.CRM2016WebAPI));
            Uri HomeRealmUri = null;



            //OrganizationServiceProxy serviceProxy; 
            System.Net.ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy serviceProxy = new Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, null);

            // This statement is required to enable early-bound type support.
            serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());

            service = (Microsoft.Xrm.Sdk.IOrganizationService)serviceProxy;
            Guid _UserID = new Guid(UserID);
            lg.AddinLogFile(Session["Filename"].ToString(), "Connected to CRM  " + service.ToString());
            serviceProxy.CallerId = _UserID;
            lg.AddinLogFile(Session["Filename"].ToString(), "CallerID set  " + _UserID);
            return service;

        }
        catch (Exception ex)
        {
            service = null;
        }
        return service;
    }
    private bool IsDuplicateExists()
    {
        bool status = false;
        try
        {
            object UUid = "null";

            if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["uuid"]) != "")
            {
                UUid = new Guid(Convert.ToString(Request.QueryString["uuid"]));
            }

            string EndDate = lblDate.Text.Trim();

            string StartDate = lblDate.Text.Split('/')[0] + "/01/" + lblDate.Text.Split('/')[2];

            string query = "exec SP_S_PERFORMANCE_CHECK @UUID='" + UUid + "'"
                                                         + ",@StartDT='" + StartDate + "'"
                                                         + ",@EndDt = '" + EndDate + "'";

            DataSet ds = clsdb.getDataSet(query);

            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    status = Convert.ToBoolean(ds.Tables[0].Rows[0]["ReturnNmb"]);
                }
            }
        }
        catch (Exception ex)
        {
            lblMessage.Text = "Error occured while checking duplicate - " + ex.Message;
            status = true;
        }
        return status;
    }

    // public static CrmService GetCrmService(string crmServerUrl, string organizationName, string CallerId)
    // {
        // // Get the CRM Users appointments
        // // Setup the Authentication Token
        // CrmAuthenticationToken token = new CrmAuthenticationToken();
        // token.AuthenticationType = 0; // Use Active Directory authentication.
        // token.OrganizationName = organizationName;
        // // string username = WindowsIdentity.GetCurrent().Name;

        // if (CallerId != "")
            // token.CallerId = new Guid(CallerId);

        // CrmService service = new CrmService();

        // if (crmServerUrl != null &&
            // crmServerUrl.Length > 0)
        // {
            // UriBuilder builder = new UriBuilder(crmServerUrl);
            // builder.Path = "//MSCRMServices//2007//CrmService.asmx";
            // service.Url = builder.Uri.ToString();
        // }

        // service.CrmAuthenticationTokenValue = token;
        // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        // //////////////////////////// impersonate service to crm user /////////////////////////////

        // // WhoAmIRequest userRequest = new WhoAmIRequest();
        // // Execute the request.
        // // WhoAmIResponse user = (WhoAmIResponse)service.Execute(userRequest);
        // // string currentuser = user.UserId.ToString();


        // //string currentuser = "62DE1F95-8203-DE11-A38C-001D09665E8F";
        // //token.CallerId = new Guid(currentuser);

        // return service;
    // }

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
        }
        else
        {
            return UserID = "27E3A8A5-2A0F-E411-9C15-0002A5443D86";
        }
    }
}
