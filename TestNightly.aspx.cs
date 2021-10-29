using Microsoft.IdentityModel.Claims;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class TestNightly : System.Web.UI.Page
{
    Logs lg = new Logs();
    public StreamWriter sw = null;
    Microsoft.Xrm.Sdk.IOrganizationService service;
    GeneralMethods clsGM = new GeneralMethods();

    protected void Page_Load(object sender, EventArgs e)
    {
        bool isAdvancedMode = false;
        /***********NightlyLoad***********************/
        isAdvancedMode = (Request.QueryString["id"] ?? String.Empty).Equals("Scheduled");

        #region LOGFILE
        DateTime dtmain = DateTime.Now;
        string LogFileName = string.Empty;
        LogFileName = "Log-" + DateTime.Now;
        LogFileName = LogFileName.Replace(":", "-");
        LogFileName = LogFileName.Replace("/", "-");
     //   LogFileName = Server.MapPath("") + @"\Logs" + "/" + LogFileName + ".txt";
        LogFileName =@"E:\Infograte\Site\ReportingCustom\Logs\" +LogFileName+".txt";
          
        sw = new StreamWriter(LogFileName);
        sw.Close();
        HttpContext.Current.Session["Filename"] = LogFileName;
        ViewState["Filename"] = LogFileName;



        LogFileName = (string)ViewState["Filename"];

        Session["Filename"] = LogFileName;

        string filenale = (string)Session["Filename"];

        lg.AddinLogFile(Session["Filename"].ToString(), "Started " + dtmain);
        #endregion
        /***********NightlyLoad***********************/


        try
        {

            if (isAdvancedMode)
            {

                string UserId = GetcurrentUser();
                lg.AddinLogFile(Session["Filename"].ToString(), "UserId " + UserId + "---" + DateTime.Now.ToShortDateString());
                service = clsGM.GetCrmService();


                Guid userid = new Guid(UserId);
                Guid teamid = new Guid("673E55EF-9A7F-DF11-8A9A-001D09665E8F");

                bool opsteamflg = IsTeamMember(teamid, userid, service);

                lg.AddinLogFile(Session["Filename"].ToString(), "opsteamflg " + opsteamflg + "---" + DateTime.Now.ToShortDateString());
                    //  Button1.Visible = false;// TESTING 
                    Button1_Click(Button1, new EventArgs());


               

            }
        }

        catch (Exception ex)
        {
            lg.AddinLogFile(Session["Filename"].ToString(), "Eror " + ex.Message.ToString() + "---" + DateTime.Now.ToShortDateString());
        }
    }
    public static bool IsTeamMember(Guid teamID, Guid userID, IOrganizationService service)
    {
        QueryExpression query = new QueryExpression("team");
        query.ColumnSet = new ColumnSet(true);
        query.Criteria.AddCondition(new ConditionExpression("teamid", ConditionOperator.Equal, teamID));
        LinkEntity link = query.AddLink("teammembership", "teamid", "teamid");
        link.LinkCriteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, userID));
        var results = service.RetrieveMultiple(query);


        if (results.Entities.Count > 0)
        {
            return true;
        }
        else
        {
            return false;
        }
    }
    protected void Button1_Click(object sender, EventArgs e)
    {


        lg.AddinLogFile(Session["Filename"].ToString(), "RUNNING SUCCESSFULLY" + "---" + DateTime.Now.ToShortDateString());
    }
    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        string sqlstr = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        //string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
        string strName = string.Empty;
        //Changed Windows to - ADFS Claims Login 8_9_2019
        if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
        {
            strName = "corp\\gbhagia";
        }
        else
        {
            IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
            strName = claimsIdentity.Name;

        }

        //string strName = @"corp\gbhagia ";// p.Identity.Name;


        //string strName = @"corp\crmadmin";// p.Identity.Name;//////////
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
}