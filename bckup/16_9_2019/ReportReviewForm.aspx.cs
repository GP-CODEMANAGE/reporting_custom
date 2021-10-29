
using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Net;
using System.Collections;
using System.Collections.Generic;
using Spire.Xls;
using System.Drawing;
using System.Data.Common;
using System.Xml;
//using CrmSdk;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Data.SqlClient;
using GemBox.Document;
using GemBox.Document.Tables;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Xls;
using Microsoft.SharePoint.Client;
using System.Security;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Text;

using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class _ReportReviewForm : System.Web.UI.Page
{
    ClientContext context;
    String sqlstr = string.Empty;
    int intResult = 0;
    bool bProceed = true;
    string strDescription;
    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = new DB();
    SqlConnection cn = null;
    public String _dbErrorMsg;
    public string strReportFiles = string.Empty;
    public string BatchGUID = string.Empty;
    public string OPSBatchGUID = string.Empty;

    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    public int numIndexPageCount = 1;  //Index page count -- if count of batch records is > 22 then it will come on next page 
    public int numIndexPageSize = 20;//22; // Size of index page 

    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    String fsReportingName = "";

    Logs lg = new Logs();
    public StreamWriter sw = null;
    public string Filename = "";

    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName, AdvisorFlag, StatusIDBatch, BatchType, lsGAorTIAHeader;
    protected void Page_Load(object sender, EventArgs e)
    {
        Response.Cache.SetCacheability(HttpCacheability.NoCache); //check cahability
        //if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["bguid"]) != "")//opsbguid
        if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["bguid"]) != "" && Convert.ToString(Request.QueryString["bguid"]) != null)
        {
            BatchGUID = "'" + Convert.ToString(Request.QueryString["bguid"]) + "'"; //"'b95cb55d-62ed-dd11-be75-001d09665e8f'";//
            //Response.Write(BatchGUID);
        }

        if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["opsbguid"]) != "" && Convert.ToString(Request.QueryString["opsbguid"]) != null)
        {
            OPSBatchGUID = "'" + Convert.ToString(Request.QueryString["opsbguid"]) + "'"; //"'b95cb55d-62ed-dd11-be75-001d09665e8f'";//
            //Response.Write(BatchGUID);
        }

        if (!IsPostBack)
        {
            Session.Remove("CurPageInBatch");
            Session.Abandon();
            if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["hhid"]) != "")
                lstHouseHold.SelectedValue = Convert.ToString(Request.QueryString["hhid"]);

            BindDropDown();
            //ddlView.SelectedIndex = 1;
            //Response.AppendHeader("Refresh", "0;URL=../BatchReport/ReportReviewForm.aspx");



            if (Convert.ToString(Request.QueryString["hhid"]) != "" && Convert.ToString(Request.QueryString["hhid"]) != null)
            {
                ddlView.SelectedValue = "0";
                lstHouseHold.SelectedValue = Convert.ToString(Request.QueryString["hhid"]);
            }
            else if (ddlView.SelectedValue == "0" && BatchGUID == "" && OPSBatchGUID == "")
            {
                ddlView.SelectedValue = "1";
                //Response.Write(ddlView.SelectedValue + "<br/><br/>");
                ddlBatchstatus.SelectedValue = "10";
                if (ddlView.SelectedValue == "1")
                {
                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    Type tp = this.GetType();
                    sb.Append("\n<script type=text/javascript>\n");
                    sb.Append("__doPostBack('ddlView', 'SelectedIndexChanged');");
                    sb.Append("</script>");
                    ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
                }

            }
            else if (BatchGUID != "" || OPSBatchGUID != "")
            {
                ddlView.SelectedValue = "1";
            }
            else
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();
                sb.Append("\n<script type=text/javascript>\n");
                sb.Append("__doPostBack('ddlView', 'SelectedIndexChanged');");
                sb.Append("</script>");
                ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
            }

            BindGridView();

        }
        if (ddlAction.SelectedValue != "8")
        {
            tblBrowse.Style.Add("display", "none");
        }
        else if (ddlAction.SelectedValue == "8")
        {
            tblBrowse.Style.Add("display", "inline");
        }

    }


    // Function to get distincts in array
    public T[] GetDistinctValues<T>(T[] array)
    {
        List<T> tmp = new List<T>();
        for (int i = 0; i < array.Length; i++)
        {
            if (tmp.Contains(array[i]))
                continue;
            tmp.Add(array[i]);
        }
        return tmp.ToArray();
    }

    public void BindGridView()
    {
        String sql = BindGrid();
        DataSet loDataset = clsDB.getDataSet(sql);

        GridView1.Columns[10].Visible = true;
        GridView1.Columns[11].Visible = true;
        GridView1.Columns[12].Visible = true;
        GridView1.Columns[14].Visible = true;
        GridView1.Columns[15].Visible = true;
        GridView1.Columns[16].Visible = true;
        GridView1.Columns[17].Visible = true;
        GridView1.Columns[18].Visible = true;
        GridView1.Columns[19].Visible = true;
        GridView1.Columns[20].Visible = true;
        GridView1.Columns[21].Visible = true;
        GridView1.Columns[22].Visible = true;
        GridView1.Columns[23].Visible = true;
        GridView1.Columns[24].Visible = true;
        GridView1.Columns[25].Visible = true;
        GridView1.Columns[26].Visible = true;
        GridView1.Columns[27].Visible = true;

        GridView1.DataSource = loDataset;
        GridView1.DataBind();

        GridView1.Columns[10].Visible = false;
        GridView1.Columns[11].Visible = false;
        GridView1.Columns[12].Visible = false;
        GridView1.Columns[14].Visible = false;
        GridView1.Columns[15].Visible = false;
        GridView1.Columns[16].Visible = false;
        GridView1.Columns[17].Visible = false;
        GridView1.Columns[18].Visible = false;
        GridView1.Columns[19].Visible = false;
        GridView1.Columns[20].Visible = false;
        GridView1.Columns[21].Visible = false;
        GridView1.Columns[22].Visible = false;
        GridView1.Columns[23].Visible = false;
        GridView1.Columns[24].Visible = false;
        GridView1.Columns[25].Visible = false;
        GridView1.Columns[26].Visible = false;
        GridView1.Columns[27].Visible = false;

        if (GridView1.Rows.Count < 1)
        {
            lblError.Text = "Record not found";
            lblError.Visible = true;
            return;
        }
        else
        {
            lblError.Visible = false;
        }
    }

    private void BindDropDown()
    {
        BindHouseHold(lstHouseHold);
        BindAdvisor(ddlAdvisor);
        BindAssociate(ddlAssociate);
        BindBatchOwner(ddlBatchOwner);
        BindRecipient(ddlRecipient);
        BindMailStatus(ddlMailStatus);
        BindBatchStatus(ddlBatchstatus);
    }

    public void BindMailStatus(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "SP_S_MAIL_STATUS";
        clsGM.getListForBindDDL(ddl, sqlstr, "Status", "ID");



        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;

        //ddl.Items.Insert(1, "Report Not Sent");
        //ddl.Items[1].Value = "9999";
        //ddl.SelectedIndex = 1;
    }

    public void BindBatchStatus(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "SP_S_REPORT_TRACKER_STATUS";
        clsGM.getListForBindDDL(ddl, sqlstr, "NameTxt", "IdNmb");


        ddl.Items.RemoveAt(0);//Remove batch status 'Handed Off'
        ddl.Items.RemoveAt(0);// Remove batch status 'Approved' 


        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;



        System.Web.UI.WebControls.ListItem itm = ddl.Items.FindByText("Sent");
        ddl.Items.Remove(itm);
        ddl.Items.Insert(ddl.Items.Count, itm);

    }

    public void BindRecipient(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "SP_S_Recipient";
        clsGM.getListForBindDDL(ddl, sqlstr, "ContactName", "ssi_contactid");



        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    public void BindBatchOwner(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "SP_S_BatchOwner";
        clsGM.getListForBindDDL(ddl, sqlstr, "Owneridname", "Ownerid");

        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    public void BindAdvisor(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "SP_S_ADVISOR";
        clsGM.getListForBindDDL(ddl, sqlstr, "OwnerIdName", "OwnerId");

        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    public void BindAssociate(DropDownList ddl)
    {
        object OwnerId = ddlAdvisor.SelectedValue == "0" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        ddl.Items.Clear();

        sqlstr = "SP_S_ASSOCIATE @OwnerId=" + OwnerId;//;///SP_S_BATCH_ASSOCIATE//
        clsGM.getListForBindDDL(ddl, sqlstr, "Ssi_SecondaryOwnerIdName", "Ssi_SecondaryOwnerId");

        if (ddl.Items.Count == 1)
        {
            if (ddl.Items[0].Value == "0")
                ddl.Items.Remove(ddl.Items[0]);
        }
        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;


    }

    public void BindHouseHold(ListBox lstBox)
    {
        object AdvisorId = ddlAdvisor.SelectedValue == "0" || ddlAdvisor.SelectedValue == "" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        object BatchId = ddlBatchOwner.SelectedValue == "0" || ddlBatchOwner.SelectedValue == "" ? "null" : "'" + ddlBatchOwner.SelectedValue + "'";
        object AssociatedId = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        object RecipientId = ddlRecipient.SelectedValue == "0" || ddlRecipient.SelectedValue == "" ? "null" : "'" + ddlRecipient.SelectedValue + "'";
        object BatchType = ddlBatchtype.SelectedValue == "0" || ddlBatchtype.SelectedValue == "" ? "null" : ddlBatchtype.SelectedValue;

        lstBox.Items.Clear();

        sqlstr = "SP_S_HouseHoldName @IncludeClassB = 1,@AdvisorId=" + AdvisorId + ",@BatchId=" + BatchId + ",@AssociateId=" + AssociatedId + ",@RecipientId=" + RecipientId + ",@BatchType=" + BatchType;
        clsGM.getListForBindListBox(lstBox, sqlstr, "Name", "Accountid");

        if (lstBox.Items.Count == 1)
        {
            if (lstBox.Items[0].Value == "0")
                lstBox.Items.Remove(lstBox.Items[0]);
        }
        lstBox.Items.Insert(0, "All");
        lstBox.Items[0].Value = "0";
        lstBox.SelectedIndex = 0;

    }

    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        // string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;

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

        //Response.Write("p.Identity.Name:" + strName + "<br/><br/>");
        //strName = HttpContext.Current.User.Identity.Name.ToString();
        //Response.Write("HttpContext.Current.User.Identity.Name:" + strName + "<br/><br/>");
        //strName = Request.ServerVariables["AUTH_USER"]; //Finding with name
        //Response.Write("AUTH_USER:" + strName + "<br/><br/>");
        //////////
        //"select top 1 internalemailaddress,systemuserid from systemuser where domainname= 'Signature\\" + strName + "'";
        sqlstr = "select top 1 internalemailaddress,systemuserid from systemuser where domainname= '" + strName + "'";
        DB clsDB = new DB();
        DataSet lodataset = clsDB.getDataSet(sqlstr);
        //Response.Write(strName + "<br/><br/>");
        //Response.Write(Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]));
        if (lodataset.Tables[0].Rows.Count > 0)
        {
            return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);
        }
        else
        {
            return UserID = "";
        }
    }

    private string GetUserID(string DomainName)
    {
        string UserID = string.Empty;
        //"select top 1 internalemailaddress,systemuserid from systemuser where domainname= 'Signature\\" + strName + "'";
        sqlstr = "select top 1 internalemailaddress,systemuserid from systemuser where domainname= '" + DomainName + "'";
        DB clsDB = new DB();
        DataSet lodataset = clsDB.getDataSet(sqlstr);

        if (lodataset.Tables[0].Rows.Count > 0)
        {
            return UserID = Convert.ToString(lodataset.Tables[0].Rows[0]["systemuserid"]);
        }
        else
        {
            return UserID = "";
        }
    }

    public string BindGrid()
    {
        String lsSQL = "";

        string UserId = GetcurrentUser() == "" ? "null" : "'" + GetcurrentUser() + "'";

        object View = ddlView.SelectedValue == "0" ? "null" : ddlView.SelectedValue;
        object Advisor = ddlAdvisor.SelectedValue == "0" || ddlAdvisor.SelectedValue == "" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        object Associate = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        object HouseHold = lstHouseHold.SelectedValue == "0" || lstHouseHold.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstHouseHold) + "'";
        object BatchType = ddlBatchtype.SelectedValue == "0" || ddlBatchtype.SelectedValue == "" ? "null" : ddlBatchtype.SelectedValue;
        object BatchOwner = ddlBatchOwner.SelectedValue == "0" || ddlBatchOwner.SelectedValue == "" ? "null" : "'" + ddlBatchOwner.SelectedValue + "'";
        object BatchStatus = ddlBatchstatus.SelectedValue == "0" || ddlBatchstatus.SelectedValue == "" ? "null" : ddlBatchstatus.SelectedValue;
        object MailStatus = ddlMailStatus.SelectedValue == "0" || ddlMailStatus.SelectedValue == "" || ddlMailStatus.SelectedValue == "9999" ? "null" : ddlMailStatus.SelectedValue;
        object Recipient = ddlRecipient.SelectedValue == "0" || ddlRecipient.SelectedValue == "" ? "null" : "'" + ddlRecipient.SelectedValue + "'";
        //OPSBatchGUID

        if (BatchGUID != "" && (HouseHold == "" || HouseHold == "null"))
        {
            lsSQL = "SP_S_REPORT_REVIEW @UserId=" + BatchGUID + ",@ViewId=" + View + ",@AdvisorId=" + Advisor + ",@AssociateId=" + Associate
                    + ",@HouseHoldIdNmbList=" + HouseHold + ",@BatchTypeId=" + BatchType + ",@OwnerId=" + BatchOwner + ",@BatchStatusId=" + BatchStatus
                    + ",@MailStatusId=" + MailStatus + ",@ReceipentId=" + Recipient;
        }
        else if (OPSBatchGUID != "" && (HouseHold == "" || HouseHold == "null"))
        {
            lsSQL = "SP_S_REPORT_REVIEW @UserId=" + OPSBatchGUID + ",@ViewId=" + View + ",@AdvisorId=" + Advisor + ",@AssociateId=" + Associate
                               + ",@HouseHoldIdNmbList=" + HouseHold + ",@BatchTypeId=" + BatchType + ",@OwnerId=" + BatchOwner + ",@BatchStatusId=" + BatchStatus
                               + ",@MailStatusId=" + MailStatus + ",@ReceipentId=" + Recipient;
        }
        else
        {
            lsSQL = "SP_S_REPORT_REVIEW @UserId=" + UserId + ",@ViewId=" + View + ",@AdvisorId=" + Advisor + ",@AssociateId=" + Associate
                     + ",@HouseHoldIdNmbList=" + HouseHold + ",@BatchTypeId=" + BatchType + ",@OwnerId=" + BatchOwner + ",@BatchStatusId=" + BatchStatus
                     + ",@MailStatusId=" + MailStatus + ",@ReceipentId=" + Recipient;
        }


        return lsSQL;
        //Response.Write(lsSQL);

    }

    protected void ddlAdvisor_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();

        BindAssociate(ddlAssociate);
        BindHouseHold(lstHouseHold);
        BindGridView();
    }
    protected void ddlAssociate_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        BindHouseHold(lstHouseHold);

        BindGridView();
    }
    protected void ddlBatchtype_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        //BindHouseHold(lstHouseHold);

        BindGridView();
    }
    protected void ddlBatchOwner_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        //BindHouseHold(lstHouseHold);

        BindGridView();
    }
    protected void ddlRecipient_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        //BindHouseHold(lstHouseHold);

        BindGridView();
    }
    protected void ddlBatchstatus_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        //BindHouseHold(lstHouseHold);

        BindGridView();
    }
    protected void ddlMailStatus_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        //BindHouseHold(lstHouseHold);

        BindGridView();
    }
    // public static CrmService GetCrmService(string crmServerUrl, string organizationName)
    // {
    // // Get the CRM Users appointments
    // // Setup the Authentication Token
    // CrmAuthenticationToken token = new CrmAuthenticationToken();
    // token.AuthenticationType = 0; // Use Active Directory authentication.
    // token.OrganizationName = organizationName;
    // // string username = WindowsIdentity.GetCurrent().Name;

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

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
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


        LogFileName = (string)ViewState["Filename"];

        Session["Filename"] = LogFileName;

        string filenale = (string)Session["Filename"];

        lg.AddinLogFile(Session["Filename"].ToString(), "Start Page Load " + dtmain);

        lg.AddinLogFile(Session["Filename"].ToString(), "Option selected " + ddlAction.SelectedItem.ToString());





        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://crm01/";
        //string crmServerURL = "http://server:5555/";
        // string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        //CrmService service = null;

        IOrganizationService service = null;

        int finalReportCreatedCount = 0;
        int selectedCount = 0;
        string test = ddlBatchOwner.SelectedValue;
        int UniqueMailingId = 0;
        lblError.Visible = true;
        lblError.Text = "";
        lblMessage.Text = "";
        DataSet loInvoiceData = null;
        bool bOpsApproveRequestFlg = false;

        try
        {
            //service = GetCrmService(crmServerUrl, orgName);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblError.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
        }

        // service.PreAuthenticate = true;
        // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        //1	Associate Approved - Pend Advisor Approval
        //2	Handed Off
        //3	Approved
        //4	Sent
        //5	Approved - Pend OPS Approval
        //6	Pend Approval
        //7	OPS Change Requested
        //8	OPS Approved
        //9	FINAL Report Created

        //<asp:ListItem Value="1">Approve</asp:ListItem>
        //<asp:ListItem Value="2">Review PDF/Batch</asp:ListItem>
        //<asp:ListItem Value="3">Request OPS Change</asp:ListItem>
        //<asp:ListItem Value="4">Un-approve</asp:ListItem>

        if (ddlAction.SelectedValue == "9")
        {
            bool bcheckBatchType = false;
            foreach (GridViewRow row in GridView1.Rows)  // To allow or disallow action logic
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string BatchStatus = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                string Type = row.Cells[26].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");

                if (chkSelectNC.Checked)
                {
                    if (Type != "4")
                    {
                        bcheckBatchType = true;
                    }
                }
            }

            if (bcheckBatchType == true)
            {
                lblError.Text = "Only Merge batches can be rejected.  Please use the action “Unapprove” for Quarterly or Monthly reports";
                lblError.Visible = true;
            }
        }



        #region Check Batch Type

        if (ddlAction.SelectedValue == "2" || ddlAction.SelectedValue == "3" || ddlAction.SelectedValue == "4" || ddlAction.SelectedValue == "7")
        {
            bool bCheck = false;

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string BatchType = row.Cells[26].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");

                if (BatchType == "4")
                {
                    if (chkSelectNC.Checked == true)
                    {
                        bCheck = true;
                    }
                }

            }

            if (bCheck == true)
            {
                lblError.Text = "This action is only allowed for quarterly batches";
                lblError.Visible = true;
                return;
            }

        }


        #endregion


        #region Check For Dat MissMatch in CHIP and GRID
        if (ddlAction.SelectedValue == "1" || ddlAction.SelectedValue == "8")
        {
            bool bCheck = false;

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string BatchStatusID = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                string ssi_batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");

                if (chkSelectNC.Checked == true)
                {
                    sqlstr = BindGrid();
                    DataSet DS = clsDB.getDataSet(sqlstr);

                    for (int i = 0; i < DS.Tables[0].Rows.Count; i++)
                    {
                        string strBatchStatusId = Convert.ToString(DS.Tables[0].Rows[i]["BatchStatusID"]);
                        string strBatchId = Convert.ToString(DS.Tables[0].Rows[i]["ssi_batchid"]);

                        if (strBatchStatusId != BatchStatusID && strBatchId == ssi_batchid)
                        {
                            BindGridView();
                            bCheck = true;
                        }

                    }
                }
            }

            if (bCheck == true)
            {
                lblError.Text = "There was inconsistency between 'CHIP' and Data in grid below" + "<br/>" + " Data has been refreshed now please perform the action again";
                lblError.Visible = true;
                return;
            }

        }
        #endregion
        // added by sasmit 10/03/2016
        #region Insert CoverLetter
        if (ddlAction.SelectedValue == "10")
        {
            int count = 0;
            bool bproceed = true;
            bool bpathproceed = true;
            bool bBatchStatus = true;
            // ssi_batch objBatch = null;

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string ssi_batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                string BatchStatusID = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                string BatchFilePath = row.Cells[17].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string BatchFileName = row.Cells[18].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
                string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + BatchFileName);
                if (chkSelectNC.Checked == true)
                {
                    count++;
                    bproceed = false;
                }

                if (chkSelectNC.Checked == true)
                {
                    if (BatchFilePath == "")
                    {
                        bpathproceed = false;
                    }
                }

                if (chkSelectNC.Checked == true)
                {
                    if (BatchStatusID == "8")
                    {
                        bBatchStatus = false;
                    }
                }
            }

            if (count > 1)
            {
                if (!bproceed)
                {
                    lblError.Text = "Merge pdf only works when you select single batch. <br/> Please select single batch.";
                    lblError.Visible = true;
                    return;
                }
            }


            if (!bpathproceed)
            {
                lblError.Text = "The current batch doesnt have pdf report to merge yet.";
                lblError.Visible = true;
                return;
            }

            if (!bBatchStatus)
            {
                lblError.Text = "Can not merge pdf once batch status is 'OPS Approved'";
                lblError.Visible = true;
                return;
            }



            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string ssi_batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                string BatchStatusID = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                string BatchFilePath = row.Cells[17].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string BatchFileName = row.Cells[18].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
                string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + BatchFileName);
                string Foldername = row.Cells[14].Text.Trim().Replace("FolderNameTxt", "").Replace("&nbsp;", "");
                if (chkSelectNC.Checked == true)
                {
                    //added ifclause on 9_5_2017 MTGBK-NO INDEX (sasmit)
                    if (Foldername.Contains("MTGBK"))
                    {
                        lblError.Text = lblError.Text + "<br/>" + "Cannot Insert Coversheet in MTGBK Type";
                        lblError.Visible = true;
                        return;
                    }
                    else if (count == 1 && BatchStatusID != "8")
                    {
                        noIndex.Text = "coversheet";

                        DateTime dtime = DateTime.Now;
                        lg.AddinLogFile(Session["Filename"].ToString(), "," + "Report Genetare start ddlAction.SelectedValue == 10" + "," + dtime + ",");

                        GenerateReport();

                        lg.AddinLogFile(Session["Filename"].ToString(), "," + "Report Genetare start ddlAction.SelectedValue == 10" + "," + dtime + ",");


                        BindGridView();
                        lblError.Text = "Cover Letter Merged Successfully.";
                        lblError.Visible = true;
                        noIndex.Text = "";
                        return;
                    }

                    else
                    {
                        lblError.Text = lblError.Text + "<br/>" + "Please select pdf file to merge";
                        lblError.Visible = true;
                        return;
                    }


                }

            }

        }
        #endregion

        if (ddlAction.SelectedValue == "8")
        {
            int count = 0;
            bool bproceed = true;
            bool bpathproceed = true;
            bool bBatchStatus = true;
            // ssi_batch objBatch = null;

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string ssi_batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                string BatchStatusID = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                string BatchFilePath = row.Cells[17].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string BatchFileName = row.Cells[18].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
                string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + BatchFileName);
                if (chkSelectNC.Checked == true)
                {
                    count++;
                    bproceed = false;
                }

                if (chkSelectNC.Checked == true)
                {
                    if (BatchFilePath == "")
                    {
                        bpathproceed = false;
                    }
                }

                if (chkSelectNC.Checked == true)
                {
                    if (BatchStatusID == "8")
                    {
                        bBatchStatus = false;
                    }
                }
            }

            if (count > 1)
            {
                if (!bproceed)
                {
                    lblError.Text = "Merge pdf only works when you select single batch. <br/> Please select single batch.";
                    lblError.Visible = true;
                    return;
                }
            }


            if (!bpathproceed)
            {
                lblError.Text = "The current batch doesnt have pdf report to merge yet.";
                lblError.Visible = true;
                return;
            }

            if (!bBatchStatus)
            {
                lblError.Text = "Can not merge pdf once batch status is 'OPS Approved'";
                lblError.Visible = true;
                return;
            }



            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string ssi_batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                string BatchStatusID = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                string BatchFilePath = row.Cells[17].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string BatchFileName = row.Cells[18].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
                string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + BatchFileName);

                if (chkSelectNC.Checked == true)
                {
                    if (FileUpload1.HasFile == true)
                    {
                        if (count == 1 && BatchStatusID != "8")
                        {

                            string str = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + BatchFileName); //FileName
                            FileUpload1.PostedFile.SaveAs(Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + FileUpload1.FileName);
                            string filename = Path.GetFileName(FileUpload1.FileName);

                            string strClientPath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + filename;
                            // File.Copy(BatchFilePath, strDirectory, true);
                            string str2 = BatchFilePath;

                            string[] str3 = new string[2];
                            if (chkPrepend.Checked)
                            {
                                str3[0] = strClientPath;
                                str3[1] = str2;
                            }
                            else
                            {
                                str3[0] = str2;
                                str3[1] = strClientPath;
                            }
                            PDFMerge pdfMerge = new PDFMerge();
                            pdfMerge.MergeFiles(str, str3);

                            System.IO.File.Copy(str, BatchFilePath, true);

                            if (System.IO.File.Exists(strClientPath))
                            {
                                System.IO.File.Delete(strClientPath);
                            }

                            BindGridView();
                            lblError.Text = "Merge Pdf Successfully.";
                            lblError.Visible = true;
                            return;
                        }
                    }
                    else
                    {
                        lblError.Text = lblError.Text + "<br/>" + "Please select pdf file to merge";
                        lblError.Visible = true;
                        return;
                    }


                }

            }

        }


        if (ddlAction.SelectedValue == "7") // Billing Complete
        {
            bool bProceed = false;
            string BatchName = string.Empty;
            string DistinctHouseHoldNames = string.Empty;
            string[] CheckString;

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string BillingHandedOff = row.Cells[11].Text.Trim().Replace("Billing Handed Off", "").Replace("&nbsp;", "");
                string Batch = row.Cells[1].Text.Trim().Replace("Batch Name", "").Replace("&nbsp;", "");

                Batch = Batch.Replace(",", " ");


                if (chkSelectNC.Checked)
                {
                    if (BillingHandedOff.ToUpper() == "FALSE" || BillingHandedOff == "")
                    {
                        if (Batch != "")
                        {
                            BatchName = BatchName + "," + Batch;
                        }
                        else
                        {
                            BatchName = Batch;
                        }
                    }

                }
            }

            BatchName = BatchName == "" ? "" : BatchName.Substring(1, BatchName.Length - 1);
            CheckString = BatchName.Split(',');

            CheckString = GetDistinctValues<string>(CheckString);


            for (int i = 0; i < CheckString.Length; i++)
            {
                DistinctHouseHoldNames = DistinctHouseHoldNames + "," + CheckString[i];
            }

            DistinctHouseHoldNames = DistinctHouseHoldNames.Substring(1, DistinctHouseHoldNames.Length - 1);

            if (DistinctHouseHoldNames != "")
            {
                lblError.Text = DistinctHouseHoldNames + "<br/>" + " can not be marked completed for billing because the current billing copy hasn't sent";
                lblError.Visible = true;
            }
        }

        DateTime dtime2 = DateTime.Now;
        // Contains = "," + "Report Genetare start ddlAction.SelectedValue == 10" + "," + dtime + ",";


        #region Update Hold Report Status
        UpdateHoldReport();
        //foreach (GridViewRow row in GridView1.Rows)
        //{
        //    CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
        //    DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");
        //    string ssi_batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");

        //    try
        //    {
        //        if (chkSelectNC.Checked == true)
        //        {
        //            //UpdateHoldReport(ssi_batchid, ddlHoldReport.SelectedValue);
        //        }
        //    }
        //    catch (System.Web.Services.Protocols.SoapException exc)
        //    {
        //        lblMessage.Text = "Submit failed, Error detail: " + exc.Detail.InnerText;
        //    }
        //    catch (Exception exc)
        //    {
        //        lblMessage.Text = "Submit failed, Error detail: " + exc.Message;
        //    }

        //}
        #endregion


        foreach (GridViewRow row in GridView1.Rows)  // To allow or disallow action logic
        {
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

            if (chkSelectNC.Checked)
            {
                selectedCount++;

                string BatchStatus = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                string Type = row.Cells[26].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");
                //added 30_8_2019 - sasmit Batch process change
                string DestinationPath = row.Cells[17].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string ConsolidatePdfFileName = row.Cells[18].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
                //added 30_8_2019 - sasmit Batch process change
                if (ddlAction.SelectedValue == "4")//Un-Approved
                {
                    if (BatchStatus == "9" || BatchStatus == "4")
                    {
                        finalReportCreatedCount++;
                    }
                }

                if (ddlAction.SelectedValue == "1" || ddlAction.SelectedValue == "9") // Approve
                {
                    //OPS Approved, OPS Change Requested, FINAL Report Created, Sent,Approved - Pend OPS Approval
                    if (BatchStatus == "8" || BatchStatus == "7" || BatchStatus == "9" || BatchStatus == "4" || BatchStatus == "5")
                    {
                        // System.Text.StringBuilder  sb = new System.Text.StringBuilder();
                        // Type tp = this.GetType();
                        // sb.Append("\n<script type=text/javascript>\n");
                        // sb.Append("\nalert('One or more of the reports you selected was already approved by you.  No changes have been made to prior approved reports')");
                        //// sb.Append("var bt = window.document.getElementById('txt1');\n");
                        // sb.Append("\n</script>");
                        // ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                        //lblError.Text = "One or more of the reports you selected was already approved by you.  No changes have been made to prior approved reports";
                        if (ddlAction.SelectedValue == "9")
                        {
                            if (BatchStatus == "8" && Type == "4")
                            {
                                lblError.Text = "One or more of the reports you selected was already approved by you.  No changes have been made to prior approved reports";
                            }
                        }
                        else if (ddlAction.SelectedValue == "1")
                        {
                            lblError.Text = "One or more of the reports selected was already approved by you OR is pending an update from OPS. No changes have been made";
                        }

                        //return;Rohit
                    }
                    else
                    {
                        if (ConsolidatePdfFileName == "" && DestinationPath == "")
                        {
                            lblError.Text = "There is no report to approve. Please choose Review PDF to create the file, before approving the report";
                        }
                    }
                }


                if (ddlAction.SelectedValue == "3" || ddlAction.SelectedValue == "9")   // Request OPS Change
                {
                    //FINAL Report Created, Sent
                    if (BatchStatus == "9" || BatchStatus == "4")
                    {
                        if (ddlAction.SelectedValue == "9" && Type == "4")
                        {
                            lblMessage.Text = "One or more of the reports you selected may have already gone out.  Please consult with OPS directly to request your change";
                            //return;
                        }
                        else if (ddlAction.SelectedValue == "3")
                        {
                            lblMessage.Text = "One or more of the reports is has already been finalized or sent – consult with OPS directly to request your change";
                            return;
                        }

                    } //OPS Approved, OPS Change Requested
                    else if ((BatchStatus == "8" || BatchStatus == "7"))
                    {
                        bOpsApproveRequestFlg = true;

                    } //Pend Approval”, “Associate Approved – Pend Advisor Approval”, “Approved – PEND OPS Approval
                    else if (BatchStatus == "6" || BatchStatus == "1" || BatchStatus == "5")
                    {
                        // continue without message
                    }
                }






                ////FINAL Report Created, Sent
                //if (ddlAction.SelectedValue == "4")  // Un Approve
                //{
                //    if (BatchStatus == "9" || BatchStatus == "4")
                //    {
                //        lblError.Text = "One or more of the reports you selected may have already gone out.  Please consult with OPS directly to request your change";
                //    }
                //}

            }
        }


        if (ddlAction.SelectedValue == "3")// Request OPS Change
        {
            if (bOpsApproveRequestFlg && Hidden1.Value != "1")
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();
                sb.Append("\n<script type=text/javascript>\n");

                sb.Append("var bt = window.document.getElementById('btnSubmit');\n");
                sb.Append("if(confirm('One or more of the reports is already owned by OPS.  Do you want to proceed with requesting a change?'))\n{");
                sb.Append("\nwindow.document.getElementById('Hidden1').value='1';");
                sb.Append(("\n bt.click();\n"));
                sb.Append("\n}");
                sb.Append("else\n{");
                sb.Append(("\nwindow.document.getElementById('Hidden1').value='0';"));
                sb.Append("\n}");
                sb.Append("</script>");
                //ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                return;
            }
            else if (bOpsApproveRequestFlg && Hidden1.Value == "1")
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();
                sb.Append("\n<script type=text/javascript>\n");
                sb.Append("\nwindow.document.getElementById('Hidden1').value=''");
                sb.Append("</script>");
                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                //return;

            }
        }

        if (ddlAction.SelectedValue == "4" && finalReportCreatedCount > 0) ////Un-Approved
        {
            lblError.Visible = true;
            lblError.Text = "One or more of the reports you selected may have already gone out.  Please consult with OPS directly to request your change";
            //lblError.Text = "One or more of the reports you selected is already in the final processing – contact OPS immediately if you would like to try to pull the report";
            //return;
        }

        //  ssi_mailrecords objMailRecords = null;
        foreach (GridViewRow row in GridView1.Rows)
        {

            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
            DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");

            string ssi_batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
            string BillingHandedOff = row.Cells[11].Text.Trim().Replace("Billing Handed Off", "").Replace("&nbsp;", "");
            AdvisorFlag = row.Cells[12].Text.Trim().Replace("AdvisorFlag", "").Replace("&nbsp;", "");

            string DestinationPath = row.Cells[17].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
            string ConsolidatePdfFileName = row.Cells[18].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
            //commented by dharamendra shifted to checkbox selected area below dated--03/may/2013
            //StatusIDBatch = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
            string AdvisorId = row.Cells[20].Text.Trim().Replace("OwnerId", "").Replace("&nbsp;", "");
            string ssi_secondaryownerid = row.Cells[21].Text.Trim().Replace("ssi_secondaryownerid", "").Replace("&nbsp;", "");
            string BatchOwnerId = row.Cells[23].Text.Trim().Replace("BatchOwnerId", "").Replace("&nbsp;", "");
            string MailStatusId = row.Cells[24].Text.Trim().Replace("MailStatusId", "").Replace("&nbsp;", "");

            string MailRecordsId = row.Cells[25].Text.Trim().Replace("ssi_mailrecordsid", "").Replace("&nbsp;", "");
            //BatchType = row.Cells[26].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");



            //2 - Handed Off
            //3 - Approved
            //4 - Sent
            //6 - Pend Approval
            //1 - Associate Approved - Pend Advisor Approval
            //5 - Approved - Pend OPS Approval
            //7 - OPS Change Requested
            //8 - OPS Approved
            //9 - FINAL Report Created

            //<asp:ListItem Value="1">Approve</asp:ListItem>
            //<asp:ListItem Value="2">Review PDF/Batch</asp:ListItem>
            //<asp:ListItem Value="3">Request OPS Change</asp:ListItem>
            //<asp:ListItem Value="4">Un-approve</asp:ListItem>

            try
            {
                if (chkSelectNC.Checked == true)
                {
                    BatchType = row.Cells[26].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");
                    StatusIDBatch = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");

                    ViewState["ReviewPdf"] = DestinationPath; // This is for action review pdf/batch (2 is for review pdf/batch)
                    ViewState["ReviewPdfName"] = ConsolidatePdfFileName;
                    if (ddlAction.SelectedValue == "1")//Approve
                    {

                        //string ConsolidatePdfFileName = Convert.ToString(ViewState["ReviewPdfName"]);
                        //string DestinationPath = Convert.ToString(ViewState["ReviewPdf"]);
                        if (ConsolidatePdfFileName != "" && DestinationPath != "")
                        {
                           
                       
                            string AsOfDate = string.Empty;
                            if (StatusIDBatch != "8" && StatusIDBatch != "9" && StatusIDBatch != "4" && StatusIDBatch != "5" && (AdvisorFlag.ToUpper() == "FALSE" || AdvisorFlag == "" || StatusIDBatch == "6"))//5 : Approved - Pend OPS Approval 
                            {

                                #region Insert Batch Review Details
                                sqlstr = "SP_S_REPORT_REVIEW_DETAIL @BatchId='" + ssi_batchid + "'";
                                DataSet loDataset = clsDB.getDataSet(sqlstr);

                                for (int j = 0; j < loDataset.Tables[0].Rows.Count; j++)
                                {
                                    //objMailRecords = new ssi_mailrecords();
                                    Entity objMailRecords = new Entity("ssi_mailrecords");

                                    //name
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["name"]) != "")
                                    {
                                        //objMailRecords.ssi_name = Convert.ToString(loDataset.Tables[0].Rows[j]["name"]);
                                        objMailRecords["ssi_name"] = Convert.ToString(loDataset.Tables[0].Rows[j]["name"]);
                                    }

                                    //Quarterly Statement
                                    // objMailRecords.ssi_mailtypeid = new Lookup();
                                    // objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                                    // objMailRecords.ssi_mailtypeid.Value = new Guid("eb776a64-cdbe-e011-a19b-0019b9e7ee05");

                                    objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid("eb776a64-cdbe-e011-a19b-0019b9e7ee05"));

                                    if (ssi_batchid != "")
                                    {
                                        // objMailRecords.ssi_batchid = new Lookup();
                                        // objMailRecords.ssi_batchid.type = EntityName.ssi_batch.ToString();
                                        // objMailRecords.ssi_batchid.Value = new Guid(ssi_batchid);

                                        objMailRecords["ssi_batchid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_batch", new Guid(Convert.ToString(ssi_batchid)));
                                    }

                                    //Mail Type
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_mailid"]) != "")
                                    {
                                        // objMailRecords.ssi_mailtypeid = new Lookup();
                                        // objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                                        // objMailRecords.ssi_mailtypeid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_mailid"]));

                                        objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_mailid"])));
                                    }


                                    //First Name
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["FirstName"]) != "")
                                    {
                                        //objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["FirstName"]);
                                        objMailRecords["ssi_ownerfname_cnt_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["FirstName"]);
                                    }

                                    //Last Name
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["LastName"]) != "")
                                    {
                                        //objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["LastName"]);
                                        objMailRecords["ssi_ownerlname_cnt_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["LastName"]);
                                    }

                                    //House Hold
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["HouseHold"]) != "")
                                    {
                                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["HouseHold"]);
                                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["HouseHold"]);
                                    }

                                    //Contact
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["Contact"]) != "")
                                    {
                                        //objMailRecords.ssi_fullname_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["Contact"]);
                                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["Contact"]);
                                    }

                                    //Address Line 1
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["AddressLine1"]) != "")
                                    {
                                        //objMailRecords.ssi_addressline1_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["AddressLine1"]);
                                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["AddressLine1"]);
                                    }

                                    //Address Line 2
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["AddressLine2"]) != "")
                                    {
                                        //objMailRecords.ssi_addressline2_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["AddressLine2"]);
                                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["AddressLine2"]);
                                    }

                                    //Address Line 3
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["AddressLine3"]) != "")
                                    {
                                        //objMailRecords.ssi_addressline3_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["AddressLine3"]);
                                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["AddressLine3"]);
                                    }

                                    //City
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["City"]) != "")
                                    {
                                        //objMailRecords.ssi_city_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["City"]);
                                        objMailRecords["ssi_city_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["City"]);
                                    }

                                    //State Or Province
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["State Or Province"]) != "")
                                    {
                                        //objMailRecords.ssi_stateprovince_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["State Or Province"]);
                                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["State Or Province"]);
                                    }


                                    //Zip Code
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["Zip Code"]) != "")
                                    {
                                        //objMailRecords.ssi_zipcode_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["Zip Code"]);
                                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["Zip Code"]);
                                    }

                                    //Country Or Region
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["Country Or Region"]) != "")
                                    {
                                        //objMailRecords.ssi_countryregion_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["Country Or Region"]);
                                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["Country Or Region"]);
                                    }

                                    //Dear
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["Dear"]) != "")
                                    {
                                        //objMailRecords.ssi_dear_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["Dear"]);
                                        objMailRecords["ssi_dear_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["Dear"]);
                                    }

                                    //Mail Preference
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["Mail Preference"]) != "")
                                    {
                                        //objMailRecords.ssi_mailpreference_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["Mail Preference"]);
                                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["Mail Preference"]);
                                    }

                                    //Salutation
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["Salutation"]) != "")
                                    {
                                        //objMailRecords.ssi_salutation_mail = Convert.ToString(loDataset.Tables[0].Rows[j]["Salutation"]);
                                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(loDataset.Tables[0].Rows[j]["Salutation"]);
                                    }

                                    //Ssi_MailingID
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["Ssi_MailingID"]) != "")
                                    {
                                        if (UniqueMailingId == 0)
                                            objMailRecords["ssi_mailingid"] = Convert.ToInt32(loDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                                        //UniqueMailingId = Convert.ToInt32(loDataset.Tables[0].Rows[j]["Ssi_MailingID"]);

                                        //objMailRecords.ssi_mailingid = new CrmNumber();
                                        //objMailRecords.ssi_mailingid.Value = UniqueMailingId; //Convert.ToInt32(loDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                                    }

                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_mailstatus"]) != "")
                                    {
                                        //objMailRecords.ssi_mailstatus = new Picklist();
                                        //objMailRecords.ssi_mailstatus.Value = Convert.ToInt32(loDataset.Tables[0].Rows[j]["ssi_mailstatus"]);

                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[0].Rows[j]["ssi_mailstatus"]));
                                    }

                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["AsOfDate"]) != "")
                                    {
                                        // objMailRecords.ssi_asofdate = new CrmDateTime();
                                        // objMailRecords.ssi_asofdate.Value = Convert.ToString(loDataset.Tables[0].Rows[j]["AsOfDate"]);
                                        // AsOfDate = Convert.ToString(loDataset.Tables[0].Rows[j]["AsOfDate"]);
                                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(loDataset.Tables[0].Rows[j]["AsOfDate"]);
                                    }

                                    //objMailRecords.ssi_batchidtxt = Convert.ToString(ssi_batchid); // BatchId Text 
                                    //objMailRecords.ssi_batchnametxt = row.Cells[1].Text.Trim(); // Batch Name

                                    objMailRecords["ssi_batchidtxt"] = Convert.ToString(ssi_batchid);
                                    objMailRecords["ssi_batchnametxt"] = row.Cells[1].Text.Trim();


                                    //HouseHold lookup
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["AccountId"]) != "")
                                    {
                                        // objMailRecords.ssi_accountid = new Lookup();
                                        // objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["AccountId"]));
                                        objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_account", new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["AccountId"])));
                                    }

                                    //Contact lookup
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["ContactId"]) != "")
                                    {
                                        // objMailRecords.ssi_contactfullnameid = new Lookup();
                                        // objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ContactId"]));
                                        objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ContactId"])));
                                    }

                                    ////ssi_LegalEntityId lookup
                                    if (Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]) != "")
                                    {
                                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                                        // objMailRecords.ssi_legalentitynameid = new Lookup();
                                        // objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]));
                                        objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_LegalEntityId"])));
                                    }


                                    // CreatedByCustomid Field 
                                    //Rohit Pawar opsbguid
                                    string UserId = string.Empty;
                                    if (BatchGUID != "")
                                    {
                                        UserId = Convert.ToString(Request.QueryString["bguid"]);
                                    }
                                    else if (OPSBatchGUID != "")
                                    {
                                        UserId = Convert.ToString(Request.QueryString["opsbguid"]);
                                    }
                                    else
                                    {
                                        UserId = GetcurrentUser();
                                    }

                                    if (UserId != "")
                                    {
                                        // objMailRecords.ssi_createdbycustomid = new Lookup();
                                        // objMailRecords.ssi_createdbycustomid.Value = new Guid(UserId);
                                        objMailRecords["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));
                                    }

                                    service.Create(objMailRecords);

                                    intResult++;
                                }
                                #endregion

                                // if Action is Review PDF/Batch or Request OPS Change

                                if (intResult > 0)
                                {

                                    //if (ddlAction.SelectedValue == "1")//Action DropDown Value 1= Approved
                                    //{
                                    //BindGridView();
                                    //lblError.Text = "Mail records inserted successfully with Report generation";
                                    //}
                                }

                                //if (BatchType != "4")
                                // GenerateReport();
                            }


                            #region DisAssociate Mail Records
                            sqlstr = "SP_S_CANCEL_MAIL_STATUS @MailStatusId='6,4' ,@BatchId='" + ssi_batchid + "'";
                            DataSet MailStatusDataset = clsDB.getDataSet(sqlstr);

                            for (int j = 0; j < MailStatusDataset.Tables[0].Rows.Count; j++)
                            {
                                // objMailRecords = new ssi_mailrecords();
                                Entity objMailRecords = new Entity("ssi_mailrecords");

                                // objMailRecords.ssi_mailrecordsid = new Key();
                                // objMailRecords.ssi_mailrecordsid.Value = new Guid(Convert.ToString(MailStatusDataset.Tables[0].Rows[j]["ssi_mailrecordsid"]));

                                objMailRecords["ssi_mailrecordsid"] = new Guid(Convert.ToString(MailStatusDataset.Tables[0].Rows[j]["ssi_mailrecordsid"]));

                                if (ssi_batchid != "")
                                {
                                    // objMailRecords.ssi_batchid = new Lookup();
                                    // objMailRecords.ssi_batchid.IsNull = true;
                                    // objMailRecords.ssi_batchid.IsNullSpecified = true;

                                    objMailRecords["ssi_batchid"] = null;

                                }

                                service.Update(objMailRecords);

                                intResult++;
                            }
                            #endregion
                            //else //if (BatchStatusID != "5")  // 5 -Approved - Pend OPS Approval
                            //{
                            // 7 - OPS Change Requested

                            BatchUpdate(ssi_batchid, AdvisorFlag, ddlHoldReport.SelectedValue, 7, BatchOwnerId, StatusIDBatch, AsOfDate, BillingHandedOff, AdvisorId);  // 7 - OPS Change Requested (Batch Status)

                            //}

                            //BindGridView();
                        }
                    }
                    else if (ddlAction.SelectedValue == "4")//Un-Approved
                    {
                        try
                        {
                            if (StatusIDBatch != "9" && StatusIDBatch != "4")
                            {
                                #region Batch Update
                                // ssi_batch objBatch = new ssi_batch();
                                Entity objBatch = new Entity("ssi_batch");

                                // objBatch.ssi_batchid = new Key();
                                // objBatch.ssi_batchid.Value = new Guid(ssi_batchid);

                                objBatch["ssi_batchid"] = new Guid(Convert.ToString(ssi_batchid));

                                // objBatch.ssi_reporttrackerstatus = new Picklist();
                                // objBatch.ssi_reporttrackerstatus.Value = 6;  //6 - Pend Approval

                                objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(6);

                                // objBatch.ssi_holdreport = new Picklist();
                                // objBatch.ssi_holdreport.IsNull = true;
                                // objBatch.ssi_holdreport.IsNullSpecified = true;

                                objBatch["ssi_holdreport"] = null;

                                // objBatch.ssi_billinghandedoff = new CrmBoolean();
                                // objBatch.ssi_billinghandedoff.Value = false;

                                objBatch["ssi_billinghandedoff"] = false;

                                // Added By Rohit Pawar 
                                // objBatch.ssi_billingcomplete = new CrmBoolean();
                                // objBatch.ssi_billingcomplete.Value = false;

                                objBatch["ssi_billingcomplete"] = false;

                                if (StatusIDBatch == "8" || StatusIDBatch == "1" || StatusIDBatch == "9" || StatusIDBatch == "4" || StatusIDBatch == "5")
                                {
                                    // objBatch.ssi_billingcontactchange = new CrmBoolean();
                                    // objBatch.ssi_billingcontactchange.Value = true;//To check email notification when status is changed to 'Pend Approval'

                                    objBatch["ssi_billingcontactchange"] = true;
                                }
                                else if (StatusIDBatch == "7")//OPS Change Requested
                                {
                                    // objBatch.ssi_opschangecomplete = new CrmBoolean();
                                    // objBatch.ssi_opschangecomplete.Value = true;//Send Email When Status is changed from 'OPS Change Requested ' to 'Pend Approval'

                                    objBatch["ssi_opschangecomplete"] = true;

                                    //below code commented on demand of jeane masa
                                    //objBatch.ssi_opsbatchguid = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl) + ":9999/BatchReport/ReportReviewForm.aspx?opsbguid=" + ssi_secondaryownerid;
                                    //objBatch.ssi_opsbatchguid = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":9999/BatchReport/ReportReviewForm.aspx";
                                    //  objBatch["ssi_opsbatchguid"] = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":9999/BatchReport/ReportReviewForm.aspx"; // commented 1_11_2019
                                    objBatch["ssi_opsbatchguid"] = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":" + AppLogic.GetParam(AppLogic.ConfigParam.CRMPortNumber).ToString() + "/BatchReport/ReportReviewForm.aspx"; // added for CRMPORT NUMBER Change 1_11_2019
                                }
                                //objBatch.ssi_batchdate = new CrmDateTime();
                                //objBatch.ssi_batchdate.Value = DateTime.Now.ToString();

                                if (ConsolidatePdfFileName != "")
                                {
                                    //objBatch.ssi_batchdisplayfilename = "";
                                    objBatch["ssi_batchdisplayfilename"] = "";
                                }

                                if (DestinationPath != "")
                                {
                                    //objBatch.ssi_batchfilename = "";
                                    objBatch["ssi_batchfilename"] = "";
                                }

                                //SecurityPrincipal assignee = new SecurityPrincipal();
                                //assignee.PrincipalId = new Guid(ssi_secondaryownerid);///HouseHold Owner ID

                                //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();
                                //targetAssign.EntityId = new Guid(ssi_batchid);
                                //targetAssign.EntityName = EntityName.ssi_batch.ToString();

                                //AssignRequest assign = new AssignRequest();
                                //assign.Assignee = assignee;
                                //assign.Target = targetAssign;

                                //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);

                                //service.Update(objBatch);




                                AssignRequest assignRequest = new AssignRequest
                                {
                                    Assignee = new EntityReference("systemuser",
                                        new Guid(ssi_secondaryownerid)),
                                    Target = new EntityReference("ssi_batch",
                                        new Guid(ssi_batchid))
                                };
                                service.Execute(assignRequest);
                                service.Update(objBatch);




                                string UserId = string.Empty;
                                if (BatchGUID != "")
                                {
                                    UserId = Convert.ToString(Request.QueryString["bguid"]);
                                }
                                else if (OPSBatchGUID != "")
                                {
                                    UserId = Convert.ToString(Request.QueryString["opsbguid"]);
                                }
                                else
                                {
                                    UserId = GetcurrentUser();// Get Current User Logged in ID
                                }

                                BatchReportStatus(6, UserId, ssi_batchid, BillingHandedOff);//Batch Report Status Log

                                #endregion

                                #region Mail Records
                                sqlstr = "SP_S_BATCH_MAILRECORDS @BatchId='" + ssi_batchid + "'";

                                DataSet NewDataset = clsDB.getDataSet(sqlstr);

                                for (int i = 0; i < NewDataset.Tables[0].Rows.Count; i++)
                                {
                                    //objMailRecords = new ssi_mailrecords();
                                    Entity objMailRecords = new Entity("ssi_mailrecords");

                                    if (Convert.ToString(NewDataset.Tables[0].Rows[i]["ssi_mailrecordsid"]) != "")
                                    {
                                        // objMailRecords.ssi_mailrecordsid = new Key();
                                        // objMailRecords.ssi_mailrecordsid.Value = new Guid(Convert.ToString(NewDataset.Tables[0].Rows[i]["ssi_mailrecordsid"]));

                                        objMailRecords["ssi_mailrecordsid"] = new Guid(Convert.ToString(NewDataset.Tables[0].Rows[i]["ssi_mailrecordsid"]));
                                    }

                                    //if (Convert.ToString(NewDataset.Tables[0].Rows[i]["ssi_mailstatus"]) != "")
                                    //{
                                    // objMailRecords.ssi_mailstatus = new Picklist();
                                    // objMailRecords.ssi_mailstatus.Value = 6;//Convert.ToInt32(NewDataset.Tables[0].Rows[i]["ssi_mailstatus"]);

                                    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(6));
                                    //}

                                    service.Update(objMailRecords);

                                    intResult++;
                                }
                                #endregion
                            }
                        }
                        //catch (System.Web.Services.Protocols.SoapException exc)
                        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                        {
                            bProceed = false;
                            strDescription = "Error occured, Error Detail1: " + exc.Detail.Message;
                            lblError.Text = strDescription;
                        }
                        catch (Exception exc)
                        {
                            bProceed = false;
                            strDescription = "Error occured, Error Detail1: " + exc.Message;
                            lblError.Text = strDescription;
                        }


                        //BindGridView();
                    }
                    else if (ddlAction.SelectedValue == "3")//Request OPS Change
                    {
                        BatchUpdate(ssi_batchid, AdvisorFlag, ddlHoldReport.SelectedValue, 7, BatchOwnerId, StatusIDBatch, "", BillingHandedOff, AdvisorId);

                        #region Mail Records Update
                        sqlstr = "SP_S_BATCH_MAILRECORDS @BatchId='" + ssi_batchid + "'";

                        DataSet NewDataset = clsDB.getDataSet(sqlstr);

                        for (int i = 0; i < NewDataset.Tables[0].Rows.Count; i++)
                        {
                            // objMailRecords = new ssi_mailrecords();
                            Entity objMailRecords = new Entity("ssi_mailrecords");

                            if (Convert.ToString(NewDataset.Tables[0].Rows[i]["ssi_mailrecordsid"]) != "")
                            {
                                // objMailRecords.ssi_mailrecordsid = new Key();
                                // objMailRecords.ssi_mailrecordsid.Value = new Guid(Convert.ToString(NewDataset.Tables[0].Rows[i]["ssi_mailrecordsid"]));

                                objMailRecords["ssi_mailrecordsid"] = new Guid(Convert.ToString(NewDataset.Tables[0].Rows[i]["ssi_mailrecordsid"]));
                            }

                            //if (Convert.ToString(NewDataset.Tables[0].Rows[i]["ssi_mailstatus"]) != "")
                            //{
                            // objMailRecords.ssi_mailstatus = new Picklist();
                            // objMailRecords.ssi_mailstatus.Value = 6;//Mail Status: Cancel

                            objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(6));
                            //}

                            service.Update(objMailRecords);
                            intResult++;
                        }
                        #endregion
                    }
                    else if (ddlAction.SelectedValue == "5")//Remove Hold
                    {
                        // ssi_batch objBatch = new ssi_batch();
                        Entity objBatch = new Entity("ssi_batch");

                        // objBatch.ssi_batchid = new Key();
                        // objBatch.ssi_batchid.Value = new Guid(ssi_batchid);

                        objBatch["ssi_batchid"] = new Guid(Convert.ToString(ssi_batchid));

                        // objBatch.ssi_holdreport = new Picklist();
                        // objBatch.ssi_holdreport.IsNull = true;
                        // objBatch.ssi_holdreport.IsNullSpecified = true;
                        objBatch["ssi_holdreport"] = null;


                        service.Update(objBatch);
                    }
                    else if (ddlAction.SelectedValue == "7")//Billing Complete
                    {
                        // ssi_batch objBatch = null;

                        if (BillingHandedOff.ToUpper() == "TRUE")
                        {
                            // objBatch = new ssi_batch();
                            Entity objBatch = new Entity("ssi_batch");

                            // objBatch.ssi_batchid = new Key();
                            // objBatch.ssi_batchid.Value = new Guid(ssi_batchid);
                            objBatch["ssi_batchid"] = new Guid(Convert.ToString(ssi_batchid));

                            // objBatch.ssi_billingcomplete = new CrmBoolean();
                            // objBatch.ssi_billingcomplete.Value = true;

                            objBatch["ssi_billingcomplete"] = true;

                            // objBatch.ssi_billingcompletemail = new CrmBoolean();
                            // objBatch.ssi_billingcompletemail.Value = true;

                            objBatch["ssi_billingcompletemail"] = true;

                            service.Update(objBatch);
                            intResult++;
                        }
                    }
                    else if (ddlAction.SelectedValue == "9")
                    {
                        // objMailRecords = new ssi_mailrecords();
                        Entity objMailRecords = new Entity("ssi_mailrecords");

                        if (BatchType == "4" && (StatusIDBatch != "8" && StatusIDBatch != "9" && StatusIDBatch != "4"))
                        {
                            if (MailRecordsId != "")
                            {
                                // objMailRecords.ssi_mailrecordsid = new Key();
                                // objMailRecords.ssi_mailrecordsid.Value = new Guid(MailRecordsId);
                                objMailRecords["ssi_mailrecordsid"] = new Guid(Convert.ToString(MailRecordsId));

                                // objMailRecords.ssi_review_reject = new CrmBoolean();
                                // objMailRecords.ssi_review_reject.Value = true;

                                objMailRecords["ssi_review_reject"] = true;

                                // objMailRecords.ssi_deleterecord_flg = new CrmBoolean();
                                // objMailRecords.ssi_deleterecord_flg.Value = true;

                                objMailRecords["ssi_deleterecord_flg"] = true;


                                // objMailRecords.ssi_initialreviewer_reject = new CrmBoolean();
                                // objMailRecords.ssi_initialreviewer_reject.Value = false;

                                objMailRecords["ssi_initialreviewer_reject"] = false;


                                string UserId = GetcurrentUser();

                                if (UserId != "")
                                {
                                    //Response.Write(UserId);
                                    // objMailRecords.ssi_rejectedbyuserid = new Lookup();
                                    // objMailRecords.ssi_rejectedbyuserid.Value = new Guid(UserId);

                                    objMailRecords["ssi_rejectedbyuserid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));
                                }


                                service.Update(objMailRecords);
                                selectedCount++;
                            }
                        }
                    }
                }
            }
            //catch (System.Web.Services.Protocols.SoapException exc)
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
            {
                bProceed = false;
                strDescription = "Error occured, Error Detail2: " + exc.Detail.Message;
                lblError.Text = strDescription;
            }
            catch (Exception exc)
            {
                bProceed = false;
                strDescription = "Error occured, Error Detail2: " + exc.Message;
                lblError.Text = strDescription;
            }
        }

        //if (ddlAction.SelectedValue != "4" && ddlAction.SelectedValue != "1")// 

        if (ddlAction.SelectedValue == "1")
        {
            //Commented 30_8_2019 - sasmit Batch process change
            //if (BatchType != "4" && StatusIDBatch == "6")
            //{

            //    DateTime dtime3 = DateTime.Now;
            //    lg.AddinLogFile(Session["Filename"].ToString(), "," + "Report Genetare start 1" + "," + dtime3 + ",");
            //    GenerateReport();
            //    lg.AddinLogFile(Session["Filename"].ToString(), "," + "Report Genetare End 1" + "," + dtime3 + "," + DateTime.Now);
            //}

            BindGridView();
            lblError.Visible = true;

            if (lblError.Text == "")
                lblError.Text = "Records Approved Successfully.";
        }
        else if (ddlAction.SelectedValue == "2" || ddlAction.SelectedValue == "3")//  Review PDF/Batch and Request OPS Change
        {
            //if Action is Review PDF/Batch or Request OPS Change
            if (ddlAction.SelectedValue == "2")
            {
                try
                {
                    if (selectedCount > 0)
                    {
                        //loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));ViewState["ReviewPdfName"]
                        string ConsolidatePdfFileName = Convert.ToString(ViewState["ReviewPdfName"]);
                        string DestinationPath = Convert.ToString(ViewState["ReviewPdf"]);
                        if (ConsolidatePdfFileName != "" && DestinationPath != "")
                        {
                            //Commented 30_8_2019 - sasmit Batch process change
                            //string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ConsolidatePdfFileName);// (Server.MapPath("") + @"\ExcelTemplate\" + ConsolidatePdfFileName);
                            //System.IO.File.Copy(strDirectory, DestinationPath, true);
                            ////File.Delete(DestinationPath);
                            ////Directory.Delete(ReportOpFolder, true);


                            //// Response.Write("<script>");
                            //string lsFileNamforFinal = ConsolidatePdfFileName;
                            ////Response.Write("window.open('" + lsFileNamforFinal + "', 'mywindow')");
                            //// Response.Write("window.open('ViewReport.aspx?" + ConsolidatePdfFileName + "', 'mywindow')");

                            ////  Response.Write("</script>");

                            //System.Text.StringBuilder sb = new System.Text.StringBuilder();
                            //Type tp = this.GetType();
                            //sb.Append("\n<script type=text/javascript>\n");
                            //sb.Append("\nwindow.open('ViewReport.aspx?" + ConsolidatePdfFileName + "', 'mywindow');");
                            //sb.Append("</script>");
                            //ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());

                           

                        }
                        else
                        {

                            DateTime dtime4 = DateTime.Now;
                            lg.AddinLogFile(Session["Filename"].ToString(), "," + "Report Genetare start 2" + "," + dtime4 + ",");
                            //  sw.WriteLine();
                            GenerateReport();
                            lg.AddinLogFile(Session["Filename"].ToString(), "," + "Report Genetare End 2" + "," + dtime4 + "," + DateTime.Now);

                        }

                    }

                }
                catch (Exception exc)
                {
                    Response.Write(exc.Message);
                }
            }
            else if (ddlAction.SelectedValue == "3")
            {
                DateTime dtime5 = DateTime.Now;
                //    sw.WriteLine();
                lg.AddinLogFile(Session["Filename"].ToString(), "," + "Report Genetare start 3" + "," + dtime5 + ",");
                GenerateReport();
                lg.AddinLogFile(Session["Filename"].ToString(), "," + "Report Genetare End 3" + "," + dtime5 + "," + DateTime.Now);
            }

            BindGridView();

            lblError.Visible = true;

            if (selectedCount > 1 && ddlAction.SelectedValue == "3") //Request OPS Change<
            {

                lblError.Text = "You have requested more than one change from OPS, please see your reports in <a href=file:///S:/BATCH%20REPORTS>S:/BATCH REPORTS</a>";

                lblError.Text = lblError.Text + "<br/>" + strReportFiles;
            }
            else if (selectedCount > 0)
                lblError.Text = "Report generated successfully ";


            string ConsolidatePdfFileName1 = Convert.ToString(ViewState["ReviewPdfName"]);
            string DestinationPath1 = Convert.ToString(ViewState["ReviewPdf"]);
            if (ConsolidatePdfFileName1 != "" && DestinationPath1 != "")
            {
                //added 30_8_2019 - sasmit Batch process change
                lblError.Text = "Report is already generated. Please un-approve and then Review PDF to generate a new report";
                lg.AddinLogFile(Session["Filename"].ToString(), "," + "Report Genetare End 2" + "," + DateTime.Now);
            }
        }
        else if (ddlAction.SelectedValue == "4")  //Un-approve
        {
            BindGridView();
            lblError.Visible = true;

            if (lblError.Text == "")
                lblError.Text = "Records Un-Approved Successfully.";
        }
        else if (ddlAction.SelectedValue == "5")  //Remove Hold
        {
            BindGridView();
            lblError.Visible = true;

            if (lblError.Text == "")
                lblError.Text = "Hold Report Updated Successfully.";
        }
        else if (ddlAction.SelectedValue == "7")
        {
            BindGridView();
            lblError.Visible = true;

            if (lblError.Text == "")
            {
                lblError.Text = "Billing Completely Successfully";
            }
            else if (intResult > 0)
            {
                lblError.Text = lblError.Text + "<br/><br/>" + "Billing Completely Successfully";
            }
        }
        else if (ddlAction.SelectedValue == "9")
        {
            if (selectedCount > 0 || BatchType == "4")
            {
                BindGridView();
                System.Threading.Thread.Sleep(20000);
                DeleteBatchAndMailRecords(service);
                if (lblError.Text == "")
                {
                    lblError.Text = lblError.Text;// "Batch and Mail Records Rejected Successfully.";
                    lblError.Visible = true;
                }
                else
                {
                    lblError.Text = lblError.Text;// "Batch and Mail Records Rejected Successfully.";
                    lblError.Visible = true;
                }

            }
            else
            {
                lblError.Text = lblError.Text;
            }
        }

        lg.AddinLogFile(Session["Filename"].ToString(), "," + "Report End" + "," + dtmain + "," + DateTime.Now);
    }
    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            string BillingHandedOff = e.Row.Cells[11].Text.Trim().Replace("Billing Handed Off", "").Replace("&nbsp;", "");
            string FileName = e.Row.Cells[17].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
            string HoldReportId = e.Row.Cells[22].Text.Trim().Replace("ssi_holdreport", "").Replace("&nbsp;", "");

            DropDownList ddlHoldReport = (DropDownList)e.Row.FindControl("ddlHoldReport");
            CheckBox chkBillingHandedOff = (CheckBox)e.Row.FindControl("chkBillingHandedOff");
            ImageButton imgApprovedFile = (ImageButton)e.Row.FindControl("imgApprovedFile");

            sqlstr = "SP_S_HOLD_REPORT";//Store Procedure to bind hold report dropdown
            clsGM.getListForBindDDL(ddlHoldReport, sqlstr, "Status", "ID");
            ddlHoldReport.Items.Insert(0, "");
            ddlHoldReport.Items[0].Value = "0";
            ddlHoldReport.SelectedIndex = 0;


            ddlHoldReport.SelectedValue = HoldReportId;


            if (BillingHandedOff.ToUpper() == "TRUE")
            {
                chkBillingHandedOff.Checked = true;
                chkBillingHandedOff.Enabled = false;
            }
            else
            {
                chkBillingHandedOff.Enabled = false;
            }

            if (FileName == "")
            {
                imgApprovedFile.Visible = false;
            }
            else if (FileName != "")
            {
                imgApprovedFile.Visible = true;
            }
        }
    }

    public void BatchUpdate(string BatchId, string AdvisorFlag, string HoldReport, int ReportTrackerStatus, string BatchOwnerId, string BatchStatusID, string AsOfDate, string BillingHandedOff, string AdvisorId)
    {

        //Response.Write("<br/>BatchUpdate function called");
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://crm01/";
        //string crmServerURL = "http://server:5555/";
        // string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        //CrmService service = null;
        IOrganizationService service = null;

        int NewBatchStatusId = 0;

        //lblError.Text = "";
        DataSet loInvoiceData = null;
        try
        {
            //service = GetCrmService(crmServerUrl, orgName);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblError.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
        }

        // service.PreAuthenticate = true;
        // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {
            // ssi_batch objBatch = new ssi_batch();
            Entity objBatch = new Entity("ssi_batch");

            string UserId = string.Empty;
            if (BatchGUID != "")
            {
                UserId = Convert.ToString(Request.QueryString["bguid"]);
            }
            else if (OPSBatchGUID != "")
            {
                UserId = Convert.ToString(Request.QueryString["opsbguid"]);
            }
            else
            {
                UserId = GetcurrentUser();// Get Current User Logged in ID
            }

            //Response.Write("<br/>UserId:" + UserId);

            // objBatch.ssi_batchid = new Key();
            // objBatch.ssi_batchid.Value = new Guid(BatchId);
            objBatch["ssi_batchid"] = new Guid(Convert.ToString(BatchId));

            if (ddlAction.SelectedValue == "3") // Ops Request Change
            {
                // objBatch.ssi_reporttrackerstatus = new Picklist();
                // objBatch.ssi_reporttrackerstatus.Value = 7;// OPS Change Requested

                objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(7));

                // objBatch.ssi_batchdisplayfilename = "";
                // objBatch.ssi_batchfilename = "";

                objBatch["ssi_batchdisplayfilename"] = "";
                objBatch["ssi_batchfilename"] = "";


                BatchReportStatus(7, UserId, BatchId, BillingHandedOff);//Batch Log Status

                //SecurityPrincipal assignee = new SecurityPrincipal();
                //assignee.PrincipalId = new Guid(GetUserID("CORP\\opsreporting"));////OPS Reporting Gresham

                //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();
                //targetAssign.EntityId = new Guid(BatchId);
                //targetAssign.EntityName = EntityName.ssi_batch.ToString();

                //AssignRequest assign = new AssignRequest();
                //assign.Assignee = assignee;
                //assign.Target = targetAssign;

                //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);




                AssignRequest assignRequest = new AssignRequest
                {
                    Assignee = new EntityReference("systemuser",
                                        new Guid("8A2AD331-109E-E411-A01F-0002A5443D86")),
                    Target = new EntityReference("ssi_batch",
                                        new Guid(BatchId))
                };
                lg.AddinLogFile(Session["Filename"].ToString(), "assignRequest1 =" + DateTime.Now);
                service.Execute(assignRequest);
                lg.AddinLogFile(Session["Filename"].ToString(), "assignRequest1 =" + DateTime.Now);





                if (HoldReport != "" && HoldReport != "0")
                {
                    // objBatch.ssi_holdreport = new Picklist();
                    // objBatch.ssi_holdreport.Value = Convert.ToInt32(HoldReport);

                    objBatch["ssi_holdreport"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(HoldReport));
                }
                else if (HoldReport == "" || HoldReport == "0")
                {
                    // objBatch.ssi_holdreport = new Picklist();
                    // objBatch.ssi_holdreport.IsNull = true;
                    // objBatch.ssi_holdreport.IsNullSpecified = true;

                    objBatch["ssi_holdreport"] = null;
                }

                if (BatchStatusID == "8" || BatchStatusID == "1" || BatchStatusID == "9" || BatchStatusID == "4" || BatchStatusID == "5")
                {
                    // objBatch.ssi_billingcontactchange = new CrmBoolean();
                    // objBatch.ssi_billingcontactchange.Value = true;

                    objBatch["ssi_billingcontactchange"] = true;
                }
                else if (BatchStatusID == "7")//OPS Change Requested
                {
                    // objBatch.ssi_opschangecomplete = new CrmBoolean();
                    // objBatch.ssi_opschangecomplete.Value = true;//Send Email When Status is changed from 'OPS Change Requested ' to 'Pend Approval'

                    objBatch["ssi_opschangecomplete"] = true;
                }


                // When status changed to OPS changed Requested
                // objBatch.ssi_opschangemail = new CrmBoolean();
                // objBatch.ssi_opschangemail.Value = true;

                objBatch["ssi_opschangemail"] = true;

                //objBatch.ssi_opsreporttracker = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":9999/BatchReport/ReportTrackerNew.aspx";
                //objBatch["ssi_opsreporttracker"] = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":9999/BatchReport/ReportTrackerNew.aspx"; // commented 1_11_2019
                objBatch["ssi_opsreporttracker"] = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":" + AppLogic.GetParam(AppLogic.ConfigParam.CRMPortNumber).ToString() + "/BatchReport/ReportTrackerNew.aspx";// added for CRMPORT NUMBER Change 1_11_2019
            }
            else if (ddlAction.SelectedValue == "1")
            {
                //Response.Write("<br/>ddlAction.SelectedValue == 1");
                //Response.Write("<br/>BatchStatusID:" + BatchStatusID);

                //below condition to change only report tracker status 

                //2 - Handed Off
                //3 - Approved
                //4 - Sent
                //6 - Pend Approval
                //1 - Associate Approved - Pend Advisor Approval
                //5 - Approved - Pend OPS Approval
                //7 - OPS Change Requested
                //8 - OPS Approved
                //9 - FINAL Report Created

                //OPS Approved, OPS Change Requested, FINAL Report Created, Sent,Associate Approved - Pend Advisor Approval, Approved - Pend OPS Approval
                if (BatchStatusID == "8" || BatchStatusID == "7" || BatchStatusID == "9" || BatchStatusID == "4" || BatchStatusID == "5") //|| BatchStatusID == "1" || BatchStatusID == "5")
                {
                    // no status change if current status is from above list
                }
                else if (BatchStatusID == "1" && AdvisorFlag.ToUpper() == "TRUE")  //Associate Approved - Pend Advisor Approval
                {
                    NewBatchStatusId = 5;  //5 - Approved - Pend OPS Approval
                    // objBatch.ssi_reporttrackerstatus = new Picklist();
                    // objBatch.ssi_reporttrackerstatus.Value = NewBatchStatusId;

                    objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(NewBatchStatusId));

                    if (HoldReport != "" && HoldReport != "0")
                    {
                        // objBatch.ssi_holdreport = new Picklist();
                        // objBatch.ssi_holdreport.Value = Convert.ToInt32(HoldReport);

                        objBatch["ssi_holdreport"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(HoldReport));

                    }
                    else if (HoldReport == "" || HoldReport == "0")
                    {
                        // objBatch.ssi_holdreport = new Picklist();
                        // objBatch.ssi_holdreport.IsNull = true;
                        // objBatch.ssi_holdreport.IsNullSpecified = true;

                        objBatch["ssi_holdreport"] = null;
                    }

                    //SecurityPrincipal assignee = new SecurityPrincipal();
                    //assignee.PrincipalId = new Guid(AdvisorId);////Advisor

                    //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();
                    //targetAssign.EntityId = new Guid(BatchId);
                    //targetAssign.EntityName = EntityName.ssi_batch.ToString();

                    //AssignRequest assign = new AssignRequest();
                    //assign.Assignee = assignee;
                    //assign.Target = targetAssign;

                    //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);


                    AssignRequest assignRequest = new AssignRequest
                    {
                        Assignee = new EntityReference("systemuser",
                        new Guid(AdvisorId)),
                        Target = new EntityReference("ssi_batch",
                        new Guid(BatchId))
                    };

                    lg.AddinLogFile(Session["Filename"].ToString(), "assignRequest2 =" + DateTime.Now);
                    service.Execute(assignRequest);
                    lg.AddinLogFile(Session["Filename"].ToString(), "assignRequest2 =" + DateTime.Now);




                }       //5 - Approved - Pend OPS Approval
                //else if (BatchStatusID == "5")// && (AdvisorFlag.ToUpper() == "FALSE" || AdvisorFlag.ToUpper() == ""))
                //{
                //    NewBatchStatusId = 8;  //8 - OPS Approved
                //    objBatch.ssi_reporttrackerstatus = new Picklist();
                //    objBatch.ssi_reporttrackerstatus.Value = NewBatchStatusId;

                //    if (HoldReport != "" && HoldReport != "0")
                //    {
                //        objBatch.ssi_holdreport = new Picklist();
                //        objBatch.ssi_holdreport.Value = Convert.ToInt32(HoldReport);
                //    }
                //    else if (HoldReport == "" || HoldReport == "0")
                //    {
                //        objBatch.ssi_holdreport = new Picklist();
                //        objBatch.ssi_holdreport.IsNull = true;
                //        objBatch.ssi_holdreport.IsNullSpecified = true;
                //    }
                //}
                else if (BatchStatusID == "6")  //Pend Approval
                {

                    //Response.Write("<br/>BatchStatusID == 6");
                    if (AdvisorFlag.ToUpper() == "TRUE")
                    {
                        NewBatchStatusId = 1; //1 - Associate Approved - Pend Advisor Approval
                        // objBatch.ssi_reporttrackerstatus = new Picklist();
                        // objBatch.ssi_reporttrackerstatus.Value = NewBatchStatusId;  //1 - Associate Approved - Pend Advisor Approval

                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(NewBatchStatusId));

                        if (HoldReport != "" && HoldReport != "0")
                        {
                            // objBatch.ssi_holdreport = new Picklist();
                            // objBatch.ssi_holdreport.Value = Convert.ToInt32(HoldReport);

                            objBatch["ssi_holdreport"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(HoldReport));
                        }
                        else if (HoldReport == "" || HoldReport == "0")
                        {
                            // objBatch.ssi_holdreport = new Picklist();
                            // objBatch.ssi_holdreport.IsNull = true;
                            // objBatch.ssi_holdreport.IsNullSpecified = true;

                            objBatch["ssi_holdreport"] = null;
                        }

                        // objBatch.ssi_billingcontactchange = new CrmBoolean();
                        // objBatch.ssi_billingcontactchange.Value = true;//Send Email to Advisor

                        objBatch["ssi_billingcontactchange"] = true;

                        if (BatchId != "")
                        {
                            //below code commented on demand of jeane masa
                            //objBatch.ssi_batchguid = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl) + ":9999/BatchReport/ReportReviewForm.aspx?bguid=" + AdvisorId;
                            //objBatch.ssi_batchguid = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":9999/BatchReport/ReportReviewForm.aspx";
                            //objBatch["ssi_batchguid"] = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":9999/BatchReport/ReportReviewForm.aspx";// commented 1_11_2019
                            objBatch["ssi_batchguid"] = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":" + AppLogic.GetParam(AppLogic.ConfigParam.CRMPortNumber).ToString() + "/BatchReport/ReportReviewForm.aspx";// added for CRMPORT NUMBER Change 1_11_2019
                        }

                        //SecurityPrincipal assignee = new SecurityPrincipal();
                        //assignee.PrincipalId = new Guid(AdvisorId);////Advisor

                        //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();
                        //targetAssign.EntityId = new Guid(BatchId);
                        //targetAssign.EntityName = EntityName.ssi_batch.ToString();

                        //AssignRequest assign = new AssignRequest();
                        //assign.Assignee = assignee;
                        //assign.Target = targetAssign;

                        //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);


                        AssignRequest assignRequest = new AssignRequest
                        {
                            Assignee = new EntityReference("systemuser",
                        new Guid(AdvisorId)),
                            Target = new EntityReference("ssi_batch",
                        new Guid(BatchId))
                        };
                        lg.AddinLogFile(Session["Filename"].ToString(), "assignRequest3 =" + DateTime.Now);
                        service.Execute(assignRequest);
                        lg.AddinLogFile(Session["Filename"].ToString(), "assignRequest3 =" + DateTime.Now);



                    }
                    else if (AdvisorFlag.ToUpper() == "FALSE" || AdvisorFlag.ToUpper() == "")
                    {

                        //Response.Write("<br/>AdvisorFlag.ToUpper() == FALSE || AdvisorFlag.ToUpper() ==");
                        NewBatchStatusId = 5; //5 - Approved - Pend OPS Approval
                        // objBatch.ssi_reporttrackerstatus = new Picklist();
                        // objBatch.ssi_reporttrackerstatus.Value = NewBatchStatusId;  //5 - Approved - Pend OPS Approval

                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(NewBatchStatusId));

                        //SecurityPrincipal assignee = new SecurityPrincipal();
                        //assignee.PrincipalId = new Guid(GetUserID("CORP\\opsreporting"));////OPS Reporting Gresham

                        //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();
                        //targetAssign.EntityId = new Guid(BatchId);
                        //targetAssign.EntityName = EntityName.ssi_batch.ToString();

                        //AssignRequest assign = new AssignRequest();
                        //assign.Assignee = assignee;
                        //assign.Target = targetAssign;

                        //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);


                        AssignRequest assignRequest = new AssignRequest
                        {
                            Assignee = new EntityReference("systemuser",
                            new Guid("8A2AD331-109E-E411-A01F-0002A5443D86")),
                            Target = new EntityReference("ssi_batch",
                            new Guid(BatchId))
                        };
                        lg.AddinLogFile(Session["Filename"].ToString(), "assignRequest4 =" + DateTime.Now);
                        service.Execute(assignRequest);
                        lg.AddinLogFile(Session["Filename"].ToString(), "assignRequest4 =" + DateTime.Now);




                    }
                    else
                    {
                        // NewBatchStatusId = 5;  //5 - Approved - Pend OPS Approval
                        // objBatch.ssi_reporttrackerstatus = new Picklist();
                        // objBatch.ssi_reporttrackerstatus.Value = NewBatchStatusId;

                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(NewBatchStatusId));
                    }

                    if (HoldReport != "" && HoldReport != "0")
                    {
                        // objBatch.ssi_holdreport = new Picklist();
                        // objBatch.ssi_holdreport.Value = Convert.ToInt32(HoldReport);

                        objBatch["ssi_holdreport"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(HoldReport));
                    }
                    else if (HoldReport == "" || HoldReport == "0")
                    {
                        // objBatch.ssi_holdreport = new Picklist();
                        // objBatch.ssi_holdreport.IsNull = true;
                        // objBatch.ssi_holdreport.IsNullSpecified = true;

                        objBatch["ssi_holdreport"] = null;
                    }
                }

                BatchReportStatus(NewBatchStatusId, UserId, BatchId, BillingHandedOff);//Batch Log Status

            }

            if (BatchStatusID == "1")   // 1- Associate Approved - Pend Advisor Approval
            {
                //AdvisorFlag.ToUpper() == "TRUE" &&

                //objBatch.ssi_reporttrackerstatus = new Picklist();
                //objBatch.ssi_reporttrackerstatus.Value = 5;  // 5 - Approved - Pend OPS Approval

                //SecurityPrincipal assignee = new SecurityPrincipal();
                //assignee.PrincipalId = new Guid(GetUserID("CORP\\opsreporting"));////OPS Reporting Gresham

                ////assignee.PrincipalId = new Guid(AdvisorId);

                //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();
                //targetAssign.EntityId = new Guid(BatchId);
                //targetAssign.EntityName = EntityName.ssi_batch.ToString();

                //AssignRequest assign = new AssignRequest();
                //assign.Assignee = assignee;
                //assign.Target = targetAssign;

                //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);

                #region UNCOMMENT
                AssignRequest assignRequest = new AssignRequest
                {
                    Assignee = new EntityReference("systemuser",
                                new Guid("8A2AD331-109E-E411-A01F-0002A5443D86")),
                    Target = new EntityReference("ssi_batch",
                                new Guid(BatchId))
                };

                service.Execute(assignRequest);

                #endregion



                // BatchReportStatus(NewBatchStatusId, UserId, BatchId);//Batch Log Status
            }


            if (AsOfDate != "")
            {
                // objBatch.ssi_batchdate = new CrmDateTime();
                // objBatch.ssi_batchdate.Value = AsOfDate;// DateTime.Now.ToString();

                objBatch["ssi_batchdate"] = Convert.ToDateTime(AsOfDate);
            }

            service.Update(objBatch);

        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            lblError.Visible = true;
            bProceed = false;
            lg.AddinLogFile(Session["Filename"].ToString(), "Error" + exc.Detail.Message + DateTime.Now);
            strDescription = "Error occured, Error Detail3: " + exc.Detail.Message;
            lblError.Text = strDescription;
        }
        catch (Exception exc)
        {
            lblError.Visible = true;
            bProceed = false;
            lg.AddinLogFile(Session["Filename"].ToString(), "Error" + exc.Message + DateTime.Now);
            strDescription = "Error occured, Error Detail3: " + exc.Message;
            lblError.Text = strDescription;
        }
    }

    public void BatchReportStatus(int Status, string UpdatedBy, string BatchId, string BillingHandedOff)
    {
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://crm01/";
        //string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        // CrmService service = null;
        IOrganizationService service = null;


        //lblError.Text = "";
        DataSet loInvoiceData = null;
        try
        {
            //service = GetCrmService(crmServerUrl, orgName);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblError.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
        }

        // service.PreAuthenticate = true;
        // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {
            // ssi_batchreportstatuslog objReportStatusLog = new ssi_batchreportstatuslog();
            Entity objReportStatusLog = new Entity("ssi_batchreportstatuslog");

            // objReportStatusLog.ssi_statusdate = new CrmDateTime();
            // objReportStatusLog.ssi_statusdate.Value = DateTime.Now.ToString();

            objReportStatusLog["ssi_statusdate"] = DateTime.Now;

            if (Status != 0)
            {
                // objReportStatusLog.ssi_status = new Picklist();
                // objReportStatusLog.ssi_status.Value = Status;

                objReportStatusLog["ssi_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(Status));
            }

            if (UpdatedBy != "")
            {
                // objReportStatusLog.ssi_updatedbyid = new Lookup();
                // objReportStatusLog.ssi_updatedbyid.Value = new Guid(UpdatedBy);

                objReportStatusLog["ssi_updatedbyid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Convert.ToString(UpdatedBy)));
            }

            if (BatchId != "")
            {
                // objReportStatusLog.ssi_batchid = new Lookup();
                // objReportStatusLog.ssi_batchid.Value = new Guid(BatchId);

                objReportStatusLog["ssi_batchid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_batch", new Guid(Convert.ToString(BatchId)));
            }

            if (BillingHandedOff.ToUpper() == "TRUE")
            {
                // objReportStatusLog.ssi_billinghandedoff = new CrmBoolean();
                // objReportStatusLog.ssi_billinghandedoff.Value = true;

                objReportStatusLog["ssi_billinghandedoff"] = true;
            }
            else
            {
                // objReportStatusLog.ssi_billinghandedoff = new CrmBoolean();
                // objReportStatusLog.ssi_billinghandedoff.Value = false;

                objReportStatusLog["ssi_billinghandedoff"] = false;
            }

            service.Create(objReportStatusLog);

        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblError.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
        }
    }

    private void GenerateReport()
    {
        string ContactFolderName = string.Empty;
        string ReportOpFolder = string.Empty;
        string ParentFolder = string.Empty;
        string TempFolderPath = string.Empty;
        string Local_ParentFolderPath = string.Empty;
        lg.AddinLogFile(Session["Filename"].ToString(), "Start Report =" + DateTime.Now);

        try
        {
            //lblError.Text = "";
            Session.Remove("CurPageInBatch");
            //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://crm01/";
            clsCombinedReports objCombinedReports = new clsCombinedReports();
            //string crmServerURL = "http://server:5555/";

            //string orgName = "GreshamPartners";
            string currentuser = null;
            //string orgName = "Webdev";
            //CrmService service = null;
            IOrganizationService service = null;

            Boolean checkrunreport = false;
            String DestinationPath = string.Empty;
            string ConsolidatePdfFileName = string.Empty;

            string ApprovedReports = AppLogic.GetParam(AppLogic.ConfigParam.ApprovedReports);// "\\\\fs01\\opsreports$\\Approved Reports\\"; //"\\\\Fs01\\shared$\\OPS REPORTS\\Approved Reports\\";
            try
            {
                //service = GetCrmService(crmServerUrl, orgName);
                service = clsGM.GetCrmService();
                strDescription = "Crm Service starts successfully";
            }
            //catch (System.Web.Services.Protocols.SoapException exc)
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
                lblError.Text = strDescription;
            }
            catch (Exception exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                lblError.Text = strDescription;
            }

            // service.PreAuthenticate = true;
            // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

            DataTable dtBatch = null;

            string[] distColName = { "Ssi_ContactIdName" };
            //DataTable distinctTable = dtBatch.DefaultView.ToTable(true, distColName);
            //  Response.Write("<br>table count:" + dtBatch.Rows.Count);
            //  Response.Write("distict from table:" + distinctTable.Rows.Count);

            DateTime dt = DateTime.Now;

            string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
            string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
            string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();
            string strMilliSecond = DateTime.Now.Millisecond.ToString().Length < 2 ? "0" + DateTime.Now.Millisecond.ToString() : DateTime.Now.Millisecond.ToString();
            //string CurrentDateTime = DateTime.Now.ToShortDateString() + " " + " " + strHour + "-" + strMinute + "-" + strSecond;

            string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

            // string strUserName = HttpContext.Current.User.Identity.Name.ToString();
            string strUserName = string.Empty;
            //Changed Windows to - ADFS Claims Login 8_9_2019
            if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
            {
                strUserName = "corp\\gbhagia";
            }
            else
            {
                IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
                strUserName = claimsIdentity.Name;

            }

            strUserName = strUserName.Substring(strUserName.IndexOf("\\") + 1);

            //UserName_YYYYMMDD_Timewhere 

            //ViewState["ParentFolder"] = CurrentDateTime.Replace(":", "-").Replace("/", "-"); // orig

            ParentFolder = strUserName + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond + strMilliSecond;

            //string ReportOpFolder = "\\\\Fs01\\_ops_C_I_R_group\\Quarterly_Reports\\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();


            //ReportOpFolder = Request.MapPath("ExcelTemplate\\BATCH REPORTS\\") + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            //if (Request.Url.AbsoluteUri.Contains("localhost"))
            //{
            //    ReportOpFolder = Request.MapPath("ExcelTemplate\\BATCH REPORTS\\") + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            //}
            //else
            //    ReportOpFolder = Request.MapPath("ExcelTemplate\\BATCH REPORTS\\") + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports);// "\\\\fs01\\opsreports$";//"\\\\Fs01\\shared$\\OPS REPORTS\\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            if (ddlAction.SelectedValue == "2" || ddlAction.SelectedValue == "3")
                ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.BatchReports);//"\\\\Fs01\\shared$\\BATCH REPORTS\\";

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                ReportOpFolder = @"C:\Reports\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            }
            else
            {
                ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports);// "\\\\fs01\\opsreports$";//"\\\\Fs01\\shared$\\OPS REPORTS\\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

                if (ddlAction.SelectedValue == "2" || ddlAction.SelectedValue == "3")
                    ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.BatchReports);//"\\\\Fs01\\shared$\\BATCH REPORTS\\";
            }


            FileInfo loCoversheetCheck;
            String ReportOpFolder1 = String.Empty;

            /*****Start :  Array declaration for PDF merge **************/
            PDFMerge PDF = new PDFMerge();
            int sourcefilecount = 0;//= dtBatch.Rows.Count + 1;
            int sourcefilecount1 = 0; // added 10/03/2016 - sasmit
            string[] SourceFileArray = null;
            string[] SourceFileArray1 = null; // added 10/03/2016 - sasmit
            /*****End   :  Array declaration for PDF merge **************/

            //ConsolidatePdfFileName = "ConsolidatedPDF" + "_" + strYear + strMonth + strDay + "_" + ".pdf";
            int NoOfBatches = 0;
            for (int j = 0; j < GridView1.Rows.Count; j++)
            {
                CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");

                if (chkBox.Checked)
                {
                    NoOfBatches++;
                }
            }

            for (int j = 0; j < GridView1.Rows.Count; j++)
            {
                CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");

                DateTime FileDateTime = DateTime.Now;

                string FileYear = FileDateTime.Year.ToString().Length < 2 ? "0" + FileDateTime.Year.ToString() : FileDateTime.Year.ToString();
                string FileMonth = FileDateTime.Month.ToString().Length < 2 ? "0" + FileDateTime.Month.ToString() : FileDateTime.Month.ToString();
                string FileDay = FileDateTime.Day.ToString().Length < 2 ? "0" + FileDateTime.Day.ToString() : FileDateTime.Day.ToString();

                string FileHour = FileDateTime.Hour.ToString().Length < 2 ? "0" + FileDateTime.Hour.ToString() : FileDateTime.Hour.ToString();
                string FileMinute = FileDateTime.Minute.ToString().Length < 2 ? "0" + FileDateTime.Minute.ToString() : FileDateTime.Minute.ToString();
                string FileSec = FileDateTime.Second.ToString().Length < 2 ? "0" + FileDateTime.Second.ToString() : FileDateTime.Second.ToString();
                string FileMiliSec = FileDateTime.Millisecond.ToString().Length < 2 ? "0" + FileDateTime.Millisecond.ToString() : FileDateTime.Millisecond.ToString();

                string CurrentTimeStamp = FileYear + "_" + FileMonth + "_" + FileDay + "_" + FileHour + "_" + FileMinute + "_" + FileSec + "_" + FileMiliSec;

                if (chkBox.Checked)
                {
                    numIndexPageCount = 1;  //Index page count -- if count of batch records is > 22 then it will come on next page 
                    numIndexPageSize = 20;//22; // Size of index page 

                    checkrunreport = true;
                    String BatchIdListTxt = Convert.ToString(GridView1.Rows[j].Cells[10].Text);

                    lg.AddinLogFile(Session["Filename"].ToString(), "Get DATATABLE" + DateTime.Now);
                    dtBatch = GetDataTable(BatchIdListTxt);
                    lg.AddinLogFile(Session["Filename"].ToString(), "DATATABLE End" + DateTime.Now);

                    //String TempName =  GridView1.Rows[j].Cells[6].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();

                    //String HHName = GridView1.Rows[j].Cells[6].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                    //string ssi_batchid = GridView1.Rows[j].Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                    String HHName = "";// GridView1.Rows[j].Cells[16].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                    string SPVFileName = string.Empty;
                    string OldHHName = GridView1.Rows[j].Cells[16].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString().Replace("&#39;", "'").ToString();

                    //string TempName = HttpContext.Current.User.Identity.Name.ToString() + "_" + 
                    double total = (double)dtBatch.Rows.Count / numIndexPageSize;
                    int liTotalPage = Convert.ToInt32(Math.Ceiling(total));
                    numIndexPageCount = numIndexPageCount + liTotalPage;

                    sourcefilecount = dtBatch.Rows.Count + (numIndexPageCount + 1);

                    sourcefilecount1 = dtBatch.Rows.Count + (numIndexPageCount + 2); // added 10/03/2016 - sasmit

                    SourceFileArray = new string[sourcefilecount];

                    SourceFileArray1 = new string[sourcefilecount1]; // added 10/03/2016 - sasmit
                    //added 30-05-2018(Sasmit- random number in folders created for each batch)
                    Random rnd = new Random();
                    string strRndNumber = Convert.ToString(rnd.Next(99999));
                    //added 30-05-2018(Sasmit- random number in folders created for each batch)
                    for (int i = 0; i < dtBatch.Rows.Count; i++)
                    {
                        if (Convert.ToString(dtBatch.Rows[i]["ssi_spvfilename"]) != "")
                        {
                            SPVFileName = Convert.ToString(dtBatch.Rows[i]["ssi_spvfilename"]);
                            HHName = Convert.ToString(dtBatch.Rows[i]["ssi_spvfilename"]);
                            HHName = HHName.Replace("/", "");
                        }
                        else
                        {
                            HHName = GridView1.Rows[j].Cells[16].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString().Replace("&#39;", "'").ToString();
                            HHName = HHName.Replace("/", "");
                        }

                        //sw.WriteLine("HHName Name " + HHName + "Start =" + DateTime.Now); 
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + ",," + DateTime.Now + ",");

                        ContactFolderName = GridView1.Rows[j].Cells[14].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString().Replace("&#39;", "'").ToString();
                        //ContactFolderName = Convert.ToString(dtBatch.Rows[i]["Ssi_ContactIdName"]).Replace(",", "");

                        //added 30-05-2018(Sasmit- random number in folders created for each batch)
                        ContactFolderName = ContactFolderName + "_" + strRndNumber;

                        //added 30-05-2018(Sasmit- random number in folders created for each batch)
                        Local_ParentFolderPath = Request.MapPath("ExcelTemplate\\TempFolder") + "\\" + ParentFolder;
                        //tempFolder at Local Path to create image,pdf files
                        TempFolderPath = Local_ParentFolderPath + "\\" + ContactFolderName;
                        if (!Directory.Exists(TempFolderPath))
                        {
                            System.IO.Directory.CreateDirectory(TempFolderPath);
                        }
                        bool isExist = System.IO.Directory.Exists(ReportOpFolder + "\\" + ParentFolder + "\\" + ContactFolderName);
                        if (!isExist)
                        {
                            System.IO.Directory.CreateDirectory(ReportOpFolder + "\\" + ParentFolder + "\\" + ContactFolderName);
                        }

                        ViewState["AsOfDate"] = Convert.ToString(dtBatch.Rows[i]["Ssi_EndAsOfDate2"]);
                        // ViewState["PdfFileName"] = HHName = Convert.ToString(dtBatch.Rows[i]["PdfFileName"]);

                        String fsAllocationGroup = Convert.ToString(dtBatch.Rows[i]["Ssi_AllocationGroup"]).Replace("'", "''");
                        String fsHouseholdName = Convert.ToString(dtBatch.Rows[i]["Ssi_HouseholdIdName"]).Replace("'", "''");
                        String fsAsofDate = Convert.ToString(dtBatch.Rows[i]["Ssi_EndAsOfDate2"]);
                        String fsSPriorDate = Convert.ToString(dtBatch.Rows[i]["Ssi_StartPriorDate1"]);
                        String fsLookthrogh = Convert.ToString(dtBatch.Rows[i]["Ssi_ConsolidateDetailLevel"]);
                        String fsContactFullname = Convert.ToString(dtBatch.Rows[i]["Ssi_ContactIdName"]);
                        String fsVersion = Convert.ToString(dtBatch.Rows[i]["Ssi_UnderlyingManagerDetail"]);

                        //overrid value of Underlying Manager Detail if Suppress manager detail is checked
                        //if (chkSuppressManagerDetail.Checked)
                        //  fsVersion = "No";

                        String fsSummaryFlag = Convert.ToString(dtBatch.Rows[i]["Ssi_SummaryDetail"]);
                        String fsAllignment = Convert.ToString(dtBatch.Rows[i]["Ssi_Alignment"]);
                        String fsDisplayContactName = Convert.ToString(dtBatch.Rows[i]["ContactName"]);
                        String fsContactId = Convert.ToString(dtBatch.Rows[i]["ssi_ContactID"]);
                        String fsKeyContactID = Convert.ToString(dtBatch.Rows[i]["ssi_keycontactId"]);
                        String fsHousholdReportTitle = Convert.ToString(dtBatch.Rows[i]["ssi_householdreporttitle"]);
                        String fsGreshReportIdName = Convert.ToString(dtBatch.Rows[i]["ssi_GreshamReportIdName"]);
                        String fsGAorTIAflag = Convert.ToString(dtBatch.Rows[i]["ssi_gaortia"]);
                        String fsCoverSheetPageTitle = Convert.ToString(dtBatch.Rows[i]["ssi_coverpagetitle"]);
                        String lsFinalTitleAfterChange = String.Empty;
                        String fsDiscretionaryFlg = Convert.ToString(dtBatch.Rows[i]["Discretionary Flag"]);
                        String fsReportRollupGroupIdName = Convert.ToString(dtBatch.Rows[i]["Ssi_ReportRollupGroupIdName"]).Replace("'", "''");
                        String fsHHreportparametersId = Convert.ToString(dtBatch.Rows[i]["Ssi_hhreportparametersId"]);
                        fsReportingName = Convert.ToString(dtBatch.Rows[i]["Ssi_ReportingName"]);


                        //added 2_1_2019 Non marketable (DYNAMO)
                        String fsReportRollupGroupId = Convert.ToString(dtBatch.Rows[i]["Ssi_ReportRollupGroupId"]);
                        String fsrHouseholdId = Convert.ToString(dtBatch.Rows[i]["Ssi_HouseholdId"]);
                        String fsFundIRR = Convert.ToString(dtBatch.Rows[i]["ssi_FundIRR"]);
                        String fsGreshamReportId = Convert.ToString(dtBatch.Rows[i]["ssi_GreshamReportId"]);

                        //added 5_20_2019 -- LegalEntity -- Title
                        String fsLegalEntityTitle = Convert.ToString(dtBatch.Rows[i]["Ssi_LegalEntityIdName"]);

                        String FundID = "";
                        if (Convert.ToString(dtBatch.Rows[i]["ssi_FundId"]) == "")
                            FundID = "";
                        else
                            FundID = Convert.ToString(dtBatch.Rows[i]["ssi_FundId"]);

                        String LegalEntity = "";
                        if (Convert.ToString(dtBatch.Rows[i]["ssi_LegalEntityId"]) == "")
                            LegalEntity = "";
                        else
                            LegalEntity = Convert.ToString(dtBatch.Rows[i]["ssi_LegalEntityId"]);




                        if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["HouseHoldReportTitle"])))
                            lsFinalTitleAfterChange = Convert.ToString(dtBatch.Rows[i]["HouseHoldReportTitle"]);

                        if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["AllocationGroupReportTitle"])))
                            lsFinalTitleAfterChange = Convert.ToString(dtBatch.Rows[i]["AllocationGroupReportTitle"]);

                        String fsFooterTxt = String.Empty;
                        if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["GreshamFooterTxt"])))
                            fsFooterTxt = Convert.ToString(dtBatch.Rows[i]["GreshamFooterTxt"]);

                        /*Change added on 31st OCT 2010*/
                        String fsReportGroupflag = "null";
                        if (Convert.ToString(dtBatch.Rows[i]["ssi_report"]) == "")
                            fsReportGroupflag = "null";
                        else
                            fsReportGroupflag = Convert.ToString(dtBatch.Rows[i]["ssi_report"]);
                        //Convert.ToString(dtBatch.Rows[i]["ssi_report"]).Replace(",", "");
                        String fsReportgroupflag2 = "null";
                        if (Convert.ToString(dtBatch.Rows[i]["ssi_report2"]) == "")
                            fsReportgroupflag2 = "null";
                        else
                            fsReportgroupflag2 = Convert.ToString(dtBatch.Rows[i]["ssi_report2"]);

                        /* END OF CHANGE*/



                        // Added By Rohit Pawar
                        // Logic to get header for commitment schedule
                        String CommitmentReportHeader = "";
                        if (fsHouseholdName != "")
                        {
                            if (lsFinalTitleAfterChange == "" && fsAllocationGroup != "")
                            {
                                CommitmentReportHeader = fsAllocationGroup;
                            }
                            else if (fsHouseholdName != "" && fsAllocationGroup == "" && lsFinalTitleAfterChange == "")
                            {
                                if (fsHousholdReportTitle != "")
                                {
                                    CommitmentReportHeader = fsHousholdReportTitle;
                                }
                                else
                                {
                                    CommitmentReportHeader = fsHouseholdName;
                                }
                            }
                            else
                            {
                                CommitmentReportHeader = lsFinalTitleAfterChange;
                            }
                        }
                        else
                        {
                            CommitmentReportHeader = "";
                        }

                        string strGUID = Guid.NewGuid().ToString();
                        strGUID = strGUID.Substring(0, 5);
                        //String lsExcleSavePath = ReportOpFolder + "\\" + ContactFolderName + "\\" + fsHouseholdName.Replace(",", "") + "_" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + strGUID + ".xls";
                        String lsExcleSavePath = ReportOpFolder + "\\" + ParentFolder + "\\" + ContactFolderName + "\\" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + lsFinalTitleAfterChange.Replace(",", "").Replace("/", "").Replace("\\", "") + "_" + Convert.ToDateTime(fsAsofDate).ToString("yyyyMMdd") + "_" + strGUID + ".xls";
                        //String lsSavePathCombReport  = ReportOpFolder + "\\" + ContactFolderName + "\\" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + lsFinalTitleAfterChange.Replace(",", "").Replace("/", "").Replace("\\", "") + "_" + Convert.ToDateTime(fsAsofDate).ToString("yyyyMMdd") + "_" + strGUID + "_Combined.pdf"; 
                        String lsCoversheet = ReportOpFolder + "\\" + ParentFolder + "\\" + ContactFolderName + "\\Coversheet.xls";
                        //String fsHouseHoldReportTitle = "";

                        //Page number logic 
                        if (i == 0)
                        {
                            dtBatch.Columns.Add("numPageNo", typeof(System.Int32));
                            dtBatch.Rows[i]["numPageNo"] = "1";
                        }


                        lg.AddinLogFile(Session["Filename"].ToString(), "Report Name " + fsGreshReportIdName);
                        bool bContinueBatch = true;

                        /** Attach Template PDF ---Static pdf logic  ***/
                        string strTemplateFilePath = Convert.ToString(dtBatch.Rows[i]["ssi_TemplateFilePath"]);
                        if (strTemplateFilePath != "")
                        {
                            string strExtension = Path.GetExtension(strTemplateFilePath);


                            #region Fetch File from Sharepoint
                            if (strTemplateFilePath.Contains("https://greshampartners.sharepoint.com") || strTemplateFilePath.Contains("http://greshampartners.sharepoint.com"))
                            {

                                string FileName = Path.GetFileName(strTemplateFilePath);
                                FileName = FileName.Replace("%20", " ");
                                // string FileName2 = HttpUtility.HtmlEncode(FileName).ToString();
                                string SharepointPath = strTemplateFilePath;
                                SharepointPath = SharepointPath.Replace("//", "/");
                                SharepointPath = SharepointPath.Replace("https:/greshampartners.sharepoint.com/clientserv/", "");
                                SharepointPath = SharepointPath.Replace("http:/greshampartners.sharepoint.com/clientserv/", "");
                                SharepointPath = SharepointPath.Replace("%20", " ");
                                SharepointPath = SharepointPath.Replace(FileName, "");

                                string LocalPath = ReportOpFolder + "\\" + ParentFolder + "\\" + ContactFolderName + "\\";

                                strTemplateFilePath = sharepointFile(FileName, SharepointPath, LocalPath);
                            }
                            #endregion
                            if (strExtension.ToString().ToLower() == ".doc" || strExtension.ToString().ToLower() == ".docx")
                            {
                                strTemplateFilePath = ConvertDocument(strTemplateFilePath, lsExcleSavePath);
                                strTemplateFilePath = strTemplateFilePath.Replace(".xls", ".pdf");
                            }
                            if (strExtension.ToString().ToLower() == ".xls" || strExtension.ToString().ToLower() == ".xlsx")
                            {
                                strTemplateFilePath = ConvertSpreadsheet(strTemplateFilePath, lsExcleSavePath);
                                strTemplateFilePath = strTemplateFilePath.Replace(".xls", ".pdf");
                            }

                            //FOR -- TESTING 
                            if (Request.Url.AbsoluteUri.Contains("localhost"))
                                strTemplateFilePath = @"C:\Reports\Commentaries.pdf";

                            if (Convert.ToString(Session["CurPageInBatch"]) == "")
                                Session["CurPageInBatch"] = "0";

                            lsExcleSavePath = strTemplateFilePath.Replace(".pdf", ".xls");
                            int numofPage = objCombinedReports.GetPageCountFromPDF(strTemplateFilePath);
                            int CurPage = Convert.ToInt32(Convert.ToString(Session["CurPageInBatch"])) + 1;
                            if (numofPage > 0)
                            {
                                numofPage--;
                                dtBatch.Rows[i]["numPageNo"] = CurPage;
                                Session["CurPageInBatch"] = numofPage + CurPage;
                                bContinueBatch = false;
                            }
                            else
                                dtBatch.Rows[i]["numPageNo"] = 0;

                        }

                        bool CombinedFileName = false;

                        /** if record is template then it will not generate report -- only static pdf will attach **/
                        /** Generate report on excel and pdf **/
                        if (bContinueBatch)
                        {
                            if (i != 0)
                            {
                                if (Session["CurPageInBatch"] != null)
                                {
                                    int CurPage = Convert.ToInt32(Convert.ToString(Session["CurPageInBatch"])) + 1;
                                    dtBatch.Rows[i]["numPageNo"] = CurPage;
                                }
                            }

                            // Generate report on excel and pdf


                            if (fsGreshReportIdName != "Asset Distribution" && fsGreshReportIdName != "Asset Distribution Comparison")
                            {
                                DateTime dtime = DateTime.Now;
                                lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + fsGreshReportIdName + "," + dtime + ",");
                                //CombinedFileName = generateCombinedPDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath.Replace(".xls", ".pdf"), fsFooterTxt, fsGreshReportIdName, LegalEntity, FundID, CommitmentReportHeader, fsGAorTIAflag, fsReportRollupGroupIdName, fsHHreportparametersId);
                                CombinedFileName = generateCombinedPDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath.Replace(".xls", ".pdf"), fsFooterTxt, fsGreshReportIdName, LegalEntity, FundID, CommitmentReportHeader, fsGAorTIAflag, fsReportRollupGroupIdName, fsHHreportparametersId, fsReportRollupGroupId, fsrHouseholdId, fsFundIRR, fsGreshamReportId, fsLegalEntityTitle, TempFolderPath);



                                lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + fsGreshReportIdName + "," + dtime + "," + DateTime.Now);

                                string fname = lsExcleSavePath.Replace(".xls", ".pdf");
                                var sess = Session["CurPageInBatch"];
                                if (sess == null)
                                {
                                    int pageno = PDF.get_pageCcount(fname);
                                    HttpContext.Current.Session["CurPageInBatch"] = pageno;
                                }

                            }
                            else
                            {
                                DateTime dtime = DateTime.Now;
                                lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + fsGreshReportIdName + "," + dtime + ",");
                                SetValuesToVariable(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg);
                                // generatesExcelsheets(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt,fsGAorTIAflag, fsDiscretionaryFlg);
                                generatePDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg, TempFolderPath);
                                CombinedFileName = true;

                                lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + fsGreshReportIdName + "," + dtime + "," + DateTime.Now);

                            }

                            loCoversheetCheck = new FileInfo(lsCoversheet);

                            ////added 18_05_2018 - CLEANUP JUNKFILES(sasmit)
                            //if (loCoversheetCheck.Exists)
                            //{
                            //    try
                            //    {
                            //        System.IO.File.Delete(lsCoversheet);
                            //        System.IO.File.Delete(lsCoversheet.Replace(".xls", ".pdf"));
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //    }
                            //}
                            //loCoversheetCheck = new FileInfo(lsCoversheet);
                            if (!loCoversheetCheck.Exists)
                            {
                                DateTime dtime = DateTime.Now;
                                //  sw.WriteLine();
                                lg.AddinLogFile(Session["Filename"].ToString(), HHName + ",Coversheet," + dtime + ",");
                                generateCoversheetPDF(fsAsofDate, lsCoversheet, fsAllocationGroup, fsHouseholdName, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange, fsCoverSheetPageTitle, fsGAorTIAflag, fsDiscretionaryFlg, TempFolderPath);
                                generatesCoverExcel(fsAsofDate, fsHouseholdName, fsAllocationGroup, lsCoversheet, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange, fsCoverSheetPageTitle, TempFolderPath);
                                lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + fsGreshReportIdName + "," + dtime + "," + DateTime.Now);
                            }
                        }
                        else
                        {
                            CombinedFileName = true;
                        }
                        /* Array fill with the PATH + Fullname of PDF*/

                        //if (i == 0)
                        //{
                        //    SourceFileArray[i] = lsCoversheet.Replace(".xls", ".pdf");
                        //    if (CombinedFileName == true)
                        //        SourceFileArray[i + 1] = lsExcleSavePath.Replace(".xls", ".pdf");
                        //}

                        // added 10/03/2016 - sasmit
                        # region coverletter Sourcearray1 filearray
                        if (i == 0 && noIndex.Text == "coversheet")
                        {
                            //FileUpload1.PostedFile.SaveAs(Server.MapPath("") + @"\ExcelTemplate\" + FileUpload1.FileName);
                            //string filename = Path.GetFileName(FileUpload1.FileName);

                            //  string strClientPath = Server.MapPath("") + @"\ExcelTemplate\" + filename;
                            string strClientPath = AppLogic.GetParam(AppLogic.ConfigParam.CoverLetter); // "\\GRPAO1-VWFS01\_ops_C_I_R_group\JM Squared\Test Planning\CoverLetters\Cover Letter.pdf";
                            strClientPath = strClientPath + "Cover Letter.pdf";

                            SourceFileArray1[0] = strClientPath;
                            SourceFileArray1[1 + i] = lsCoversheet.Replace(".xls", ".pdf");
                            for (int PageCnt = 1; PageCnt < numIndexPageCount; PageCnt++)
                            {
                                //9-5-2017 added ifclause MTGBK-NO Index - sasmit
                                if (!ContactFolderName.Contains("MTGBK"))
                                {
                                    SourceFileArray1[1 + i + PageCnt] = (Server.MapPath("") + @"\ExcelTemplate\Blank.pdf");
                                }
                            }
                            //  SourceFileArray1[2 + i] = (Server.MapPath("") + @"\ExcelTemplate\Blank.pdf");


                            if (CombinedFileName == true)
                                SourceFileArray1[i + (numIndexPageCount + 1)] = lsExcleSavePath.Replace(".xls", ".pdf");


                        }
                        #endregion
                        else if (i == 0)
                        {
                            SourceFileArray[i] = lsCoversheet.Replace(".xls", ".pdf");
                            for (int PageCnt = 1; PageCnt < numIndexPageCount; PageCnt++)
                            {
                                if (!ContactFolderName.Contains("MTGBK"))
                                {
                                    SourceFileArray[i + PageCnt] = (Server.MapPath("") + @"\ExcelTemplate\Blank.pdf");
                                }
                            }
                            if (CombinedFileName == true)
                                SourceFileArray[i + (numIndexPageCount)] = lsExcleSavePath.Replace(".xls", ".pdf");
                            //else if (CombinedFileName == false)
                            //{
                            //lblError.Text = "No Record Found";
                            //lblError.Visible = true;
                            //return;
                            //}
                        }
                        else
                        {
                            if (CombinedFileName == true && noIndex.Text == "coversheet")
                                SourceFileArray1[i + 1 + (numIndexPageCount)] = lsExcleSavePath.Replace(".xls", ".pdf");
                            else if (CombinedFileName == true)
                                SourceFileArray[i + (numIndexPageCount)] = lsExcleSavePath.Replace(".xls", ".pdf");
                        }

                        /* Array fill with the PATH + Fullname of PDF*/
                    }

                    // Consolidate File Logic NEW
                    DateTime dtAsOfDate = Convert.ToDateTime(ViewState["AsOfDate"]);

                    strYear = dtAsOfDate.Year.ToString().Length < 2 ? "0" + dtAsOfDate.Year.ToString() : dtAsOfDate.Year.ToString();
                    strMonth = dtAsOfDate.Month.ToString().Length < 2 ? "0" + dtAsOfDate.Month.ToString() : dtAsOfDate.Month.ToString();
                    strDay = dtAsOfDate.Day.ToString().Length < 2 ? "0" + dtAsOfDate.Day.ToString() : dtAsOfDate.Day.ToString();

                    ConsolidatePdfFileName = HHName + "_" + strYear + "-" + strMonth + strDay + "_" + CurrentTimeStamp + ".pdf";
                    ConsolidatePdfFileName = GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);

                    //string DisplayFileName = HHName + "_" + strYear + "-" + strMonth + strDay + ".pdf";

                    string DisplayFileName = HHName + " " + strYear + "-" + strMonth + strDay + ".pdf";

                    string OldDisplayFileName = OldHHName + " " + strYear + "-" + strMonth + strDay + ".pdf";



                    if (SPVFileName != "")
                    {
                        DisplayFileName = GeneralMethods.RemoveSpecialCharacters(DisplayFileName);
                        DisplayFileName = DisplayFileName;//.Replace(" Family", "").Replace(",", "");
                    }
                    else
                    {
                        DisplayFileName = GeneralMethods.RemoveSpecialCharacters(DisplayFileName);
                        DisplayFileName = DisplayFileName.Replace(" Family", "").Replace(",", "");
                    }


                    if (!System.IO.File.Exists(ReportOpFolder + "\\" + ParentFolder + "\\" + ConsolidatePdfFileName))
                        System.IO.File.Copy(ReportOpFolder + "\\" + ParentFolder + "\\" + ContactFolderName + "\\Coversheet.pdf", ReportOpFolder + "\\" + ParentFolder + "\\" + ConsolidatePdfFileName);


                    DestinationPath = ReportOpFolder + "\\" + GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);


                    if (ContactFolderName.Contains("MTGBK")) //generate without coversheet
                    {
                        string[] target = new string[sourcefilecount - (numIndexPageCount)];
                        Array.Copy(SourceFileArray, (numIndexPageCount), target, 0, sourcefilecount - (numIndexPageCount));

                        DateTime dtime = DateTime.Now;
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + ",Merge," + dtime + ",");
                        PDF.MergeFiles(DestinationPath, target);
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + ",Merge," + "," + dtime + "," + DateTime.Now);
                    }
                    else if (noIndex.Text == "coversheet") // added 10/03/2016 - sasmit
                    {
                        DateTime dtime = DateTime.Now;
                        //   sw.WriteLine(HHName + ", Merge Coversheet," + dtime + ",");
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + ", Merge Coversheet," + dtime + ",");
                        PDF.MergeFiles1(DestinationPath, SourceFileArray1);
                        // PDF.CombineMultiplePDFs(DestinationPath, SourceFileArray1);
                        //  filecount1.Text = "merge done at " + DestinationPath + "--" + DateTime.Now + "------------- End of path";
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + ", Merge Coversheet," + "," + dtime + "," + DateTime.Now);

                    }
                    else  //generate with coversheet
                    {
                        DateTime dtime = DateTime.Now;
                        //  sw.WriteLine(HHName + ", Merge Coversheet," + dtime + ",");
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + ", Merge Coversheet," + dtime + ",");
                        PDF.MergeFiles(DestinationPath, SourceFileArray);
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + ", Merge Coversheet," + "," + dtime + "," + DateTime.Now);
                    }

                    lg.AddinLogFile(Session["Filename"].ToString(), "ContactFolderName" + ContactFolderName + DateTime.Now);
                    lg.AddinLogFile(Session["Filename"].ToString(), "noIndex.Text" + noIndex.Text + DateTime.Now);
                    //added 9_5_2017 (MTGBK-NO INDEX)-SASMIT
                    if (ContactFolderName.Contains("MTGBK"))
                    {
                        DateTime dtime = DateTime.Now;
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + ",MTGBK Index," + "," + dtime + "," + DateTime.Now);
                        System.IO.File.Copy(DestinationPath, ApprovedReports + ConsolidatePdfFileName, true);
                        // System.IO.File.Copy(DestinationPath, DestinationPath, true);

                        ////added  31-july-2018 sasmit(ops folder delete issue)
                        //if (ContactFolderName != "")
                        //{
                        //    if (Directory.Exists(ReportOpFolder + "\\" + ParentFolder))
                        //    {
                        //        System.IO.Directory.Delete(ReportOpFolder + "\\" + ParentFolder, true);
                        //    }

                        //}
                        ////delete tempfolder creted at local Directory
                        //if (Directory.Exists(Local_ParentFolderPath))
                        //{
                        //    Directory.Delete(Local_ParentFolderPath, true);
                        //}

                        Session.Remove("CurPageInBatch");
                        strReportFiles = strReportFiles + "<br/>" + "<a href=file:" + AppLogic.GetParam(AppLogic.ConfigParam.OutPutReports) + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1).Replace(" ", "%20") + ">" + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1) + " </a>";

                    }
                    else if (noIndex.Text == "coversheet") // added 10/03/2016 - sasmit
                    {
                        DateTime dtime = DateTime.Now;
                        //  sw.WriteLine();

                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + ", Coversheet Index," + dtime + ",");
                        string DestinationPath1 = objCombinedReports.addPageIndex1(DestinationPath, dtBatch, TempFolderPath);
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + ", Coversheet Index," + "," + dtime + "," + DateTime.Now);
                        //sourcearray.Text = "Error in  " + Session["Error"].ToString();//Testing

                        // sourcearray.Text = DestinationPath1;
                        //filecount1.Text = filecount1 + "with cover" + Session["Rowcount"].ToString();//Testing
                        //filecount2.Text = "without cover" + Session["EndRow"].ToString();//Testing

                        System.IO.File.Copy(DestinationPath1, ApprovedReports + ConsolidatePdfFileName, true);  //Test PAth 
                        // File.Copy(DestinationPath1,@"C:\Reports\" + ConsolidatePdfFileName, true);   //Local PAth 
                        System.IO.File.Copy(DestinationPath1, DestinationPath, true);

                        ////added  31-july-2018 sasmit(ops folder delete issue)
                        //if (ContactFolderName != "")
                        //{
                        //    if (Directory.Exists(ReportOpFolder + "\\" + ParentFolder))
                        //    {
                        //        System.IO.Directory.Delete(ReportOpFolder + "\\" + ParentFolder, true);
                        //    }

                        //}
                        ////delete tempfolder creted at local Directory
                        //if (Directory.Exists(Local_ParentFolderPath))
                        //{
                        //    Directory.Delete(Local_ParentFolderPath, true);
                        //}

                        // //added 18_05_2018 - CLEANUP JUNKFILES(sasmit)
                        //System.IO.File.Delete(DestinationPath1);

                        //string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + ConsolidatePdfFileName);

                        //File.Copy(DestinationPath, strDirectory, true);

                        Session.Remove("CurPageInBatch");
                        strReportFiles = strReportFiles + "<br/>" + "<a href=file:" + AppLogic.GetParam(AppLogic.ConfigParam.OutPutReports) + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1).Replace(" ", "%20") + ">" + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1) + " </a>";
                        //sourcearray.Text = "File at :" + strReportFiles + DateTime.Now +"end of path "; //testing

                    }
                    else
                    {
                        DateTime dtime = DateTime.Now;
                        //   sw.WriteLine(HHName + ", Index," + dtime + ",");
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + ", Index," + dtime + ",");

                        // File.Copy(DestinationPath, ApprovedReports + ConsolidatePdfFileName, true);
                        string DestinationPath1 = objCombinedReports.addPageIndex(DestinationPath, dtBatch, TempFolderPath);
                        lg.AddinLogFile(Session["Filename"].ToString(), HHName + "," + ", Index," + "," + dtime + "," + DateTime.Now);
                        System.IO.File.Copy(DestinationPath1, ApprovedReports + ConsolidatePdfFileName, true);
                        System.IO.File.Copy(DestinationPath1, DestinationPath, true);

                        ////added  31-july-2018 sasmit(ops folder delete issue)
                        //if (ContactFolderName != "")
                        //{
                        //    if (Directory.Exists(ReportOpFolder + "\\" + ParentFolder))
                        //    {
                        //        System.IO.Directory.Delete(ReportOpFolder + "\\" + ParentFolder, true);
                        //    }

                        //}
                        ////delete tempfolder creted at local Directory
                        //if (Directory.Exists(Local_ParentFolderPath))
                        //{
                        //    Directory.Delete(Local_ParentFolderPath, true);
                        //}

                        // //added 18_05_2018 - CLEANUP JUNKFILES(sasmit)
                        //System.IO.File.Delete(DestinationPath1);

                        Session.Remove("CurPageInBatch");
                        strReportFiles = strReportFiles + "<br/>" + "<a href=file:" + AppLogic.GetParam(AppLogic.ConfigParam.OutPutReports) + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1).Replace(" ", "%20") + ">" + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1) + " </a>";

                    }
                    if (ddlAction.SelectedValue == "10")  //Insert Coversheet  added 10/03/2016 - sasmit
                    {
                        #region Region to update Batch File Name & Batch File Display Name
                        //code to update updatedate in batch ety of crm
                        //ssi_batch objBatch = new ssi_batch();
                        Entity objBatch = new Entity("ssi_batch");

                        if (BatchIdListTxt != "")
                        {
                            // objBatch.ssi_batchid = new Key();
                            // objBatch.ssi_batchid.Value = new Guid(BatchIdListTxt);

                            objBatch["ssi_batchid"] = new Guid(Convert.ToString(BatchIdListTxt));
                        }

                        if (DestinationPath != "")
                        {
                            //objBatch.ssi_batchdisplayfilename = DisplayFileName;
                            objBatch["ssi_batchdisplayfilename"] = DisplayFileName;
                        }

                        if (ConsolidatePdfFileName != "")
                        {
                            //objBatch.ssi_batchfilename = DestinationPath;
                            objBatch["ssi_batchfilename"] = DestinationPath;
                        }

                        if (BatchIdListTxt != "")
                        {
                            service.Update(objBatch);
                        }


                        #endregion
                    }
                   // else if (ddlAction.SelectedValue == "1")  // Approve
                    else if (ddlAction.SelectedValue == "2")  //Review Pdf
                    {
                        #region Region to update Batch File Name & Batch File Display Name
                        //code to update updatedate in batch ety of crm
                        //ssi_batch objBatch = new ssi_batch();
                        Entity objBatch = new Entity("ssi_batch");

                        if (BatchIdListTxt != "")
                        {
                            // objBatch.ssi_batchid = new Key();
                            // objBatch.ssi_batchid.Value = new Guid(BatchIdListTxt);

                            objBatch["ssi_batchid"] = new Guid(Convert.ToString(BatchIdListTxt));
                        }

                        if (DestinationPath != "")
                        {
                            // objBatch.ssi_batchdisplayfilename = DisplayFileName;
                            objBatch["ssi_batchdisplayfilename"] = DisplayFileName;
                        }

                        if (ConsolidatePdfFileName != "")
                        {
                            // objBatch.ssi_batchfilename = DestinationPath;
                            objBatch["ssi_batchfilename"] = DestinationPath;
                        }

                        if (BatchIdListTxt != "")
                        {
                            service.Update(objBatch);
                        }


                        #endregion
                    }

                }
            }

            ////////////////////////////////////

            // if (ddlAction.SelectedValue != "1" && ddlAction.SelectedValue != "10")//Approved and Insert Coversheet
            if (ddlAction.SelectedValue != "2" && ddlAction.SelectedValue != "10")//Approved and Insert Coversheet
            {
                if (NoOfBatches == 1)
                {
                    string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ConsolidatePdfFileName);

                    System.IO.File.Copy(DestinationPath, strDirectory, true);
                    System.IO.File.Delete(DestinationPath);
                    //Directory.Delete(ReportOpFolder, true);

                    try
                    {
                        //loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

                        // Response.Write("<script>");
                        string lsFileNamforFinal = "./ExcelTemplate/TempFolder/" + ConsolidatePdfFileName;
                        //Response.Write("window.open('" + lsFileNamforFinal + "', 'mywindow')");
                        // Response.Write("window.open('ViewReport.aspx?" + ConsolidatePdfFileName + "', 'mywindow')");

                        //  Response.Write("</script>");

                        System.Text.StringBuilder sb = new System.Text.StringBuilder();
                        Type tp = this.GetType();
                        sb.Append("\n<script type=text/javascript>\n");
                        sb.Append("\nwindow.open('ViewReport.aspx?" + ConsolidatePdfFileName + "', 'mywindow');");
                        sb.Append("</script>");
                        ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());

                    }
                    catch (Exception exc)
                    {
                        Response.Write(exc.Message);
                        lg.AddinLogFile(Session["Filename"].ToString(), exc.Message.ToString() + DateTime.Now);
                    }
                }
                else
                {
                    //strReportFiles = "<br/>" + strReportFiles + "<a href=file:///" + DestinationPath.Replace("\\\\", "\\").Replace("\\", "//").Replace(" ", "%20") + ">" + DestinationPath.Replace("\\\\", "\\").Replace("\\", "//").Replace(" ", "%20") + "</a>";


                }
            }
            else if (ddlAction.SelectedValue == "1")
            {
                //File.Copy(DestinationPath, ApprovedReports + ConsolidatePdfFileName);
            }

        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            ////added  31-july-2018 sasmit(ops folder delete issue)
            //if (ContactFolderName != "")
            //{
            //    if (Directory.Exists(ReportOpFolder + "\\" + ContactFolderName))
            //    {
            //        Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
            //    }
            //    //delete tempfolder creted at local Directory
            //    if (Directory.Exists(Local_ParentFolderPath))
            //    {
            //        Directory.Delete(Local_ParentFolderPath, true);
            //    }
            //}
            lg.AddinLogFile(Session["Filename"].ToString(), exc.Message.ToString() + DateTime.Now);
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblError.Text = strDescription;
        }
        catch (Exception ex)
        {
            ////added  31-july-2018 sasmit(ops folder delete issue)
            //if (ContactFolderName != "")
            //{
            //    if (Directory.Exists(ReportOpFolder + "\\" + ContactFolderName))
            //    {
            //        Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
            //    }
            //    //delete tempfolder creted at local Directory
            //    if (Directory.Exists(Local_ParentFolderPath))
            //    {
            //        Directory.Delete(Local_ParentFolderPath, true);
            //    }
            //}
            lg.AddinLogFile(Session["Filename"].ToString(), ex.Message.ToString() + DateTime.Now);
            lblError.Text = "Error Generating Report " + ex.ToString();
        }
        finally
        {
            //added  31-july-2018 sasmit(ops folder delete issue)
            if (Directory.Exists(ReportOpFolder + "\\" + ParentFolder))
            {
                lg.AddinLogFile(Session["Filename"].ToString(), "FOLDER DELETE--> " + ReportOpFolder + "\\" + ParentFolder + "-----" + DateTime.Now);
                Directory.Delete(ReportOpFolder + "\\" + ParentFolder, true);
            }
            //delete tempfolder creted at local Directory
            if (Directory.Exists(Local_ParentFolderPath))
            {
                lg.AddinLogFile(Session["Filename"].ToString(), "FOLDER DELETE--> " + Local_ParentFolderPath + "-----" + DateTime.Now);
                Directory.Delete(Local_ParentFolderPath, true);
            }
            //  sw.Close();
        }
        lg.AddinLogFile(Session["Filename"].ToString(), "Report END " + DateTime.Now);
    }
    public string sharepointFile(string FileName, string path, string finalPath)
    {
        string Value = null;


        string siteUrl = "https://greshampartners.sharepoint.com/clientserv";
        context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();

        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
        //foreach (var c in "51ngl3malt") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);
        string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword);
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);

        Web site = context.Web;

        // Folder subFoldercol = site.GetFolderByServerRelativeUrl("Documents" + "/"+"_Test Files");
        // Folder subFoldercol = site.GetFolderByServerRelativeUrl(path.ToLower().Replace("clientserv/", ""));
        Folder subFoldercol = site.GetFolderByServerRelativeUrl(path);
        // Microsoft.SharePoint.Client.File subfile = site.GetFileByServerRelativeUrl("Anziano" + "/" + Path);
        ListCollection collList = site.Lists;

        //  FolderCollection fcolection = subFoldercol.Folders;
        Microsoft.SharePoint.Client.FileCollection fcolection = subFoldercol.Files;
        context.Load(fcolection);
        context.Load(collList);
        context.ExecuteQuery();
        foreach (Microsoft.SharePoint.Client.File f in fcolection)
        {

            string FileNAME = f.Name.ToString();
            if (FileName == FileNAME)
            {
                FileCopy(f, finalPath);
                Value = finalPath + "\\" + FileName;
                break;
            }
            else
            {
                Value = null;
            }
        }
        return Value;
    }
    public void FileCopy(Microsoft.SharePoint.Client.File files1, string finalPath)
    {
        // -- Get fIle and copy to Destination
        Stream filestrem = getFile(files1);
        string fileName = System.IO.Path.GetFileName(files1.Name);
        // string filepath = System.IO.Path.Combine(Test, fileName);
        string filepath = System.IO.Path.Combine(finalPath, fileName);
        // FileStream fileStream = System.IO.File.Create(filepath, (int)filestrem.Length); // Test Local PAth
        FileStream fileStream = System.IO.File.Create(filepath, (int)filestrem.Length); // Original PAth
        // Initialize the bytes array with the stream length and then fill it with data 
        byte[] bytesInStream = new byte[filestrem.Length];
        filestrem.Read(bytesInStream, 0, bytesInStream.Length);
        // Use write method to write to the file specified above 
        fileStream.Write(bytesInStream, 0, bytesInStream.Length);

        fileStream.Close();
    }
    public Stream getFile(Microsoft.SharePoint.Client.File files1)
    {
        context.Load(files1);
        ClientResult<Stream> stream = files1.OpenBinaryStream();
        context.ExecuteQuery();
        return this.ReadFully(stream.Value);
    }
    private Stream ReadFully(Stream input)
    {
        byte[] buffer = new byte[16 * 1024];
        using (MemoryStream ms = new MemoryStream())
        {
            int read;
            while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                ms.Write(buffer, 0, read);
            }
            return new MemoryStream(ms.ToArray()); ;
        }
    }
    public bool generateCombinedPDF(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate,
        String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment,
        String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFooterTxt, String ReportName,
        String LegalEntityId, String FundId, String CommitmentReportHeader, String GAorTIAflag, String ReportRollupGroupIdName, String fsHHreportparametersId, String fsReportRollupGroupId, String fsHouseholdId, String fsFundIRR, String fsGreshamReportId, String fsLegalEntityTitle, String TempFolderPath)
    {
        clsCombinedReports objCombinedReports = new clsCombinedReports();

        objCombinedReports.HouseHoldValue = "";
        objCombinedReports.HouseHoldText = fsHouseholdName;
        objCombinedReports.AllocationGroupValue = "";
        objCombinedReports.AllocationGroupText = fsAllocationGroup;
        objCombinedReports.AsOfDate = fsAsofDate;
        objCombinedReports.lsFamiliesName = fsHouseholdName;
        objCombinedReports.lsDateName = "";
        objCombinedReports.LegalEntityId = LegalEntityId;
        objCombinedReports.FundId = FundId;
        objCombinedReports.FooterText = lsFooterTxt;
        objCombinedReports.CommitmentReportHeader = CommitmentReportHeader;
        objCombinedReports.GreshamAdvisedFlag = GAorTIAflag;
        objCombinedReports.ReportRollupGroupIdName = ReportRollupGroupIdName;
        objCombinedReports.PriorDate = fsSPriorDate;

        //added 2_1_2019 - Non Marketable(DYNAMO)
        objCombinedReports.ReportRollupGroupId = fsReportRollupGroupId;
        objCombinedReports.HouseholdId = fsHouseholdId;
        objCombinedReports.FundIRR = fsFundIRR;
        objCombinedReports.HHParameterTxt = fsHHreportparametersId;
        objCombinedReports.ReportingID = fsGreshamReportId;
        objCombinedReports.ReportName = ReportName;

        //added 8_14_2019 batch Issue(Mixing of Reports)
        objCombinedReports.TempFolderPath = TempFolderPath;

        //added 5_20_2019 -- LegalEntity -- Title
        if (objCombinedReports.ReportingID.ToUpper() == "AFD08C8B-2E25-E911-8106-000D3A1C025B" || objCombinedReports.ReportingID.ToUpper() == "806E4D33-1D29-E911-8106-000D3A1C025B" || objCombinedReports.ReportingID.ToUpper() == "90D6C145-1D29-E911-8106-000D3A1C025B" || objCombinedReports.ReportingID.ToUpper() == "A47E365E-1D29-E911-8106-000D3A1C025B") //Private Equity Performance||Private REal Asset Performance||Outside Private Equity Performance||Outside Private REal Asset Performance
        {
            if (fsLegalEntityTitle != "")
            {
                objCombinedReports.CommitmentReportHeader = fsLegalEntityTitle;
            }
        }

        if (fsReportingName != "")
            objCombinedReports.ReportingName = fsReportingName;

        if (ReportName == "Client Goals" || ReportName == "Absolute Returns" || ReportName == "Capital Protection" || ReportName == "Short Term Performance")
        {
            string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";


            SqlConnection Gresham_con = new SqlConnection(Gresham_String);
            String HHRPIDListTxt = Convert.ToString(fsHHreportparametersId);
            string greshamquery = "[SP_S_HH_PARAMETER_ASSETCLASS] @HHParameterListTxt='" + HHRPIDListTxt + "'";

            SqlDataAdapter dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            DataSet ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);

            if (ds_gresham.Tables[0].Rows.Count > 0)
            {
                string _assetclass = "";
                for (int i = 0; i < ds_gresham.Tables[0].Rows.Count; i++)
                {
                    _assetclass = _assetclass + "," + ds_gresham.Tables[0].Rows[i]["sas_name"].ToString();
                }

                _assetclass = _assetclass.Substring(1, _assetclass.Length - 1);
                objCombinedReports.AssetClassCSV = _assetclass;
            }
        }


        string filepdfname = objCombinedReports.MergeReports(fsFinalLocation, ReportName);

        if (filepdfname == "")
        {
            return false;
        }
        else
            return true;

    }




    public void generatePDF(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFooterTxt, String fsGAorTIAflag, String fsDiscretionaryFlg, String TempFolderPath)
    {
        clsCombinedReports objCombinedReports = new clsCombinedReports();

        liPageSize = 29;
        DataSet lodataset; DB clsDB = new DB();
        lodataset = null;

        String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, fsGAorTIAflag, fsDiscretionaryFlg);
        // Response.Write(lsSQL);
        lodataset = clsDB.getDataSet(lsSQL);
        DataSet loInsertblankRow = lodataset.Copy();
        lodataset.Tables[0].Clear();
        lodataset.Clear();
        lodataset = null;
        lodataset = loInsertblankRow.Clone();
        int liBlankCounter = 1;

        for (int liBlankRow = 0; liBlankRow < loInsertblankRow.Tables[0].Rows.Count; liBlankRow++)
        {
            if (liBlankRow != 0 && loInsertblankRow.Tables[0].Rows[liBlankRow]["_Ssi_BoldFlg"].ToString().ToUpper() == "TRUE" || loInsertblankRow.Tables[0].Rows[liBlankRow]["_Ssi_SuperBoldFlg"].ToString().ToUpper() == "TRUE")
            {
                //if (!String.IsNullOrEmpty(fsSPriorDate) && loInsertblankRow.Tables[0].Rows.Count - 1 != liBlankRow)
                if (loInsertblankRow.Tables[0].Rows.Count - 1 != liBlankRow)
                {
                    DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                    newCustomersRow[0] = "test";
                    newCustomersRow[1] = "test";
                    lodataset.Tables[0].Rows.Add(newCustomersRow);
                    liBlankCounter = liBlankCounter + 1;
                }
                else if (Convert.ToString(loInsertblankRow.Tables[0].Rows[liBlankRow][0]) == "NET WORTH")
                {
                    DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                    newCustomersRow[0] = "test";
                    newCustomersRow[1] = "test";
                    lodataset.Tables[0].Rows.Add(newCustomersRow);
                    liBlankCounter = liBlankCounter + 1;
                }
                else if (fsAllignment != "Horizontal")
                {
                    DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                    newCustomersRow[0] = "test";
                    newCustomersRow[1] = "test";
                    lodataset.Tables[0].Rows.Add(newCustomersRow);
                    liBlankCounter = liBlankCounter + 1;
                }
            }
            lodataset.Tables[0].ImportRow(loInsertblankRow.Tables[0].Rows[liBlankRow]);
        }
        lodataset.AcceptChanges();
        DataSet loInsertdataset = lodataset.Copy();
        for (int liNewdataset = lodataset.Tables[0].Columns.Count - 1; liNewdataset > -1; liNewdataset--)
        {
            if (lodataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") || lodataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
            {
                loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[liNewdataset]);
            }
        }
        //    loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[1]);
        loInsertdataset.AcceptChanges();

        //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 10);
        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 30, 30, 31, 8);//10,10
                                                                                                                              //  String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
        String ls = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        document.Open();


        lsTotalNumberofColumns = loInsertdataset.Tables[0].Columns.Count + "";
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, loInsertdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
        iTextSharp.text.Cell loCell = new Cell();
        setTableProperty(loTable);
        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
        int liTotalPage = (loInsertdataset.Tables[0].Rows.Count / liPageSize);
        int liCurrentPage = 0;
        if (loInsertdataset.Tables[0].Rows.Count % liPageSize != 0)
        {
            liTotalPage = liTotalPage + 1;
        }
        else
        {
            liPageSize = 28;
            liTotalPage = liTotalPage + 1;
        }

        //check the length of the column name to set the pagesize.
        for (int j = 0; j < loInsertdataset.Tables[0].Columns.Count; j++)
        {
            if (loInsertdataset.Tables[0].Columns[j].ColumnName.Length > 30)
            {
                liPageSize = 28;
            }
        }

        for (int liRowCount = 0; liRowCount < loInsertdataset.Tables[0].Rows.Count; liRowCount++)
        {
            if (liRowCount % liPageSize == 0)
            {
                document.Add(loTable);

                if (liRowCount != 0)
                {
                    liCurrentPage = liCurrentPage + 1;
                    //document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                    document.NewPage();
                    objCombinedReports.SetTotalPageCount("Asset Distribution");
                }


                setHeader(document, loInsertdataset);
                loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, loInsertdataset.Tables[0].Rows.Count);   // 2 rows, 2 columns           
                setTableProperty(loTable);
            }

            int colsize = loInsertdataset.Tables[0].Columns.Count;
            for (int liColumnCount = 0; liColumnCount < colsize; liColumnCount++)
            {
                iTextSharp.text.Chunk lochunk = new Chunk();
                String lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                try
                {
                    if (liColumnCount == loInsertdataset.Tables[0].Columns.Count - 1 && fsAllignment == "Horizontal")
                    {
                        lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lsFormatedString));
                    }
                    else
                    {
                        lsFormatedString = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));

                    }
                }
                catch
                {

                }

                //changed on 02/25/2011
                //lochunk = new Chunk(lsFormatedString, Font8Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                lochunk = new Chunk(lsFormatedString, Font7Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                loCell = new iTextSharp.text.Cell();
                loCell.Border = 0;
                loCell.NoWrap = true;
                // loCell.VerticalAlignment=0;
                loCell.VerticalAlignment = 5;

                setGreyBorder(lodataset, loCell, liRowCount);
                loCell.Leading = 6f;//6
                loCell.UseBorderPadding = true;

                //  if (lodataset.Tables[0].Rows[liRowCount]["_Ssi_TabFlg"].ToString() == "True" && lodataset.Tables[0].Rows[liRowCount]["_Ssi_UnderlineFlg"].ToString() != "True")


                if (liColumnCount != 0)
                {
                    loCell.HorizontalAlignment = 2;
                }


                /*=========START WITH BOLD AND SUPERBOLD FLAG========*/
                if (checkTrue(lodataset, liRowCount, "_Ssi_BoldFlg") || checkTrue(lodataset, liRowCount, "_Ssi_SuperBoldFlg"))
                {
                    lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                    try
                    {
                        if (liColumnCount == loInsertdataset.Tables[0].Columns.Count - 1 && fsAllignment == "Horizontal")
                        {
                            lsFormatedString = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(lsFormatedString));
                        }
                        else
                        {
                            lsFormatedString = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(lsFormatedString));

                        }
                    }
                    catch
                    {

                    }

                    //changed on 02/25/2011
                    //lochunk = new Chunk(lsFormatedString, Font9Bold());
                    lochunk = new Chunk(lsFormatedString, Font8Bold());

                    if (!lodataset.Tables[0].Rows[liRowCount][0].ToString().Contains("NET CHANGE"))
                    {
                        //changed on 02/25/2011
                        //lochunk = new Chunk(lsFormatedString, Font9Bold());
                        lochunk = new Chunk(lsFormatedString, Font8Bold());
                        //loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216); // change by abhi
                        loCell.BackgroundColor = new iTextSharp.text.Color(191, 191, 191);
                        if (lsFormatedString.Length > 25)
                        {
                            if (checkTrue(lodataset, liRowCount, "_Ssi_BoldFlg"))
                            {
                                //decrease columncount by 1 to adjust the Colspan. eg: NON-INVESTMENT ASSETS/LOOK-THROUGHS
                                loCell.Colspan = 2;
                                colsize = colsize - 1;
                            }
                        }
                        setBottomWidthWhite(loCell);

                    } /*=========IF END OF BOLD AND SUPERBOLD FLAG========*/
                    else
                    {
                        if (lodataset.Tables[0].Rows[liRowCount][0].ToString() == "NET CHANGE")
                        {
                            setGreyBorder(loCell);
                            //added on 28Feb2011 to change font size for total
                            if (liColumnCount != 0)
                            {
                                lochunk = new Chunk(lsFormatedString, Font7Bold());
                            }
                        }
                    }

                    if (lodataset.Tables[0].Rows[liRowCount][0].ToString().Contains("NET CHANGE %"))
                    {

                        lsFormatedString = Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][liColumnCount]);
                        try
                        {
                            lsFormatedString = String.Format("{0:#,###0.0%;(#,###0.0%)}", Convert.ToDecimal(lsFormatedString) / 100);
                        }
                        catch
                        {

                        }
                        //changed on 02/25/2011
                        //lochunk = new Chunk(lsFormatedString, Font9Bold());
                        lochunk = new Chunk(lsFormatedString, Font8Bold());
                        //added on 28Feb2011 to change font size for total
                        if (liColumnCount != 0)
                        {
                            lochunk = new Chunk(lsFormatedString, Font7Bold());
                        }
                    }
                }
                else
                {
                    if (liColumnCount == 0 && !checkTrue(lodataset, liRowCount, "_Ssi_UnderlineFlg"))
                    {
                        String abc = "          " + lodataset.Tables[0].Rows[liRowCount][1].ToString();
                        //changed on 02/25/2011
                        //lochunk = new Chunk(abc, Font9Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                        lochunk = new Chunk(abc, Font7Whitecheck(Convert.ToString(loInsertdataset.Tables[0].Rows[liRowCount][0])));
                    }
                }
                if (checkTrue(lodataset, liRowCount, "_Ssi_TabFlg") && !checkTrue(lodataset, liRowCount, "_Ssi_UnderlineFlg"))
                {
                    if (liColumnCount == 0)
                    {
                        String abc = "          " + "          " + lodataset.Tables[0].Rows[liRowCount][1].ToString();
                        //changed on 02/25/2011
                        //lochunk = new Chunk(abc, Font8Grey());
                        lochunk = new Chunk(abc, Font7Grey());
                    }
                    else
                    {
                        //changed on 02/25/2011
                        //lochunk = new Chunk(lsFormatedString, Font8Grey());
                        lochunk = new Chunk(lsFormatedString, Font7Grey());
                    }
                }

                //CONDITION FOR SUPERBOLDFLAG
                checkTrue(lodataset, liRowCount, "_Ssi_SuperBoldFlg", loCell, new iTextSharp.text.Color(183, 221, 232));
                //====added on 28Feb2011 to change font size for total====
                if (checkTrue(lodataset, liRowCount, "_Ssi_SuperBoldFlg"))
                {
                    if (liColumnCount != 0)
                    {
                        lochunk = new Chunk(lsFormatedString, Font7Bold());
                    }
                }
                /*=====END=====*/

                if (checkTrue(lodataset, liRowCount, "_Ssi_UnderlineFlg"))
                {
                    if (liColumnCount == 0)
                    {
                        String abc = "          " + "          " + "Total";
                        //changed on 02/25/2011
                        //lochunk = new Chunk(abc, Font8Normal());
                        lochunk = new Chunk(abc, Font7Normal());
                    }
                    setTopWidthBlack(loCell);
                    setBottomWidthWhite(loCell);

                }
                loCell.Add(lochunk);
                loTable.AddCell(loCell);
            }

            if (liRowCount == loInsertdataset.Tables[0].Rows.Count - 1)
            {
                document.Add(loTable);
                liCurrentPage = liCurrentPage + 1;
                //document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));
                objCombinedReports.SetTotalPageCount("Asset Distribution");
            }
        }

        document.Close();

        FileInfo loFile = new FileInfo(ls);
        try
        {
            loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        }
        catch { }
    }

    public void SetValuesToVariable(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFinalReportTitle, String lsFooterTxt, String fsGAorTIAflag, String fsDiscretionaryFlg)
    {
        String lsfamilyName = fsHouseholdName;
        int liCommaCounter = lsfamilyName.IndexOf(",");
        int liSpaceCounter = lsfamilyName.LastIndexOf(" ");
        if (liCommaCounter > 0 && liSpaceCounter > 0)
            lsfamilyName = lsfamilyName.Substring(0, liCommaCounter) + " " + lsfamilyName.Substring(liSpaceCounter);
        else
            lsfamilyName = lsfamilyName;

        if (!String.IsNullOrEmpty(fsAllocationGroup))
        {
            lsfamilyName = fsAllocationGroup;
        }
        if (!String.IsNullOrEmpty(lsFinalReportTitle))
            lsfamilyName = lsFinalReportTitle;

        //Set for Pdf
        if (fsAllignment != "Horizontal")
            lsDistributionName = "Asset Distribution Comparison";
        else
            lsDistributionName = "Asset Distribution";

        lsFamiliesName = lsfamilyName;
        lsDateName = Convert.ToDateTime(fsAsofDate).ToString("MMMM dd, yyyy") + "";

        if (fsGAorTIAflag == "GA")
        {
            if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                lsGAorTIAHeader = "GRESHAM ADVISED ASSETS - DISCRETIONARY";
            else
                lsGAorTIAHeader = "GRESHAM ADVISED ASSETS";
        }
        else
        {
            if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                lsGAorTIAHeader = "TOTAL INVESTMENT ASSETS - DISCRETIONARY";
            else
                lsGAorTIAHeader = "TOTAL INVESTMENT ASSETS";
        }
    }

    public void generateCoversheetPDF(String lsDateString, String fsFinalLocation, String fsAllocationGroup, String fsHouseholdName, String fsContactId, DataTable foTable, String fsKeyContactID, String fsHouseHoldTitle, String fsContactFullname, String fsDisplayContactName, String lsFinalReportTitle, String lsCoverSheetPageTitle, String fsGAorTIAflag, String fsDiscretionaryFlg, String TempFolderPath)
    {
        int TotalReportCount = foTable.Rows.Count;
        int UpperspaceCount = 0;
        int RptTitleCount = 0;
        int MainTitleLengthCount = 0;

        String lsfamilyName = fsHouseholdName;
        int liCommaCounter = lsfamilyName.IndexOf(",");
        int liSpaceCounter = lsfamilyName.LastIndexOf(" ");
        if (liCommaCounter > 0 && liSpaceCounter > 0)
            lsfamilyName = lsfamilyName.Substring(0, liCommaCounter) + " " + lsfamilyName.Substring(liSpaceCounter);
        else
            lsfamilyName = lsfamilyName;

        if (!String.IsNullOrEmpty(fsAllocationGroup))
        {
            lsfamilyName = fsAllocationGroup;
        }

        lsfamilyName = "";

        if (fsKeyContactID == fsContactId)
        {
            //lsfamilyName = fsHouseHoldTitle;
            //if (!String.IsNullOrEmpty(fsAllocationGroup))
            //    lsfamilyName = fsAllocationGroup;
            if (!String.IsNullOrEmpty(lsFinalReportTitle))
                lsfamilyName = lsFinalReportTitle;
        }
        else
        {
            lsfamilyName = "Reports for " + fsDisplayContactName;
        }

        //if (!String.IsNullOrEmpty(lsFinalReportTitle))
        //    lsfamilyName = lsFinalReportTitle;

        if (lsCoverSheetPageTitle != "")
        {
            lsfamilyName = lsCoverSheetPageTitle;
        }

        MainTitleLengthCount = lsfamilyName.Length;


        if (TotalReportCount > 0 && TotalReportCount < 6)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
            else
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }

        }
        else if (TotalReportCount >= 6 && TotalReportCount < 9)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
            else
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
        }
        else if (TotalReportCount >= 9 && TotalReportCount < 11)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 5;
                RptTitleCount = 10;
            }
            else
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
        }
        else if (TotalReportCount >= 11 && TotalReportCount < 13)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
            else
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
        }
        else if (TotalReportCount >= 11 && TotalReportCount < 18)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
            else
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
        }
        else
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
            else
            {
                UpperspaceCount = 10;
                RptTitleCount = 10;
            }
        }

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 80, 80, 31, 5);
        // String ls = Server.MapPath("") + "/a" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
        String ls = TempFolderPath + "/a" + Guid.NewGuid().ToString() + ".pdf";

        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        String lsDateTime = DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
        document.Open();
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(2);
        loTable.Width = 100;
        int[] headerwidths = { 39, 45 }; //{ 47, 35 }
        loTable.SetWidths(headerwidths);
        loTable.Border = 0;

        iTextSharp.text.Cell loCell = new Cell();
        Chunk loChunk = new Chunk();
        for (int liCounter = 0; liCounter < 13; liCounter++)//13//7
        {
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.Colspan = 2;
            loCell.HorizontalAlignment = 1;
            loCell.Border = 0;
            loTable.AddCell(loCell);

        }

        loCell = new Cell();
        loChunk = new Chunk(lsfamilyName, setFontsAll(26, 0, 0));//setFontsAll(26, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 1;
        if (MainTitleLengthCount >= 54)
        {
            loCell.Leading = 25f;
        }
        loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(Convert.ToDateTime(lsDateString).ToString("MMMM dd, yyyy") + "", setFontsAll(12, 0, 1));
        loCell.Add(loChunk);
        loCell.Leading = 25f;
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 1;
        loTable.AddCell(loCell);


        for (int liCounter = 0; liCounter < 2; liCounter++)//4
        {
            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.Colspan = 2;
            loCell.HorizontalAlignment = 1;
            loCell.Border = 0;
            loTable.AddCell(loCell);
        }

        int rowcount = foTable.Rows.Count;
        int rowdiff = 0;
        int j = 0;
        //for (int liCounter = 0; liCounter < RptTitleCount; liCounter++)
        //{
        //    rowdiff = RptTitleCount - rowcount;
        //    if (liCounter >= rowdiff)
        //    {
        //        if (fsContactId == Convert.ToString(foTable.Rows[j]["ssi_ContactID"]).Replace(",", ""))
        //        {
        //            loCell = new Cell();
        //            loChunk = new Chunk("dev", Font8Whitecheck("test"));
        //            loCell.Add(loChunk);
        //            loCell.Colspan = 0;
        //            loCell.HorizontalAlignment = 0;
        //            loCell.Leading = 0.3f;//0.7f
        //            loCell.Border = 1;
        //            loTable.AddCell(loCell);

        //            loCell = new Cell();
        //            String lsAllocationGroupNEW = Convert.ToString(foTable.Rows[j]["Ssi_AllocationGroup"]);

        //            String lsFinalTitleAfterChange = String.Empty;
        //            if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["HouseHoldReportTitle"])))
        //                lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["HouseHoldReportTitle"]);

        //            if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["AllocationGroupReportTitle"])))
        //                lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["AllocationGroupReportTitle"]);

        //            fsGAorTIAflag = Convert.ToString(foTable.Rows[j]["ssi_gaortia"]);
        //            fsDiscretionaryFlg = Convert.ToString(foTable.Rows[j]["Discretionary Flag"]);

        //            String ReportName = Convert.ToString(foTable.Rows[j]["ssi_GreshamReportIdName"]);
        //            if (ReportName == "Client Goals" || ReportName == "Absolute Returns" || ReportName == "Capital Protection")
        //            {
        //                if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["Ssi_HouseholdIdName"])))
        //                {
        //                    lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["Ssi_HouseholdIdName"]);
        //                }
        //            }

        //            if (fsGAorTIAflag == "GA")
        //            {
        //                if (fsDiscretionaryFlg.ToUpper() == "TRUE")
        //                    loChunk = new Chunk("GA " + Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
        //                else
        //                    loChunk = new Chunk("GA " + Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
        //            }
        //            else
        //            {
        //                if (fsDiscretionaryFlg.ToUpper() == "TRUE")
        //                    loChunk = new Chunk(Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));//setFontsAll(10, 0, 1));
        //                else
        //                    loChunk = new Chunk(Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));//setFontsAll(10, 0, 1));
        //            }
        //            loChunk = new Chunk("dev", Font8Whitecheck("test"));
        //            loCell.Add(loChunk);
        //            loCell.Colspan = 1;
        //            loCell.Border = 0;
        //            loCell.Width = 45;//20                    
        //            loCell.HorizontalAlignment = 0;
        //            loTable.AddCell(loCell);
        //            j++;
        //        }
        //    }
        //    else
        //    {
        //        if (liCounter == rowdiff - 1)
        //        {
        //            loCell = new Cell();
        //            loChunk = new Chunk("dev", Font8Whitecheck("test"));
        //            loCell.Add(loChunk);
        //            loCell.Colspan = 0;
        //            loCell.Leading = 1f;
        //            loCell.HorizontalAlignment = 0;
        //            loCell.Border = 1;
        //            loTable.AddCell(loCell);

        //            loCell = new Cell();
        //            // loChunk = new Chunk("Reports included:", setFontsAll(10, 0, 1));
        //            loChunk = new Chunk("dev", Font8Whitecheck("test"));
        //            loCell.Add(loChunk);
        //            loCell.Colspan = 1;
        //            loCell.Border = 0;
        //            loCell.HorizontalAlignment = 0;
        //            loTable.AddCell(loCell);
        //        }
        //        else
        //        {
        //            loCell = new Cell();
        //            loChunk = new Chunk("dev", Font8Whitecheck("test"));
        //            loCell.Add(loChunk);
        //            loCell.Colspan = 2;
        //            loCell.HorizontalAlignment = 1;
        //            loCell.Border = 0;
        //            loTable.AddCell(loCell);
        //        }
        //    }

        //}

        for (int liCounter1 = 0; liCounter1 < 14; liCounter1++)
        {
            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.Colspan = 2;
            loCell.HorizontalAlignment = 1;
            loCell.Border = 0;
            loTable.AddCell(loCell);

        }


        loCell = new Cell();
        //loChunk = new Chunk("The values shown for the current period and the prior period are subject to the availability of information. In particular, certain non-marketable investments such as commercial real estate and private equity holdings do not provide frequent valuations. In these and other cases, we have either carried the investments at cost or used the general partner's most recent valuation estimates adjusted for subsequent investments or distributions. \"Prior Period Net Worth\" includes the most recent manager provided updated balances, some of which may remain estimated values.", setFontsAll(8, 0, 1, new iTextSharp.text.Color(150, 150, 150)));
        loChunk = new Chunk("The values shown for the current period and the prior period are subject to the availability of information. In particular, certain non-marketable investments such as commercial real estate and private equity holdings do not provide frequent valuations. In these and other cases, we have either carried the investments at cost or used the general partner's most recent quarterly valuation estimates adjusted for subsequent investments or distributions.", setFontsAll(8, 0, 1, new iTextSharp.text.Color(150, 150, 150)));
        loCell.Add(loChunk);
        loCell.Leading = 9f;
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loTable.AddCell(loCell);
        int liFindRow = foTable.Rows.Count * 2;
        //for (int liCounterww = 0; liCounterww < 19 - liFindRow; liCounterww++)
        for (int liCounterww = 0; liCounterww < 3; liCounterww++)
        {
            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.Colspan = 2;
            loCell.HorizontalAlignment = 0;
            loCell.Leading = 5f;
            loCell.Border = 0;
            loTable.AddCell(loCell);
        }

        loCell = new Cell();
        loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        loCell.Add(loChunk);
        loCell.BorderWidth = 0;
        loCell.Colspan = 2;
        loCell.HorizontalAlignment = 2;
        // loTable.AddCell(loCell); //Commented -- FooterLogic

        document.Add(loTable);

        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(@"C:\AdventReport\images\Gresham_Logo.png"); //(Server.MapPath("") + @"\images\Gresham_Logo.png");
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        document.Add(png);
        document.Close();
        try
        {
            FileInfo loFile = new FileInfo(ls);
            loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        }
        catch { }
    }

    public void generatesCoverExcel(String fsAsofDate, String fsHouseholdName, String fsAllocationGroup, String fsFinalLocation, String fsContactID, DataTable foTable, String fsKeyContactID, String fsHouseHoldTitle, String fsContactFullname, String fsDisplayContactName, String lsFinalReportTitle, String lsCoverSheetPageTitle, String TempFolderPath)
    {
        Random rand = new Random();
        rand.Next();

        //   String lsFileNamforFinalXls = Convert.ToString(rand.Next()) + System.DateTime.Now.ToString("MMddyyHHmmss") + ".xls";
        String lsFileNamforFinalXls = Guid.NewGuid().ToString() + ".xls";
        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\coversheet.xls");
        //string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls);
        //string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls.Replace("xls", "xml"));
        string strDirectory = TempFolderPath + "\\" + lsFileNamforFinalXls;
        string strDirectory2 = TempFolderPath + "\\" + lsFileNamforFinalXls.Replace("xls", "xml");

        FileInfo loFile = new FileInfo(strDirectory1);
        loFile.CopyTo(strDirectory, true);

        #region StyleUsing Spire.xls
        Workbook workbook = new Workbook();
        workbook.LoadFromFile(strDirectory);

        //Gets first worksheet
        Worksheet sheetCover = workbook.Worksheets[0];

        String lsfamilyName = fsHouseholdName;
        int liCommaCounter = lsfamilyName.IndexOf(",");
        int liSpaceCounter = lsfamilyName.LastIndexOf(" ");
        if (liCommaCounter > 0 && liSpaceCounter > 0)
            lsfamilyName = lsfamilyName.Substring(0, liCommaCounter) + " " + lsfamilyName.Substring(liSpaceCounter);
        else
            lsfamilyName = lsfamilyName;
        if (!String.IsNullOrEmpty(fsAllocationGroup))
        {
            lsfamilyName = fsAllocationGroup;
        }


        //Set for Pdf

        lsfamilyName = "";
        if (fsKeyContactID == fsContactID)
        {
            //lsfamilyName = fsHouseHoldTitle;
            //if (!String.IsNullOrEmpty(fsAllocationGroup))
            //    lsfamilyName = fsAllocationGroup;
            if (!String.IsNullOrEmpty(lsFinalReportTitle))
                lsfamilyName = lsFinalReportTitle;
        }
        else
        {
            lsfamilyName = "Reports for " + fsDisplayContactName;
        }

        //if (!String.IsNullOrEmpty(lsFinalReportTitle))
        //    lsfamilyName = lsFinalReportTitle;

        if (lsCoverSheetPageTitle != "")
        {
            lsfamilyName = lsCoverSheetPageTitle;
        }

        sheetCover.Range["A21"].Text = lsfamilyName;

        sheetCover.Range["A23"].Text = Convert.ToDateTime(fsAsofDate).ToString("MMMM dd, yyyy") + "";
        sheetCover.Range[1, 1, 500, 1].ColumnWidth = 23.1;
        sheetCover.Range["A21"].RowHeight = 37;

        int liK = 31;//35

        for (int liCounter = 0; liCounter < foTable.Rows.Count; liCounter++)
        {
            //CheckBox chkBox = (CheckBox)gvList.Rows[liCounter].FindControl("chkbSelectBatch");

            //if (chkBox.Checked && fsContactID == Convert.ToString(foTable.Rows[liCounter]["ssi_ContactID"]).Replace(",", ""))
            if (fsContactID == Convert.ToString(foTable.Rows[liCounter]["ssi_ContactID"]).Replace(",", ""))
            {

                String lsShhetNumber = "K" + liK;
                String lsAllocationGroupNEW = Convert.ToString(foTable.Rows[liCounter]["Ssi_AllocationGroup"]);

                /*if (!String.IsNullOrEmpty(lsAllocationGroupNEW))
                {
                    sheetCover.Range[lsShhetNumber].Text = Convert.ToString(foTable.Rows[liCounter]["ssi_greshamreportidname"]) + ": " + lsAllocationGroupNEW;

                }
                else
                {
                    sheetCover.Range[lsShhetNumber].Text = Convert.ToString(foTable.Rows[liCounter]["ssi_greshamreportidname"]) + ": " + Convert.ToString(foTable.Rows[liCounter]["ssi_householdreporttitle"]);


                }*/
                //	sheetCover.Range[lsShhetNumber].Text = Convert.ToString(foTable.Rows[liCounter]["ssi_greshamreportidname"]) + ": " + lsFamiliesName;

                String lsFinalTitleAfterChange = String.Empty;
                if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[liCounter]["HouseHoldReportTitle"])))
                    lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[liCounter]["HouseHoldReportTitle"]);

                if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[liCounter]["AllocationGroupReportTitle"])))
                    lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[liCounter]["AllocationGroupReportTitle"]);

                sheetCover.Range[lsShhetNumber].Text = Convert.ToString(foTable.Rows[liCounter]["ssi_greshamreportidname"]) + ": " + lsFinalTitleAfterChange;
                sheetCover.Range[lsShhetNumber].RowHeight = 15;

                liK = liK + 1;
            }


        }
        workbook.SaveAsXml(strDirectory2);
        workbook = null;
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.Load(strDirectory2);
        XmlElement businessEntities = xmlDoc.DocumentElement;
        XmlNode loNode = businessEntities.LastChild;
        businessEntities.RemoveChild(loNode);
        foreach (XmlNode lxNode in businessEntities)
        {
            if (lxNode.Name == "ss:Worksheet")
            {
                foreach (XmlNode lxPagingNode in lxNode.ChildNodes)
                {
                    if (lxPagingNode.Name == "x:WorksheetOptions")
                    {
                        foreach (XmlNode lxPagingSetup in lxPagingNode.ChildNodes)
                        {
                            if (lxPagingSetup.Name == "x:PageSetup")
                            {
                                //  lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&C&0022Frutiger 55 Roman,Regular0022&8 Page &P of &N &R&0022Frutiger 55 Roman,italic0022&8  &KD8D8D8&D, &T";
                                try
                                {
                                    if (!lxNode.Attributes[0].InnerText.ToLower().Contains("cover"))
                                        lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&C&\"Frutiger 55 Roman,Regular\"&8Page &P of &N&R&\"Frutiger 55 Roman,Italic\"&8&KD8D8D8&D,&T";
                                    else
                                        lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&R&\"Frutiger 55 Roman,Italic\"&8&KD8D8D8&D,&T";

                                }
                                catch { }
                            }
                        }
                    }

                }
            }

            if (lxNode.Name == "ss:Styles")
            {
                foreach (XmlNode lxNodes in lxNode.ChildNodes)
                {
                    try
                    {

                        foreach (XmlNode lxNodess in lxNodes.ChildNodes)
                        {
                            if (lxNodess.Name == "ss:Interior")
                            {
                                if (lxNodess.Attributes[0].InnerText == "#33CCCC")
                                    lxNodess.Attributes[0].InnerText = "#B7DDE8";

                                if (lxNodess.Attributes[0].InnerText == "#C0C0C0")
                                    lxNodess.Attributes[0].InnerText = "#D8D8D8";

                            }
                        }

                        foreach (XmlNode lxNodess in lxNodes.ChildNodes)
                        {
                            if (lxNodess.Name == "ss:Borders")
                            {
                                foreach (XmlNode lxNodessss in lxNodess.ChildNodes)
                                {
                                    if (lxNodessss.Attributes["ss:Color"].InnerText == "#C0C0C0")
                                    {
                                        lxNodessss.Attributes["ss:Color"].InnerText = "#F2F2F2";
                                    }
                                }

                            }
                        }





                    }
                    catch
                    {
                    }
                }
            }
        }

        xmlDoc.Save(strDirectory2);
        xmlDoc = null;
        loFile = null;
        loFile = new FileInfo(strDirectory);
        loFile.Delete();
        loFile = new FileInfo(strDirectory2);
        // loFile.CopyTo(strDirectory, true);
        loFile.CopyTo(fsFinalLocation, true);
        loFile = null;
        loFile = new FileInfo(strDirectory2);
        loFile.Delete();
        #endregion




    }

    private DataTable GetDataTable(String BatchIdListTxt)
    {
        string greshamquery;
        int totalCount = 0;
        //string ReportOpFolder2 = ConfigurationManager.AppSettings.Keys[1].ToString();
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";


        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        cmd.CommandTimeout = 400;
        SqlDataAdapter dagersham = new SqlDataAdapter();
        DataSet ds_gresham = new DataSet();

        try
        {
            object PriorDate = "null"; //txtPriorDate.Text == "" ? "null" : "'" + txtPriorDate.Text + "'";
            object EndDate = "null";//txtEndDate.Text == "" ? "null" : "'" + txtEndDate.Text + "'";

            object NoComparison = "null"; //chkNoComparison.Checked == false ? 0 : 1;
            //greshamquery = "sp_s_batch @BatchIdListTxt='" + BatchIdListTxt + "',@PriorDT=" + PriorDate + ",@EndDT=" + EndDate + ",@NoComparisonLineFlg=" + Convert.ToBoolean(chkNoComparison.Checked);
            greshamquery = "SP_S_BATCH @BatchIdListTxt='" + BatchIdListTxt + "',@PriorDT=" + PriorDate + ",@EndDT=" + EndDate + ",@NoComparisonLineFlg=" + NoComparison;

            dagersham = new SqlDataAdapter(greshamquery, Gresham_con);
            ds_gresham = new DataSet();
            dagersham.Fill(ds_gresham);
            totalCount = ds_gresham.Tables[0].Rows.Count;
            // Response.Write("Batch: " + DateTime.Now.ToString());
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            totalCount = 0;
            Response.Write("sp_s_batch sp fails error desc:" + exc.Detail.Message);
            // LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }
        catch (Exception exc)
        {
            bProceed = false;
            totalCount = 0;
            Response.Write("sp_S_batch sp fails error desc:" + exc.Message);
            //LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }

        return ds_gresham.Tables[0];
    }


    private string ConvertDocument(string strSourcePath, string strDestPath)
    {
        try
        {

            ComponentInfo.SetLicense("D7OT-O3KE-PMVU-IXWZ");
            //ComponentInfo.FreeLimitReached += (sender1, e1) => e1.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            DocumentModel document = DocumentModel.Load(strSourcePath);

            document.Save(strDestPath.Replace(".xls", ".pdf"));

            return strDestPath.Replace(".pdf", ".xls");


        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
            return "";
        }
    }

    private string ConvertSpreadsheet(string strSourcePath, string strDestPath)
    {
        try
        {

            SpreadsheetInfo.SetLicense("E43Y-7VYO-CTN8-X97J");
            // ComponentInfo.FreeLimitReached += (sender1, e1) => e1.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
            ExcelFile document = ExcelFile.Load(strSourcePath);

            document.Save(strDestPath.Replace(".xls", ".pdf"));

            return strDestPath.Replace(".pdf", ".xls");


        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
            return "";
        }
    }

    public string getFinalSp(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsGAorTIAflag, String fsDiscretionaryFlg)
    {
        String lsSQL = "";
        if (!String.IsNullOrEmpty(fsAllocationGroup))
        {
            lsSQL = "SP_R_Advent_Report_Allocation_NEW_GA @AllocationGroupNameTxt='" + fsAllocationGroup + "', ";
        }
        else
        {
            lsSQL = "SP_R_Advent_Report_Other_NEW_GA";
        }

        lsSQL = lsSQL + " @UUID = '" + System.Guid.NewGuid().ToString() + "'," +
                "@HouseholdName = '" + fsHouseholdName + "',";

        if (!String.IsNullOrEmpty(fsAsofDate))
        {
            lsSQL += "@EndAsofDate = '" + Convert.ToDateTime(fsAsofDate).ToShortDateString() + "',";
        }
        else
        {
            lsSQL += "@EndAsofDate = " + "null" + ",";
        }
        if (!String.IsNullOrEmpty(fsSPriorDate))
        {
            lsSQL += "@StartAsofDate = '" + Convert.ToDateTime(fsSPriorDate).ToShortDateString() + "',";
        }
        else
        {
            lsSQL += "@StartAsofDate = " + "null" + ",";
        }

        if (!String.IsNullOrEmpty(fsGAorTIAflag))
        {
            lsSQL += "@PositionGAFlagTxt = '" + fsGAorTIAflag + "',";
        }
        else
        {
            lsSQL += "@PositionGAFlagTxt = " + "null" + ",";
        }

        if (fsDiscretionaryFlg.ToUpper() == "TRUE")
            fsDiscretionaryFlg = "1";
        else if (fsDiscretionaryFlg.ToUpper() == "FALSE")
            fsDiscretionaryFlg = "0";
        else
            fsDiscretionaryFlg = "null";

        lsSQL += "@LookThruDetailTxt = '" + fsLookthrogh.Replace("'", "''") + "'," +
                 "@ContactFullNameTxt = '" + fsContactFullname.Replace("'", "''") + "'," +
                 "@VersionTxt = '" + fsVersion.Replace("'", "''") + "'," +
                 "@summaryflgtxt = '" + fsSummaryFlag + "'," +
                 "@ReportType = '" + fsAllignment + "'," +
                 "@ReportGroupFlg = " + fsReportGroupflag +
                 ",@Report2GroupFlg = " + fsReportgroupflag2 +
                 ",@DiscretionaryFlg = " + fsDiscretionaryFlg;

        //if (chkNoComparison.Checked)
        //    lsSQL = lsSQL + ",@ComparisonFlg = 1";

        //  Response.Write("<br><br><br>" + lsSQL + "<br><br><br>");
        return lsSQL;
    }

    public void setTableProperty(iTextSharp.text.Table fotable)
    {
        //int[] headerwidths = { 28, 9, 9, 9, 9, 9, 9, 9, 7 };

        setWidthsoftable(fotable);

        //fotable.Width = 100;
        fotable.Alignment = 1;
        fotable.Border = 0;
        fotable.Cellspacing = 0;
        fotable.Cellpadding = 3;
        fotable.Locked = false;

    }
    public void setWidthsoftable(iTextSharp.text.Table fotable)
    {

        switch (lsTotalNumberofColumns)
        {
            case "2":
                int[] headerwidths2 = { 30, 9 };
                fotable.SetWidths(headerwidths2);
                fotable.Width = 40;
                break;
            case "3":
                int[] headerwidths3 = { 30, 9, 9 };
                fotable.SetWidths(headerwidths3);
                fotable.Width = 49;
                break;
            case "4":
                int[] headerwidths4 = { 30, 9, 9, 9 };
                fotable.SetWidths(headerwidths4);
                fotable.Width = 58;
                break;
            case "5":
                int[] headerwidths5 = { 30, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths5);
                fotable.Width = 67;
                break;
            case "6":
                int[] headerwidths6 = { 30, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths6);
                fotable.Width = 76;
                break;
            case "7":
                int[] headerwidths7 = { 30, 9, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths7);
                fotable.Width = 85;
                break;
            case "8":
                int[] headerwidths8 = { 30, 9, 9, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths8);
                fotable.Width = 94;
                break;
            case "9":
                int[] headerwidths9 = { 25, 9, 9, 9, 9, 9, 9, 9, 9 };
                fotable.SetWidths(headerwidths9);
                fotable.Width = 97;
                break;

            case "10":
                int[] headerwidths10 = { 25, 8, 8, 8, 8, 8, 8, 8, 8, 8 };
                fotable.SetWidths(headerwidths10);
                fotable.Width = 97; break;
            case "11":
                //int[] headerwidths11 = { 25, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7 };
                int[] headerwidths11 = { 25, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8 };
                fotable.SetWidths(headerwidths11);
                fotable.Width = 95; break;
            case "12":
                int[] headerwidths12 = { 25, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7 };
                fotable.SetWidths(headerwidths12);
                fotable.Width = 102; break;
            case "13":
                int[] headerwidths13 = { 30, 9 };
                fotable.SetWidths(headerwidths13);
                fotable.Width = 39; break;
            case "14":
                int[] headerwidths14 = { 30, 9 };
                fotable.SetWidths(headerwidths14);
                fotable.Width = 39;
                break;
            case "15":
                int[] headerwidths15 = { 30, 9 };
                fotable.SetWidths(headerwidths15);
                fotable.Width = 39;
                break;
            case "16":
                int[] headerwidths16 = { 30, 9 };
                fotable.SetWidths(headerwidths16);
                fotable.Width = 39;
                break;
            case "17":
                int[] headerwidths17 = { 30, 9 };
                fotable.SetWidths(headerwidths17);
                fotable.Width = 39;
                break;
            case "18":
                int[] headerwidths18 = { 30, 9 };
                fotable.SetWidths(headerwidths18);
                fotable.Width = 39;
                break;
            case "19":
                int[] headerwidths19 = { 30, 9 };
                fotable.SetWidths(headerwidths19);
                fotable.Width = 39;
                break;
            case "20":
                int[] headerwidths20 = { 30, 9 };
                fotable.SetWidths(headerwidths20);
                fotable.Width = 39;
                break;

        }
    }
    public Boolean checkTrue(DataSet foDataset, int fiRowCount, String fsField)
    {
        Boolean lblReturn = false;
        if (foDataset.Tables[0].Rows[fiRowCount][fsField].ToString().ToUpper() == "TRUE")
        {
            lblReturn = true;
        }
        return lblReturn;

    }
    public iTextSharp.text.Font Font9Normal()
    {
        return setFontsAll(9, 0, 0);
    }
    public iTextSharp.text.Font Font1Normal()
    {
        return setFontsAll(1, 0, 0);
    }
    public iTextSharp.text.Font Font8Normal()
    {
        return setFontsAll(8, 0, 0);
    }

    public iTextSharp.text.Font Font7Normal()
    {
        return setFontsAll(7, 0, 0);
    }

    public iTextSharp.text.Font Font8GreyItalic()
    {
        return setFontsAll(8, 0, 1, new iTextSharp.text.Color(216, 216, 216));
    }

    public iTextSharp.text.Font Font7GreyItalic()
    {
        return setFontsAll(7, 0, 1, new iTextSharp.text.Color(216, 216, 216));
    }
    public iTextSharp.text.Font Font8Grey()
    {
        return setFontsAll(8, 0, 0, new iTextSharp.text.Color(175, 175, 175));
        //return setFontsAll(9, 0, 0, new iTextSharp.text.Color(175, 175, 175));
    }

    public iTextSharp.text.Font Font7Grey()
    {
        //return setFontsAll(7, 0, 0, new iTextSharp.text.Color(175, 175, 175));
        //return setFontsAll(7, 0, 0, new iTextSharp.text.Color(165, 165, 165));
        return setFontsAll(7, 0, 0, new iTextSharp.text.Color(0, 102, 153));
    }

    public iTextSharp.text.Font Font8Whitecheck(String fsTest)
    {
        if (fsTest == "test")
            return setFontsAll(8, 0, 0, new iTextSharp.text.Color(255, 255, 255));
        else
            return setFontsAll(8, 0, 0);
    }

    public iTextSharp.text.Font Font7Whitecheck(String fsTest)
    {
        if (fsTest == "test")
            return setFontsAll(7, 0, 0, new iTextSharp.text.Color(255, 255, 255));
        else
            return setFontsAll(7, 0, 0);
    }

    public iTextSharp.text.Font Font9Whitecheck(String fsTest)
    {
        if (fsTest == "test")
            return setFontsAll(9, 0, 0, new iTextSharp.text.Color(255, 255, 255));
        else
            return setFontsAll(9, 0, 0);
    }
    public iTextSharp.text.Font Font9Bold()
    {
        return setFontsAll(9, 1, 0);
    }

    public iTextSharp.text.Font Font8Bold()
    {
        return setFontsAll(8, 1, 0);
    }

    public iTextSharp.text.Font Font7Bold()
    {
        return setFontsAll(7, 1, 0);
    }

    public void checkTrue(DataSet foDataset, int fiRowCount, String fsField, Cell foCell, iTextSharp.text.Color foColor)
    {

        if (foDataset.Tables[0].Rows[fiRowCount][fsField].ToString().ToUpper() == "TRUE")
        {
            foCell.BackgroundColor = foColor;
        }


    }
    public iTextSharp.text.Table addFooter(String lsDateTime, int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt)
    {

        iTextSharp.text.Table fotable = new iTextSharp.text.Table(2, 1);
        fotable.Width = 90;
        fotable.Border = 0;
        int[] headerwidths = { 50, 40 };
        fotable.SetWidths(headerwidths);
        fotable.Cellpadding = 0;
        Cell loCell = new Cell();
        Chunk loChunk = new Chunk();

        for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData; liCounter++)
        {
            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.HorizontalAlignment = 2;
            loCell.BorderWidth = 0;
            loCell.Add(loChunk);
            fotable.AddCell(loCell);

            loCell = new Cell();
            loChunk = new Chunk("dev", Font8Whitecheck("test"));
            loCell.Add(loChunk);
            loCell.BorderWidth = 0;
            loCell.HorizontalAlignment = 2;
            fotable.AddCell(loCell);
        }
        if (footerflg)
        {
            loCell = new Cell();
            //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
            loChunk = new Chunk(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
            //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
            loCell.Leading = 8f;
            loCell.HorizontalAlignment = 0;
            loCell.Colspan = 2;
            loCell.BorderWidth = 0;
            loCell.Add(loChunk);
            fotable.AddCell(loCell);
        }


        loCell = new Cell();
        //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
        loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font7Normal());
        loCell.Leading = 15f;//25f
        loCell.HorizontalAlignment = 2;
        loCell.BorderWidth = 0;
        loCell.Add(loChunk);
        fotable.AddCell(loCell);

        loCell = new Cell();
        //loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        loChunk = new Chunk(lsDateTime, Font7GreyItalic());
        loCell.Add(loChunk);
        loCell.Leading = 15f;//25f
        loCell.BorderWidth = 0;
        loCell.HorizontalAlignment = 2;
        fotable.AddCell(loCell);
        //fotable.TableFitsPage = true;

        return fotable;
    }

    public void setHeader(Document foDocument, DataSet loInsertdataset)
    {
        iTextSharp.text.Table loTable = new iTextSharp.text.Table(loInsertdataset.Tables[0].Columns.Count, 4);   // 2 rows, 2 columns        
        setTableProperty(loTable);
        Chunk loParagraph = new Chunk();


        //     Chunk lochunk = new Chunk(lsFamiliesName, iTextSharp.text.FontFactory.GetFont("frutigerce-roman", BaseFont.CP1252, BaseFont.EMBEDDED, 14, iTextSharp.text.Font.BOLD));
        Chunk lochunk = new Chunk(lsFamiliesName, setFontsAll(14, 1, 0));
        // loParagraph.Chunks.Add(lochunk);
        iTextSharp.text.Cell loCell = new Cell();
        loCell.Add(lochunk);
        loCell.Colspan = loInsertdataset.Tables[0].Columns.Count;
        loCell.HorizontalAlignment = 1;


        lochunk = new Chunk("\n" + lsGAorTIAHeader, setFontsAll(10, 0, 0));
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + lsDistributionName, setFontsAll(12, 1, 0, new iTextSharp.text.Color(System.Drawing.ColorTranslator.FromHtml("#31869B"))));
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
        loCell.Add(lochunk);
        loCell.Border = 0;
        //   loCell.Add(loParagraph);
        loCell.Leading = 13F;
        loTable.AddCell(loCell);



        Boolean lbCheckFoMarket = false;
        for (int liColumnCount = 0; liColumnCount < loInsertdataset.Tables[0].Columns.Count; liColumnCount++)
        {
            if (liColumnCount == 0)
            {
                //changed on 02/25/2011
                //lochunk = new Chunk("", setFontsAll(9, 1, 0));
                lochunk = new Chunk("", setFontsAll(7, 1, 0));
            }
            else
            {
                //changed on 02/25/2011
                lochunk = new Chunk(Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Replace(" Market Value", ""), setFontsAll(7, 1, 0));
                //lochunk = new Chunk(Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Replace(" Market Value", ""), setFontsAll(9, 1, 0));
                if (loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName.Contains(" Market Value"))
                    lbCheckFoMarket = true;

            }
            loCell = new Cell();

            loCell.Add(lochunk);
            loCell.Border = 0;
            loCell.NoWrap = true;//true;

            if (liColumnCount != 0)
            {
                loCell.HorizontalAlignment = 2;
            }
            if (Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Contains(" "))
            {
                loCell.Leading = 10f;//8
                loCell.MaxLines = 5;
                //loCell.Leading = 9f;
            }
            loCell.Leading = 10f;//8
            loCell.VerticalAlignment = 6;//5 ,6 bottom : WASTE VALUES - 3,4
            loTable.AddCell(loCell);

        }


        //loCell = new Cell("");
        //lochunk = new Chunk("Market Value", FontFactory.GetFont(lsStringName, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 9, Font.BOLD));
        if (lbCheckFoMarket)
        {
            for (int liColumnCount = 0; liColumnCount < loInsertdataset.Tables[0].Columns.Count; liColumnCount++)
            {
                //Response.Write("<br>"+liColumnCount + "<br>");
                loCell.Border = 0;
                loCell.NoWrap = true;

                loCell = new Cell();
                if (liColumnCount != 0)
                {
                    loCell.HorizontalAlignment = 2;
                }
                if (Convert.ToString(loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName).Contains(" "))
                {
                    loCell.NoWrap = false;
                }
                if (loInsertdataset.Tables[0].Columns[liColumnCount].ColumnName.Contains(" Market Value"))
                {
                    // Response.Write("<br>" + liColumnCount + " In<br>");
                    //changed on 02/25/2011
                    //lochunk = new Chunk("Market Value", setFontsAll(9, 1, 0));
                    lochunk = new Chunk("Market Value", setFontsAll(7, 1, 0));
                }
                else
                {
                    //Response.Write("<br>" + liColumnCount + " Out<br>");
                    //changed on 02/25/2011
                    //lochunk = new Chunk("", setFontsAll(9, 1, 0));
                    lochunk = new Chunk("", setFontsAll(7, 1, 0));

                }
                loCell.Add(lochunk);
                loCell.Border = 0;
                loCell.NoWrap = true;
                loCell.Leading = 6f;
                loTable.AddCell(loCell);
            }
        }

        //loCell = new Cell();
        //loCell.Add(lochunk);
        //loCell.Border = 0;
        //loCell.NoWrap = true;
        //loTable.AddCell(loCell);
        //loCell = new Cell("");

        //loCell.Border = 0;
        //loCell.NoWrap = true;
        //loTable.AddCell(loCell);

        foDocument.Add(loTable);
        //iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(@"C:\AdventReport\images\Gresham_Logo.png");
        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        foDocument.Add(png);
    }

    public void setGreyBorder(DataSet foDataset, Cell foCell, int fiRowCount)
    {
        try
        {
            if (checkTrue(foDataset, fiRowCount, "_Ssi_UnderlineFlg") || checkTrue(foDataset, fiRowCount, "_Ssi_BoldFlg") || checkTrue(foDataset, fiRowCount, "_Ssi_SuperBoldFlg"))
            {
                setBottomWidthWhite(foCell);
            }
            if (checkTrue(foDataset, fiRowCount + 1, "_Ssi_UnderlineFlg") || checkTrue(foDataset, fiRowCount + 1, "_Ssi_BoldFlg") || checkTrue(foDataset, fiRowCount + 1, "_Ssi_SuperBoldFlg"))
            {
                setBottomWidthWhite(foCell);
            }
            else
            {
                foCell.BorderWidthBottom = 0.1F;
                //foCell.BorderColorBottom = new iTextSharp.text.Color(242, 242, 242);
                //  foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216);// change by abhi
                foCell.BorderColorBottom = new iTextSharp.text.Color(191, 191, 191);
            }
        }
        catch { }
    }

    public void setGreyBorder(Cell foCell)
    {

        foCell.BorderWidthBottom = 0.1F;
        //foCell.BorderColorBottom = new iTextSharp.text.Color(242, 242, 242);
        //  foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216);// change by abhi
        foCell.BorderColorBottom = new iTextSharp.text.Color(191, 191, 191);
    }
    public void setBottomWidthWhite(Cell foCell)
    {
        foCell.BorderWidthBottom = 0;
        foCell.BorderColorBottom = new iTextSharp.text.Color(255, 255, 255);
    }

    public void setTopWidthBlack(Cell foCell)
    {
        foCell.BorderColor = iTextSharp.text.Color.BLACK;
        foCell.Border = iTextSharp.text.Rectangle.TOP_BORDER;
        foCell.BorderWidth = 0.1F;
    }

    public iTextSharp.text.Font setFontsAll(int size, int bold, int italic, iTextSharp.text.Color foColor)
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = Server.MapPath(".");
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD, foColor);
        }
        if (italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL, foColor);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC, foColor);
        }
        return font;
        #endregion
    }
    public iTextSharp.text.Font setFontsAll(int size, int bold, int italic)
    {
        #region WITH OLD FONTS FROM FRUTIGER
        //string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //if (bold == 1)
        //{
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        //}
        //if (italic == 1)
        //{
        //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        //}
        //if (bold == 1 && italic == 1)
        //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        //return font; 
        #endregion

        #region WITH NEW FONTS FROM FRUTIGER
        string fontpath = Server.MapPath(".");
        //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
        BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
        iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        if (bold == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
        }
        if (italic == 1)
        {
            //FTI_____.PFM
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
        }
        if (bold == 1 && italic == 1)
        {
            customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
        }
        return font;
        #endregion
    }
    protected void imgApprovedFile_Click(object sender, ImageClickEventArgs e)
    {
        GridViewRow r = (GridViewRow)((DataControlFieldCell)((ImageButton)sender).Parent).Parent;
        int rowIndex = Convert.ToInt32(r.RowIndex);

        string str18 = GeneralMethods.RemoveSpecialCharacters(GridView1.Rows[rowIndex].Cells[18].Text);

        if (GridView1.Rows[rowIndex].Cells[17].Text != "")
        {
            string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + str18);

            System.IO.File.Copy(GridView1.Rows[rowIndex].Cells[17].Text, strDirectory, true);
            //Directory.Delete(ReportOpFolder, true);

            try
            {
                //Response.Write("<script>");
                //string lsFileNamforFinal = "./ExcelTemplate/" + GridView1.Rows[rowIndex].Cells[18].Text;
                //Response.Write("window.open('ViewReport.aspx?" + GridView1.Rows[rowIndex].Cells[18].Text + "', 'mywindow')");
                //Response.Write("</script>");
                Session["id"] = str18;// GridView1.Rows[rowIndex].Cells[18].Text;

                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();
                sb.Append("\n<script type=text/javascript>\n");
                sb.Append("\nwindow.open('ViewReport.aspx?" + str18 + "', 'mywindow');");
                sb.Append("</script>");
                ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());



            }
            catch (Exception exc)
            {
                Response.Write(exc.Message);
            }
        }
    }



    protected void ddlAction_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindGridView();
    }
    protected void lstHouseHold_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();

        BindGridView();
    }
    protected void ddlView_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        BindBatchOwner(ddlBatchOwner);

        BindGridView();
    }

    private void ClearControls()
    {
        lblError.Text = "";
        lblMessage.Text = "";
    }


    protected void btnRefresh_Click(object sender, EventArgs e)
    {
        lblError.Text = "";
        lblMessage.Text = "";
        BindGridView();
    }

    public void UpdateHoldReport()
    {
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://crm01/";
        //string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        //CrmService service = null;
        IOrganizationService service = null;
        lblMessage.Text = "";
        lblError.Text = "";
        try
        {
            //service = GetCrmService(crmServerUrl, orgName);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        // service.PreAuthenticate = true;
        // service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {

            ////////////////////

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");
                string ssi_batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                string HoldReport = ddlHoldReport.SelectedValue;

                try
                {
                    if (chkSelectNC.Checked == true)
                    {
                        // UpdateHoldReport(ssi_batchid, ddlHoldReport.SelectedValue);
                        //ssi_batch objBatch = new ssi_batch();
                        Entity objBatch = new Entity("ssi_batch");

                        // objBatch.ssi_batchid = new Key();
                        // objBatch.ssi_batchid.Value = new Guid(ssi_batchid);

                        objBatch["ssi_batchid"] = new Guid(Convert.ToString(ssi_batchid));


                        if (HoldReport != "" && HoldReport != "0")
                        {
                            // objBatch.ssi_holdreport = new Picklist();
                            // objBatch.ssi_holdreport.Value = Convert.ToInt32(HoldReport);

                            objBatch["ssi_holdreport"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(HoldReport));
                        }
                        else if (HoldReport == "" || HoldReport == "0")
                        {
                            // objBatch.ssi_holdreport = new Picklist();
                            // objBatch.ssi_holdreport.IsNull = true;
                            // objBatch.ssi_holdreport.IsNullSpecified = true;

                            objBatch["ssi_holdreport"] = null;
                        }

                        service.Update(objBatch);

                    }
                }
                //catch (System.Web.Services.Protocols.SoapException exc)
                catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
                {
                    lblMessage.Text = "Submit failed, Error detail: " + exc.Detail.Message;
                }
                catch (Exception exc)
                {
                    lblMessage.Text = "Submit failed, Error detail: " + exc.Message;
                }

            }


            //////////////////////



        }
        //catch (System.Web.Services.Protocols.SoapException exc)
        catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.Message;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }
    }


    private void DeleteBatchAndMailRecords(IOrganizationService service)
    {
        string DelBatch = "";
        bool chk = false;
        foreach (GridViewRow row in GridView1.Rows)
        {
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
            DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");

            string Batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
            string MailRecordsDelete = row.Cells[27].Text.Trim().Replace("ssi_mailrecords_del", "").Replace("&nbsp;", "");
            string BatchType = row.Cells[26].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");
            string BatchStatusId = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");


            try
            {
                if (ddlAction.SelectedValue == "9")//Approve
                {
                    if (MailRecordsDelete.ToUpper() == "TRUE")
                    {
                        if (BatchType == "4" && (BatchStatusId != "8" && BatchStatusId != "9" && BatchStatusId != "4"))
                        {
                            if (Batchid != "")
                            {
                                System.Threading.Thread.Sleep(20000);

                                string sqlMailRecords = "SP_D_MailRecord_Batch @BatchIdList='" + Batchid + "'";
                                //string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords);
                                //string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords, EntityName.ssi_mailrecords, service);
                                string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords, "ssi_mailrecords", service);

                                string sqlBatch = "SP_D_Batch @BatchIdList='" + Batchid + "'";
                                //DelBatch = clsDB.DeleteRecord(sqlBatch);
                                // DelBatch = clsDB.DeleteRecord(sqlBatch, EntityName.ssi_batch, service);
                                DelBatch = clsDB.DeleteRecord(sqlBatch, "ssi_batch", service);

                                chk = true;
                            }
                        }
                    }
                }
            }
            //catch (System.Web.Services.Protocols.SoapException exc)
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> exc)
            {
                bProceed = false;
                strDescription = "Error occured, Error Detail4: " + exc.Detail.Message;
                lblError.Text = strDescription;
            }
            catch (Exception exc)
            {
                bProceed = false;
                strDescription = "Error occured, Error Detail4: " + exc.Message;
                lblError.Text = strDescription;
            }
        }


        if (ddlAction.SelectedValue == "9")//Approve
        {
            BindGridView();
            if (chk == true)
            {
                lblError.Text = lblError.Text + "<br/><br/>Batch and Mail Records Rejected Successfully.";
                //lblError.Visible = true;
            }
        }
    }


}
