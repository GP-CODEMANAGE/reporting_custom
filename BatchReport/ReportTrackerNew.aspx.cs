
using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Security.Principal;
using System.Data.SqlClient;
using System.Collections;
using System.Collections.Generic;
//using CrmSdk;
using System.IO;
using Spire.Xls;
using System.Data.Common;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using GemBox.Document;
using GemBox.Document.Tables;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Xls;
using Microsoft.SharePoint.Client;
using System.Security;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using System.Threading;
using Microsoft.IdentityModel.Claims;
using System.Net.Mail;

public partial class ReportTrackerNew : System.Web.UI.Page
{
    ClientContext context;
    String sqlstr = string.Empty;
    DB clsDB = new DB();
    GeneralMethods clsGM = new GeneralMethods();
    public StreamWriter sw = null;
    bool bProceed = true;
    string strDescription;
    bool bMarkAllRecordsSent = false;
    string greshamquery;
    int totalCount = 0;
    int successcount = 0;
    int intResult = 0;
    public String _dbErrorMsg;
    string MailRecordsIdListTxt = "";
    public string strReportFiles = string.Empty;
    public string HoldReasonValue = string.Empty;
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    public int numIndexPageCount = 1;  //Index page count -- if count of batch records is > 22 then it will come on next page 
    public int numIndexPageSize = 20;//22; // Size of index page 
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    String fsReportingName = "";
    DataTable dtMail = null;
    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName, lsGAorTIAHeader;
    // int successcount = 0;
    GeneralMethods GM = new GeneralMethods();
    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {
            Session.Remove("CurPageInBatch");
            ddlBatchType.SelectedValue = "5";
            Bindddls();
            BindGridView();
        }
        if (ddlAction.SelectedValue != "11")
        {
            tblBrowse.Style.Add("display", "none");
        }
        else if (ddlAction.SelectedValue == "11")
        {
            tblBrowse.Style.Add("display", "inline");
        }
    }

    private void BindHousehold(ListBox lstBox)
    {

        object AdvisorId = "null";// ddlAdvisor.SelectedValue == "0" || ddlAdvisor.SelectedValue == "" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        object BatchId = "null";// ddlBatchOwner.SelectedValue == "0" || ddlBatchOwner.SelectedValue == "" ? "null" : "'" + ddlBatchOwner.SelectedValue + "'";
        object AssociatedId = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        object RecipientId = "null";// ddlRecipient.SelectedValue == "0" || ddlRecipient.SelectedValue == "" ? "null" : "'" + ddlRecipient.SelectedValue + "'";
        object BatchType = "null";// ddlBatchtype.SelectedValue == "0" || ddlBatchtype.SelectedValue == "" ? "null" : ddlBatchtype.SelectedValue;

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

    private void Bindddls()
    {
        BindAssociate(ddlAssociate);
        BindHousehold(lstHouseHold);

        sqlstr = "[SP_S_MAIL_QUEUE_INTERNAL_BILLING]";
        BindDropdown(ddlInternalBillingContact, sqlstr, "FullName", "systemuserid");

        sqlstr = "[SP_S_REPORT_TRACKER_STATUS]";
        BindDropdownNew(ddlReportTracker, sqlstr, "NameTxt", "IdNmb");
        ddlReportTracker.SelectedValue = "10"; //Report Not Sent

        sqlstr = "[SP_S_SENDVIA]";
        BindDropdownNew(ddlSendVia, sqlstr, "status", "status");

        BindMailType(ddlMailtype);

    }

    public void BindAssociate(DropDownList ddl)
    {
        //object OwnerId = ddlAdvisor.SelectedValue == "0" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        ddl.Items.Clear();

        sqlstr = "SP_S_BATCH_ASSOCIATE";//SP_S_ASSOCIATE @OwnerId=" + OwnerId;////
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
    public void BindMailType(DropDownList ddl)
    {


        ddl.Items.Clear();
        sqlstr = "SP_S_MAILTYPE";
        clsGM.getListForBindDDL(ddl, sqlstr, "ssi_name", "ssi_mailid");

        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;

    }
    private void BindDropdown(DropDownList ddl, string sqlstr, string TextField, string ValueField)
    {
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();

        // ddl.Items.Clear();
        dagersham = new SqlDataAdapter(sqlstr, Gresham_con);
        ds_gresham = new DataSet();
        dagersham.Fill(ds);

        ddl.DataTextField = TextField;
        ddl.DataValueField = ValueField;

        ddl.DataSource = ds;
        ddl.DataBind();

        if (ddl.ClientID == "ddlInternalBillingContact")
        {
            ddl.Items.Insert(0, "All");
            ddl.Items[0].Value = "0";

            //ddl.Items.Insert(1, "Not Null");
            //ddl.Items[1].Value = "0";
        }
        else
        {
            ddl.Items.Insert(0, "Select");
            ddl.Items[0].Value = "0";
        }
    }

    private void BindDropdownNew(DropDownList ddl, string sqlstr, string TextField, string ValueField)
    {
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";

        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();

        ddl.Items.Clear();
        dagersham = new SqlDataAdapter(sqlstr, Gresham_con);
        ds_gresham = new DataSet();
        dagersham.Fill(ds);

        ddl.DataTextField = TextField;
        ddl.DataValueField = ValueField;

        ddl.DataSource = ds;
        ddl.DataBind();



        if (ddl.ID == "ddlReportTracker")
        {
            ddl.Items.RemoveAt(0);//Remove batch status 'Handed Off'
            ddl.Items.RemoveAt(0);// Remove batch status 'Approved'
        }

        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;

        if (ddl.ID == "ddlReportTracker")
        {
            System.Web.UI.WebControls.ListItem itm = ddl.Items.FindByText("Sent");
            ddl.Items.Remove(itm);
            ddl.Items.Insert(ddl.Items.Count, itm);
        }
    }

    private void BindSecondaryOwner()
    {
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
        SqlConnection Gresham_con = new SqlConnection(Gresham_String);
        SqlCommand cmd = new SqlCommand();
        SqlDataAdapter dagersham = new SqlDataAdapter();
        SqlDataAdapter da_CRM;
        DataSet ds_gresham = new DataSet();
        DataSet ds = new DataSet();

        string sqlstr = "[SP_S_CAUpdated_HouseHold_SecondaryOwner]";
        dagersham = new SqlDataAdapter(sqlstr, Gresham_con);
        ds_gresham = new DataSet();
        dagersham.Fill(ds);

        ddlBatchOwner.DataTextField = "Ssi_SecondaryOwnerIdName";
        ddlBatchOwner.DataValueField = "Ssi_SecondaryOwnerId";

        ddlBatchOwner.DataSource = ds;
        ddlBatchOwner.DataBind();

        ddlBatchOwner.Items.Insert(0, "All");
        ddlBatchOwner.Items[0].Value = "0";
        ddlBatchOwner.SelectedIndex = 0;

    }

    private void BindGridView()
    {
        sqlstr = GetReportData();
        DataSet loDataset = clsDB.getDataSet(sqlstr);

        GridView1.Columns[18].Visible = true;
        GridView1.Columns[19].Visible = true;
        GridView1.Columns[20].Visible = true;
        GridView1.Columns[21].Visible = true;
        GridView1.Columns[22].Visible = true;
        GridView1.Columns[24].Visible = true;
        GridView1.Columns[25].Visible = true;
        GridView1.Columns[26].Visible = true;
        GridView1.Columns[27].Visible = true;
        GridView1.Columns[28].Visible = true;
        GridView1.Columns[29].Visible = true;
        GridView1.Columns[30].Visible = true;
        GridView1.Columns[31].Visible = true;
        GridView1.Columns[32].Visible = true;
        GridView1.Columns[33].Visible = true;
        GridView1.Columns[34].Visible = true;
        GridView1.Columns[35].Visible = true;
        GridView1.Columns[36].Visible = true;
        GridView1.Columns[37].Visible = true;
        GridView1.Columns[38].Visible = true;
        GridView1.Columns[39].Visible = true;
        GridView1.Columns[40].Visible = true;
        GridView1.Columns[41].Visible = true;

        GridView1.Columns[42].Visible = true;



        GridView1.DataSource = loDataset;
        GridView1.DataBind();

        GridView1.Columns[18].Visible = false;
        GridView1.Columns[19].Visible = false;
        GridView1.Columns[20].Visible = false;
        GridView1.Columns[21].Visible = false;
        GridView1.Columns[22].Visible = false;
        GridView1.Columns[24].Visible = false;
        GridView1.Columns[25].Visible = false;
        GridView1.Columns[26].Visible = false;
        GridView1.Columns[27].Visible = false;
        GridView1.Columns[28].Visible = false;
        GridView1.Columns[29].Visible = false;
        GridView1.Columns[30].Visible = false;
        GridView1.Columns[31].Visible = false;
        GridView1.Columns[32].Visible = false;
        GridView1.Columns[33].Visible = false;
        GridView1.Columns[34].Visible = false;
        GridView1.Columns[35].Visible = false;
        GridView1.Columns[36].Visible = false;
        GridView1.Columns[37].Visible = false;
        GridView1.Columns[38].Visible = false;
        GridView1.Columns[39].Visible = false;
        GridView1.Columns[40].Visible = false;
        GridView1.Columns[41].Visible = false;

        GridView1.Columns[42].Visible = false;


        if (GridView1.Rows.Count < 1)
        {
            lblMessage.Text = "Record not found";
            lblMessage.Visible = true;
            return;
        }
        else
        {
            //lblMessage.Visible = false;
        }

    }

    private String GetReportData()
    {
        object BatchType = ddlBatchType.SelectedValue == "0" || ddlBatchType.SelectedValue == "" ? "null" : "'" + ddlBatchType.SelectedValue + "'";
        object MailType = ddlMailtype.SelectedValue == "0" || ddlMailtype.SelectedValue == "" ? "null" : "'" + ddlMailtype.SelectedValue + "'";
        object BatchOwner = ddlBatchOwner.SelectedValue == "0" ? "null" : ddlBatchOwner.SelectedValue;

        object AssociateId = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        object HouseHoldId = lstHouseHold.SelectedValue == "0" || lstHouseHold.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstHouseHold) + "'";

        object SendVia = ddlSendVia.SelectedValue == "0" || ddlSendVia.SelectedValue == "" ? "null" : "'" + ddlSendVia.SelectedValue + "'";
        object ReportTrackerStatus = ddlReportTracker.SelectedValue == "0" || ddlReportTracker.SelectedValue == "" || ddlReportTracker.SelectedValue == "999" ? "null" : ddlReportTracker.SelectedValue;

        object InternalBillingContact = ddlInternalBillingContact.SelectedValue == "0" || ddlInternalBillingContact.SelectedValue == "" ? "null" : "'" + ddlInternalBillingContact.SelectedValue + "'";

        object BillingHandedOff = ddlBillingCopyHandedOff.SelectedValue == "All" ? "null" : ddlBillingCopyHandedOff.SelectedValue;  //chkbxBillingCopyHandedOff.Checked ? 1 : 0;

        sqlstr = "exec SP_S_REPORT_TRACKER @BatchOwnerId=" + BatchOwner
                                                          + ",@AssociateId=" + AssociateId
                                                          + ",@HouseHoldIdNmbList=" + HouseHoldId

                                                          + ",@SentViaTxt=" + SendVia
                                                          + ",@TrackerStatusId=" + ReportTrackerStatus
                                                          + ",@BillingContactId=" + InternalBillingContact
                                                          + ",@BillingCopyId=" + BillingHandedOff
                                                          + ",@BatchType=" + BatchType
                                                          + ",@MailTypeId=" + MailType;
        return sqlstr;
    }

    /// <summary>
    /// Set up the CRM Service.
    /// </summary>
    /// <param name="organizationName">My Organization</param>
    /// <returns>CrmService configured with AD Authentication</returns>
    //public static CrmService GetCrmService(string crmServerUrl, string organizationName)
    //{
    //    // Get the CRM Users appointments
    //    // Setup the Authentication Token
    //    CrmAuthenticationToken token = new CrmAuthenticationToken();
    //    token.AuthenticationType = 0; // Use Active Directory authentication.
    //    token.OrganizationName = organizationName;
    //    // string username = WindowsIdentity.GetCurrent().Name;

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

    protected void ddlAssociate_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        lblError.Text = "";
        BindHousehold(lstHouseHold);

        sqlstr = "[SP_S_MAIL_QUEUE_INTERNAL_BILLING]";
        BindDropdown(ddlInternalBillingContact, sqlstr, "FullName", "systemuserid");

        sqlstr = "[SP_S_REPORT_TRACKER_STATUS]";
        BindDropdownNew(ddlReportTracker, sqlstr, "NameTxt", "IdNmb");

        sqlstr = "[SP_S_SENDVIA]";
        BindDropdownNew(ddlSendVia, sqlstr, "status", "status");

        ddlBillingCopyHandedOff.SelectedIndex = 0;

        BindGridView();

    }
    protected void ddlBatchOwner_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        lblError.Text = "";

        BindHousehold(lstHouseHold);
        BindGridView();
    }
    protected void ddlSendVia_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        lblError.Text = "";
        //BindHousehold(lstHouseHold);
        BindGridView();
    }
    protected void ddlReportTracker_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        lblError.Text = "";
        //BindHousehold(lstHouseHold);
        BindGridView();
    }
    protected void ddlInternalBillingContact_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        lblError.Text = "";
        //BindHousehold(lstHouseHold);
        BindGridView();
    }
    protected void ddlBillingCopyHandedOff_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        lblError.Text = "";
        //BindHousehold(lstHouseHold);
        BindGridView();
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

    private void HouseholdsAffected()
    {
        bool bProceed = false;
        string HouseholdName = string.Empty;
        string DistinctHouseHoldNames = string.Empty;
        string[] CheckString;

        foreach (GridViewRow row in GridView1.Rows)
        {
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
            DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");
            string HouseHold = row.Cells[4].Text.Trim().Replace("HouseHold", "").Replace("&nbsp;", "");

            if (chkSelectNC.Checked)
            {
                if (ddlHoldReport.SelectedValue != "" && ddlHoldReport.SelectedValue != "0")
                {
                    if (HouseHold != "")
                    {
                        HouseholdName = HouseholdName + "," + HouseHold;
                    }
                    else
                    {
                        HouseholdName = HouseHold;
                    }
                }
            }
        }

        HouseholdName = HouseholdName == "" ? "" : HouseholdName.Substring(1, HouseholdName.Length - 1);
        CheckString = HouseholdName.Split(',');

        CheckString = GetDistinctValues<string>(CheckString);


        for (int i = 0; i < CheckString.Length; i++)
        {
            DistinctHouseHoldNames = DistinctHouseHoldNames + "," + CheckString[i];
        }

        DistinctHouseHoldNames = DistinctHouseHoldNames.Substring(1, DistinctHouseHoldNames.Length - 1);

        if (DistinctHouseHoldNames != "")
        {
            lblMessage.Text = DistinctHouseHoldNames + " affected";
            lblMessage.Visible = true;
        }
    }

    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;
        lblMessage.Text = "";
        lblError.Text = "";
        int UniqueMailingId = 0;
        int finalReportCreatedCount = 0;
        int otherThanFinalReportCnt = 0;
        int selectedCount = 0;
        bool bUnapprove = true;
        bool bContinue = true;
        // DataTable dtMail = null;
        lblMessage.Text = "";
        lblError.Text = "";

        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
            // sw.WriteLine("step 1 ");
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
            //  sw.WriteLine(strDescription);
            //LogMessage(sw, service, strDescription, 62, "GeneralError");
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        string BatchIdListTxt = "";


        if (ddlAction.SelectedValue == "12")
        {
            bool bBatchCheck = false;
            dtMail = new DataTable();
            dtMail.Columns.Add("Batchid");
            dtMail.Columns.Add("BatchName");
            dtMail.Columns.Add("CheckBoxStatus");
            dtMail.Columns.Add("BillingInvoiceid");
            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string BatchStatus = row.Cells[26].Text.Trim().Replace("ssi_reporttrackerstatus", "").Replace("&nbsp;", "");
                string BatchType = row.Cells[39].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");
                string Batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                string BatchName = row.Cells[2].Text.Trim().Replace("Batch Name", "").Replace("&nbsp;", "");
                string BillingInvoiceid = row.Cells[42].Text.Trim().Replace("ssi_billinginvoiceid", "").Replace("&nbsp;", ""); ;

                if (chkSelectNC.Checked == true)
                {
                    DataRow dr1 = dtMail.NewRow();
                    //  dtOrder.NewRow();
                    dr1["Batchid"] = Batchid;
                    dr1["BatchName"] = BatchName;
                    dr1["CheckBoxStatus"] = chkSelectNC.Checked.ToString();
                    dr1["BillingInvoiceid"] = BillingInvoiceid;
                    dtMail.Rows.Add(dr1);
                    if (BatchType != "4")
                    {
                        bBatchCheck = true;
                    }
                }

            }

            if (bBatchCheck == true)
            {
                lblMessage.Text = "Only Merge batches can be rejected.  Please use the action “Unapprove” for Quarterly or Monthly reports";
                lblMessage.Visible = true;
                return;
            }

        }



        #region Check Batch Type

        if (ddlAction.SelectedValue == "4" || ddlAction.SelectedValue == "7" || ddlAction.SelectedValue == "8")
        {
            bool bCheck = false;

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string BatchType = row.Cells[39].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");

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
                lblMessage.Text = "This action is only allowed for quarterly batches";
                lblMessage.Visible = true;
                return;
            }

        }


        #endregion

        #region Insert CoverLetter
        if (ddlAction.SelectedValue == "13")
        {
            int count = 0;
            bool bproceed = true;
            bool bpathproceed = true;
            bool bBatchStatus = true;
            //  ssi_batch objBatch = null;
            Entity objBatch = null;

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string ssi_batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");

                string BatchFilePath = row.Cells[21].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string BatchFileName = row.Cells[22].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
                string BatchStatusID = row.Cells[26].Text.Trim().Replace("ssi_reporttrackerstatus", "").Replace("&nbsp;", "");

                // string ssi_batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                // string BatchStatusID = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                //  string BatchFilePath = row.Cells[17].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                //  string BatchFileName = row.Cells[18].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
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

                string ssi_batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");

                string BatchFilePath = row.Cells[21].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string BatchFileName = row.Cells[22].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
                string BatchStatusID = row.Cells[26].Text.Trim().Replace("ssi_reporttrackerstatus", "").Replace("&nbsp;", "");
                string Foldername = row.Cells[28].Text.Trim().Replace("FolderNameTxt", "").Replace("&nbsp;", "");

                //string ssi_batchid = row.Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                //string BatchStatusID = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                //string BatchFilePath = row.Cells[17].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                //string BatchFileName = row.Cells[18].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
                string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + BatchFileName);

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

                        GenerateReport();



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

        if (ddlAction.SelectedValue == "11")
        {
            int count = 0;
            bool bproceed = true;
            bool bpathproceed = true;
            // ssi_batch objBatch = null;
            Entity objBatch = null;
            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string ssi_batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");

                string BatchFilePath = row.Cells[21].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string BatchFileName = row.Cells[22].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
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

            }

            if (count > 1)
            {
                if (!bproceed)
                {
                    lblMessage.Text = "Merge pdf only works when you select single batch. <br/> Please select single batch.";
                    lblMessage.Visible = true;
                    return;
                }
            }


            if (!bpathproceed)
            {
                lblMessage.Text = "The current batch doesnt have pdf report to merge yet.";
                lblMessage.Visible = true;
                return;
            }



            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string ssi_batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");

                string BatchFilePath = row.Cells[21].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string BatchFileName = row.Cells[22].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
                string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + BatchFileName);

                if (chkSelectNC.Checked == true)
                {
                    if (FileUpload1.HasFile == true)
                    {
                        if (count == 1)
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
                            lblMessage.Text = "Merge Pdf Successfully.";
                            lblMessage.Visible = true;
                            return;
                        }
                    }
                    else
                    {
                        lblMessage.Text = lblMessage.Text + "<br/>" + "Please select pdf file to merge";
                        lblMessage.Visible = true;
                        return;
                    }


                }

            }

        }



        #region Check For Dat MissMatch in CHIP and GRID
        if (ddlAction.SelectedValue == "1")
        {
            bool bCheck = false;

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string BatchStatus = row.Cells[26].Text.Trim().Replace("ssi_reporttrackerstatus", "").Replace("&nbsp;", "");
                string ssi_batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");

                if (chkSelectNC.Checked == true)
                {
                    sqlstr = GetReportData();
                    DataSet DS = clsDB.getDataSet(sqlstr);

                    for (int i = 0; i < DS.Tables[0].Rows.Count; i++)
                    {
                        string strBatchStatusId = Convert.ToString(DS.Tables[0].Rows[i]["ssi_reporttrackerstatus"]);
                        string strBatchId = Convert.ToString(DS.Tables[0].Rows[i]["ssi_batchid"]);


                        if (strBatchStatusId != BatchStatus && strBatchId == ssi_batchid)
                        {
                            BindGridView();
                            bCheck = true;
                        }

                    }
                }
            }

            if (bCheck == true)
            {
                lblMessage.Text = "There was inconsistency between 'CHIP' and Data in grid below" + "<br/>" + " Data has been refreshed now please perform the action again";
                lblMessage.Visible = true;
                return;
            }

        }
        #endregion


        #region Update Hold Report Status

        UpdateHoldReport();

        //foreach (GridViewRow row in GridView1.Rows)
        //{
        //    CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
        //    DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");
        //    string ssi_batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");

        //    try
        //    {

        //            if (chkSelectNC.Checked == true)
        //            {
        //                UpdateHoldReport(ssi_batchid, ddlHoldReport.SelectedValue);
        //            }
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

                string BatchStatus = row.Cells[26].Text.Trim().Replace("ssi_reporttrackerstatus", "").Replace("&nbsp;", "");

                if (ddlAction.SelectedValue == "4" || ddlAction.SelectedValue == "12")//OPS Change Requested
                {
                    //9	FINAL Report Created   || 4	Sent
                    if (BatchStatus == "9" || BatchStatus == "4")
                    {
                        finalReportCreatedCount++;
                    }

                }

                if (ddlAction.SelectedValue == "6")
                {
                    if (BatchStatus != "9")
                    {
                        otherThanFinalReportCnt++;
                    }
                }

                if (ddlAction.SelectedValue == "7")//OPS Change Requested
                {
                    //9	FINAL Report Created   || 4	Sent
                    if (BatchStatus == "9" || BatchStatus == "4")
                    {
                        bUnapprove = false;
                    }
                }

                if (ddlAction.SelectedValue == "8")
                {
                    //9	FINAL Report Created   || 4	Sent
                    if (BatchStatus != "9" && BatchStatus != "4")
                    {
                        bContinue = false;
                    }
                }
            }
        }

        if (ddlAction.SelectedValue == "4")  ////OPS Change Requested
        {
            if (finalReportCreatedCount > 0 && Hidden1.Value != "1")
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();
                sb.Append("\n<script type=text/javascript>\n");

                sb.Append("var bt = window.document.getElementById('btnSumbitTop');\n");
                sb.Append("if(confirm('The report you selected has been finalized or sent  Do you want to continue requesting a change?'))\n{");
                sb.Append("\nwindow.document.getElementById('Hidden1').value='1';");
                sb.Append(("\nbt.click();\n"));
                sb.Append("\n}");
                sb.Append("else\n{");
                sb.Append(("\nwindow.document.getElementById('Hidden1').value='0';"));
                sb.Append("\n}");
                sb.Append("</script>");
                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                return;
            }
            else if (finalReportCreatedCount > 0 && Hidden1.Value == "1")
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


        if (ddlAction.SelectedValue == "12")  ////Reject
        {
            if (finalReportCreatedCount > 0 && Hidden1.Value != "1")
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();
                sb.Append("\n<script type=text/javascript>\n");

                sb.Append("var bt = window.document.getElementById('btnSumbitTop');\n");
                sb.Append("if(confirm('One or more of the reports you selected has already been finalized or sent. Do you want to continue rejecting the report(s)'))\n{");
                sb.Append("\nwindow.document.getElementById('Hidden1').value='1';");
                sb.Append(("\nbt.click();\n"));
                sb.Append("\n}");
                sb.Append("else\n{");
                sb.Append(("\nwindow.document.getElementById('Hidden1').value='0';"));
                sb.Append("\n}");
                sb.Append("</script>");
                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                return;
            }
            else if (finalReportCreatedCount > 0 && Hidden1.Value == "1")
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

        if (ddlAction.SelectedValue == "6") //Mark Sent
        {
            string FinalReviewList = string.Empty;
            bool checkMailPref = false;
            int intResult = 0;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            Type tp = this.GetType();
            sb.Append("\n<script type=text/javascript>\n");


            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                if (chkSelectNC.Checked == true)
                {
                    intResult++;
                    string ssi_reviewreqdbyid = row.Cells[36].Text.Trim().Replace("ssi_reviewreqdbyid", "").Replace("&nbsp;", "");
                    string ssi_reviewreqdbyidName = row.Cells[37].Text.Trim().Replace("ssi_reviewreqdbyidname", "").Replace("&nbsp;", "");
                    string MailPref = row.Cells[7].Text.Trim().Replace("Send VIA", "").Replace("&nbsp;", "");
                    string BatchName = row.Cells[2].Text.Trim().Replace("Batch Name", "").Replace("&nbsp;", "");

                    if (MailPref.ToUpper() != "EMAIL" && ssi_reviewreqdbyid != "")
                    {
                        if (FinalReviewList != "")
                        {
                            FinalReviewList = FinalReviewList + "\\r\\r\\r\\r" + "Batch Name: " + BatchName + "\\r" + "Final Reviewers Name: " + ssi_reviewreqdbyidName;
                        }
                        else
                        {
                            FinalReviewList = "Batch Name: " + BatchName + "\\r" + "Final Reviewers Name: " + ssi_reviewreqdbyidName;
                        }

                        checkMailPref = true;
                    }

                }
            }

            if (intResult > 0)
            {
                bMarkAllRecordsSent = true;
            }

            if (otherThanFinalReportCnt > 0)
            {
                lblMessage.Visible = true;
                lblMessage.Text = "One or more of your reports were not marked as sent because the final report(s) has/have not been created";

                return;
            }

            if (checkMailPref == true)
            {
                sb.Append("\n alert('" + FinalReviewList + "');");
                sb.Append("</script>");
                ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                //return;
            }
        }

        if (ddlAction.SelectedValue == "7") //Un-approve
        {
            if (!bUnapprove)
            {
                if (Hidden1.Value != "1")
                {
                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    Type tp = this.GetType();
                    sb.Append("\n<script type=text/javascript>\n");

                    sb.Append("var bt = window.document.getElementById('btnSumbitTop');\n");
                    sb.Append("if(confirm('One or more of the reports you selected has already been finalized or sent.  Do you want to continue un-approving the report(s)'))\n{");
                    sb.Append("\nwindow.document.getElementById('Hidden1').value='1';");
                    sb.Append(("\nbt.click();\n"));
                    sb.Append("\n}");
                    sb.Append("else\n{");
                    sb.Append(("\nwindow.document.getElementById('Hidden1').value='0';"));
                    sb.Append("\n}");
                    sb.Append("</script>");
                    ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                    return;
                }
                else if (Hidden1.Value == "1")
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
        }

        if (ddlAction.SelectedValue == "8")  // 8 - Send Billing Copy
        {
            if (!bContinue)
            {
                lblMessage.Text = "Billing Notification cannot be sent for reports that have not been finalized or sent";
                return;
            }
        }

        if (ddlAction.SelectedValue == "5")  // Create Final Report
        {
            string BatchType = string.Empty;
            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                BatchType = row.Cells[39].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");
                if (chkSelectNC.Checked)
                {
                    if (BatchIdListTxt == "")
                        BatchIdListTxt = batchid;
                    else
                        BatchIdListTxt = BatchIdListTxt + "," + batchid;
                }
            }

            Session["BatchIdList"] = BatchIdListTxt;

            string csname2 = "ClientScript";
            System.Text.StringBuilder cstext2 = new System.Text.StringBuilder();
            cstext2.Append("<script type=\"text/javascript\"> ");
            cstext2.Append("window.open('MailQueue.aspx?btypeid=" + BatchType + "') </");//?bidlist=" + BatchIdListTxt + "'
            cstext2.Append("script>");
            RegisterClientScriptBlock(csname2, cstext2.ToString());

            //return;
        }

        //ssi_mailrecords objMailRecords = null;
        Entity objMailRecords = null;

        foreach (GridViewRow row in GridView1.Rows)
        {
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
            //Need to ask to jeanne about updating values in dropdown 'Hold Report' on action 'OPS Change Requested'
            DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");
            string ssi_batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
            string ssi_mailrecordsid = row.Cells[19].Text.Trim().Replace("ssi_mailrecordsid", "").Replace("&nbsp;", "");
            string AdvisorApproval = row.Cells[20].Text.Trim().Replace("Advisor Approval", "").Replace("&nbsp;", "");

            string DestinationPath = row.Cells[21].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
            string ConsolidatePdfFileName = row.Cells[22].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
            string ssi_batchdate = row.Cells[24].Text.Trim().Replace("ssi_batchdate", "").Replace("&nbsp;", "");
            string ssi_secondaryownerid = row.Cells[25].Text.Trim().Replace("ssi_secondaryownerid", "").Replace("&nbsp;", "");
            string BatchStatusId = row.Cells[26].Text.Trim().Replace("ssi_reporttrackerstatus", "").Replace("&nbsp;", "");


            string BatchAsOfDate = row.Cells[29].Text.Trim().Replace("As Of Date", "").Replace("&nbsp;", "");
            string BatchOwnerId = row.Cells[30].Text.Trim().Replace("OwnerId", "").Replace("&nbsp;", "");
            string RecipientId = row.Cells[31].Text.Trim().Replace("contactid", "").Replace("&nbsp;", "");
            string AdvisorId = row.Cells[33].Text.Trim().Replace("hhownerid", "").Replace("&nbsp;", "");
            string InternalBillingContact = row.Cells[34].Text.Trim().Replace("Ssi_InternalBillingContactId", "").Replace("&nbsp;", "");
            string BillingHandedOff = row.Cells[35].Text.Trim().Replace("Ssi_BillingHandedOff", "").Replace("&nbsp;", "");
            string ssi_reviewreqdbyid = row.Cells[36].Text.Trim().Replace("ssi_reviewreqdbyid", "").Replace("&nbsp;", "");
            string MailPref = row.Cells[7].Text.Trim().Replace("Send VIA", "").Replace("&nbsp;", "");



            string MailRecordsId = row.Cells[38].Text.Trim().Replace("ssi_mailrecordsid", "").Replace("&nbsp;", "");
            string BatchType = row.Cells[39].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");
            string TypeID = row.Cells[41].Text.Trim().Replace("BatchType", "").Replace("&nbsp;", "");


            HoldReasonValue = ddlHoldReport.SelectedValue;

            try
            {
                if (chkSelectNC.Checked == true)
                {

                    //if (ddlAction.SelectedValue == "8")  // 8 - Send Billing Copy
                    //{

                    //}
                    //else
                    //{
                    BatchUpdate(ssi_batchid, ddlAction.SelectedValue, AdvisorApproval, ssi_secondaryownerid, BatchStatusId, AdvisorId, InternalBillingContact, Convert.ToBoolean(BillingHandedOff), HoldReasonValue, MailPref, ssi_reviewreqdbyid, BatchOwnerId, TypeID);

                    //if ((ddlAction.SelectedValue == "1" || ddlAction.SelectedValue == "2" || ddlAction.SelectedValue == "3") && BatchStatusId == "6")
                    if (ddlAction.SelectedValue == "1" && BatchStatusId == "6")
                    {

                        #region Insert Batch Review Details
                        sqlstr = "SP_S_REPORT_REVIEW_DETAIL @BatchId='" + ssi_batchid + "'";
                        DataSet loDataset = clsDB.getDataSet(sqlstr);

                        for (int j = 0; j < loDataset.Tables[0].Rows.Count; j++)
                        {
                            //objMailRecords = new ssi_mailrecords();
                            objMailRecords = new Entity("ssi_mailrecords");


                            //Quarterly Statement EB776A64-CDBE-E011-A19B-0019B9E7EE05
                            //objMailRecords.ssi_mailtypeid = new Lookup();
                            //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                            //objMailRecords.ssi_mailtypeid.Value = new Guid("eb776a64-cdbe-e011-a19b-0019b9e7ee05");
                            objMailRecords["ssi_mailtypeid"] = new EntityReference("ssi_mail", new Guid("eb776a64-cdbe-e011-a19b-0019b9e7ee05"));


                            //name
                            if (Convert.ToString(loDataset.Tables[0].Rows[j]["name"]) != "")
                            {
                                //objMailRecords.ssi_name = Convert.ToString(loDataset.Tables[0].Rows[j]["name"]);
                                objMailRecords["ssi_name"] = Convert.ToString(loDataset.Tables[0].Rows[j]["name"]);
                            }

                            if (ssi_batchid != "")
                            {
                                //objMailRecords.ssi_batchid = new Lookup();
                                //objMailRecords.ssi_batchid.type = EntityName.ssi_batch.ToString();
                                //objMailRecords.ssi_batchid.Value = new Guid(ssi_batchid);
                                objMailRecords["ssi_batchid"] = new EntityReference("ssi_batch", new Guid(ssi_batchid));
                            }
                            ////Mail Type
                            //if (Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_mailid"]) != "")
                            //{
                            //    objMailRecords.ssi_mailtypeid = new Lookup();
                            //    objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                            //    objMailRecords.ssi_mailtypeid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_mailid"]));
                            //}


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
                                    UniqueMailingId = Convert.ToInt32(loDataset.Tables[0].Rows[j]["Ssi_MailingID"]);

                                //objMailRecords.ssi_mailingid = new CrmNumber();
                                //objMailRecords.ssi_mailingid.Value = UniqueMailingId; //Convert.ToInt32(loDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                                objMailRecords["ssi_mailingid"] = UniqueMailingId;
                            }

                            if (Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_mailstatus"]) != "")
                            {
                                //objMailRecords.ssi_mailstatus = new Picklist();
                                //objMailRecords.ssi_mailstatus.Value = Convert.ToInt32(loDataset.Tables[0].Rows[j]["ssi_mailstatus"]);
                                objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[0].Rows[j]["ssi_mailstatus"]));
                            }

                            if (Convert.ToString(loDataset.Tables[0].Rows[j]["AsOfDate"]) != "")
                            {
                                //objMailRecords.ssi_asofdate = new CrmDateTime();
                                //objMailRecords.ssi_asofdate.Value = Convert.ToString(loDataset.Tables[0].Rows[j]["AsOfDate"]);

                                objMailRecords["ssi_asofdate"] = Convert.ToDateTime(loDataset.Tables[0].Rows[j]["AsOfDate"]);

                            }

                            //objMailRecords.ssi_batchidtxt = Convert.ToString(ssi_batchid); // BatchId Text 
                            objMailRecords["ssi_batchidtxt"] = Convert.ToString(ssi_batchid);
                            //objMailRecords.ssi_batchnametxt = row.Cells[2].Text.Trim(); // Batch Name
                            objMailRecords["ssi_batchnametxt"] = row.Cells[2].Text.Trim();



                            //HouseHold lookup
                            if (Convert.ToString(loDataset.Tables[0].Rows[j]["AccountId"]) != "")
                            {
                                //objMailRecords.ssi_accountid = new Lookup();
                                //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["AccountId"]));
                                objMailRecords["ssi_accountid"] = new EntityReference("account", new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["AccountId"])));
                            }

                            //Contact lookup
                            if (Convert.ToString(loDataset.Tables[0].Rows[j]["ContactId"]) != "")
                            {
                                //objMailRecords.ssi_contactfullnameid = new Lookup();
                                //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ContactId"]));
                                objMailRecords["ssi_contactfullnameid"] = new EntityReference("contact", new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ContactId"])));
                            }

                            ////ssi_LegalEntityId lookup
                            if (Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]) != "")
                            {
                                //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                                //objMailRecords.ssi_legalentitynameid = new Lookup();
                                //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]));
                                objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(loDataset.Tables[0].Rows[j]["ssi_LegalEntityId"])));
                            }


                            // Created By Custom Field 
                            //Rohit Pawar
                            string Userid = GetcurrentUser();

                            if (Userid != "")
                            {
                                //objMailRecords.ssi_createdbycustomid = new Lookup();
                                //objMailRecords.ssi_createdbycustomid.Value = new Guid(Userid);
                                objMailRecords["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));

                            }

                            service.Create(objMailRecords);
                            intResult++;
                        }
                        #endregion


                        #region DisAssociate Mail Records
                        sqlstr = "SP_S_CANCEL_MAIL_STATUS @MailStatusId='6,4' ,@BatchId='" + ssi_batchid + "'";
                        DataSet MailStatusDataset = clsDB.getDataSet(sqlstr);

                        for (int j = 0; j < MailStatusDataset.Tables[0].Rows.Count; j++)
                        {
                            // objMailRecords = new ssi_mailrecords();
                            objMailRecords = new Entity("ssi_mailrecords");

                            //objMailRecords.ssi_mailrecordsid = new Key();
                            //objMailRecords.ssi_mailrecordsid.Value = new Guid(Convert.ToString(MailStatusDataset.Tables[0].Rows[j]["ssi_mailrecordsid"]));
                            objMailRecords["ssi_mailrecordsid"] = new Guid(Convert.ToString(MailStatusDataset.Tables[0].Rows[j]["ssi_mailrecordsid"]));

                            if (ssi_batchid != "")
                            {
                                //objMailRecords.ssi_batchid = new Lookup();
                                //objMailRecords.ssi_batchid.IsNull = true;
                                //objMailRecords.ssi_batchid.IsNullSpecified = true;
                                objMailRecords["ssi_batchid"] = null;

                            }

                            service.Update(objMailRecords);
                            intResult++;
                        }
                        #endregion
                        // if Action is Review PDF/Batch or Request OPS Change
                        if (BatchType != "4")
                        {
                            if (DestinationPath == "" && ConsolidatePdfFileName == "")
                            {
                                GenerateReport();
                                lblMessage.Text = "Mail records inserted successfully with Report generation";
                            }
                        }

                    }
                    else if (ddlAction.SelectedValue == "4") //Batch or Request OPS Change
                    {
                        //GenerateReport();
                    }
                    else if (ddlAction.SelectedValue == "12")//Reject
                    {
                        //objMailRecords = new ssi_mailrecords();
                        objMailRecords = new Entity("ssi_mailrecords");
                        if (BatchType == "4" && (BatchStatusId != "9" && BatchStatusId != "4"))
                        {
                            if (MailRecordsId != "")
                            {
                                //objMailRecords.ssi_mailrecordsid = new Key();
                                //objMailRecords.ssi_mailrecordsid.Value = new Guid(MailRecordsId);
                                objMailRecords["ssi_mailrecordsid"] = new Guid(MailRecordsId);

                                //objMailRecords.ssi_review_reject = new CrmBoolean();
                                //objMailRecords.ssi_review_reject.Value = true;
                                objMailRecords["ssi_review_reject"] = true;

                                //objMailRecords.ssi_deleterecord_flg = new CrmBoolean();
                                //objMailRecords.ssi_deleterecord_flg.Value = true;
                                objMailRecords["ssi_deleterecord_flg"] = true;

                                //objMailRecords.ssi_initialreviewer_reject = new CrmBoolean();
                                //objMailRecords.ssi_initialreviewer_reject.Value = false;
                                objMailRecords["ssi_initialreviewer_reject"] = false;

                                string UserId = GetcurrentUser();

                                if (UserId != "")
                                {
                                    //objMailRecords.ssi_rejectedbyuserid = new Lookup();
                                    //objMailRecords.ssi_rejectedbyuserid.Value = new Guid(UserId);
                                    objMailRecords["ssi_rejectedbyuserid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));
                                }


                                service.Update(objMailRecords);
                                selectedCount++;
                            }
                        }
                    }

                    if (BatchStatusId != "6")
                    {
                        MailRecordsUpdate(ssi_batchid, ddlAction.SelectedValue, ssi_batchdate, MailPref, ssi_reviewreqdbyid);
                    }

                    //}

                    //Batch Change Log
                    //BatchChangeLog(BatchAsOfDate, Convert.ToInt32(BatchStatusId), BatchOwnerId, ssi_batchid, RecipientId);
                }
            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                lblMessage.Text = "Submit failed, Error detail: " + exc.Detail.InnerText;
            }
            catch (Exception exc)
            {
                lblMessage.Text = "Submit failed, Error detail: " + exc.Message;
            }
        }


        if (ddlAction.SelectedValue == "4")// 'OPS Change Requested' Action dropdown
        {
            GenerateReport();
            HouseholdsAffected();//Shows number of households affected.
            BindGridView();

            if (selectedCount > 1) // 
            {
                lblMessage.Visible = true;
                lblMessage.Text = lblMessage.Text + "<br/>You have requested more than one change from OPS, please see your reports in <a href=file:///S:/BATCH%20REPORTS>S:/BATCH REPORTS</a>";

                lblMessage.Text = lblMessage.Text + "<br/>" + strReportFiles;
            }
        }
        else if (ddlAction.SelectedValue == "12")
        {
            if (selectedCount > 0)
            {

                BindGridView();
                System.Threading.Thread.Sleep(15000);
                DeleteBatchAndMailRecords(service);
                //  BindGridView();
                lblMessage.Text = "Batch and Mail Records Rejected Successfully.";
                lblMessage.Visible = true;
            }
        }
        else if (ddlAction.SelectedValue == "6")
        {
            if (bMarkAllRecordsSent == true)
            {
                ActionMarkAllSent();
                lblMessage.Visible = true;
                BindGridView();
                lblMessage.Text = "Records Updated Successfully ";
            }
        }
        else
        {
            lblMessage.Visible = true;
            BindGridView();
            lblMessage.Text = "Records Updated Successfully ";
        }
    }


    public void SendEmail(string BatchName)
    {
        try
        {
            string mailmessage = string.Empty;
            MailMessage myMessage = new MailMessage();
            // SmtpClient SMTPSERVER = new SmtpClient();

            string EmailID = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
            string Password = AppLogic.GetParam(AppLogic.ConfigParam.Password);
            string SMTPHost = AppLogic.GetParam(AppLogic.ConfigParam.SMTPHost);
            string ToEmailIDs2 = AppLogic.GetParam(AppLogic.ConfigParam.ToEmailIDbillingReject);

            // string ToEmailIDs = "skane@infograte.com|jmasa@greshampartners.com|bfeeny@greshampartners.com";
            int Port = Convert.ToInt32(AppLogic.GetParam(AppLogic.ConfigParam.Port));

            myMessage.From = new MailAddress(EmailID);
            string[] strTo = ToEmailIDs2.Split('|');


            for (int i = 0; i < strTo.Length; i++)
            {
                if (strTo[i] != "")
                {
                    myMessage.To.Add(new MailAddress(strTo[i]));
                }
            }

            // myMessage.Bcc.Add("auto-emails@infograte.com");

            myMessage.CC.Add("skane@infograte.com");


            // string str = "" + LEFolderName + "";

            string Server = AppLogic.GetParam(AppLogic.ConfigParam.Server);


            if (Server.ToLower() == "test")
                myMessage.Subject = "TEST - Billing Invoice rejected for " + BatchName;
            else if (Server.ToLower() == "prod")
                myMessage.Subject = "Billing Invoice rejected for " + BatchName;

            myMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
            myMessage.Body = "Billing for: '" + BatchName + "' was rejected. Please regenerate the invoice.";


            myMessage.IsBodyHtml = true;

            SmtpClient SMTPSERVER = new SmtpClient(SMTPHost, Port);
            SMTPSERVER.DeliveryMethod = SmtpDeliveryMethod.Network;


            //SMTPSERVER.EnableSsl = false; for office 365
            SMTPSERVER.EnableSsl = true;
            // smtp.EnableSsl = true;
            SMTPSERVER.UseDefaultCredentials = true;
            System.Net.NetworkCredential basicAuthenticationInfo = new System.Net.NetworkCredential(EmailID, Password);
            SMTPSERVER.Credentials = basicAuthenticationInfo;
            SMTPSERVER.Send(myMessage);
            //lblEmail.Text = "Send";
            myMessage.Dispose();
            myMessage = null;
            SMTPSERVER = null;
            mailmessage = null;
            //  Response.Write("Email send sucessful");
        }
        catch (Exception ex)
        {
            string strDescription = "Error sending Mail :" + ex.Message.ToString();
            //commented on 12_4_2018 Jscalise nolonger in process
            lblError.Text = lblError.Text + "  ," + "Send Error" + ex.Message.ToString();
            // lblEmail.Text = "Send Error" + ex.Message.ToString();
            Response.Write("Email error" + ex.ToString());
            //LogMessage(sw, strDescription);
        }

    }

    private void ActionMarkAllSent()
    {
        int Count = 0;
        foreach (GridViewRow row in GridView1.Rows)
        {
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
            string ssi_mailrecordsId = row.Cells[19].Text.Trim().Replace("ssi_mailrecordsId", "").Replace("&nbsp;", "");
            string MailType = row.Cells[2].Text.Trim().Replace("Mail Type", "").Replace("&nbsp;", "");
            string ssi_reviewreqdbyid = row.Cells[36].Text.Trim().Replace("Ssi_ReviewReqdById", "").Replace("&nbsp;", "");

            if (chkSelectNC.Checked)
            {

                // if (MailType.ToUpper() != "Quarterly Statement".ToUpper()) //Commented --- Report Date and report by should update for all mail type.
                //{
                Count++;
                if (MailRecordsIdListTxt == "")
                    MailRecordsIdListTxt = ssi_mailrecordsId;
                else
                    MailRecordsIdListTxt = MailRecordsIdListTxt + "," + ssi_mailrecordsId;

                updateSentData(ssi_reviewreqdbyid, ssi_mailrecordsId);
                //}
            }
        }

        string[] Test = MailRecordsIdListTxt.Split(',');

        Session["MailRecordsIdList"] = MailRecordsIdListTxt;

        //if (Count > 0)
        //{
        //    string csname2 = "ClientScript";
        //    System.Text.StringBuilder cstext2 = new System.Text.StringBuilder();
        //    cstext2.Append("<script type=\"text/javascript\"> ");
        //    cstext2.Append("var myObject = window.open('SenderPopUp.aspx','win2','toolbar=0,status=no,resizable=yes,menubar=0,scrollbars=1,width=700,height=250,TOP=150,left=100');");
        //    cstext2.Append(" myObject.focus();");
        //    cstext2.Append("</script>");
        //    RegisterClientScriptBlock(csname2, cstext2.ToString());
        //}

    }

    protected void updateSentData(string updateUserId, string ssi_mailrecordsId)
    {
        int intResult = 0;
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);
        string orgName = "GreshamPartners";

        IOrganizationService service = null;

        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";

        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        // ssi_mailrecords objMailRecords = null;
        Entity objAccount = null;



        try
        {
            //if (Session["MailRecordsIdList"] != null)
            //{

            //   ViewState["MailRecordsIdListTxt"] = Convert.ToString(Session["MailRecordsIdList"]);

            string UserId = "";

            if (string.IsNullOrEmpty(updateUserId))
                UserId = GetcurrentUser();
            else
                UserId = updateUserId;

            //  string[] strMailRecordsId = Convert.ToString(ViewState["MailRecordsIdListTxt"]).Split(',');
            // strMailRecordsId = GetDistinctValues<string>(strMailRecordsId);

            //objMailRecords = new ssi_mailrecords();
            Entity objMailRecords = new Entity("ssi_mailrecords");


            //objMailRecords.ssi_mailrecordsid = new Key();
            //objMailRecords.ssi_mailrecordsid.Value = new Guid(ssi_mailrecordsId);
            objMailRecords["ssi_mailrecordsid"] = new Guid(ssi_mailrecordsId);

            //objMailRecords.ssi_sentbyid = new Lookup();
            //objMailRecords.ssi_sentbyid.Value = new Guid(UserId);
            objMailRecords["ssi_sentbyid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));

            //objMailRecords.ssi_sentdate = new CrmDateTime();
            //objMailRecords.ssi_sentdate.Value = DateTime.Today.ToString("MM/dd/yyyy");
            objMailRecords["ssi_sentdate"] = DateTime.Now;

            service.Update(objMailRecords);
            intResult++;

            //}

        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            // lblMessage.Text = "Submit failed, Error detail: " + exc.Detail.InnerText;
        }
        catch (Exception exc)
        {
            // lblMessage.Text = "Submit failed, Error detail: " + exc.Message;
        }
    }


    public void UpdateHoldReport()
    {
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;
        lblMessage.Text = "";
        lblError.Text = "";
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {

            ////////////////////

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");
                string ssi_batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                string HoldReport = ddlHoldReport.SelectedValue;

                try
                {
                    if (chkSelectNC.Checked == true)
                    {
                        // UpdateHoldReport(ssi_batchid, ddlHoldReport.SelectedValue);
                        //ssi_batch objBatch = new ssi_batch();
                        Entity objBatch = new Entity("ssi_batch");

                        //objBatch.ssi_batchid = new Key();
                        //objBatch.ssi_batchid.Value = new Guid(ssi_batchid);
                        objBatch["ssi_batchid"] = new Guid(ssi_batchid);


                        if (HoldReport != "" && HoldReport != "0")
                        {
                            //objBatch.ssi_holdreport = new Picklist();
                            //objBatch.ssi_holdreport.Value = Convert.ToInt32(HoldReport);
                            objBatch["ssi_holdreport"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(HoldReport));

                        }
                        else if (HoldReport == "" || HoldReport == "0")
                        {
                            //objBatch.ssi_holdreport = new Picklist();
                            //objBatch.ssi_holdreport.IsNull = true;
                            //objBatch.ssi_holdreport.IsNullSpecified = true;
                            objBatch["ssi_holdreport"] = null;

                        }

                        service.Update(objBatch);
                    }
                }
                catch (System.Web.Services.Protocols.SoapException exc)
                {
                    lblMessage.Text = "Submit failed, Error detail: " + exc.Detail.InnerText;
                }
                catch (Exception exc)
                {
                    lblMessage.Text = "Submit failed, Error detail: " + exc.Message;
                }

            }


            //////////////////////



        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }
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


    public void BatchUpdate(string BatchId, string Action, string AdvisorApproval, string ssi_secondaryownerid, string BatchStatusId, string AdvisorId, string InternalBillingContact, bool BillingHandedOff, string HoldReasonValu, string MailPref, string ssi_ReviewReqByid, string BatchOwnerId, string BatchTypeID)
    {
        // Define Status
        int ReportTrackerStatus = 0;
        //<asp:ListItem Value="1">OPS Approve</asp:ListItem>
        //<asp:ListItem Value="2">Associated Approved-Pend Advisor Approval</asp:ListItem>
        //<asp:ListItem Value="3">Approved-Pend OPS Approval</asp:ListItem>
        //<asp:ListItem Value="4">OPS Change Requested</asp:ListItem>
        //<asp:ListItem Value="5">Create Final Report</asp:ListItem>
        //<asp:ListItem Value="6">Mark Sent</asp:ListItem>
        //<asp:ListItem Value="7">Un-approve</asp:ListItem>
        //<asp:ListItem Value="8">Send Billing Copy</asp:ListItem>
        //<asp:ListItem Value="9">Remove Hold</asp:ListItem>
        //<asp:ListItem Value="10">Update Hold</asp:ListItem>

        //1	Associate Approved - Pend Advisor Approval
        //2	Handed Off
        //3	Approved
        //4	Sent
        //5	Approved - Pend OPS Approval
        //6	Pend Approval
        //7	OPS Change Requested
        //8	OPS Approved
        //9	FINAL Report Created

        switch (Action)
        {
            case "1":  //Approve
                if (BatchStatusId == "6")  //6	Pend Approval
                {
                    if (AdvisorApproval.ToUpper() == "TRUE")
                    {
                        ReportTrackerStatus = 1;// batch status is 'Associate Approved – Pend Advisor Approval'
                    }
                    else if (AdvisorApproval.ToUpper() == "FALSE" || AdvisorApproval == "")
                    {
                        ReportTrackerStatus = 5;// 5 - 'Approved Pend OPS Approval'
                    }
                }
                else if (BatchStatusId == "1")  //1	Associate Approved - Pend Advisor Approval
                {
                    ReportTrackerStatus = 5;//5 - 'Approved Pend OPS Approval'
                }
                else if (BatchStatusId == "5") //5 - 'Approved Pend OPS Approval'
                {
                    ReportTrackerStatus = 8;//8 - 'OPS Approved'
                    InsertIntoWireExecution(BatchId, BatchTypeID);
                }
                break;
            case "2":
                //if (AdvisorApproval.ToUpper() == "TRUE")
                //{
                //    ReportTrackerStatus = 1;// batch status is 'Associate Approved – Pend Advisor Approval'
                //}
                //else if (AdvisorApproval.ToUpper() == "FALSE" || AdvisorApproval== "")
                //{
                //    ReportTrackerStatus = 5;// batch status is 'Approved Pend OPS Approval'
                //}
                break;
            case "3":
                //ReportTrackerStatus = 5;// batch status is 'Approved Pend OPS Approval'
                break;
            case "4":  // OPS Change Requested
                ReportTrackerStatus = 7;// 7 - 'OPS Change Requested'
                break;

            case "6":  // Mark Sent
                ReportTrackerStatus = 4;// batch status is 'Sent'
                break;
            case "7": // Un-Approve
                ReportTrackerStatus = 6;// 6 - 'Pend Approval'
                break;
            case "8": // Send Billing Copy
                ReportTrackerStatus = 888; // Send Billing Copy ( Action not a status)
                break;
            case "9": //Remove Hold Action 
                ReportTrackerStatus = 999; //Remove Hold Action  ( Action not a status)
                break;
            case "10": //Update Hold Action 
                ReportTrackerStatus = 111; //Remove Hold Action  ( Action not a status)
                break;
        }

        BatchReportTrackerStatus(ReportTrackerStatus, BatchId, ssi_secondaryownerid, BatchStatusId, AdvisorApproval, AdvisorId, InternalBillingContact, Convert.ToBoolean(BillingHandedOff), HoldReasonValue, MailPref, ssi_ReviewReqByid, BatchOwnerId);

    }

    public void BatchReportTrackerStatus(int ReportTrackerStatus, string BatchId, string ssi_secondaryownerid, string BatchStatusId, string AdvisorApproval, string AdvisorId, string InternalBillingContactId, bool BillingHandedOff, string HoldReasonValue, string MailPref, string ssi_ReviewReqByid, string BatchOwnerId)
    {

        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;
        lblMessage.Text = "";
        lblError.Text = "";
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {
            //bool BillingHandedOff = false;
            //ssi_batch objBatch = new ssi_batch();
            Entity objBatch = new Entity("ssi_batch");

            //objBatch.ssi_batchid = new Key();
            //objBatch.ssi_batchid.Value = new Guid(BatchId);
            objBatch["ssi_batchid"] = new Guid(BatchId);


            // if (ReportTrackerStatus == 888 && BillingHandedOff == true)// Send Billing Copy Action ( 888 is not a status it is passed for condition check)

            if (ReportTrackerStatus == 888)
            {
                //objBatch.ssi_reporttrackerstatus = new Picklist();
                //objBatch.ssi_reporttrackerstatus.Value = Convert.ToInt32(BatchStatusId);
                objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(BatchStatusId));

                //objBatch.ssi_sendemailib = new CrmBoolean();
                //objBatch.ssi_sendemailib.Value = true;
                objBatch["ssi_sendemailib"] = true;

                //objBatch.ssi_billinghandedoff = new CrmBoolean();
                //objBatch.ssi_billinghandedoff.Value = true;
                objBatch["ssi_billinghandedoff"] = true;
                BillingHandedOff = true;
            }
            else if (ReportTrackerStatus == 999)  //Remove Hold Action ( 999 is not a status it is passed for condition check)
            {
                //objBatch.ssi_holdreport = new Picklist();
                //objBatch.ssi_holdreport.IsNull = true;
                //objBatch.ssi_holdreport.IsNullSpecified = true;
                objBatch["ssi_holdreport"] = null;
            }
            else if (ReportTrackerStatus == 111)
            {
                if (HoldReasonValue != "" && HoldReasonValue != "0")
                {
                    //objBatch.ssi_holdreport = new Picklist();
                    //objBatch.ssi_holdreport.Value = Convert.ToInt32(HoldReasonValue);
                    objBatch["ssi_holdreport"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(HoldReasonValue));
                }
                else if (HoldReasonValue == "" || HoldReasonValue == "0")
                {
                    //objBatch.ssi_holdreport = new Picklist();
                    //objBatch.ssi_holdreport.IsNull = true;
                    //objBatch.ssi_holdreport.IsNullSpecified = true;
                    objBatch["ssi_holdreport"] = null;
                }
            }
            else
            {
                //objBatch.ssi_reporttrackerstatus = new Picklist();
                //objBatch.ssi_reporttrackerstatus.Value = ReportTrackerStatus; //OPS Approve;
                objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(ReportTrackerStatus));
            }

            if (ReportTrackerStatus == 7)  // OPS Change Requested
            {
                //objBatch.ssi_batchdisplayfilename = "";
                objBatch["ssi_batchdisplayfilename"] = "";

                //objBatch.ssi_batchfilename = "";
                objBatch["ssi_batchfilename"] = "";

                if (BatchStatusId == "8" || BatchStatusId == "1" || BatchStatusId == "9" || BatchStatusId == "4" || BatchStatusId == "5" || BatchStatusId == "6")
                {
                    //objBatch.ssi_billingcontactchange = new CrmBoolean();
                    //objBatch.ssi_billingcontactchange.Value = true;//To check email notification when status is changed to 'Pend Approval'
                    objBatch["ssi_billingcontactchange"] = true;
                }
                else if (BatchStatusId == "7")//OPS Change Requested
                {
                    //objBatch.ssi_opschangecomplete = new CrmBoolean();
                    //objBatch.ssi_opschangecomplete.Value = true;//Send Email When Status is changed from 'OPS Change Requested ' to 'Pend Approval'
                    objBatch["ssi_opschangecomplete"] = true;
                }

                // When status changed to OPS changed Requested

                //objBatch.ssi_opschangemail = new CrmBoolean();
                //objBatch.ssi_opschangemail.Value = true;
                objBatch["ssi_opschangemail"] = true;

                //objBatch.ssi_opsreporttracker = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":9999/BatchReport/ReportTrackerNew.aspx";
                // objBatch["ssi_opsreporttracker"] = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":9999/BatchReport/ReportTrackerNew.aspx";	// commented 1_11_2019
                objBatch["ssi_opsreporttracker"] = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":" + AppLogic.GetParam(AppLogic.ConfigParam.CRMPortNumber) + "/BatchReport/ReportTrackerNew.aspx";// added for CRMPORT NUMBER Change 1_11_2019

            }
            else if (ReportTrackerStatus == 6) //'Pend Approval'
            {
                //objBatch.ssi_holdreport = new Picklist();
                //objBatch.ssi_holdreport.IsNull = true;
                //objBatch.ssi_holdreport.IsNullSpecified = true;
                objBatch["ssi_holdreport"] = null;

                //objBatch.ssi_batchdisplayfilename = "";
                objBatch["ssi_batchdisplayfilename"] = "";

                //objBatch.ssi_batchfilename = "";
                objBatch["ssi_batchfilename"] = "";

                //Added By Rohit Pawar
                //objBatch.ssi_billingcomplete = new CrmBoolean();
                //objBatch.ssi_billingcomplete.Value = false;
                objBatch["ssi_billingcomplete"] = false;

                if (BatchStatusId == "8" || BatchStatusId == "1" || BatchStatusId == "9" || BatchStatusId == "4" || BatchStatusId == "5")
                {
                    //objBatch.ssi_billingcontactchange = new CrmBoolean();
                    //objBatch.ssi_billingcontactchange.Value = true;//To check email notification when status is changed to 'Pend Approval'
                    objBatch["ssi_billingcontactchange"] = true;
                }
                else if (BatchStatusId == "7")//OPS Change Requested
                {
                    //below code commented on demand of jeane masa
                    //objBatch.ssi_opsbatchguid = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl) + ":9999/BatchReport/ReportReviewForm.aspx?opsbguid=" + ssi_secondaryownerid;
                    //objBatch.ssi_opsbatchguid = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":9999/BatchReport/ReportReviewForm.aspx";
                    //objBatch.ssi_opschangecomplete = new CrmBoolean();
                    //objBatch.ssi_opschangecomplete.Value = true;//Send Email When Status is changed from 'OPS Change Requested ' to 'Pend Approval'
                    objBatch["ssi_opschangecomplete"] = true;
                }
                else
                {
                    //objBatch.ssi_billinghandedoff = new CrmBoolean();
                    //objBatch.ssi_billinghandedoff.Value = false;
                    objBatch["ssi_billinghandedoff"] = false;
                    BillingHandedOff = false; ;
                }

                //SecurityPrincipal assignee = new SecurityPrincipal();
                //assignee.PrincipalId = new Guid(ssi_secondaryownerid);///HouseHold Owner ID

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
                        new Guid(ssi_secondaryownerid)),
                    Target = new EntityReference("ssi_batch",
                        new Guid(BatchId))
                };



                service.Execute(assignRequest);





            }
            else if (ReportTrackerStatus == 1) // 1 - 'Associate Approved – Pend Advisor Approval'
            {
                #region Update batchguid similar to report review form
                //added 5_7_2020 - Jeanne Demanded for this
                if (BatchStatusId == "6")
                {
                    if (BatchId != "")
                    {
                        objBatch["ssi_batchguid"] = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Remove(AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl).Length - 1) + ":" + AppLogic.GetParam(AppLogic.ConfigParam.CRMPortNumber).ToString() + "/BatchReport/ReportReviewForm.aspx";
                    }
                }
                #endregion
                ssi_secondaryownerid = GetUserID("CORP\\opsreporting");  //OPS Reporting, Gresham

                //SecurityPrincipal assignee = new SecurityPrincipal();
                //assignee.PrincipalId = new Guid(ssi_secondaryownerid);///HouseHold Owner ID

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
                        new Guid(ssi_secondaryownerid)),
                    Target = new EntityReference("ssi_batch",
                        new Guid(BatchId))
                };

                service.Execute(assignRequest);



            }

            //////////////////////// new logic //////////////////////////////////////

            if (ddlAction.SelectedValue == "7") // Un Approve
            {
                //SecurityPrincipal assignee = new SecurityPrincipal();
                //assignee.PrincipalId = new Guid(ssi_secondaryownerid);///HouseHold Owner ID

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
                        new Guid(ssi_secondaryownerid)),
                    Target = new EntityReference("ssi_batch",
                        new Guid(BatchId))
                };


                service.Execute(assignRequest);







            } //“Associate Approved – Pend Advisor Approval”: or Action =“OPS Change Requested || 
            else if (BatchStatusId == "1" || ReportTrackerStatus == 7 || (BatchStatusId == "6" && AdvisorApproval.ToUpper() != "TRUE"))
            {
                ssi_secondaryownerid = GetUserID("CORP\\opsreporting");  //OPS Reporting, Gresham

            }
            else if (BatchStatusId == "6" && AdvisorApproval.ToUpper() == "TRUE")
            {
                ssi_secondaryownerid = AdvisorId;
                //objBatch.ssi_billingcontactchange = new CrmBoolean();
                //objBatch.ssi_billingcontactchange.Value = true;// to send Email Notofication to Advisor
                objBatch["ssi_billingcontactchange"] = true;
            }

            if (BatchStatusId == "6" || BatchStatusId == "1" || ReportTrackerStatus == 7)
            {
                //ssi_secondaryownerid = "e5e49f9c-9e04-e111-b3cd-0019b9e7ee05";  //OPS Reporting, Gresham

                //SecurityPrincipal assignee = new SecurityPrincipal();
                //assignee.PrincipalId = new Guid(ssi_secondaryownerid);///HouseHold Owner ID

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
                        new Guid(ssi_secondaryownerid)),
                    Target = new EntityReference("ssi_batch",
                        new Guid(BatchId))
                };


                service.Execute(assignRequest);


            }

            //////////////////////// End of new logic //////////////////////////////////////
            if (ddlAction.SelectedValue != "7")
            {
                if ((HoldReasonValue != "" && HoldReasonValue != "0") && ddlAction.SelectedValue != "9")
                {
                    //objBatch.ssi_holdreport = new Picklist();
                    //objBatch.ssi_holdreport.Value = Convert.ToInt32(HoldReasonValue);
                    objBatch["ssi_holdreport"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(HoldReasonValue));


                }
            }
            else if (HoldReasonValue == "" || HoldReasonValue == "0")
            {
                //objBatch.ssi_holdreport = new Picklist();
                //objBatch.ssi_holdreport.IsNull = true;
                //objBatch.ssi_holdreport.IsNullSpecified = true;
                objBatch["ssi_holdreport"] = null;
            }
            if (BatchStatusId == "6")
            {

            }
            //if (ddlAction.SelectedValue == "6")// Action 'Mark Sent'
            //{
            //    // Sent E-mail to Reviewer Required by
            //    if (MailPref.ToUpper() == "EMAIL" && ssi_ReviewReqByid != "")
            //    {
            //        objBatch.ssi_sendrrbymail = new CrmBoolean();
            //        objBatch.ssi_sendrrbymail.Value = true;
            //    }
            //}
            service.Update(objBatch);

            string UserId = GetcurrentUser();

            BatchReportStatus(ReportTrackerStatus, UserId, BatchId, BillingHandedOff);

        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }
    }

    public void MailRecordsStatus(int MailStatus, string BatchId, string ssi_batchdate)
    {
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;

        lblMessage.Text = "";
        lblError.Text = "";
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {
            sqlstr = "SP_S_BATCH_MAILRECORDS @BatchId='" + BatchId + "'";
            DataSet NewDataset = clsDB.getDataSet(sqlstr);
            for (int i = 0; i < NewDataset.Tables[0].Rows.Count; i++)
            {
                #region Update Mail Records
                //ssi_mailrecords objMailRecords = new ssi_mailrecords();
                Entity objMailRecords = new Entity("ssi_mailrecords");
                if (Convert.ToString(NewDataset.Tables[0].Rows[i]["Ssi_mailrecordsid"]) != "")
                {
                    //objMailRecords.ssi_mailrecordsid = new Key();
                    //objMailRecords.ssi_mailrecordsid.Value = new Guid(Convert.ToString(NewDataset.Tables[0].Rows[i]["Ssi_mailrecordsid"]));
                    objMailRecords["ssi_mailrecordsid"] = new Guid(Convert.ToString(NewDataset.Tables[0].Rows[i]["Ssi_mailrecordsid"]));
                }

                //objMailRecords.ssi_mailstatus = new Picklist();
                //objMailRecords.ssi_mailstatus.Value = MailStatus;
                objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(MailStatus);

                // objMailRecords.ssi_asofdate = new CrmDateTime();
                // objMailRecords.ssi_asofdate.Value = ssi_batchdate;

                if (Convert.ToInt32(NewDataset.Tables[0].Rows[i]["Ssi_mailstatus"]) != 6)//mail status is 'Canceled'
                {
                    service.Update(objMailRecords);
                }
                #endregion

                #region Mail Records Status Change Log

                //ssi_mailrecordstatuschangelog objMailRecordsChangeLog = new ssi_mailrecordstatuschangelog();

                //// As Of Date
                //if (Convert.ToString(NewDataset.Tables[0].Rows[i]["As Of Date"]) != "")
                //{
                //    objMailRecordsChangeLog.ssi_asofdate = new CrmDateTime();
                //    objMailRecordsChangeLog.ssi_asofdate.Value = Convert.ToString(NewDataset.Tables[0].Rows[i]["As Of Date"]);
                //}

                ////Mail ID
                //if (Convert.ToString(NewDataset.Tables[0].Rows[i]["Ssi_mailingId"]) != "")
                //{
                //    objMailRecordsChangeLog.ssi_mailid = new CrmNumber();
                //    objMailRecordsChangeLog.ssi_mailid.Value = Convert.ToInt32(NewDataset.Tables[0].Rows[i]["Ssi_mailingId"]);
                //}

                ////Mail Type
                //if (Convert.ToString(NewDataset.Tables[0].Rows[i]["Ssi_mailTypeId"]) != "")
                //{
                //    objMailRecordsChangeLog.ssi_mailtypeid = new Lookup();
                //    objMailRecordsChangeLog.ssi_mailtypeid.Value = new Guid(Convert.ToString(NewDataset.Tables[0].Rows[i]["Ssi_mailTypeId"]));
                //}

                ////Recipient
                //if (Convert.ToString(NewDataset.Tables[0].Rows[i]["contactid"]) != "")
                //{
                //    objMailRecordsChangeLog.ssi_recipientid = new Lookup();
                //    objMailRecordsChangeLog.ssi_recipientid.Value = new Guid(Convert.ToString(NewDataset.Tables[0].Rows[i]["contactid"]));
                //}

                //// Mailing Status
                //if (Convert.ToString(NewDataset.Tables[0].Rows[i]["Ssi_mailstatus"]) != "")
                //{
                //    objMailRecordsChangeLog.ssi_status = new Picklist();
                //    objMailRecordsChangeLog.ssi_status.Value = Convert.ToInt32(NewDataset.Tables[0].Rows[i]["Ssi_mailstatus"]);
                //}

                //// Mail Record Owner id not found
                //// Currently Used is Batch OwnerId
                //if (Convert.ToString(NewDataset.Tables[0].Rows[i]["OwnerId"]) != "")
                //{
                //    objMailRecordsChangeLog.ssi_mailrecordownerid = new Lookup();
                //    objMailRecordsChangeLog.ssi_mailrecordownerid.Value = new Guid(Convert.ToString(NewDataset.Tables[0].Rows[i]["OwnerId"]));
                //}

                //// Date Time Of Change
                //objMailRecordsChangeLog.ssi_datetimeofchange = new CrmDateTime();
                //objMailRecordsChangeLog.ssi_datetimeofchange.Value = DateTime.Now.ToString();

                //// Person Who did Changes
                //string UserId = GetcurrentUser();
                //if (UserId != "")
                //{
                //    objMailRecordsChangeLog.ssi_mailrecordmodifiedbyid = new Lookup();
                //    objMailRecordsChangeLog.ssi_mailrecordmodifiedbyid.Value = new Guid(UserId);
                //}


                //service.Create(objMailRecordsChangeLog);

                #endregion
            }
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

    }

    public void MailRecordsUpdate(string BatchId, string Action, string ssi_batchdate, string MailPref, string ssi_ReviewReqByid)
    {
        if (BatchId != "" && Action != "")
        {
            int MailStatus;

            switch (Action)
            {
                case "1": // Ops Approve
                    MailStatus = 7; // mail status is 'Approved'
                    MailRecordsStatus(MailStatus, BatchId, ssi_batchdate);
                    break;
                //case "2":
                //    MailStatus = 5;// mail status is 'Pend Approval'
                //    MailRecordsStatus(MailStatus, BatchId, ssi_batchdate);
                //    break;
                //case "5":
                //    MailStatus = 7;// mail status is 'Approved'
                //    MailRecordsStatus(MailStatus, BatchId, ssi_batchdate);
                //    break;
                case "4":
                    MailStatus = 6;// mail status is 'Canceled'
                    MailRecordsStatus(MailStatus, BatchId, ssi_batchdate);
                    break;
                case "6":
                    //if (MailPref.ToUpper() == "EMAIL" && ssi_ReviewReqByid != "")
                    //{
                    //    MailStatus = 3;// mail status is 'Sent to Final Reviewer'
                    //}
                    //else
                    //{
                    MailStatus = 4;// mail status is 'Sent'
                    //}
                    MailRecordsStatus(MailStatus, BatchId, ssi_batchdate);
                    break;
                case "7":
                    MailStatus = 6;// mail status is 'Canceled'
                    MailRecordsStatus(MailStatus, BatchId, ssi_batchdate);
                    break;
            }
        }
    }

    public void BatchReportStatus(int Status, string UpdatedBy, string BatchId, bool BillingHandedOff)
    {
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;

        lblMessage.Text = "";
        lblError.Text = "";
        DataSet loInvoiceData = null;
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {
            //ssi_batchreportstatuslog objReportStatusLog = new ssi_batchreportstatuslog();
            Entity objReportStatusLog = new Entity("ssi_batchreportstatuslog");

            //objReportStatusLog.ssi_statusdate = new CrmDateTime();
            //objReportStatusLog.ssi_statusdate.Value = DateTime.Now.ToString();
            objReportStatusLog["ssi_statusdate"] = DateTime.Now;

            if (Status != 0 && Status != 888 && Status != 999)
            {
                //objReportStatusLog.ssi_status = new Picklist();
                //objReportStatusLog.ssi_status.Value = Status;
                objReportStatusLog["ssi_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(Status);
            }

            if (UpdatedBy != "")
            {
                //objReportStatusLog.ssi_updatedbyid = new Lookup();
                //objReportStatusLog.ssi_updatedbyid.Value = new Guid(UpdatedBy);
                objReportStatusLog["ssi_updatedbyid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UpdatedBy));
            }

            if (BatchId != "")
            {
                //objReportStatusLog.ssi_batchid = new Lookup();
                //objReportStatusLog.ssi_batchid.Value = new Guid(BatchId);
                objReportStatusLog["ssi_batchid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_batch", new Guid(BatchId));
            }

            if (BillingHandedOff == true && (Status != 0 && Status != 888))//True
            {
                //objReportStatusLog.ssi_billinghandedoff = new CrmBoolean();
                //objReportStatusLog.ssi_billinghandedoff.Value = true;
                objReportStatusLog["ssi_billinghandedoff"] = true;
            }
            else if (Status != 0 && Status != 888)
            {
                //objReportStatusLog.ssi_billinghandedoff = new CrmBoolean();
                //objReportStatusLog.ssi_billinghandedoff.Value = false;
                objReportStatusLog["ssi_billinghandedoff"] = false;
            }

            if (Status != 0 && Status != 888)
            {
                service.Create(objReportStatusLog);
            }
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }
    }

    private string GetcurrentUser()
    {
        //// to find windows user 
        string UserID = string.Empty;
        System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        //  string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
        string strName = string.Empty;
        // Changed Windows to -ADFS Claims Login 8_9_2019
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


    protected void imgApprovedFile_Click(object sender, ImageClickEventArgs e)
    {
        GridViewRow r = (GridViewRow)((DataControlFieldCell)((ImageButton)sender).Parent).Parent;
        int rowIndex = Convert.ToInt32(r.RowIndex);

        string str22 = GeneralMethods.RemoveSpecialCharacters(GridView1.Rows[rowIndex].Cells[22].Text);

        if (GridView1.Rows[rowIndex].Cells[21].Text != "")
        {
            string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + str22);

            System.IO.File.Copy(GridView1.Rows[rowIndex].Cells[21].Text, strDirectory, true);
            //Directory.Delete(ReportOpFolder, true);

            try
            {
                string lsFileNamforFinal = "./ExcelTemplate/TempFolder/" + str22;
                Session["id"] = str22;
                //Response.Write("<script>");
                //Response.Write("window.open('ViewReport.aspx?" + GridView1.Rows[rowIndex].Cells[22].Text + "', 'mywindow')");
                //Response.Write("</script>");
                //GridView1.Rows[rowIndex].Cells[22].Text
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();
                sb.Append("\n<script type=text/javascript>\n");
                sb.Append("\nwindow.open('ViewReport.aspx?" + str22 + "', 'mywindow')");
                sb.Append("</script>");
                ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
            }
            catch (Exception exc)
            {
                Response.Write(exc.Message);
            }
        }
    }
    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            string BillingHandedOff = e.Row.Cells[27].Text.Trim().Replace("Billing Handed Off", "").Replace("&nbsp;", "");
            string ssi_holdreport = e.Row.Cells[32].Text.Trim().Replace("ssi_holdreport", "").Replace("&nbsp;", "");

            CheckBox chkBillingHandedOff = (CheckBox)e.Row.FindControl("chkBillingHandedOff");
            DropDownList ddlHoldReport = (DropDownList)e.Row.FindControl("ddlHoldReport");

            sqlstr = "SP_S_HOLD_REPORT";//Store Procedure to bind hold report dropdown
            clsGM.getListForBindDDL(ddlHoldReport, sqlstr, "Status", "ID");
            ddlHoldReport.Items.Insert(0, "");
            ddlHoldReport.Items[0].Value = "0";
            ddlHoldReport.SelectedIndex = 0;

            //Set Value to Hold Report Parameter
            ddlHoldReport.SelectedValue = ssi_holdreport;


            if (BillingHandedOff.ToUpper() == "X")
            {
                chkBillingHandedOff.Checked = true;
                chkBillingHandedOff.Enabled = false;
            }
            else
            {
                chkBillingHandedOff.Enabled = false;
            }


            if (BillingHandedOff.ToUpper() == "X")
            {
                chkBillingHandedOff.Checked = true;
                chkBillingHandedOff.Enabled = false;
            }
            else
            {
                chkBillingHandedOff.Enabled = false;
            }


            string FileName = e.Row.Cells[21].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
            ImageButton imgApprovedFile = (ImageButton)e.Row.FindControl("imgApprovedFile");

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
    protected void lstHouseHold_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        lblError.Text = "";
        BindGridView();
    }

    public void BatchChangeLog(string BatchAsOfDate, int Status, string BatchOwnerId, string BatchId, string RecipientId)
    {
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;


        lblMessage.Text = "";
        lblError.Text = "";
        DataSet loInvoiceData = null;
        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        try
        {
            //ssi_batchchangelog objBatchChangeLog = new ssi_batchchangelog();
            Entity objBatchChangeLog = new Entity("ssi_batchchangelog");

            if (BatchAsOfDate != "")
            {
                //objBatchChangeLog.ssi_batchasofdate = new CrmDateTime();
                //objBatchChangeLog.ssi_batchasofdate.Value = BatchAsOfDate;
                objBatchChangeLog["ssi_batchasofdate"] = Convert.ToDateTime(BatchAsOfDate);
            }

            if (BatchId != "")
            {
                //objBatchChangeLog.ssi_batchid = new Lookup();
                //objBatchChangeLog.ssi_batchid.Value = new Guid(BatchId);
                objBatchChangeLog["ssi_batchid"] = new Guid(BatchId);
            }

            if (RecipientId != "")
            {
                //objBatchChangeLog.ssi_recipientid = new Lookup();
                //objBatchChangeLog.ssi_recipientid.Value = new Guid(RecipientId);
                objBatchChangeLog["ssi_recipientid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(RecipientId));
            }

            if (Status != 0)
            {
                //objBatchChangeLog.ssi_status = new Picklist();
                //objBatchChangeLog.ssi_status.Value = Status;

                objBatchChangeLog["ssi_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(Status);
            }


            if (BatchOwnerId != "")
            {
                //objBatchChangeLog.ssi_batchownerid = new Lookup();
                //objBatchChangeLog.ssi_batchownerid.Value = new Guid(BatchOwnerId);
                objBatchChangeLog["ssi_batchownerid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(BatchOwnerId));

                //SecurityPrincipal assignee = new SecurityPrincipal();
                //assignee.PrincipalId = new Guid(BatchOwnerId);

                //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();
                //targetAssign.EntityId = new Guid(BatchId);
                //targetAssign.EntityName = EntityName.ssi_batchchangelog.ToString();

                //AssignRequest assign = new AssignRequest();
                //assign.Assignee = assignee;
                //assign.Target = targetAssign;

                //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);
            }

            //objBatchChangeLog.ssi_datetimeofchange = new CrmDateTime();
            //objBatchChangeLog.ssi_datetimeofchange.Value = DateTime.Now.ToString();
            objBatchChangeLog["ssi_datetimeofchange"] = DateTime.Now;


            string UserId = GetcurrentUser();

            if (UserId != "")
            {
                //objBatchChangeLog.ssi_batchmodifiedbyid = new Lookup();
                //objBatchChangeLog.ssi_batchmodifiedbyid.Value = new Guid(UserId);
                objBatchChangeLog["ssi_batchmodifiedbyid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));
            }


            service.Create(objBatchChangeLog);
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }
    }

    #region PDF Report Generation

    private void GenerateReport()
    {
        string ReportOpFolder = string.Empty;
        string ContactFolderName = string.Empty;

        string ParentFolder = string.Empty;
        string TempFolderPath = string.Empty;
        string Local_ParentFolderPath = string.Empty;
        clsCombinedReports objCombinedReports = new clsCombinedReports();
        if (strReportFiles != "")
        {
            return;
        }

        try
        {
            lblMessage.Text = "";
            lblError.Text = "";
            Session.Remove("CurPageInBatch");
            string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
            //string crmServerURL = "http://server:5555/";

            string orgName = "GreshamPartners";
            string currentuser = null;
            //string orgName = "Webdev";
            IOrganizationService service = null;
            Boolean checkrunreport = false;
            String DestinationPath = string.Empty;
            string ConsolidatePdfFileName = string.Empty;

            string ApprovedReports = AppLogic.GetParam(AppLogic.ConfigParam.ApprovedReports);//"\\\\fs01\\opsreports$\\Approved Reports\\";
            try
            {
                service = GM.GetCrmService();
                strDescription = "Crm Service starts successfully";
            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
                lblMessage.Text = strDescription;
            }
            catch (Exception exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                lblMessage.Text = strDescription;
            }

            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

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

            ParentFolder = strUserName + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond + strMilliSecond;


            ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports);//"\\\\fs01\\opsreports$";//"\\\\Fs01\\shared$\\OPS REPORTS\\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            //if (ddlAction.SelectedValue == "2" || ddlAction.SelectedValue == "3")
            //    ReportOpFolder = "\\\\Fs01\\shared$\\BATCH REPORTS\\";


            if (ddlAction.SelectedValue == "4")
                ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports);// "\\\\fs01\\opsreports$";//"\\\\Fs01\\shared$\\OPS REPORTS\\";

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                ReportOpFolder = @"C:\Reports\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            }
            else
            {
                ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports);//"\\\\fs01\\opsreports$";//"\\\\Fs01\\shared$\\OPS REPORTS\\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

                //if (ddlAction.SelectedValue == "1" || ddlAction.SelectedValue == "2" || ddlAction.SelectedValue == "3")
                //    ReportOpFolder = "\\\\Fs01\\shared$\\BATCH REPORTS\\";
            }


            if (ddlAction.SelectedValue == "4" && !Request.Url.AbsoluteUri.Contains("localhost"))
                ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.BatchReports);//"\\\\Fs01\\shared$\\BATCH REPORTS\\";



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

                string FileSeconds = FileDateTime.Second.ToString().Length < 2 ? "0" + FileDateTime.Second.ToString() : FileDateTime.Second.ToString();
                string FileMiliSec = FileDateTime.Millisecond.ToString().Length < 2 ? "0" + FileDateTime.Millisecond.ToString() : FileDateTime.Millisecond.ToString();

                string CurrentTimeStamp = FileYear + "_" + FileMonth + "_" + FileDay + "_" + FileHour + "_" + FileMinute + "_" + FileSeconds + "_" + FileMiliSec;

                if (chkBox.Checked)
                {
                    numIndexPageCount = 1;  //Index page count -- if count of batch records is > 22 then it will come on next page 
                    numIndexPageSize = 20;// 22; // Size of index page 

                    checkrunreport = true;
                    String BatchIdListTxt = Convert.ToString(GridView1.Rows[j].Cells[18].Text);
                    dtBatch = GetDataTable(BatchIdListTxt);

                    //String TempName =  GridView1.Rows[j].Cells[6].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();

                    //String HHName = GridView1.Rows[j].Cells[6].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                    //string ssi_batchid = GridView1.Rows[j].Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                    String HHName = "";
                    string SPVFileName = string.Empty;
                    string OldHHName = GridView1.Rows[j].Cells[4].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString().Replace("&#39;", "'").ToString();
                    OldHHName = OldHHName.Replace("/", "");
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
                            HHName = GridView1.Rows[j].Cells[4].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString().Replace("&#39;", "'").ToString();
                            HHName = HHName.Replace("/", "");
                        }


                        ContactFolderName = GridView1.Rows[j].Cells[28].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString().Replace("&#39;", "'").ToString();
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
                            //Response.Write("Folder: " + ReportOpFolder + "\\" + ContactFolderName);
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


                        String ssi_FooterLocation = String.Empty;
                        if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["ssi_FooterLocation"])))
                            ssi_FooterLocation = Convert.ToString(dtBatch.Rows[i]["ssi_FooterLocation"]);


                        String Ssi_GreshamClientFooter = String.Empty;
                        if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["Ssi_GreshamClientFooter"])))
                            Ssi_GreshamClientFooter = Convert.ToString(dtBatch.Rows[i]["Ssi_GreshamClientFooter"]);

                        String ClientFooterTxt = String.Empty;
                        if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["ClientFooterTxt"])))
                            ClientFooterTxt = Convert.ToString(dtBatch.Rows[i]["ClientFooterTxt"]);

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

                        bool bContinueBatch = true;

                        /** Attach Template PDF ---Static pdf logic  ***/
                        string strTemplateFilePath = Convert.ToString(dtBatch.Rows[i]["ssi_TemplateFilePath"]);
                        if (strTemplateFilePath != "")
                        {
                            string strExtension = Path.GetExtension(strTemplateFilePath);


                            #region Fetch File from Sharepoint
                            // if (strTemplateFilePath.Contains("https://greshampartners.sharepoint.com") || strTemplateFilePath.Contains("http://greshampartners.sharepoint.com"))
                            if (strTemplateFilePath.Contains(AppLogic.GetParam(AppLogic.ConfigParam.SharepointURL)) || strTemplateFilePath.Contains(AppLogic.GetParam(AppLogic.ConfigParam.httpSharepointURL)))
                            {

                                string FileName = Path.GetFileName(strTemplateFilePath);
                                FileName = FileName.Replace("%20", " ");
                                // string FileName2 = HttpUtility.HtmlEncode(FileName).ToString();
                                string SharepointPath = strTemplateFilePath;
                                SharepointPath = SharepointPath.Replace("//", "/");
                               // SharepointPath = SharepointPath.Replace("https:/greshampartners.sharepoint.com/clientserv/", "");
                                //SharepointPath = SharepointPath.Replace("http:/greshampartners.sharepoint.com/clientserv/", "");
                                SharepointPath = SharepointPath.Replace(AppLogic.GetParam(AppLogic.ConfigParam.clientservURL) + "/", "");
                                SharepointPath = SharepointPath.Replace(AppLogic.GetParam(AppLogic.ConfigParam.httpclientservURL) + "/ ", "");


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
                            //Page number logic                            
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
                                // CombinedFileName = generateCombinedPDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath.Replace(".xls", ".pdf"), fsFooterTxt, fsGreshReportIdName, LegalEntity, FundID, CommitmentReportHeader, fsGAorTIAflag, fsReportRollupGroupIdName, fsHHreportparametersId);
                                CombinedFileName = generateCombinedPDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath.Replace(".xls", ".pdf"), fsFooterTxt, fsGreshReportIdName, LegalEntity, FundID, CommitmentReportHeader, fsGAorTIAflag, fsReportRollupGroupIdName, fsHHreportparametersId, fsReportRollupGroupId, fsrHouseholdId, fsFundIRR, fsGreshamReportId, fsLegalEntityTitle, TempFolderPath,ssi_FooterLocation,ClientFooterTxt,Ssi_GreshamClientFooter);

                            }
                            else
                            {
                                SetValuesToVariable(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg);
                                // generatesExcelsheets(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg);
                                generatePDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg, TempFolderPath, ssi_FooterLocation, ClientFooterTxt, Ssi_GreshamClientFooter);
                                CombinedFileName = true;
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
                                generateCoversheetPDF(fsAsofDate, lsCoversheet, fsAllocationGroup, fsHouseholdName, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange, fsCoverSheetPageTitle, fsGAorTIAflag, fsDiscretionaryFlg, TempFolderPath);
                                generatesCoverExcel(fsAsofDate, fsHouseholdName, fsAllocationGroup, lsCoversheet, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange, fsCoverSheetPageTitle, TempFolderPath);
                            }
                        }
                        else
                        {
                            CombinedFileName = true;
                        }
                        /* Array fill with the PATH + Fullname of PDF*/

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
                            //9-5-2017 added ifclause MTGBK-NO Index - sasmit
                            if (!ContactFolderName.Contains("MTGBK"))
                            {
                                SourceFileArray1[2 + i] = (Server.MapPath("") + @"\ExcelTemplate\Blank.pdf");
                            }

                            if (CombinedFileName == true)
                                SourceFileArray1[i + (numIndexPageCount + 1)] = lsExcleSavePath.Replace(".xls", ".pdf");


                        }
                        #endregion
                        else if (i == 0)
                        {
                            SourceFileArray[i] = lsCoversheet.Replace(".xls", ".pdf");
                            for (int PageCnt = 1; PageCnt < numIndexPageCount; PageCnt++)
                            {
                                //9-5-2017 added ifclause MTGBK-NO Index - sasmit
                                if (!ContactFolderName.Contains("MTGBK"))
                                {
                                    SourceFileArray[i + PageCnt] = (Server.MapPath("") + @"\ExcelTemplate\Blank.pdf");
                                }
                            }
                            if (CombinedFileName == true)
                                SourceFileArray[i + (numIndexPageCount)] = lsExcleSavePath.Replace(".xls", ".pdf");
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
                        PDF.MergeFiles(DestinationPath, target);
                    }
                    else if (noIndex.Text == "coversheet") // added 10/03/2016 - sasmit
                    {

                        PDF.MergeFiles1(DestinationPath, SourceFileArray1);
                        //string DestinationPath1 = objCombinedReports.addPageIndex1(DestinationPath, dtBatch);
                        //File.Copy(DestinationPath1, DestinationPath, true);
                    }
                    else  //generate with coversheet
                    {
                        PDF.MergeFiles(DestinationPath, SourceFileArray);


                    }
                    //added 9_5_2017 (MTGBK-NO INDEX)-SASMIT
                    if (ContactFolderName.Contains("MTGBK"))
                    {
                        DateTime dtime = DateTime.Now;

                        System.IO.File.Copy(DestinationPath, ApprovedReports + ConsolidatePdfFileName, true);
                        // System.IO.File.Copy(DestinationPath, DestinationPath, true);

                        ////added  31-july-2018 sasmit(ops folder delete issue)
                        //if (ContactFolderName != "")
                        //{
                        //    System.IO.Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
                        //}
                        Session.Remove("CurPageInBatch");
                        strReportFiles = strReportFiles + "<br/>" + "<a href=file:" + AppLogic.GetParam(AppLogic.ConfigParam.OutPutReports) + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1).Replace(" ", "%20") + ">" + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1) + " </a>";

                    }
                    else if (noIndex.Text == "coversheet") // added 10/03/2016 - sasmit
                    {

                        string DestinationPath1 = objCombinedReports.addPageIndex1(DestinationPath, dtBatch, TempFolderPath);

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
                        //    System.IO.Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
                        //}

                        ////added 18_05_2018 - CLEANUP JUNKFILES(sasmit)
                        //System.IO.File.Delete(DestinationPath1);

                        //string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + ConsolidatePdfFileName);

                        //File.Copy(DestinationPath, strDirectory, true);

                        Session.Remove("CurPageInBatch");
                        strReportFiles = strReportFiles + "<br/>" + "<a href=file:" + AppLogic.GetParam(AppLogic.ConfigParam.OutPutReports) + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1).Replace(" ", "%20") + ">" + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1) + " </a>";
                        //sourcearray.Text = "File at :" + strReportFiles + DateTime.Now +"end of path "; //testing
                    }
                    else
                    {
                        // File.Copy(DestinationPath, ApprovedReports + ConsolidatePdfFileName, true);
                        string DestinationPath1 = objCombinedReports.addPageIndex(DestinationPath, dtBatch, TempFolderPath);
                        System.IO.File.Copy(DestinationPath1, ApprovedReports + ConsolidatePdfFileName, true);
                        System.IO.File.Copy(DestinationPath1, DestinationPath, true);
                        ////added  31-july-2018 sasmit(ops folder delete issue)
                        //if (ContactFolderName != "")
                        //{
                        //    System.IO.Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
                        //}
                        ////added 18_05_2018 - CLEANUP JUNKFILES(sasmit)
                        //System.IO.File.Delete(DestinationPath1);

                        Session.Remove("CurPageInBatch");
                        strReportFiles = strReportFiles + "<br/>" + "<a href=file:" + AppLogic.GetParam(AppLogic.ConfigParam.OutPutReports) + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1).Replace(" ", "%20") + ">" + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1) + " </a>";
                    }
                    if (ddlAction.SelectedValue == "13")  //Insert Coversheet  added 10/03/2016 - sasmit
                    {
                        #region Region to update Batch File Name & Batch File Display Name
                        //code to update updatedate in batch ety of crm
                        //ssi_batch objBatch = new ssi_batch();
                        Entity objBatch = new Entity("ssi_batch");

                        if (BatchIdListTxt != "")
                        {
                            //objBatch.ssi_batchid = new Key();
                            //objBatch.ssi_batchid.Value = new Guid(BatchIdListTxt);
                            objBatch["ssi_batchid"] = new Guid(BatchIdListTxt);
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
                    if (ddlAction.SelectedValue == "1")  // Approve
                    {
                        #region Region to update Batch File Name & Batch File Display Name
                        //code to update updatedate in batch ety of crm
                        //ssi_batch objBatch = new ssi_batch();
                        Entity objBatch = new Entity("ssi_batch");
                        if (BatchIdListTxt != "")
                        {
                            //objBatch.ssi_batchid = new Key();
                            //objBatch.ssi_batchid.Value = new Guid(BatchIdListTxt);
                            objBatch["ssi_batchid"] = new Guid(BatchIdListTxt);
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


                    //    else //generate with coversheet
                    //    {
                    //        PDF.MergeFiles(DestinationPath, SourceFileArray);
                    //        string DestinationPath1 = objCombinedReports.addPageIndex(DestinationPath, dtBatch);
                    //        File.Copy(DestinationPath1, DestinationPath, true);
                    //        chkerror.Text = DestinationPath + noIndex.Text;
                    //    }

                    //    System.IO.Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
                    //    Session.Remove("CurPageInBatch");
                    //    strReportFiles = strReportFiles + "<br/>" + "<a href=file:" + AppLogic.GetParam(AppLogic.ConfigParam.OutPutReports) + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1).Replace(" ", "%20") + ">" + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1) + " </a>";
                    //}
                }
            }

            ////////////////////////////////////

            if (ddlAction.SelectedValue == "4") //OPS Change Requested
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

                        string lsFileNamforFinal = "./ExcelTemplate/TempFolder/" + ConsolidatePdfFileName;

                        //Response.Write("<script>");
                        //Response.Write("window.open('ViewReport.aspx?" + ConsolidatePdfFileName + "', 'mywindow')");
                        //Response.Write("</script>");

                        System.Text.StringBuilder sb = new System.Text.StringBuilder();
                        Type tp = this.GetType();
                        sb.Append("\n<script type=text/javascript>\n");
                        sb.Append("\nwindow.open('ViewReport.aspx?" + ConsolidatePdfFileName + "', 'mywindow')");
                        sb.Append("</script>");
                        ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
                    }
                    catch (Exception exc)
                    {
                        Response.Write(exc.Message);
                    }
                }
            }
            else //if (ddlAction.SelectedValue == "1")
            {
                System.IO.File.Copy(DestinationPath, ApprovedReports + ConsolidatePdfFileName, true);
            }

            //File.Copy(DestinationPath, ApprovedReports + ConsolidatePdfFileName);
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            ////added  31-july-2018 sasmit(ops folder delete issue)
            //if (ContactFolderName != "")
            //{
            //    if (Directory.Exists(ReportOpFolder + "\\" + ContactFolderName))
            //    {
            //        Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
            //    }
            //}
            Response.Write("Error Occured" + exc.Message.ToString());
            bProceed = false;
            strDescription = "Error Generating Report, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
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
            //}
            Response.Write("Error Occured" + ex.Message.ToString());
            lblMessage.Text = "Error Generating Report " + ex.ToString();
        }
        finally
        {
            //added  31-july-2018 sasmit(ops folder delete issue)
            if (Directory.Exists(ReportOpFolder + "\\" + ParentFolder))
            {
                Directory.Delete(ReportOpFolder + "\\" + ParentFolder, true);
            }
            //delete tempfolder creted at local Directory
            if (Directory.Exists(Local_ParentFolderPath))
            {
                Directory.Delete(Local_ParentFolderPath, true);
            }
        }
    }
    public string sharepointFile(string FileName, string path, string finalPath)
    {
        string Value = null;

        string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientservURL);
       // string siteUrl = "https://greshampartners.sharepoint.com/clientserv";
        context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();

        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
        //foreach (var c in "51ngl3malt") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);

        string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID);
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword);
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
        String LegalEntityId, String FundId, String CommitmentReportHeader, String GAorTIAflag, String ReportRollupGroupIdName, String fsHHreportparametersId, String fsReportRollupGroupId, String fsHouseholdId, String fsFundIRR, String fsGreshamReportId, String fsLegalEntityTitle, String TempFolderPath, String FooterLocation, String ClientFootertext, String Ssi_GreshamClientFooter)
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

        objCombinedReports.Footerlocation = FooterLocation;

        objCombinedReports.ClientFooterTxt = ClientFootertext;

        objCombinedReports.Ssi_GreshamClientFooter = Ssi_GreshamClientFooter;


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

    public void generatePDF(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFooterTxt, String fsGAorTIAflag, String fsDiscretionaryFlg, String TempFolderPath, String FooterLocation, String ClientFootertext, String Ssi_GreshamClientFooter)
    {
        clsCombinedReports objCombinedReports = new clsCombinedReports();
        liPageSize = 28;//commented on 07/01/2020 as confirmed by sir
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
                                                                                                                              // String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
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
                    //  document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty));//Commented -- FooterLogic
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
                        //loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);  changes by abhi
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
                // document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));//Commented -- FooterLogic

                document.Add(addFooter(liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt, FooterLocation, ClientFootertext, Ssi_GreshamClientFooter));
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

        //            String ReportName = Convert.ToString(foTable.Rows[j]["ssi_GreshamReportIdName"]);
        //            if (ReportName == "Client Goals" || ReportName == "Absolute Returns" || ReportName == "Capital Protection")
        //            {
        //                if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["Ssi_HouseholdIdName"])))
        //                {
        //                    lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["Ssi_HouseholdIdName"]);
        //                }
        //            } 


        //            fsGAorTIAflag = Convert.ToString(foTable.Rows[j]["ssi_gaortia"]);
        //            fsDiscretionaryFlg = Convert.ToString(foTable.Rows[j]["Discretionary Flag"]);


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
        //                    loChunk = new Chunk(Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
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
        //            loChunk = new Chunk("Reports included:", setFontsAll(10, 0, 1));
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
        // loTable.AddCell(loCell);//Commented -- FooterLogic

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

        // String lsFileNamforFinalXls = System.DateTime.Now.ToString("MMddyyHHmmss") + ".xls";
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

                sheetCover.Range[lsShhetNumber].Text = Convert.ToString(foTable.Rows[liCounter]["ssi_greshamreportidname"]).Replace("v2.1", "") + ": " + lsFinalTitleAfterChange;
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
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            totalCount = 0;
            Response.Write("sp_s_batch sp fails error desc:" + exc.Detail.InnerText);
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
    public iTextSharp.text.Table addFooter(String lsDateTime, int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt, String footerLocation, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        iTextSharp.text.Table fotable = new iTextSharp.text.Table(2, 1);
        fotable.Width = 90;
        fotable.Border = 0;
        int[] headerwidths = { 50, 40 };
        fotable.SetWidths(headerwidths);
        fotable.Cellpadding = 0;
        Cell loCell = new Cell();
        Chunk loChunk = new Chunk();
        // footerLocation = "End of Report";
        int EndOfReportPageCnt = 4;
        if (footerflg)
        {
            if (Ssi_GreshamClientFooter == "2")
            {
                FooterTxt = ClientFooterTxt;
                footerLocation = "100000000";
                if (footerLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                    }





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



                    for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData - EndOfReportPageCnt; liCounter++)
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



                    #endregion
                }
                else
                {
                    #region Footer on Default

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


                    #endregion
                }
            }
            else if (Ssi_GreshamClientFooter == "3")
            {
                if (footerLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                    }





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



                    for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData - EndOfReportPageCnt; liCounter++)
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


                    loCell = new Cell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Chunk(ClientFooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loCell.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.Add(loChunk);
                    fotable.AddCell(loCell);



                    #endregion
                }
                else
                {
                    #region Footer on Default

                    FooterTxt = FooterTxt + "\n" + ClientFooterTxt;
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


                    #endregion
                }
            }


            else if (Ssi_GreshamClientFooter == "1")
            {
                if (footerLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                    }





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



                    for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData - EndOfReportPageCnt; liCounter++)
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



                    #endregion
                }
                else
                {
                    #region Footer on Default

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


                    #endregion
                }
            }

            else if (Ssi_GreshamClientFooter == "4")
            {

                #region For NONE
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
                #endregion

            }



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



        //if (footerflg)
        //{
        //    loCell = new Cell();
        //    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
        //    loChunk = new Chunk(FooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        //    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
        // //   loCell.Leading = 8f;
        //    loCell.HorizontalAlignment = 0;
        //    loCell.Colspan = 2;
        //    loCell.BorderWidth = 0;
        //    loCell.Add(loChunk);
        //    fotable.AddCell(loCell);
        //}



        //fotable.TableFitsPage = true;

        return fotable;
    }

    public iTextSharp.text.Table addFooter(int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt, String footerLocation, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {

        iTextSharp.text.Table fotable = new iTextSharp.text.Table(2, 1);
        fotable.Width = 90;
        fotable.Border = 0;
        int[] headerwidths = { 50, 40 };
        fotable.SetWidths(headerwidths);
        fotable.Cellpadding = 0;
        Cell loCell = new Cell();
        Chunk loChunk = new Chunk();
        int EndOfReportPageCnt = 4;

        if (footerflg)
        {
            if (Ssi_GreshamClientFooter == "2")
            {
                FooterTxt = ClientFooterTxt;
                footerLocation = "100000000";
                if (footerLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                    }





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



                    for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData - EndOfReportPageCnt; liCounter++)
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



                    #endregion
                }
                else
                {
                    #region Footer on Default

                    if (liLastPageData != 0)
                    {

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
                    }


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


                    #endregion
                }
            }
            else if (Ssi_GreshamClientFooter == "3")
            {
                if (footerLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                    }





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



                    for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData - EndOfReportPageCnt; liCounter++)
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


                    loCell = new Cell();
                    //loChunk = new Chunk("Footer testing INPROGRESS", Font8Normal());
                    loChunk = new Chunk(ClientFooterTxt, setFontsAll(7, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    //loChunk = new Chunk(FooterTxt, setFontsAll(9, 0, 0, new iTextSharp.text.Color(150, 150, 150)));
                    loCell.Leading = 8f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Colspan = 2;
                    loCell.BorderWidth = 0;
                    loCell.Add(loChunk);
                    fotable.AddCell(loCell);



                    #endregion
                }
                else
                {
                    #region Footer on Default

                    FooterTxt = ClientFooterTxt + "\n" + FooterTxt;

                    if (liLastPageData != 0)
                    {
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
                    }


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


                    #endregion
                }
            }

            else if (Ssi_GreshamClientFooter == "1")
            {
                if (footerLocation == "100000001")
                {
                    #region Footer on End Report


                    for (int i = 0; i < EndOfReportPageCnt; i++)
                    {
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                        loCell = new Cell();
                        loChunk = new Chunk("dev", Font8Whitecheck("test"));
                        loCell.HorizontalAlignment = 2;
                        loCell.BorderWidth = 0;
                        loCell.Add(loChunk);
                        fotable.AddCell(loCell);
                    }





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



                    for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData - EndOfReportPageCnt; liCounter++)
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



                    #endregion
                }
                else
                {
                    #region Footer on Default

                    if (liLastPageData != 0)
                    {

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
                    }


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


                    #endregion
                }
            }

            else if (Ssi_GreshamClientFooter == "4")
            {
                #region For NONE
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
                #endregion

            }
        }


        //loCell = new Cell();
        ////loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font8Normal());
        //loChunk = new Chunk("Page " + liCurrentPage + " of " + liTotalPages, Font7Normal());
        //loCell.Leading = 15f;//25f
        //loCell.HorizontalAlignment = 2;
        //loCell.BorderWidth = 0;
        //loCell.Add(loChunk);
        //fotable.AddCell(loCell);

        //loCell = new Cell();
        ////loChunk = new Chunk(lsDateTime, Font8GreyItalic());
        //loChunk = new Chunk(lsDateTime, Font7GreyItalic());
        //loCell.Add(loChunk);
        //loCell.Leading = 15f;//25f
        //loCell.BorderWidth = 0;
        //loCell.HorizontalAlignment = 2;
        //fotable.AddCell(loCell);
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
                //foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216); changes by abhi 
                foCell.BorderColorBottom = new iTextSharp.text.Color(191, 191, 191);
            }
        }
        catch { }
    }

    public void setGreyBorder(Cell foCell)
    {

        foCell.BorderWidthBottom = 0.1F;
        //foCell.BorderColorBottom = new iTextSharp.text.Color(242, 242, 242);
        // foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216);changes by abhi 
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


    #endregion

    protected void btnRefresh_Click(object sender, EventArgs e)
    {
        lblMessage.Text = "";
        lblError.Text = "";
        BindGridView();
    }

    protected void ddlMailtype_SelectedIndexChanged(object sender, EventArgs e)
    {
        // ClearControls();
        //BindHouseHold(lstHouseHold);
        lblMessage.Text = "";
        BindGridView();
    }

    protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlBatchType.SelectedValue != "")
        {
            BindGridView();
        }
    }

    private void DeleteBatchAndMailRecords(IOrganizationService service)
    {
        foreach (GridViewRow row in GridView1.Rows)
        {
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
            DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");

            string Batchid = row.Cells[18].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
            string BatchName = row.Cells[2].Text.Trim().Replace("Batch Name", "").Replace("&nbsp;", "");
            string MailRecordsDelete = row.Cells[40].Text.Trim().Replace("ssi_mailrecords_del", "").Replace("&nbsp;", "");
            string BatchType = row.Cells[39].Text.Trim().Replace("BatchTypeID", "").Replace("&nbsp;", "");
            string BillingInvoiceid = row.Cells[42].Text.Trim().Replace("ssi_billinginvoiceid", "").Replace("&nbsp;", ""); ;
            try
            {
                if (ddlAction.SelectedValue == "12")//Approve
                {
                    if (MailRecordsDelete.ToUpper() == "TRUE")
                    {
                        if (BatchType == "4")
                        {
                            if (Batchid != "")
                            {
                                System.Threading.Thread.Sleep(20000);

                                string sqlMailRecords = "SP_D_MailRecord_Batch @BatchIdList='" + Batchid + "'";
                                //string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords);
                                string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords, "ssi_mailrecords", service);

                                string sqlBatch = "SP_D_Batch @BatchIdList='" + Batchid + "'";
                                //string DelBatch = clsDB.DeleteRecord(sqlBatch);
                                string DelBatch = clsDB.DeleteRecord(sqlBatch, "ssi_batch", service);
                            }
                        }
                    }
                }


            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                bProceed = false;
                strDescription = "Error occured, Error Detail: " + exc.Detail.InnerText;
                lblMessage.Text = strDescription;
            }
            catch (Exception exc)
            {
                bProceed = false;
                strDescription = "Error occured, Error Detail: " + exc.Message;
                lblMessage.Text = strDescription;
            }
        }

        foreach (DataRow dr in dtMail.Rows)
        {
            string BatchName = Convert.ToString(dr["BatchName"]);
            string BillingInvoiceid = Convert.ToString(dr["BillingInvoiceid"]);
            #region Update completed flag on billing invoice

            //  string strbillingComplteFlag = "SP_S_MailRecordsTempID_List @MailIDList=" + ViewState["MailId"].ToString();// +",@LegalEntityNameID='" + LegalEntityId + "',@ContactFullnameID='" + ContactId + "'";

            Entity objBillingInvoice = new Entity("ssi_billinginvoice");
            //if (chkSelectNC.Checked)
            //{
            if (BillingInvoiceid != "")
            {
                objBillingInvoice["ssi_billinginvoiceid"] = new Guid(Convert.ToString(BillingInvoiceid));

                objBillingInvoice["ssi_completed"] = false;


                objBillingInvoice["ssi_invoicedate"] = null;

                service.Update(objBillingInvoice);
                SendEmail(BatchName);
                dtMail.Dispose();
                // SendEmail(BatchName);
            }
            //}

            //}


            #endregion

        }
        if (ddlAction.SelectedValue == "12")//Approve
        {
            BindGridView();
        }
    }



    private void InsertIntoWireExecution(string BatchId, string BatchTypeId)
    {
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;
        int finalReportCreatedCount = 0;
        int selectedCount = 0;
        string test = ddlBatchOwner.SelectedValue;
        int UniqueMailingId = 0;
        lblMessage.Text = "";
        lblError.Text = "";
        DataSet loInvoiceData = null;
        bool bOpsApproveRequestFlg = false;

        try
        {
            service = GM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblMessage.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblMessage.Text = strDescription;
        }

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        //ssi_wireexecution ObjWireExe = null;
        Entity ObjWireExe = null;

        if (BatchId != "")
        {
            if (BatchTypeId == "1") // Capital Call 
            {
                #region Capital Call

                string strsql = "SP_S_CapitalCall_WireExecution @BatchIdList='" + BatchId + "'";
                DataSet WireExeDataset = clsDB.getDataSet(strsql);

                for (int i = 0; i < WireExeDataset.Tables[0].Rows.Count; i++)
                {
                    //ObjWireExe = new ssi_wireexecution();
                    ObjWireExe = new Entity("ssi_wireexecution");
                    //ssi_name
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["Name"]) != "")
                    {
                        //ObjWireExe.ssi_name = Convert.ToString(WireExeDataset.Tables[0].Rows[i]["Name"]);
                        ObjWireExe["ssi_name"] = Convert.ToString(WireExeDataset.Tables[0].Rows[i]["Name"]);
                    }


                    //ssi_typeid
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_typeid"]) != "")
                    {
                        //ObjWireExe.ssi_type = new Picklist();
                        //ObjWireExe.ssi_type.Value = Convert.ToInt32(WireExeDataset.Tables[0].Rows[i]["ssi_typeid"]);
                        ObjWireExe["ssi_type"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(WireExeDataset.Tables[0].Rows[i]["ssi_typeid"]));
                    }


                    //ssi_legalentitynameid
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_legalentityid"]) != "")
                    {
                        //ObjWireExe.ssi_legalentityid = new Lookup();
                        //ObjWireExe.ssi_legalentityid.Value = new Guid(Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_legalentityid"]));
                        ObjWireExe["ssi_legalentityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_legalentityid"])));
                    }

                    //ssi_Householdid
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_Householdid"]) != "")
                    {
                        //ObjWireExe.ssi_householdid = new Lookup();
                        //ObjWireExe.ssi_householdid.Value = new Guid(Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_Householdid"]));
                        ObjWireExe["ssi_householdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_Householdid"])));
                    }

                    //ssi_totaladjustedcommitment
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_totaladjustedcommitment"]) != "")
                    {
                        //ObjWireExe.ssi_totaladjustedcommitment = new CrmMoney();
                        //ObjWireExe.ssi_totaladjustedcommitment.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_totaladjustedcommitment"]);
                        ObjWireExe["ssi_totaladjustedcommitment"] = new Money(Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_totaladjustedcommitment"]));
                    }


                    //ssi_capitalcalltotalpercentcalled
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcalltotalpercentcalled"]) != "")
                    {
                        //ObjWireExe.ssi_capitalcallpercentcalled = new CrmDecimal();
                        //ObjWireExe.ssi_capitalcallpercentcalled.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcalltotalpercentcalled"]);
                        ObjWireExe["ssi_capitalcallpercentcalled"] = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcalltotalpercentcalled"]);
                    }

                    //ssi_amount
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_amount"]) != "")
                    {
                        //ObjWireExe.ssi_amount = new CrmMoney();
                        //ObjWireExe.ssi_amount.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_amount"]);
                        ObjWireExe["ssi_amount"] = new Money(Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_amount"]));
                    }

                    //ssi_capitalcallpriorcalls
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpriorcalls"]) != "")
                    {
                        //ObjWireExe.ssi_capitalcallpriorcalls = new CrmMoney();
                        //ObjWireExe.ssi_capitalcallpriorcalls.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpriorcalls"]);
                        ObjWireExe["ssi_capitalcallpriorcalls"] = new Money(Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpriorcalls"]));
                    }


                    //ssi_capitalcallpercentpriorcalls
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentpriorcalls"]) != "")
                    {
                        //ObjWireExe.ssi_capitalcallpercentpriorcalls = new CrmDecimal();
                        //ObjWireExe.ssi_capitalcallpercentpriorcalls.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentpriorcalls"]);
                        ObjWireExe["ssi_capitalcallpercentpriorcalls"] = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentpriorcalls"]);
                    }


                    //ssi_capitalcallpercentpriorcalls
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallcalledtodate"]) != "")
                    {
                        //ObjWireExe.ssi_capitalcallcalledtodate = new CrmMoney();
                        //ObjWireExe.ssi_capitalcallcalledtodate.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallcalledtodate"]);
                        ObjWireExe["ssi_capitalcallcalledtodate"] = new Money(Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallcalledtodate"]));
                    }

                    //ssi_capitalcallpercentpriorcalls
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentcalledtodate"]) != "")
                    {
                        //ObjWireExe.ssi_capitalcallpercentcalledtodate = new CrmDecimal();
                        //ObjWireExe.ssi_capitalcallpercentcalledtodate.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentcalledtodate"]);
                        ObjWireExe["ssi_capitalcallpercentcalledtodate"] = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentcalledtodate"]);
                    }

                    //ssi_commitmentremainingcommitment
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitment"]) != "")
                    {
                        //ObjWireExe.ssi_commitmentremainingcommitment = new CrmMoney();
                        //ObjWireExe.ssi_commitmentremainingcommitment.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitment"]);
                        ObjWireExe["ssi_commitmentremainingcommitment"] = new Money(Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitment"]));
                    }


                    //ssi_commitmentremainingcommitmentpercent
                    if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitmentpercent"]) != "")
                    {
                        //    ObjWireExe.ssi_commitmentremainingcommitmentpercent = new CrmDecimal();
                        //    ObjWireExe.ssi_commitmentremainingcommitmentpercent.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitmentpercent"]);
                        ObjWireExe["ssi_commitmentremainingcommitmentpercent"] = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitmentpercent"]);
                    }

                    service.Create(ObjWireExe);
                    selectedCount++;

                }

                #endregion
            }
            else if (BatchTypeId == "2")//Distribution
            {
                #region Distribution

                string strsql = "SP_S_Distribution_WireExecution  @BatchIdList='" + BatchId + "'";
                DataSet WireExeDistributionDataset = clsDB.getDataSet(strsql);

                for (int j = 0; j < WireExeDistributionDataset.Tables[0].Rows.Count; j++)
                {

                    //ObjWireExe = new ssi_wireexecution();
                    ObjWireExe = new Entity("ssi_wireexecution");

                    //ssi_name
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["Name"]) != "")
                    {
                        //ObjWireExe.ssi_name = Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["Name"]);
                        ObjWireExe["ssi_name"] = Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["Name"]);
                    }


                    //Type
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_typeid"]) != "")
                    {
                        //ObjWireExe.ssi_type = new Picklist();
                        //ObjWireExe.ssi_type.Value = Convert.ToInt32(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_typeid"]);
                        ObjWireExe["ssi_type"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_typeid"]));
                    }

                    //ssi_LegalEntityid
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_LegalEntityid"]) != "")
                    {
                        //ObjWireExe.ssi_legalentityid = new Lookup();
                        //ObjWireExe.ssi_legalentityid.Value = new Guid(Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_LegalEntityid"]));
                        ObjWireExe["ssi_legalentityid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_LegalEntityid"])));
                    }

                    //ssi_Householdid
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_Householdid"]) != "")
                    {
                        //ObjWireExe.ssi_householdid = new Lookup();
                        //ObjWireExe.ssi_householdid.Value = new Guid(Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_Householdid"]));
                        ObjWireExe["ssi_householdid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_Householdid"])));
                    }


                    //ssi_totalcommitment
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_totalcommitment"]) != "")
                    {
                        //ObjWireExe.ssi_totalcommitment = new CrmMoney();
                        //ObjWireExe.ssi_totalcommitment.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_totalcommitment"]);
                        ObjWireExe["ssi_totalcommitment"] = new Money(Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_totalcommitment"]));
                    }


                    //ssi_amount
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_amount"]) != "")
                    {
                        //ObjWireExe.ssi_amount = new CrmMoney();
                        //ObjWireExe.ssi_amount.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_amount"]);
                        ObjWireExe["ssi_amount"] = new Money(Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_amount"]));

                    }


                    //ssi_distributionpercent
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpercent"]) != "")
                    {
                        //ObjWireExe.ssi_distributionpercent = new CrmDecimal();
                        //ObjWireExe.ssi_distributionpercent.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpercent"]);
                        ObjWireExe["ssi_distributionpercent"] = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpercent"]);
                    }


                    //ssi_distributionpriordistributions
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributions"]) != "")
                    {
                        //ObjWireExe.ssi_distributionpriordistributions = new CrmMoney();
                        //ObjWireExe.ssi_distributionpriordistributions.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributions"]);
                        ObjWireExe["ssi_distributionpriordistributions"] = new Money(Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributions"]));
                    }

                    //ssi_distributionpriordistributionspercent
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributionspercent"]) != "")
                    {
                        //ObjWireExe.ssi_distributionpriordistributionspercent = new CrmDecimal();
                        //ObjWireExe.ssi_distributionpriordistributionspercent.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributionspercent"]);
                        ObjWireExe["ssi_distributionpriordistributionspercent"] = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributionspercent"]);
                    }

                    //ssi_distributionstodate
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionstodate"]) != "")
                    {
                        //ObjWireExe.ssi_distributionstodate = new CrmMoney();
                        //ObjWireExe.ssi_distributionstodate.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionstodate"]);
                        ObjWireExe["ssi_distributionstodate"] = new Money(Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionstodate"]));
                    }

                    //ssi_DistributionstoDatePct
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_DistributionstoDatePct"]) != "")
                    {
                        //ObjWireExe.ssi_distributionstodatepct = new CrmDecimal();
                        //ObjWireExe.ssi_distributionstodatepct.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_DistributionstoDatePct"]);
                        ObjWireExe["ssi_distributionstodatepct"] = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_DistributionstoDatePct"]);
                    }


                    //ssi_capitalcallcalledtodate
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallcalledtodate"]) != "")
                    {
                        //ObjWireExe.ssi_capitalcallcalledtodate = new CrmMoney();
                        //ObjWireExe.ssi_capitalcallcalledtodate.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallcalledtodate"]);
                        ObjWireExe["ssi_capitalcallcalledtodate"] = new Money(Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallcalledtodate"]));
                    }


                    //ssi_commitmentremainingcommitment
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_commitmentremainingcommitment"]) != "")
                    {
                        //ObjWireExe.ssi_commitmentremainingcommitment = new CrmMoney();
                        //ObjWireExe.ssi_commitmentremainingcommitment.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_commitmentremainingcommitment"]);
                        ObjWireExe["ssi_commitmentremainingcommitment"] = new Money(Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_commitmentremainingcommitment"]));
                    }


                    //ssi_capitalcallpercentcalledtodate
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallpercentcalledtodate"]) != "")
                    {
                        //ObjWireExe.ssi_capitalcallpercentcalledtodate = new CrmDecimal();
                        //ObjWireExe.ssi_capitalcallpercentcalledtodate.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallpercentcalledtodate"]);
                        ObjWireExe["ssi_capitalcallpercentcalledtodate"] = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallpercentcalledtodate"]);
                    }


                    //ssi_distributionsclassbfeeadjustment
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsclassbfeeadjustment"]) != "")
                    {
                        //ObjWireExe.ssi_distributionsclassbfeeadjustement = new CrmMoney();
                        //ObjWireExe.ssi_distributionsclassbfeeadjustement.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsclassbfeeadjustment"]);
                        ObjWireExe["ssi_distributionsclassbfeeadjustement"] = new Money(Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsclassbfeeadjustment"]));
                    }



                    //ssi_distributionsactualcashdistributions 
                    if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsactualcashdistributions"]) != "")
                    {
                        //ObjWireExe.ssi_distributionsactualcashdistributions = new CrmMoney();
                        //ObjWireExe.ssi_distributionsactualcashdistributions.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsactualcashdistributions"]);
                        ObjWireExe["ssi_distributionsactualcashdistributions"] = new Money(Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsactualcashdistributions"]));
                    }

                    service.Create(ObjWireExe);
                    selectedCount++;

                }



                #endregion
            }
        }
    }

}
