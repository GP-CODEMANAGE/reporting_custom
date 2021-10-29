
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
using System.IO;
using System.Data.Common;
using Spire.Xls;
using System.Drawing;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Data.SqlClient;
using System.Text;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using System.Security;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Crm.Sdk;
using List = Microsoft.SharePoint.Client.List;
using System.Net.Mail;

using System.Threading;
using Microsoft.IdentityModel.Claims;
using System.Text.RegularExpressions;

public partial class MailQueue : System.Web.UI.Page
{
    string LogFileName = string.Empty;
    sharepoint sp = new sharepoint();
    Logs lg = new Logs();
    public StreamWriter sw = null;
    String sqlstr = string.Empty;
    DB clsDB = new DB();
    GeneralMethods clsGM = new GeneralMethods();
    bool bProceed = true;
    bool bMarkAllRecordsSent = false;
    string strDescription;
    public String _dbErrorMsg;
    public string ReviewRequiredBy = string.Empty;
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName;
    string MailRecordsIdListTxt = "";
    public string Message = string.Empty;
    GeneralMethods GM = new GeneralMethods();
    bool bListFetched = false;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {

            BindDropdownandListBox();
            string BillingFlg = string.Empty;
            //BillingFlg = Convert.ToString(Request.QueryString["bflag"]);
            bool bflag = false;
            //if (BillingFlg != null)
            //{
            //    // bflag = Convert.ToBoolean(BillingFlg);
            //    bflag = true;
            //}



            //  BindDropdownandListBox();

            string BatchIdListTxt = string.Empty;
            string BatchType = string.Empty;


            if (Request.QueryString.Count > 0 && Convert.ToString(Request.QueryString["btypeid"]) != "")
            {
                BatchType = Convert.ToString(Request.QueryString["btypeid"]);
                BillingFlg = Convert.ToString(Request.QueryString["bflag"]);

                // Response.Write("BillingFlg" + BillingFlg) ;
                //BillingFlg = "true";
            }
            if (BillingFlg != "")
                bflag = Convert.ToBoolean(BillingFlg);

            //  Response.Write(bflag);

            //  Response.Write(BillingFlg);

            // bListFetched = LoadList(bflag);

            bListFetched = LoadList();
            ViewState["bListFetched"] = bListFetched;
            if (Session["BatchIdList"] != null)
            {
                ViewState["BatchIdListTxt"] = Session["BatchIdList"].ToString(); //Convert.ToString(Request.QueryString["bidlist"]);
                ViewState["BillingFlg"] = bflag;


                lstMailStatus.SelectedValue = "7";
                if (BatchType == "4")
                {
                    ddlType.SelectedValue = "4";
                }
                else
                {
                    ddlType.SelectedValue = "5";
                }

                ddlAsofDate.SelectedIndex = 0;
            }
            else
            {
                ViewState["BatchIdListTxt"] = null;
                ViewState["BillingFlg"] = null;
                lstMailStatus.SelectedValue = "10";
                ddlType.SelectedValue = "5";
            }

            BindGridView();


            //  Response.Write("viewstate value" + bflag);
            chkMailingSheets.Enabled = false;
            chkReportSeperator.Enabled = false;



        }
        lblError.Text = "";
        //added 30_8_2019 TEst Flag auto checked if its local or test site
        string Server = AppLogic.GetParam(AppLogic.ConfigParam.Server);
        if (Server.ToLower() != "prod")
        {
            chkTest.Checked = true;
        }
    }

    private void BindGridView()
    {
        sqlstr = GetData();
        DataSet ds_table = clsDB.getDataSet(sqlstr);

        GridView1.Columns[12].Visible = true;
        GridView1.Columns[13].Visible = true;
        GridView1.Columns[14].Visible = true;
        GridView1.Columns[15].Visible = true;
        GridView1.Columns[16].Visible = true;
        GridView1.Columns[17].Visible = true;
        GridView1.Columns[18].Visible = true;
        GridView1.Columns[19].Visible = true;
        GridView1.Columns[20].Visible = true;
        GridView1.Columns[21].Visible = true;
        GridView1.Columns[23].Visible = true;
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
        //New Sharepoint CS site Changes - added 22_2_2019
        GridView1.Columns[40].Visible = true;
        GridView1.Columns[41].Visible = true;
        GridView1.Columns[42].Visible = true;
        GridView1.Columns[43].Visible = true;
        GridView1.Columns[44].Visible = true;
        GridView1.Columns[45].Visible = true;
        //added 11_9_2019 -- change Maling Names to ID
        GridView1.Columns[46].Visible = true;

        clsGM.SortGridView(GridView1, sqlstr, Convert.ToString(ViewState["sortExpression"]), Convert.ToString(ViewState["Direction"]));

        GridView1.Columns[12].Visible = false;
        GridView1.Columns[13].Visible = false;
        GridView1.Columns[14].Visible = false;
        GridView1.Columns[15].Visible = false;
        GridView1.Columns[16].Visible = false;
        GridView1.Columns[17].Visible = false;
        GridView1.Columns[18].Visible = false;
        GridView1.Columns[19].Visible = false;
        GridView1.Columns[20].Visible = false;
        GridView1.Columns[21].Visible = false;
        GridView1.Columns[23].Visible = false;
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
        //New Sharepoint CS site Changes - added 22_2_2019
        GridView1.Columns[40].Visible = false;
        GridView1.Columns[41].Visible = false;
        GridView1.Columns[42].Visible = false;
        GridView1.Columns[43].Visible = false;

        GridView1.Columns[44].Visible = false;
        GridView1.Columns[45].Visible = false;

        //added 11_9_2019 -- change Maling Names to ID
        GridView1.Columns[46].Visible = false;
        //GridView1.DataSource = ds_table;
        //GridView1.DataBind();
    }

    private String GetData()
    {
        object Type = ddlType.SelectedValue == "0" || ddlType.SelectedValue == "" ? "null" : "'" + ddlType.SelectedValue + "'";
        object AsOfDate = ddlAsofDate.SelectedValue == "0" ? "null" : "'" + ddlAsofDate.SelectedValue + "'";
        object MailID = lstMailId.SelectedValue == "0" || lstMailId.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstMailId) + "'";
        object MailTypeId = lstMailType.SelectedValue == "0" || lstMailType.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstMailType) + "'";
        object HouseHold = ddlHousehold.SelectedValue == "0" || ddlHousehold.SelectedValue == "" ? "null" : "'" + ddlHousehold.SelectedValue + "'";
        object Associate = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        object Advisor = ddlAdvisor.SelectedValue == "0" || ddlAdvisor.SelectedValue == "" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        object MailPreference = lstMailPreference.SelectedValue == "0" || lstMailPreference.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstMailPreference) + "'";
        object MailStatus = lstMailStatus.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstMailStatus) + "'";

        object Salutation = ddlSalutationPref.SelectedValue == "0" || ddlSalutationPref.SelectedValue == "" ? "null" : "'" + ddlSalutationPref.SelectedItem.Text + "'";
        object CreatedBy = ddlCreatedBy.SelectedValue == "0" || ddlCreatedBy.SelectedValue == "" || ddlCreatedBy.SelectedItem.Text == "''" ? "null" : "'" + ddlCreatedBy.SelectedItem.Text + "'";
        object CreatedOn = txtCreatedOn.Text == "" ? "null" : "'" + txtCreatedOn.Text + "'";


        sqlstr = "SP_S_REPORT_MAIL_QUEUE @AsOfDate=" + AsOfDate +
                                         ",@MailIDNmbList=" + MailID +
                                         ",@MailTypeIdNmbList=" + MailTypeId +
                                         ",@HouseHoldId=" + HouseHold +
                                         ",@AssociateId=" + Associate +
                                         ",@AdvisorId=" + Advisor +
                                         ",@MailPreferenceTxtList=" + MailPreference +
                                         ",@MailStatusIdNmbList=" + MailStatus +
                                         ",@SalutationTxt=" + Salutation +
                                         ",@CreatedBy=" + CreatedBy +
                                         ",@CreatedOn=" + CreatedOn +
                                        ",@BatchType=" + Type;

        if (ViewState["BatchIdListTxt"] != null)
        {
            MailStatus = "7";

            sqlstr = "SP_S_REPORT_MAIL_QUEUE ";
            sqlstr = sqlstr + " @MailStatusIdNmbList=" + MailStatus + ",@BatchIDList='" + Convert.ToString(ViewState["BatchIdListTxt"]) + "'";
        }

        return sqlstr;
    }

    private void BindDropdownandListBox()
    {
        BindMailId(lstMailId);
        BindMailType(lstMailType);
        BindMailPreference(lstMailPreference);
        BindMailStatus(lstMailStatus);
        BindHousehold(ddlHousehold);
        BindAssociate(ddlAssociate);
        BindAdvisor(ddlAdvisor);
        BindCreatedBy(ddlCreatedBy);
        BindAsOfDate(ddlAsofDate);
    }

    public void BindAsOfDate(DropDownList ddl)
    {
        //ddl.Items.Clear();
        sqlstr = "SP_S_MAIL_QUEUE_ASOFDATE";

        clsGM.getListForBindDDL(ddl, sqlstr, "ssi_AsOfDate", "ssi_AsOfDate");

        //Changed 6/4/2020 jeanne request
        //ddl.Items.Insert(0, "All");
        ddl.Items.Insert(0, "All in Last 12m");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 1;
    }

    public void BindMailId(ListBox lstBox)
    {
        sqlstr = "SP_S_MAIL_QUEUE_MAILID";
        clsGM.getListForBindListBox(lstBox, sqlstr, "ssi_MailingID", "ssi_MailingID");
       
        //Changed 6/4/2020 jeanne request
       // lstBox.Items.Insert(0, "All");
        lstBox.Items.Insert(0, "All in Last 12m");
        lstBox.Items[0].Value = "0";
        lstBox.SelectedIndex = 0;
    }

    private void DisableMailType()
    {
        string MailType = clsGM.GetMultipleSelectedItemsFromListBox(lstMailType);

        string[] Type = MailType.Split(',');

        for (int i = 0; i < Type.Length; i++)
        {
            // General Mailing ID                                      // Prospect Mailing ID                                //Smart Mailing ID
            if (Type[i] == "99b74584-e2d3-e011-a19b-0019b9e7ee05" || Type[i] == "c71108da-e1d3-e011-a19b-0019b9e7ee05" || Type[i] == "c10ba3b7-e1d3-e011-a19b-0019b9e7ee05")
            {
                ddlHousehold.SelectedIndex = 0;
                ddlHousehold.Enabled = false;

                ddlAssociate.SelectedIndex = 0;
                ddlAssociate.Enabled = false;

                ddlAdvisor.SelectedIndex = 0;
                ddlAdvisor.Enabled = false;
            }
            else if (Type[i] != "99b74584-e2d3-e011-a19b-0019b9e7ee05" || Type[i] != "c71108da-e1d3-e011-a19b-0019b9e7ee05" || Type[i] != "c10ba3b7-e1d3-e011-a19b-0019b9e7ee05")
            {

                ddlHousehold.SelectedIndex = 0;
                ddlHousehold.Enabled = true;

                ddlAssociate.SelectedIndex = 0;
                ddlAssociate.Enabled = true;

                ddlAdvisor.SelectedIndex = 0;
                ddlAdvisor.Enabled = true;
            }
        }

    }

    public void BindMailType(ListBox lstBox)
    {
        sqlstr = "SP_S_MAIL_LKUP";
        clsGM.getListForBindListBox(lstBox, sqlstr, "ssi_name", "ssi_mailid");

        lstBox.Items.Insert(0, "All");
        lstBox.Items[0].Value = "0";
        lstBox.SelectedIndex = 0;

        DisableMailType();
    }

    public void BindMailPreference(ListBox lstBox)
    {
        sqlstr = "SP_S_SENDVIA";
        clsGM.getListForBindListBox(lstBox, sqlstr, "status", "status");

        lstBox.Items.Insert(0, "All");
        lstBox.Items[0].Value = "0";
        lstBox.SelectedIndex = 0;
    }

    public void BindMailStatus(ListBox lstBox)
    {
        lstBox.Items.Clear();
        sqlstr = "SP_S_MAIL_STATUS";
        clsGM.getListForBindListBox(lstBox, sqlstr, "Status", "ID");


        lstBox.Items.Insert(0, "All - Excluding Canceled and sent");
        lstBox.Items[0].Value = "0";
        lstBox.SelectedIndex = 0;
    }

    public void BindHousehold(DropDownList ddl)
    {
        object AdvisorId = ddlAdvisor.SelectedValue == "0" || ddlAdvisor.SelectedValue == "" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        object AssociatedId = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        sqlstr = "SP_S_MAIL_QUEUE_HOUSEHOLD @IncludeClassB = 1,@AdvisorId=" + AdvisorId + ",@AssociateId=" + AssociatedId;
        clsGM.getListForBindDDL(ddl, sqlstr, "Name", "Accountid");

        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    public void BindAssociate(DropDownList ddl)
    {
        object AdvisorId = ddlAdvisor.SelectedValue == "0" || ddlAdvisor.SelectedValue == "" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        sqlstr = "SP_S_MAIL_QUEUE_ASSOCIATE @OwnerId=" + AdvisorId;
        clsGM.getListForBindDDL(ddl, sqlstr, "Ssi_SecondaryOwnerIdName", "Ssi_SecondaryOwnerId");

        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    public void BindAdvisor(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "SP_S_MAIL_QUEUE_ADVISOR";
        clsGM.getListForBindDDL(ddl, sqlstr, "OwnerIdName", "OwnerId");

        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }

    public void BindCreatedBy(DropDownList ddl)
    {
        ddl.Items.Clear();
        sqlstr = "SP_S_MAIL_QUEUE_CREATEDBY";
        clsGM.getListForBindDDL(ddl, sqlstr, "CreatedByTxt", "CreatedById");

        ddl.Items.Insert(0, "All");
        ddl.Items[0].Value = "0";
        ddl.SelectedIndex = 0;
    }
    protected void lstMailType_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
        DisableMailType();
    }
    protected void btnSerach_Click(object sender, EventArgs e)
    {
        BindGridView();
    }

    public SortDirection GridViewSortDirection
    {
        get
        {
            if (ViewState["sortDirection"] == null)
                ViewState["sortDirection"] = SortDirection.Descending;
            return (SortDirection)ViewState["sortDirection"];
        }
        set { ViewState["sortDirection"] = value; }
    }


    protected void GridView1_Sorting(object sender, GridViewSortEventArgs e)
    {
        if (GridViewSortDirection == SortDirection.Ascending && (ViewState["sortExpression"] == null || ViewState["sortExpression"].ToString() == e.SortExpression))
        {

            GridViewSortDirection = SortDirection.Descending;
            ViewState["Direction"] = "DESC";
            ViewState["sortExpression"] = e.SortExpression;
        }
        else
        {
            GridViewSortDirection = SortDirection.Ascending;
            ViewState["Direction"] = "ASC";
            ViewState["sortExpression"] = e.SortExpression;
        }

        BindGridView();
    }

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





    private void ExporttoCSV() // Export Selected Row of grid to csv
    {
        sqlstr = GetData();
        DataSet ds_table = clsDB.getDataSet(sqlstr);

        DataTable table = ds_table.Tables[0].Clone();

        for (int j = 0; j < GridView1.Rows.Count; j++)
        {
            CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");
            string MailID = GridView1.Rows[j].Cells[1].Text.Trim().Replace("Mail ID", "").Replace("&nbsp;", "");
            string MailType = GridView1.Rows[j].Cells[2].Text.Trim().Replace("Mail Type", "").Replace("&nbsp;", "");
            string AsOfDate = GridView1.Rows[j].Cells[3].Text.Trim().Replace("As Of Date", "").Replace("&nbsp;", "");
            string RelatedBatch = GridView1.Rows[j].Cells[4].Text.Trim().Replace("Related Batch", "").Replace("&nbsp;", "");
            string MailStatus = GridView1.Rows[j].Cells[5].Text.Trim().Replace("Mail Status", "").Replace("&nbsp;", "");
            string Receipent = GridView1.Rows[j].Cells[6].Text.Trim().Replace("Receipent", "").Replace("&nbsp;", "");
            string MailingAddressEmail = GridView1.Rows[j].Cells[7].Text.Trim().Replace("Mailing Address/ Email", "").Replace("&nbsp;", "");
            string MailPreference = GridView1.Rows[j].Cells[8].Text.Trim().Replace("Mail Preference", "").Replace("&nbsp;", "");
            string SalutationPreference = GridView1.Rows[j].Cells[9].Text.Trim().Replace("Salutation Preference", "").Replace("&nbsp;", "");
            string CreatedBy = GridView1.Rows[j].Cells[10].Text.Trim().Replace("CreatedBy", "").Replace("&nbsp;", "");
            string Createdon = GridView1.Rows[j].Cells[11].Text.Trim().Replace("Createdon", "").Replace("&nbsp;", "");

            if (chkBox.Checked)
            {
                DataRow addrow = table.NewRow();
                addrow["Mail ID"] = MailID;
                addrow["Mail Type"] = MailType;
                addrow["As Of Date"] = AsOfDate;
                addrow["Related Batch"] = RelatedBatch;
                addrow["Mail Status"] = MailStatus;
                addrow["Receipent"] = Receipent;
                addrow["Mailing Address/ Email"] = MailingAddressEmail;
                addrow["Mail Preference"] = MailPreference;
                addrow["Salutation Preference"] = SalutationPreference;
                addrow["CreatedBy"] = CreatedBy;
                addrow["Createdon"] = Createdon;
                table.Rows.Add(addrow);
            }
        }

        table.AcceptChanges();

        //copy the structure of the data table
        DataTable toExcel = table.Copy();

        //set http contex 
        HttpContext context = HttpContext.Current;

        for (int k = 0; k < toExcel.Columns.Count - 2; k++)
        {
            context.Response.Write(toExcel.Columns[k].ColumnName + ",");
        }

        context.Response.Write(Environment.NewLine);

        //Loop through rows and output
        foreach (DataRow row in toExcel.Rows)
        {
            for (int i = 0; i < toExcel.Columns.Count - 2; i++)
            {
                context.Response.Write(row[i].ToString().Replace(",", string.Empty) + ",");
            }
            context.Response.Write(Environment.NewLine);
        }


        //output the csv file

        context.Response.ContentType = "text/csv";

        context.Response.AppendHeader("Content-Disposition", "attachment; filename=Fed-Ex export file.csv");

        context.Response.End();

    }


    private void ExportGridToCSV()
    {
        sqlstr = GetData();
        DataSet ds_table = clsDB.getDataSet(sqlstr);

        Response.Clear();

        Response.Buffer = true;

        Response.AddHeader("content-disposition",

         "attachment;filename=" + lstMailType.SelectedItem.Text + ".csv");

        Response.Charset = "";

        Response.ContentType = "application/text";

        GridView1.AllowPaging = false;

        //copy the structure of the data table
        DataTable toExcel = ds_table.Tables[0];

        //set http contex 
        HttpContext context = HttpContext.Current;

        for (int k = 0; k < toExcel.Columns.Count - 2; k++)
        {
            context.Response.Write(toExcel.Columns[k].ColumnName + ",");
        }

        context.Response.Write(Environment.NewLine);

        //Loop through rows and output
        foreach (DataRow row in toExcel.Rows)
        {
            for (int i = 0; i < toExcel.Columns.Count - 2; i++)
            {
                context.Response.Write(row[i].ToString().Replace(",", string.Empty) + ",");
            }
            context.Response.Write(Environment.NewLine);
        }


        //output the csv file

        context.Response.ContentType = "text/csv";

        context.Response.AppendHeader("Content-Disposition", "attachment; filename=Fed-Ex export file.csv");

        context.Response.End();
    }

    protected void btnSubmit_Click(object sender, EventArgs e)
    {

        #region Log
        DateTime dtmain = DateTime.Now;

        LogFileName = "Log-" + DateTime.Now;
        LogFileName = LogFileName.Replace(":", "-");
        LogFileName = LogFileName.Replace("/", "-");
        LogFileName = Server.MapPath("") + @"\Logs" + "/" + LogFileName + ".txt";
        sw = new StreamWriter(LogFileName);
        //  sw.Close();

        //  lg.AddinLogFile(LogFileName, "Start Page Load " + dtmain);

        //  lg.AddinLogFile(LogFileName, "Option selected " + ddlAction.SelectedItem.ToString());
        #endregion
        string ParentFolder = string.Empty;
        string TempFolderPath = string.Empty;
        try
        {
            // bool billingFlag = Convert.ToBoolean(ViewState["BillingFlg"]);
            string NewFolderWarningMsg = "";
            //bool bListFetched = LoadList();
            bListFetched = (bool)ViewState["bListFetched"];
            if (bListFetched)
            {
                #region Create TempFolder
                string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
                string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
                string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();
                string strMilliSecond = DateTime.Now.Millisecond.ToString().Length < 2 ? "0" + DateTime.Now.Millisecond.ToString() : DateTime.Now.Millisecond.ToString();

                string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
                string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

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

                TempFolderPath = Request.MapPath("ExcelTemplate\\TempFolder") + "\\" + ParentFolder;
                //tempFolder at Local Path to create image,pdf files                
                if (!Directory.Exists(TempFolderPath))
                {
                    System.IO.Directory.CreateDirectory(TempFolderPath);
                }
                #endregion
                if (ddlAction.SelectedValue == "3") //Generate single PDF Report.
                {
                    string MailType = string.Empty;
                    string BatchTypeId = string.Empty;
                    bool bProceed = true;

                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                        if (chkSelectNC.Checked == true)
                        {
                            MailType = row.Cells[2].Text;
                            BatchTypeId = row.Cells[31].Text;

                            if (MailType != "Quarterly Statement" && BatchTypeId != "4")
                            {
                                bProceed = false;
                            }
                        }
                    }

                    if (bProceed)
                    {
                        ViewState["ConsolidatedSinglePDF"] = null;

                        bool status = GenerateConsolidatedPDF(TempFolderPath);
                        if (status == true)
                            lblError.Text = "Report generated successfully";

                        try
                        {
                            if (ViewState["ConsolidatedSinglePDF"] != null)
                            {
                                Random rnd = new Random();
                                string strRndNumber = Convert.ToString(rnd.Next(5));
                                string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + strRndNumber;
                                string FileName = "ConsolidatedSinglePDF_" + strGUID;

                                string ls = Convert.ToString(ViewState["ConsolidatedSinglePDF"]);
                                // String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + FileName + ".pdf";
                                // String fsFinalLocation = TempFolderPath + "\\" + FileName + ".pdf";
                                String fsFinalLocation = Request.MapPath("ExcelTemplate\\TempFolder") + "\\" + FileName + ".pdf";


                                FileInfo loFile = new FileInfo(ls);
                                loFile.CopyTo(fsFinalLocation.Replace(".xls", ".pdf"), true);
                                ViewState["ConsolidatedSinglePDF"] = null;

                                //Response.Write("<script>");
                                //// string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + FileName + ".pdf";
                                //string lsFileNamforFinalXls = TempFolderPath + "//" + FileName + ".pdf";
                                //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
                                //Response.Write("</script>");

                                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                Type tp = this.GetType();
                                sb.Append("\n<script type=text/javascript>\n");
                                sb.Append("\nwindow.open('ViewReport.aspx?" + FileName + ".pdf" + "', 'mywindow');");
                                sb.Append("</script>");
                                ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
                            }
                        }
                        catch (Exception exc)
                        {
                            Response.Write(exc.Message);
                        }

                    }
                    else
                        lblError.Text = "Please select only batch related mail type";

                    return;
                }


                if (ddlAction.SelectedValue == "10") //Generate single PDF Report.
                {
                    string MailType = string.Empty;
                    string BatchTypeId = string.Empty;
                    bool bProceed = true;

                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                        if (chkSelectNC.Checked == true)
                        {
                            MailType = row.Cells[2].Text;
                            BatchTypeId = row.Cells[31].Text;

                            if (MailType == "Quarterly Statement" && BatchTypeId == "4")
                            {
                                bProceed = false;
                            }

                        }
                    }

                    if (bProceed)
                    {
                        bool status = InsertMailingsheet(TempFolderPath);
                        if (status == true)
                            lblError.Text = "Report generated successfully";
                    }
                    else
                        lblError.Text = "Please select only batch related mail type";

                    return;
                }

                if (ddlAction.SelectedValue == "9") //Create Individual PDFs –Grouped by Household and Recipient
                {
                    string MailType = string.Empty;
                    string BatchTypeId = string.Empty;
                    bool bProceed = true;

                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                        if (chkSelectNC.Checked == true)
                        {
                            MailType = row.Cells[2].Text;
                            BatchTypeId = row.Cells[31].Text;

                            if (MailType != "Quarterly Statement" && BatchTypeId != "4")
                            {
                                bProceed = false;
                            }
                        }
                    }

                    if (bProceed)
                    {
                        GroupBy(TempFolderPath);
                        //bool status = GenerateMergeTypeConsolidatedPDF();
                        //if (status == true)
                        lblError.Text = "Report generated successfully";
                    }
                    else
                        lblError.Text = "Please select batch having type Merge.";
                    return;
                }



                if (ddlAction.SelectedValue == "7") //Send Email to Associate and Mark Sent
                {
                    string MailType = string.Empty;
                    bool bProceed = true;

                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                        if (chkSelectNC.Checked == true)
                        {
                            MailType = row.Cells[8].Text;

                            string strMailType = "EMAIL|REGULAR MAIL AND EMAIL|FEDEX - 2DAY AND EMAIL|FEDEX - OVERNIGHT AND EMAIL";
                            string[] strType = strMailType.Split('|');


                            for (int i = 0; i < strType.Length; i++)
                            {
                                if (MailType.ToUpper().Contains("EMAIL"))
                                {
                                    bProceed = true;
                                }
                                else
                                {
                                    bProceed = false;
                                }
                            }



                            //if (MailType.ToUpper() !=  || MailType.ToUpper() != "" || MailType.ToUpper() != "" || MailType.ToUpper() != "")
                            //{
                            //    bProceed = false;
                            //}
                        }
                    }

                    if (bProceed)
                    {
                        // Continue
                    }
                    else if (bProceed == false)
                    {
                        lblError.Text = "Please select only Mail preference Email records to Send Email";
                        return;
                    }
                }

                if (ddlAction.SelectedValue == "5")//Mark all records Canceled && Mark All Records Sent
                {
                    string MailType = string.Empty;
                    string BatchStatus = string.Empty;// Added By Rohit
                    string MailStatus = string.Empty;// Added By Rohit
                    bool bProceed = true;

                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                        if (chkSelectNC.Checked == true)
                        {
                            MailType = row.Cells[2].Text;
                            BatchStatus = row.Cells[24].Text.Trim().Replace("Ssi_reporttrackerstatus", "").Replace("&nbsp;", "");//Batch Status
                            MailStatus = row.Cells[25].Text.Trim().Replace("Mail StatusID", "").Replace("&nbsp;", "");//Mail Status

                            //if (MailType == "Quarterly Statement")
                            //{
                            //    bProceed = false;
                            //}
                            //else
                            if ((BatchStatus == "4" || BatchStatus == "9") || (MailStatus == "2" || MailStatus == "4" || MailStatus == "8"))
                            {        // Sent             //Final Report Created  //Printed           //Sent               //Created  
                                lblError.Text = "One or more of the reports you selected already has a finalized report sent out.  You cannot cancel those records";
                                return;
                            }
                        }
                    }

                    if (!bProceed)
                    {
                        lblError.Text = "This Action is not allowed on Batch related records";
                        return;
                    }

                }

                if (ddlAction.SelectedValue == "6")//Mark All Records Sent
                {
                    string MailType = string.Empty;
                    string BatchStatus = string.Empty;// Added By Rohit
                    bool bProceed = true;



                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                        if (chkSelectNC.Checked == true)
                        {
                            MailType = row.Cells[2].Text;
                            BatchStatus = row.Cells[24].Text.Trim().Replace("Ssi_reporttrackerstatus", "").Replace("&nbsp;", "");//Batch Status

                            if (BatchStatus != "9" && BatchStatus != "")//'Final Report Created' Added By Rohit 
                            {
                                lblError.Text = "One or more of the reports either has not been Approved or Sent.You cannot mark those records Sent";
                                return;
                            }
                        }
                    }


                    string FinalReviewList = string.Empty;
                    bool checkMailPref = false;

                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    Type tp = this.GetType();
                    sb.Append("\n<script type=text/javascript>\n");


                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                        if (chkSelectNC.Checked == true)
                        {
                            string ssi_reviewreqdbyid = row.Cells[23].Text.Trim().Replace("Ssi_ReviewReqdById", "").Replace("&nbsp;", "");
                            string ssi_reviewreqdbyidName = row.Cells[22].Text.Trim().Replace("ReviewReqdBy", "").Replace("&nbsp;", "");
                            string MailPref = row.Cells[8].Text.Trim().Replace("Mail Preference", "").Replace("&nbsp;", "");
                            string BatchName = row.Cells[4].Text.Trim().Replace("Related Batch", "").Replace("&nbsp;", "");

                            //if (MailPref.ToUpper().Contains("EMAIL") != "EMAIL" && ssi_reviewreqdbyid != "")
                            if (MailPref.ToUpper().Contains("EMAIL") && ssi_reviewreqdbyid != "")
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

                    if (checkMailPref == true)
                    {
                        if (ddlType.SelectedValue == "4" || ddlType.SelectedValue == "5")//condition added on 23 jan 2015 to solve the popup window goes behind due to this alert.
                        {
                            Session["MailQueueAlert"] = null;
                            sb.Append("\n alert('" + FinalReviewList + "');");
                            sb.Append("</script>");
                            ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                            //return;
                        }
                        else
                        {
                            Session["MailQueueAlert"] = FinalReviewList;
                        }
                    }

                }

                if (ddlAction.SelectedValue == "1" || ddlAction.SelectedValue == "11")//Save Reports to Sharepoint and client folder(1) & Save Reports to SharePoint only (11)
                {
                    string BatchName = string.Empty;
                    string MailAddress = string.Empty;
                    string ContactAddress = string.Empty;

                    string MailType = string.Empty;
                    string BatchStatus = string.Empty;// Added By Rohit
                    string DifferAddress = string.Empty;
                    bool bProceed = false;
                    bool OPSApproved = false;
                    bool checkAddress = false;

                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    Type tp = this.GetType();
                    sb.Append("\n<script type=text/javascript>\n");

                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                        if (chkSelectNC.Checked == true)
                        {
                            BatchStatus = row.Cells[24].Text.Trim().Replace("Ssi_reporttrackerstatus", "").Replace("&nbsp;", "");//Batch Status
                            string MailPref = (string)row.Cells[8].Text;
                            string HoldReport = row.Cells[27].Text.Trim().Replace("ssi_holdreport", "").Replace("&nbsp;", "");

                            BatchName = row.Cells[4].Text.Trim().Replace("Related Batch", "").Replace("&nbsp;", "");
                            MailAddress = row.Cells[7].Text.Trim().Replace("Mailing Address/Email", "").Replace("&nbsp;", "");
                            ContactAddress = row.Cells[30].Text.Trim().Replace("Contact Address/ Email", "").Replace("&nbsp;", "");

                            #region Check Address
                            if ((MailAddress != ContactAddress) && Hidden1.Value != "1" && BatchStatus == "8")
                            {
                                if (MailPref.ToUpper().Contains("EMAIL") && MailPref.ToUpper() != "Client Portal".ToUpper())
                                {
                                    if (DifferAddress != "")
                                    {
                                        DifferAddress = DifferAddress + "\\r\\n\\r\\n" + "Batch Name :" + BatchName + "\\r\\n" + "Mail Records Address :" + MailAddress + "\\r\\n" + "Contact Address :" + ContactAddress;
                                    }
                                    else
                                    {
                                        DifferAddress = "Batch Name :" + BatchName + "\\r\\n" + "Mail Records Address :" + MailAddress + "\\r\\n" + "Contact Address :" + ContactAddress;
                                    }

                                    checkAddress = true;
                                }
                            }

                            #endregion

                            if ((ddlAction.SelectedValue == "1" || ddlAction.SelectedValue == "11") && (BatchStatus == "8" || BatchStatus == "9" || BatchStatus == "4" && MailPref.ToUpper().Contains("EMAIL")))// 8 : 'OPS Approved' && HoldReport == ""
                            {
                                OPSApproved = true;
                            }

                            if ((BatchStatus == "9" || BatchStatus == "4") && Hidden1.Value != "1")//'Final Report Created' , 'Sent' Added By Rohit 
                            {
                                bProceed = true;
                            }
                            else if ((BatchStatus == "9" || BatchStatus == "4") && Hidden1.Value == "1")
                            {
                                //System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                //Type tp = this.GetType();
                                //sb.Append("\n<script type=text/javascript>\n");
                                sb.Append("\nwindow.document.getElementById('Hidden1').value=''");
                                // sb.Append("</script>");
                                // ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                                //return;
                            }
                        }
                    }


                    if (bProceed)
                    {
                        sb.Append("var bt = window.document.getElementById('btnSubmit');\n");
                        sb.Append("if(confirm('One or more of the reports you selected already has a finalized report; Do you want to continue?'))\n{");
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

                    if (checkAddress)
                    {

                        DifferAddress = DifferAddress.Replace("\n", " ");

                        sb.Append("\n alert('" + DifferAddress + "');");
                        sb.Append("</script>");
                        ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());

                        //sb.Append("var bt = window.document.getElementById('btnSubmit');\n");
                        //sb.Append("if(alert('" + DifferAddress + "');\n{");
                        //sb.Append("\nwindow.document.getElementById('Hidden1').value='1';");
                        //sb.Append(("\nbt.click();\n"));
                        //sb.Append("\n}");
                        //sb.Append("else\n{");
                        //sb.Append(("\nwindow.document.getElementById('Hidden1').value='0';"));
                        //sb.Append("\n}");
                        //sb.Append("</script>");
                        //ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                        //return;
                    }



                    if (!OPSApproved)
                    {
                        sb.Append("\n alert('One or more of your reports were not created because they have not been approved'); ");
                        sb.Append("</script>");
                        ClientScript.RegisterStartupScript(tp, "Script", sb.ToString());
                        return;
                    }



                }

                if (ddlAction.SelectedValue == "8") //Mark Print
                {
                    string MailType = string.Empty;
                    string chkBatchName = string.Empty;
                    bool bCheckMail = false;

                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                        chkBatchName = row.Cells[4].Text.Trim().Replace("Related Batch", "").Replace("&nbsp;", "");
                        if (chkSelectNC.Checked == true)
                        {
                            MailType = row.Cells[2].Text;

                            if (MailType.ToUpper() == "QUARTERLY STATEMENT")
                            {
                                if (chkBatchName != "")
                                {
                                    chkBatchName = chkBatchName + "<br/>" + chkBatchName;
                                }
                                else
                                {
                                    chkBatchName = chkBatchName;
                                }

                                bCheckMail = true;
                            }
                        }
                    }

                    if (bCheckMail == true)
                    {
                        Message = "";
                        Message = "The following records are related to batch.Those records cannot be updated with this action : <br/>" + chkBatchName;
                    }
                }


                #region Create CRM Service
                string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
                                                                                           //string crmServerURL = "http://server:5555/";
                string orgName = "GreshamPartners";
                //string orgName = "Webdev";
                IOrganizationService service = null;

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
                    lblError.Text = strDescription;
                }
                catch (Exception exc)
                {
                    bProceed = false;
                    strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                    lblError.Text = strDescription;
                }

                //service.PreAuthenticate = true;
                //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

                #endregion
                string fileName = "Output.CSV";
                // ssi_mailrecords objMailRecords = null;
                Entity objMailRecords = null;

                DataSet ds = null;
                string MailIdNmbListTxt = string.Empty;
                string strExistingFiles = string.Empty;
                string strNewFiles = string.Empty;
                string strFolderExists = string.Empty;
                int Count = 0;

                List<string> SucessClientName = new List<string>();

                //List<string> SucessClientName1 = new List<string>();

                List<string> SucessFileName = new List<string>();
                List<string> FailClientName = new List<string>();
                List<string> FailFileName = new List<string>();
                List<string> ExistingFileName = new List<string>();
                List<string> ListBatchName = new List<string>();

                foreach (GridViewRow row in GridView1.Rows)
                {
                    // lg.AddinLogFile(LogFileName, "foreach loop  " );
                    CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                    string ssi_batchid = row.Cells[12].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                    string ssi_mailrecordsId = row.Cells[13].Text.Trim().Replace("Ssi_reporttrackerstatus", "").Replace("&nbsp;", "");
                    string BatchStatus = row.Cells[24].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                    string ssi_secondaryownerid = row.Cells[26].Text.Trim().Replace("ssi_secondaryownerid", "").Replace("&nbsp;", "");
                    string ssi_ReviewReqByid = row.Cells[23].Text.Trim().Replace("Ssi_ReviewReqdById", "").Replace("&nbsp;", "");
                    string ssi_internalbillingcontactid = row.Cells[29].Text.Trim().Replace("ssi_internalbillingcontactid", "").Replace("&nbsp;", "");
                    string HoldReport = row.Cells[27].Text.Trim().Replace("ssi_holdreport", "").Replace("&nbsp;", "");
                    string BillingHandedOff = row.Cells[28].Text.Trim().Replace("ssi_BillingHandedOff", "").Replace("&nbsp;", "");
                    string MailAddress = row.Cells[7].Text.Trim().Replace("Mailing Address/Email", "").Replace("&nbsp;", "");
                    string ContactAddress = row.Cells[30].Text.Trim().Replace("Contact Address/ Email", "").Replace("&nbsp;", "");
                    string MailType = row.Cells[2].Text.Trim().Replace("Mail Type", "").Replace("&nbsp;", "");
                    ReviewRequiredBy = ssi_ReviewReqByid; // pass the value to batch status log function
                                                          // string FinalSharepointFolder = (string)row.Cells[20].Text == "&nbsp;" || (string)row.Cells[20].Text == "" ? "" : (string)row.Cells[20].Text; // Commented New Sharepoint Client Services site Changes 

                    //New Sharepoint Client Services site Changes 
                    string ssi_CSCiteUUID = row.Cells[40].Text.Replace("&nbsp;", "");
                    string LegalEntityUUID = row.Cells[41].Text.Replace("&nbsp;", "");
                    string ssi_SPLEFolder = row.Cells[42].Text.Replace("&nbsp;", "").Replace("&#39;", "'").Replace("/", "_").Replace("#", "No.").Replace("*", "").Replace(":", "").Replace("<", "").Replace(">", "").Replace("?", "").Replace("\"", "").Replace("|", "");
                    string ssi_SPSiteType = row.Cells[43].Text.Replace("&nbsp;", "");


                    //billing changes

                    string ssi_billinginvoiceid = row.Cells[44].Text.Replace("&nbsp;", "");
                    string ssi_billingid = row.Cells[45].Text.Replace("&nbsp;", "");


                    string ClientPortalPAth = (string)row.Cells[19].Text == "&nbsp;" || (string)row.Cells[19].Text == "" ? "" : (string)row.Cells[19].Text;

                    string BatchFileName = (string)row.Cells[21].Text.Replace("&#39;", "'").ToString();




                    string ssi_clientportalname = row.Cells[39].Text;
                    ssi_clientportalname = ssi_clientportalname.Replace("&nbsp;", "");
                    ssi_clientportalname = ssi_clientportalname.Replace("&#39;", "'");

                    if (BatchFileName.Contains("#"))
                    {
                        //string[] rndremove = BatchFileName.Split('.');
                        //if (rndremove.Length == 2)
                        //    BatchFileName = rndremove[0].Remove(rndremove[0].Length - 4, 4) + rndremove[1];
                        BatchFileName = BatchFileName.Remove(BatchFileName.Length - 9, 5);
                    }

                    try
                    {
                        if (chkSelectNC.Checked == true)
                        {
                            if (ssi_CSCiteUUID == "" && ClientPortalPAth == "" && ClientPortalPAth == null && ssi_CSCiteUUID == null)
                            {
                                lblError3.Visible = true;
                                lblError3.Text = "ClientServices and Clientportal Empty, Contact Administrator";

                            }
                            string UserId = GetcurrentUser();
                            int intStatus = 0;
                            string MailPref = (string)row.Cells[8].Text;
                            //objMailRecords = new ssi_mailrecords();
                            objMailRecords = new Entity("ssi_mailrecords");


                            //objMailRecords.ssi_mailrecordsid = new Key();
                            //objMailRecords.ssi_mailrecordsid.Value = new Guid(ssi_mailrecordsId);
                            objMailRecords["ssi_mailrecordsid"] = new Guid(ssi_mailrecordsId);

                            //OPS Approved
                            if (ddlAction.SelectedValue == "1" && (BatchStatus == "8" || BatchStatus == "9" || BatchStatus == "4"))// Save Reports to SharePoint and Client Portal //&& HoldReport == ""
                            {
                                // lg.AddinLogFile(LogFileName, "Save Reports to SharePoint and Client Portal ");
                                #region OLD Sharepoint Code commented after New client Sevices Site 
                                //#region Update Batch Status to 'Final Report Sent'

                                ////ssi_batch objBatch = new ssi_batch();
                                //Entity objBatch = new Entity("ssi_batch");

                                ////objBatch.ssi_batchid = new Key();
                                ////objBatch.ssi_batchid.Value = new Guid(ssi_batchid);
                                //objBatch["ssi_batchid"] = new Guid(ssi_batchid);

                                //if (FinalSharepointFolder != "")
                                //{
                                //    //objBatch.ssi_sharepointreportfolderfinal = FinalSharepointFolder.Replace(" Family", "").Replace(",", "%2C").Replace(" ", "%20").Replace("'", "%27").Replace("&#39;", "'").ToString();
                                //    objBatch["ssi_sharepointreportfolderfinal"] = FinalSharepointFolder.Replace(" Family", "").Replace(",", "%2C").Replace(" ", "%20").Replace("'", "%27").Replace("&#39;", "'").ToString();
                                //}
                                //if (row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace("&nbsp;", "") != "")
                                //{
                                //    //objBatch.ssi_sharepointemaillink = (string)row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                //    objBatch["ssi_sharepointemaillink"] = (string)row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                //}
                                //else
                                //{
                                //    //objBatch.ssi_sharepointemaillink = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(" ", "%20");
                                //    objBatch["ssi_sharepointemaillink"] = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(" ", "%20");
                                //}

                                //if (MailPref.ToUpper().Contains("EMAIL"))
                                //{
                                //    //objBatch.ssi_finalreportcreatedflag = new CrmBoolean();
                                //    //objBatch.ssi_finalreportcreatedflag.Value = true;
                                //    objBatch["ssi_finalreportcreatedflag"] = true;

                                //    updateSentData(ssi_secondaryownerid, ssi_mailrecordsId);
                                //}
                                //else
                                //{
                                //    //objBatch.ssi_sendemailib = new CrmBoolean();
                                //    //objBatch.ssi_sendemailib.Value = true;
                                //}

                                ////objBatch.ssi_sendemailib = new CrmBoolean();
                                ////objBatch.ssi_sendemailib.Value = true;
                                //objBatch["ssi_sendemailib"] = true;


                                //if (MailPref.ToUpper().Contains("EMAIL") && ssi_ReviewReqByid != "")
                                //{
                                //    // Send Email to Review Required By 
                                //    //objBatch.ssi_reviewrequiredbyid = new Lookup();
                                //    //objBatch.ssi_reviewrequiredbyid.type = EntityName.systemuser.ToString();
                                //    //objBatch.ssi_reviewrequiredbyid.Value = new Guid(ssi_ReviewReqByid);
                                //    objBatch["ssi_reviewrequiredbyid"] = new EntityReference("systemuser", new Guid(ssi_ReviewReqByid));

                                //    //objBatch.ssi_sendrrbymail = new CrmBoolean();
                                //    //objBatch.ssi_sendrrbymail.Value = true;
                                //    objBatch["ssi_sendrrbymail"] = true;

                                //    //objBatch.ssi_reporttrackerstatus = new Picklist();
                                //    //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Sent'
                                //    objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                //    // intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                //    intStatus = 4;



                                //    //BatchReportStatus(intStatus, UserId, ssi_batchid, BillingHandedOff);
                                //    BillingHandedOff = "true";
                                //}
                                //else if (MailPref.ToUpper().Contains("EMAIL") && ReviewRequiredBy == "")
                                //{
                                //    // Send Email to Associate
                                //    //objBatch.ssi_reporttrackerstatus = new Picklist();
                                //    //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Sent'
                                //    objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                //    //objBatch.ssi_sendmailassociate = new CrmBoolean();
                                //    //objBatch.ssi_sendmailassociate.Value = true;
                                //    objBatch["ssi_sendmailassociate"] = true;

                                //    //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                //    intStatus = 4;
                                //    //BatchReportStatus(intStatus, UserId, ssi_batchid, BillingHandedOff);
                                //    BillingHandedOff = "true";
                                //}
                                //else
                                //{
                                //    int trackerstatus;
                                //    if (ssi_internalbillingcontactid != "")
                                //    {
                                //        //objBatch.ssi_billinghandedoff = new CrmBoolean();
                                //        //objBatch.ssi_billinghandedoff.Value = true;
                                //        objBatch["ssi_billinghandedoff"] = true;
                                //        BillingHandedOff = "true";
                                //    }
                                //    else
                                //    {
                                //        //objBatch.ssi_billinghandedoff = new CrmBoolean();
                                //        //objBatch.ssi_billinghandedoff.Value = false;
                                //        objBatch["ssi_billinghandedoff"] = false;
                                //        BillingHandedOff = "false";
                                //    }


                                //    if (MailPref.ToUpper() == "Client Portal".ToUpper())
                                //    {
                                //        //objBatch.ssi_reporttrackerstatus = new Picklist();
                                //        //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Final Report Sent'
                                //        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                //        trackerstatus = 4;

                                //        //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                //        updateSentData(ssi_ReviewReqByid, ssi_mailrecordsId);
                                //    }
                                //    else
                                //    {
                                //        //objBatch.ssi_reporttrackerstatus = new Picklist();
                                //        //objBatch.ssi_reporttrackerstatus.Value = 9;// Batch status 'Final Report Sent'
                                //        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(9);
                                //        trackerstatus = 9;
                                //    }

                                //    //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                //    intStatus = trackerstatus;

                                //}

                                //if (ssi_ReviewReqByid != "" && MailPref.ToUpper().Contains("EMAIL"))
                                //{
                                //    lblError.Text = "Some of the reports you selected need to be given to the person designated in Review required by.";
                                //}

                                //service.Update(objBatch);

                                //bool billingHandedOff = Convert.ToBoolean(BillingHandedOff);
                                //BatchReportStatus(intStatus, UserId, ssi_batchid, billingHandedOff);

                                //#endregion

                                //#region Update Mail status to 'Created'
                                //if (MailPref.ToUpper().Contains("EMAIL") && ssi_ReviewReqByid != "")
                                //{
                                //    //objMailRecords.ssi_mailstatus = new Picklist();
                                //    //objMailRecords.ssi_mailstatus.Value = 3;//mail status 'Sent to FINAL Reviewer'
                                //    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(3);
                                //}
                                //else if ((MailPref.ToUpper().Contains("EMAIL") || MailPref.ToUpper() == "Client Portal".ToUpper()) && ssi_ReviewReqByid == "")
                                //{
                                //    //objMailRecords.ssi_mailstatus = new Picklist();
                                //    //objMailRecords.ssi_mailstatus.Value = 4;//mail status 'Sent to FINAL Reviewer'
                                //    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                //}
                                //else if (MailPref.ToUpper() == "Client Portal".ToUpper() && ssi_ReviewReqByid != "")
                                //{
                                //    //objMailRecords.ssi_mailstatus = new Picklist();
                                //    //objMailRecords.ssi_mailstatus.Value = 4;//mail status 'Sent to FINAL Reviewer'
                                //    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                //}
                                //else
                                //{
                                //    //objMailRecords.ssi_mailstatus = new Picklist();
                                //    //objMailRecords.ssi_mailstatus.Value = 8;//mail status 'Created'
                                //    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(8);

                                //}

                                //#region Update Address not used
                                ////if (MailAddress != ContactAddress)
                                ////{
                                ////    string[] NewAddress = ContactAddress.Split('\n');
                                ////    //Address1
                                ////    if (NewAddress[0] != "")
                                ////    {
                                ////        objMailRecords.ssi_addressline1_mail = NewAddress[0];
                                ////    }
                                ////    //Address2
                                ////    if (NewAddress[1] != "")
                                ////    {
                                ////        objMailRecords.ssi_addressline2_mail = NewAddress[1];
                                ////    }
                                ////    //Address3
                                ////    if (NewAddress[2] != "")
                                ////    {
                                ////        objMailRecords.ssi_addressline3_mail = NewAddress[2];
                                ////    }
                                ////    //City
                                ////    if (NewAddress[3] != "")
                                ////    {
                                ////        objMailRecords.ssi_city_mail = NewAddress[3];
                                ////    }

                                ////    // State Province
                                ////    if (NewAddress[4] != "")
                                ////    {
                                ////        objMailRecords.ssi_stateprovince_mail = NewAddress[4];
                                ////    }

                                ////    // Zip Code
                                ////    if (NewAddress[5] != "")
                                ////    {
                                ////        objMailRecords.ssi_zipcode_mail = NewAddress[5];
                                ////    }

                                ////}
                                //#endregion

                                //service.Update(objMailRecords);
                                //#endregion

                                //string BatchName = string.Empty;
                                //string SubFolderName = string.Empty;
                                //string DestinationFileName = string.Empty;
                                //string SubFolder = string.Empty;

                                //BatchName = (string)row.Cells[4].Text;
                                //SubFolderName = (string)row.Cells[32].Text;

                                ////\\GRPAO1-VWFS01\shared$\Mail Merge\Completed Mailings
                                ////\\GRPAO1-VWFS01\shared$\Mail Merge\Completed Mailings
                                //SubFolder = "\\\\GRPAO1-VWFS01\\shared$\\Mail Merge\\Completed Mailings\\" + SubFolderName;
                                ////SubFolder = "\\\\GRPAO1-VWFS01\\opsreports$\\Mail Merge\\Completed Mailings\\" + SubFolderName;

                                //string BatchFilePath = (string)row.Cells[18].Text;
                                //string AsOfDate = (string)row.Cells[3].Text;
                                //string[] date = AsOfDate.Split(new char[] { '/' });
                                //string year = date[2];
                                ////string BatchFileName = (string)row.Cells[21].Text;  // BatchFilePath.Substring(BatchFilePath.LastIndexOf("\\") + 1); // 
                                //string SPVFileName = row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace("&nbsp;", "");

                                ////BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "");
                                //if (chkTest.Checked == true)
                                //{
                                //    if (SPVFileName != "")
                                //    {
                                //        BatchFileName = SPVFileName.Replace(".pdf", "_Test.pdf");
                                //    }
                                //    else
                                //    {
                                //        BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(".pdf", "_Test.pdf");
                                //    }
                                //}
                                //else
                                //{
                                //    if (SPVFileName != "")
                                //    {
                                //        BatchFileName = SPVFileName;
                                //    }
                                //    else
                                //    {
                                //        BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "");
                                //    }
                                //}


                                //string ClientFolder = (string)row.Cells[19].Text == "&nbsp;" || (string)row.Cells[19].Text == "" ? "" : (string)row.Cells[19].Text;
                                //string SharepointFolder = (string)row.Cells[20].Text == "&nbsp;" || (string)row.Cells[20].Text == "" ? "" : (string)row.Cells[20].Text;

                                //ClientFolder = ClientFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();
                                //SharepointFolder = SharepointFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();


                                //string ClientFolderFilePath = ClientFolder + "\\" + BatchFileName;
                                ////string SharepointFolderFilePath = SharepointFolder + "\\" + BatchFileName;

                                ////string ClientFolderFilePath = ClientFolder + "//" + BatchFileName;
                                //string SharepointFolderFilePath = SharepointFolder + "//" + BatchFileName;

                                //string SubFolderFilePath = SubFolder;

                                ////string SharepointFolder = "C:\\Reports" + "\\" + BatchFileName;

                                //ClientFolderFilePath = ClientFolderFilePath.Replace("%20", " ");
                                //SharepointFolderFilePath = SharepointFolderFilePath.Replace("%20", " ");

                                //string strSubFolderPath = SubFolderFilePath + "\\" + BatchFileName;
                                //strSubFolderPath = strSubFolderPath.Replace("%20", " ").Replace("&#39;", "'").ToString();

                                //#region Not in Use

                                ////"\\\\sp02\\DavWWWRoot\\ClientServ\\Documents\\Clients\\Active\\Anathan\\Correspondence\\Quarterly AXYS Reports\\" + BatchFileName;

                                ////string SharepointFolder = "\\\\GRPAO1-VWFS01\\_ops_C_I_R_group\\AdventReport\\" + "Anathan Family-Anathan G_2011-0930_2011_11_25_07_48.pdf";

                                ////try
                                ////{

                                ////if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && SharepointFolder != "" && SharepointFolder != "&nbsp;")
                                ////{
                                ////    if (File.Exists(SharepointFolder))
                                ////    {
                                ////        if (strExistingFiles == "")
                                ////            strExistingFiles = "<br/> <a href='" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + " </a>";
                                ////        else
                                ////            strExistingFiles = strExistingFiles + ",<br/><a href='" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + " </a>";
                                ////    }
                                ////    else
                                ////    {
                                ////        if (strNewFiles == "")
                                ////            strNewFiles = "<br/> <a href='" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "</a>";
                                ////        else
                                ////            strNewFiles = strNewFiles + ",<br/><a href='" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//") + "'>" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "</a>";
                                ////    }

                                ////   //File.Copy(BatchFilePath, SharepointFolder, true);
                                ////   // Response.Write("<br/>File Copied: " + SharepointFolder +"<br/>");
                                ////}
                                ////else
                                ////    SharepointFolder = "";

                                ////if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && ClientFolder != "&nbsp;" && ClientFolder != "")
                                ////{
                                ////    if (File.Exists(ClientFolder))
                                ////    {
                                ////        if (strExistingFiles == "")
                                ////            strExistingFiles = "<br/><a href='" + "http" + ":" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + " </a>";
                                ////        else
                                ////            strExistingFiles = strExistingFiles + ",<br/> <a href='" + "http" + "://" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "</a>";
                                ////    }
                                ////    else
                                ////    {
                                ////        if (strNewFiles == "")
                                ////            strNewFiles = "<br/> <a href='" + "http" + "://" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "</a>";
                                ////        else
                                ////            strNewFiles = strNewFiles + ",<br/><a href='" + "http" + "://" + ClientFolder.Replace(" ", "%20").Replace("\\", "//") + "'>" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "</a>";
                                ////    }
                                ////    //File.Copy(BatchFilePath, ClientFolder, true);
                                ////   // Response.Write("<br/>File Copied: " + SharepointFolder+"<br/>");
                                ////}
                                ////else
                                ////    ClientFolder = "";

                                ////if (strExistingFiles != "")
                                ////    lblError.Text = lblError.Text + "<br/>Batch Report Saved to Sharepoint folder, " + strExistingFiles + "<br/> above files are overwritten.";
                                ////else if (strNewFiles != "")
                                ////    lblError.Text = lblError.Text + "<br/>New Batch Report added to Sharepoint folder, " + strNewFiles;
                                ////else
                                ////    lblError.Text = lblError.Text + "<br/>Batch Report Saved to Sharepoint folder";

                                //#endregion
                                ////string Batch_Namedf1er = row.Cells[37].Text;
                                ////string BatchType1222222 = row.Cells[38].Text;
                                //// string Batch_Name1er = row.Cells[39].Text;


                                //string Batch_Name1 = row.Cells[38].Text;

                                //if ((BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                //{
                                //    if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && SharepointFolderFilePath != "" && SharepointFolderFilePath != "&nbsp;")
                                //    {
                                //        BatchName = (string)row.Cells[4].Text;
                                //        if (SharepointFolder != "")
                                //        {


                                //            if (SharepointFolder.ToLower().Contains("clientserv"))
                                //            {
                                //                string newsharepointPath = sharepointFolderPath(SharepointFolder);
                                //                string sharepointpath = newsharepointPath;
                                //                newsharepointPath = "https://greshampartners.sharepoint.com/clientserv/" + newsharepointPath;



                                //                // CopyFile(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);

                                //                bool IsSharepointFolderExists = CheckFolderPathExists(SharepointFolder);
                                //                if (IsSharepointFolderExists)
                                //                {
                                //                    try
                                //                    {
                                //                        if (checkSharepouintFileExist(sharepointpath, BatchFileName))
                                //                        {
                                //                            if (strExistingFiles == "")
                                //                            {
                                //                                strExistingFiles = "<br/>" + newsharepointPath + "/" + BatchFileName;
                                //                            }
                                //                            else
                                //                                strExistingFiles = strExistingFiles + ",<br/>" + newsharepointPath + "/" + BatchFileName;


                                //                            // ExistingFileName.Add(newsharepointPath + "/" + BatchFileName);
                                //                        }

                                //                        CopyFilenew(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);

                                //                        //     strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;

                                //                        SucessClientName.Add(Batch_Name1);
                                //                        SucessFileName.Add(newsharepointPath);

                                //                    }
                                //                    catch (Exception exc)
                                //                    {
                                //                        // strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                        Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                //               " to " + SharepointFolderFilePath + "<br/>" + exc.Message + ", " + exc.StackTrace);

                                //                        //FailClientName.Add();
                                //                        //FailFileName.Add

                                //                    }

                                //                }
                                //                else
                                //                {
                                //                    strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                }
                                //            }
                                //            else if (SharepointFolder.ToLower().Contains("clientportal"))
                                //            {
                                //                string HouseHoldName = row.Cells[15].Text;
                                //                HouseHoldName = HouseHoldName.Replace(" Family", "");
                                //                HouseHoldName = HouseHoldName.Replace(" family", "");
                                //                HouseHoldName = HouseHoldName.Replace(" FAMILY", "");
                                //                //LastNameSort = GridView1.Rows[j].Cells[36].Text;//36
                                //                //var Name = HouseHold.Split(' ');
                                //                //string HouseHoldName = Name[0];

                                //                ssi_clientportalname = ssi_clientportalname.Replace(" Family", "");
                                //                ssi_clientportalname = ssi_clientportalname.Replace(" family", "");
                                //                ssi_clientportalname = ssi_clientportalname.Replace(" FAMILY", "");

                                //                string BatchType1 = row.Cells[37].Text;


                                //                //string newsharepointPath =sharepointFolderPath(SharepointFolder);
                                //                string newsharepointPath = "Documents taxonomy";
                                //                //string sharepointpath = newsharepointPath;
                                //                newsharepointPath = "https://greshampartners.sharepoint.com/ClientPortal/" + newsharepointPath;

                                //                //string BatchType = row.Cells[37].Text;
                                //                //string Batch_Name = row.Cells[38].Text;
                                //                // CopyFile(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);

                                //                try
                                //                {
                                //                    string FileName = Path.GetFileName(SharepointFolderFilePath);
                                //                    // bool result=  CopyFiletoSharepoint(SharepointFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, HouseHoldName, year, SharepointFolderFilePath);
                                //                    //  bool result = CopyFiletoSharepoint(SharepointFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, ssi_clientportalname, year, SharepointFolderFilePath);
                                //                    bool result = CopyFiletoSharepoint(SharepointFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, HouseHoldName, year, SharepointFolderFilePath, ssi_clientportalname);
                                //                    if (result)
                                //                    {
                                //                        // strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                        SucessClientName.Add(Batch_Name1);
                                //                        SucessFileName.Add(newsharepointPath);
                                //                    }
                                //                    else
                                //                    {
                                //                        // strFolderExists = strFolderExists + "<br/>Can not File in Client Portal Tag Missing, Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                        string strFailClientName = "";
                                //                        if (ssi_clientportalname != "")
                                //                        {
                                //                            strFailClientName = ssi_clientportalname;
                                //                        }
                                //                        else
                                //                        {
                                //                            strFailClientName = HouseHoldName;
                                //                        }
                                //                        FailClientName.Add(strFailClientName);
                                //                        //  FailClientName.Add(ssi_clientportalname + "," + HouseHoldName);
                                //                        FailFileName.Add(newsharepointPath);
                                //                        ListBatchName.Add(Batch_Name1);

                                //                    }
                                //                }
                                //                catch (Exception exc)
                                //                {
                                //                    // strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                    Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                //           " to " + SharepointFolderFilePath + "<br/>" + exc.Message + ", " + exc.StackTrace);
                                //                }


                                //            }
                                //            else
                                //            {
                                //                if (System.IO.File.Exists(SharepointFolderFilePath))
                                //                {
                                //                    if (strExistingFiles == "")
                                //                        strExistingFiles = "<br/>" + SharepointFolderFilePath;
                                //                    else
                                //                        strExistingFiles = strExistingFiles + ",<br/>" + SharepointFolderFilePath;
                                //                    // ExistingFileName.Add(SharepointFolderFilePath + "/" + BatchFileName);

                                //                    try
                                //                    {
                                //                        System.IO.File.Copy(BatchFilePath, SharepointFolderFilePath, true);
                                //                        // strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + SharepointFolder;
                                //                        SucessClientName.Add(BatchName);
                                //                        SucessFileName.Add(SharepointFolder);
                                //                    }
                                //                    catch (Exception ex)
                                //                    {
                                //                        Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                //                            " to " + SharepointFolderFilePath + "<br/>" + ex.Message + ", " + ex.StackTrace);
                                //                    }
                                //                    //SetFileReadAccess(SharepointFolderFilePath, true);
                                //                    //Response.Write("<br/>File Copied: " + SharepointFolder +"<br/>");
                                //                }
                                //            }
                                //        }
                                //        else
                                //        {
                                //            if (SharepointFolder.ToLower().Contains("clientserv"))
                                //            {
                                //                string newsharepointPath = sharepointFolderPath(SharepointFolder);
                                //                SharepointFolder = "https://greshampartners.sharepoint.com/clientserv/" + newsharepointPath;
                                //            }

                                //            if (strFolderExists == "")
                                //                strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + SharepointFolder;
                                //            else
                                //            {
                                //                strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + SharepointFolder;
                                //            }
                                //        }
                                //    }
                                //    else
                                //        SharepointFolderFilePath = "";
                                //}
                                //else if (SharepointFolder != "")
                                //{
                                //    BatchName = (string)row.Cells[4].Text;

                                //    //if (strFolderExists == "")
                                //    //    strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + SharepointFolder;
                                //    //else
                                //    //    strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + SharepointFolder;
                                //}

                                ////---------------------- Clientfolder Region
                                //if ((BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                //{

                                //    if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && ClientFolderFilePath != "&nbsp;" && ClientFolderFilePath != "")
                                //    {



                                //        if (ClientFolder.ToLower().Contains("clientserv"))
                                //        {
                                //            // CopyFile(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);
                                //            string newsClientFolder = sharepointFolderPath(ClientFolder);
                                //            string newClientsharepointPath = newsClientFolder;
                                //            newsClientFolder = "https://greshampartners.sharepoint.com/clientserv/" + newsClientFolder;



                                //            bool isClientFolderExists = CheckFolderPathExists(ClientFolder);
                                //            if (isClientFolderExists)
                                //            {
                                //                try
                                //                {
                                //                    if (checkSharepouintFileExist(newClientsharepointPath, BatchFileName))
                                //                    {
                                //                        if (strExistingFiles == "")
                                //                            strExistingFiles = "<br/>" + newsClientFolder + "/" + BatchFileName;
                                //                        else
                                //                            strExistingFiles = strExistingFiles + ",<br/>" + newsClientFolder + "/" + BatchFileName;


                                //                        //ExistingFileName.Add(newsClientFolder + "/" + BatchFileName);
                                //                    }


                                //                    CopyFilenew(ClientFolder, Path.GetFileName(ClientFolderFilePath), BatchFilePath);
                                //                    //  strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Client folder path: " + newsClientFolder;
                                //                    SucessClientName.Add(Batch_Name1);
                                //                    SucessFileName.Add(newsClientFolder);
                                //                }
                                //                catch (Exception exce)
                                //                {
                                //                    //  strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsClientFolder;
                                //                    Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                //            " to " + ClientFolderFilePath + "<br/>" + exce.Message + ", " + exce.StackTrace);
                                //                }

                                //            }
                                //            else
                                //            {
                                //                strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsClientFolder;
                                //            }

                                //        }
                                //        else if (ClientFolder.ToLower().Contains("clientportal"))
                                //        {
                                //            string HouseHoldName = row.Cells[15].Text;
                                //            //   string HouseHold = row.Cells[15].Text;
                                //            HouseHoldName = HouseHoldName.Replace(" Family", "");
                                //            HouseHoldName = HouseHoldName.Replace(" family", "");
                                //            HouseHoldName = HouseHoldName.Replace(" FAMILY", "");

                                //            //var Name = HouseHold.Split(' ');
                                //            //string HouseHoldName = Name[0];

                                //            string BatchType1 = row.Cells[37].Text;
                                //            //string Batch_Name1 = row.Cells[38].Text;

                                //            // CopyFile(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);
                                //            string newsClientFolder = "Documents taxonomy";
                                //            string newClientsharepointPath = newsClientFolder;
                                //            newsClientFolder = "https://greshampartners.sharepoint.com/ClientPortal/" + newsClientFolder;


                                //            try
                                //            {
                                //                string FileName = Path.GetFileName(SharepointFolderFilePath);
                                //                //bool result=  CopyFiletoSharepoint(ClientFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, HouseHoldName, year, SharepointFolderFilePath);\
                                //                //    bool result = CopyFiletoSharepoint(SharepointFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, ssi_clientportalname, year, SharepointFolderFilePath);
                                //                bool result = CopyFiletoSharepoint(SharepointFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, HouseHoldName, year, SharepointFolderFilePath, ssi_clientportalname);
                                //                if (result)
                                //                {
                                //                    // strFolderExists = strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsClientFolder;
                                //                    SucessClientName.Add(Batch_Name1);
                                //                    SucessFileName.Add(newsClientFolder);
                                //                }
                                //                else
                                //                {
                                //                    // strFolderExists = strFolderExists + "<br/>Can not File in Client Portal Tag Missing, Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsClientFolder;
                                //                    //  FailClientName.Add(BatchName);
                                //                    //FailClientName.Add(ssi_clientportalname + "," + BatchName);
                                //                    //FailFileName.Add(newsClientFolder);
                                //                    // FailClientName.Add(ssi_clientportalname + "," + HouseHoldName);
                                //                    string strFailClientName = "";
                                //                    if (ssi_clientportalname != "")
                                //                    {
                                //                        strFailClientName = ssi_clientportalname;
                                //                    }
                                //                    else
                                //                    {
                                //                        strFailClientName = HouseHoldName;
                                //                    }
                                //                    FailClientName.Add(strFailClientName);
                                //                    FailFileName.Add(newsClientFolder);
                                //                    ListBatchName.Add(Batch_Name1);
                                //                }

                                //            }
                                //            catch (Exception exce)
                                //            {
                                //                //  strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsClientFolder;
                                //                Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                //        " to " + ClientFolderFilePath + "<br/>" + exce.Message + ", " + exce.StackTrace);
                                //            }



                                //        }
                                //        else
                                //        {
                                //            try
                                //            {
                                //                if (System.IO.File.Exists(ClientFolderFilePath))
                                //                {
                                //                    if (strExistingFiles == "")
                                //                        strExistingFiles = "<br/>" + ClientFolderFilePath;
                                //                    else
                                //                        strExistingFiles = strExistingFiles + ",<br/>" + ClientFolderFilePath;


                                //                    // ExistingFileName.Add(SharepointFolderFilePath + "/" + BatchFileName);
                                //                    System.IO.File.Copy(BatchFilePath, ClientFolderFilePath, true);
                                //                    strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Client folder path: " + ClientFolder;
                                //                }

                                //                //System.IO.File.Copy(BatchFilePath, ClientFolderFilePath, true);
                                //                //strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Client folder path: " + ClientFolder;
                                //            }
                                //            catch (Exception ex)
                                //            {
                                //                Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                //                         " to " + ClientFolderFilePath + "<br/>" + ex.Message + ", " + ex.StackTrace);

                                //            }
                                //        }
                                //        //SetFileReadAccess(ClientFolderFilePath, true);
                                //        //Response.Write("<br/>File Copied: " + ClientFolder + "<br/>");
                                //    }
                                //    else
                                //        ClientFolderFilePath = "";
                                //}
                                //else if (ClientFolder != "")
                                //{

                                //    BatchName = (string)row.Cells[4].Text;

                                //    if (strFolderExists == "")
                                //    {
                                //        strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Client folder path: " + ClientFolder;
                                //    }
                                //    else
                                //        strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Client folder path: " + ClientFolder;
                                //}

                                ////---------------------- SubFolder Region
                                //if (Directory.Exists(SubFolderFilePath) && (BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                //{

                                //    if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && SubFolderFilePath != "&nbsp;" && SubFolderFilePath != "")
                                //    {
                                //        if (System.IO.File.Exists(SubFolderFilePath))
                                //        {
                                //            if (strExistingFiles == "")
                                //                strExistingFiles = "<br/>" + strSubFolderPath;
                                //            else
                                //                strExistingFiles = strExistingFiles + ",<br/>" + strSubFolderPath;

                                //            //ExistingFileName.Add(SharepointFolderFilePath + "/" + BatchFileName);
                                //        }

                                //        try
                                //        {
                                //            System.IO.File.Copy(BatchFilePath, strSubFolderPath, true);
                                //        }
                                //        catch (Exception ex)
                                //        { }
                                //        //SetFileReadAccess(ClientFolderFilePath, true);
                                //        //Response.Write("<br/>File Copied: " + ClientFolder + "<br/>");
                                //    }

                                //}
                                //else if (SubFolder != "")
                                //{
                                //    Directory.CreateDirectory(SubFolderFilePath);
                                //    if (Directory.Exists(SubFolderFilePath) && (BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                //    {
                                //        if (System.IO.File.Exists(strSubFolderPath))
                                //        {
                                //            if (strExistingFiles == "")
                                //                strExistingFiles = "<br/>" + strSubFolderPath;
                                //            else
                                //                strExistingFiles = strExistingFiles + ",<br/>" + strSubFolderPath;

                                //            //ExistingFileName.Add(SharepointFolderFilePath + "/" + BatchFileName);
                                //        }

                                //        try
                                //        {
                                //            System.IO.File.Copy(BatchFilePath, strSubFolderPath, true);
                                //        }
                                //        catch (Exception ex)
                                //        { }

                                //    }
                                //}

                                #endregion

                                string BatchName = string.Empty;
                                string SubFolderName = string.Empty;
                                string DestinationFileName = string.Empty;
                                string SubFolder = string.Empty;

                                BatchName = (string)row.Cells[4].Text;
                                SubFolderName = (string)row.Cells[32].Text;
                                string CompletedMailings = AppLogic.GetParam(AppLogic.ConfigParam.CompletedMailings);
                                //\\GRPAO1-VWFS01\shared$\Mail Merge\Completed Mailings
                                //\\GRPAO1-VWFS01\shared$\Mail Merge\Completed Mailings
                                // SubFolder = "\\\\GRPAO1-VWFS01\\shared$\\Mail Merge\\Completed Mailings\\" + SubFolderName;
                                SubFolder = CompletedMailings + SubFolderName; // shared drive Changes- 7_4_2019
                                                                               //SubFolder = "\\\\GRPAO1-VWFS01\\opsreports$\\Mail Merge\\Completed Mailings\\" + SubFolderName;

                                string BatchFilePath = (string)row.Cells[18].Text;

                                // string BatchFilePath = Server.MapPath("") + @"\ExcelTemplate\";



                                string AsOfDate = (string)row.Cells[3].Text;
                                string[] date = AsOfDate.Split(new char[] { '/' });
                                string year = date[2];

                                // New Sharepoint Client Services site Changes 
                                DateTime dtAsOfDate = Convert.ToDateTime(AsOfDate);
                                string Quarter = GetQuarter(dtAsOfDate);

                                //string BatchFileName = (string)row.Cells[21].Text;  // BatchFilePath.Substring(BatchFilePath.LastIndexOf("\\") + 1); // 
                                string SPVFileName = row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace("&nbsp;", "");

                                //BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "");
                                if (chkTest.Checked == true)
                                {
                                    if (SPVFileName != "")
                                    {
                                        // New Sharepoint Client Services site Changes 
                                        //   BatchFileName = SPVFileName.Replace(".pdf", "_Test.pdf");
                                        BatchFileName = "zzTest_" + SPVFileName;
                                    }
                                    else
                                    {
                                        // New Sharepoint Client Services site Changes 
                                        //BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(".pdf", "_Test.pdf");
                                        BatchFileName = "zzTest_" + BatchFileName.Replace(" Family", "").Replace(",", "");
                                    }
                                }
                                else
                                {
                                    if (SPVFileName != "")
                                    {
                                        BatchFileName = SPVFileName;
                                    }
                                    else
                                    {
                                        BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "");
                                    }
                                }


                                string ClientFolder = (string)row.Cells[19].Text == "&nbsp;" || (string)row.Cells[19].Text == "" ? "" : (string)row.Cells[19].Text;
                                //Commented New Sharepoint Client Services site Changes 
                                //  string SharepointFolder = (string)row.Cells[20].Text == "&nbsp;" || (string)row.Cells[20].Text == "" ? "" : (string)row.Cells[20].Text;

                                ClientFolder = ClientFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();
                                //Commented New Sharepoint Client Services site Changes 
                                //SharepointFolder = SharepointFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();


                                string ClientFolderFilePath = ClientFolder + "\\" + BatchFileName;
                                //string SharepointFolderFilePath = SharepointFolder + "\\" + BatchFileName;

                                //string ClientFolderFilePath = ClientFolder + "//" + BatchFileName;
                                //Commented New Sharepoint Client Services site Changes 
                                //string SharepointFolderFilePath = SharepointFolder + "//" + BatchFileName;

                                string SubFolderFilePath = SubFolder;

                                //string SharepointFolder = "C:\\Reports" + "\\" + BatchFileName;

                                ClientFolderFilePath = ClientFolderFilePath.Replace("%20", " ");

                                //Commented New Sharepoint Client Services site Changes 
                                // SharepointFolderFilePath = SharepointFolderFilePath.Replace("%20", " ");

                                string strSubFolderPath = SubFolderFilePath + "\\" + BatchFileName;
                                strSubFolderPath = strSubFolderPath.Replace("%20", " ").Replace("&#39;", "'").ToString();

                                DataSet dtDocTax = null;
                                DataTable dtHouseholdUUID = null;
                                DataTable dtLegalEntityUUID = null;
                                DataTable dtActiveClient = null;
                                string ClientServicesLink = string.Empty;
                                bool bResultClientPortal = false;
                                string NewCSSharepointFolderPath = string.Empty;
                                string FinalClientServiceLink = string.Empty;
                                bool bLegalEntityFolder = false;
                                bool IsSharepointSiteExists = false;
                                string Batch_Name1 = row.Cells[38].Text;
                                string newsClientFolder = "";




                                if ((BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                {
                                    //Commented New Sharepoint Client Services site Changes 
                                    //  if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && SharepointFolderFilePath != "" && SharepointFolderFilePath != "&nbsp;")
                                    if (BatchFilePath != "&nbsp;" && BatchFilePath != "") //&& ssi_CSCiteUUID != "" && ssi_CSCiteUUID != "&nbsp;")
                                    {
                                        BatchName = (string)row.Cells[4].Text;
                                        //Commented New Sharepoint Client Services site Changes 
                                        // if (SharepointFolder != "")
                                        //if (ssi_CSCiteUUID != "")
                                        if (ssi_CSCiteUUID != "" && ssi_CSCiteUUID != "&nbsp;")
                                        {

                                            if (ssi_SPSiteType == "100000000" && ssi_SPLEFolder != "")//Household and check if legalEntity folder isnt Empty
                                            {
                                                bLegalEntityFolder = true;
                                            }
                                            if (bLegalEntityFolder)// For LegalEntity Folder inside Household Library
                                            {
                                                dtActiveClient = (DataTable)ViewState["dtActiveClientList"];
                                                //dtDocTax = (DataSet)ViewState["dtDocumentTaxonomy"];

                                                //dtHouseholdUUID = dtDocTax.Tables[0];
                                                //dtLegalEntityUUID = dtDocTax.Tables[1];


                                                NewCSSharepointFolderPath = sp.FetchNewSpURL(dtActiveClient, ssi_CSCiteUUID);


                                                IsSharepointSiteExists = CheckLEgalEntityFolderExist(NewCSSharepointFolderPath, "PublishedDocuments", ssi_SPLEFolder);

                                                if (IsSharepointSiteExists)
                                                {
                                                    try
                                                    {
                                                        // if (checkSharepouintFileExist(sharepointpath, BatchFileName))
                                                        if (CheckFileExistinLegalEntity(NewCSSharepointFolderPath, "PublishedDocuments", ssi_SPLEFolder, BatchFileName))
                                                        {
                                                            lg.AddinLogFile(LogFileName, "CheckFileExistinLegalEntity nside");
                                                            if (strExistingFiles == "")
                                                                strExistingFiles = "<br/>" + NewCSSharepointFolderPath + "/PublishedDocuments/" + ssi_SPLEFolder + "/" + BatchFileName;
                                                            else
                                                                strExistingFiles = strExistingFiles + ",<br/>" + NewCSSharepointFolderPath + "/PublishedDocuments/" + ssi_SPLEFolder + "/" + BatchFileName;

                                                            //ExistingFileName.Add(newsharepointPath + "/" + BatchFileName);

                                                        }

                                                        // New Sharepoint Client Services site Changes 
                                                        //  CopyFilenew(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);
                                                        string BatchType1 = row.Cells[37].Text;
                                                        lg.AddinLogFile(LogFileName, "CopyFileinLegalEntityFolder before");
                                                        // ClientServicesLink = CopyFilenewCS(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, year, BatchType1, Quarter, LegalEntityUUID, dtLegalEntityUUID);

                                                        //if (!billingFlag)
                                                        //   ClientServicesLink = CopyFileinLegalEntityFolder(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, "PublishedDocuments/" + ssi_SPLEFolder, year, BatchType1, Quarter, LegalEntityUUID, billingFlag);
                                                        ClientServicesLink = CopyFileinLegalEntityFolder(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, "PublishedDocuments/" + ssi_SPLEFolder, year, BatchType1, Quarter, ssi_billingid);

                                                        //else
                                                        //    ClientServicesLink = CopyFileinLegalEntityFolder1(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, "ComplianceDocuments/" + ssi_SPLEFolder, year, BatchType1, Quarter);

                                                        lg.AddinLogFile(LogFileName, "ClientServicesLink " + ClientServicesLink);
                                                        if (ClientServicesLink != "")
                                                        {
                                                            // ClientServicesLink = "https://greshampartners.sharepoint.com" + ClientServicesLink;
                                                            ClientServicesLink = AppLogic.GetParam(AppLogic.ConfigParam.SharepointURL)+ ClientServicesLink;

                                                            FinalClientServiceLink = ClientServicesLink.Replace(BatchFileName, "");
                                                            SucessClientName.Add(Batch_Name1);
                                                            SucessFileName.Add(FinalClientServiceLink);

                                                            FinalClientServiceLink = FinalClientServiceLink.Remove(FinalClientServiceLink.Length - 1);// remove last "/" from URL
                                                            FinalClientServiceLink = HttpUtility.UrlPathEncode(FinalClientServiceLink); // Encode String to URL
                                                        }
                                                    }
                                                    catch (Exception exc)
                                                    {
                                                        lg.AddinLogFile(LogFileName, "DErrorr " + exc.Message.ToString());
                                                        // strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                                        Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                               " to " + NewCSSharepointFolderPath + "/PublishedDocuments/" + ssi_SPLEFolder + "<br/>" + exc.Message + ", " + exc.StackTrace);
                                                    }
                                                }
                                                else
                                                {

                                                    //CopyFilenewCS(ss)

                                                    string BatchType1 = row.Cells[37].Text;

                                                    // ClientServicesLink = CopyFilenewCS(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, year, BatchType1, Quarter, LegalEntityUUID, billingFlag);

                                                    ClientServicesLink = CopyFilenewCS(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, year, BatchType1, Quarter, ssi_billingid);

                                                    // strFolderExists = strFolderExists + "<br/>File Not saved to Client Services <br/>Legal Entity Folder not found of the below batch<br/> Batch Name: " + Batch_Name1;
                                                    // string ClientServicesFinalLink = "https://greshampartners.sharepoint.com" + ClientServicesLink.Trim().Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                                    string ClientServicesFinalLink = AppLogic.GetParam(AppLogic.ConfigParam.SharepointURL) + ClientServicesLink.Trim().Replace(" ", "%20").Replace("&#39;", "'").ToString();


                                                    //FinalClientServiceLink = "https://greshampartners.sharepoint.com" + ClientServicesLink.Replace(BatchFileName, "");

                                                    //FinalClientServiceLink = FinalClientServiceLink.Remove(FinalClientServiceLink.Length - 1);// remove last "/" from URL
                                                    //FinalClientServiceLink = HttpUtility.UrlPathEncode(FinalClientServiceLink); // Encode String to URL

                                                    //Response.Write("Before:" + ClientServicesFinalLink);

                                                    //Regex r = new Regex(@"(https?://[^\s]+)");
                                                    //string  myString = r.Replace(NewCSSharepointFolderPath, "<a href=\"$1\">$1</a>");

                                                    // strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + NewCSSharepointFolderPath + "/" + "PublishedDocuments";

                                                    //if (ClientServicesLink != "")
                                                    //{
                                                    //    ClientServicesLink = "https://greshampartners.sharepoint.com" + ClientServicesLink;

                                                    //    FinalClientServiceLink = ClientServicesLink.Replace(BatchFileName, "");
                                                    //    SucessClientName.Add(Batch_Name1);
                                                    //    SucessFileName.Add(FinalClientServiceLink);

                                                    //    FinalClientServiceLink = FinalClientServiceLink.Remove(FinalClientServiceLink.Length - 1);// remove last "/" from URL
                                                    //    FinalClientServiceLink = HttpUtility.UrlPathEncode(FinalClientServiceLink); // Encode String to URL
                                                    //}


                                                    //  Uri myUri = new Uri(NewCSSharepointFolderPath + "/" + "PublishedDocuments", UriKind.Absolute);
                                                    //  stringy = stringy + "<a href='http://www.youtube.com/" + gin.Value + "'>ClickMe</a>";


                                                    string link = NewCSSharepointFolderPath + "/" + "PublishedDocuments";

                                                    // strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + NewCSSharepointFolderPath + "/" + "PublishedDocuments";

                                                    // strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + "<a href='" + link + "'>" + link + "</a>";


                                                    strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + "<a href='" + link + "' target=_blank >" + link + "</a>";

                                                    //  strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + myUri;

                                                    // Response.Write(FinalClientServiceLink);
                                                    //// strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + FinalClientServiceLink;

                                                    SendEmail(ssi_SPLEFolder, ClientServicesFinalLink);

                                                    // Response.Write("after:" + ClientServicesFinalLink);
                                                }
                                            }
                                            else
                                            {
                                                // New Sharepoint Client Services site Changes 
                                                dtActiveClient = (DataTable)ViewState["dtActiveClientList"];
                                                //dtDocTax = (DataSet)ViewState["dtDocumentTaxonomy"];

                                                //dtHouseholdUUID = dtDocTax.Tables[0];
                                                //dtLegalEntityUUID = dtDocTax.Tables[1];



                                                // NewCSSharepointFolderPath = sp.FetchSharepointLink(dtHouseholdUUID, dtActiveClient, ssi_CSCiteUUID);
                                                NewCSSharepointFolderPath = sp.FetchNewSpURL(dtActiveClient, ssi_CSCiteUUID);
                                                //Commented New Sharepoint Client Services site Changes 
                                                // if (SharepointFolder != "")
                                                //bool IsSharepointFolderExists = CheckFolderPathExists(SharepointFolder);



                                                IsSharepointSiteExists = CheckNewCSSiteExists(NewCSSharepointFolderPath, ssi_billingid);
                                                if (IsSharepointSiteExists)
                                                {
                                                    try
                                                    {
                                                        // if (checkSharepouintFileExist(sharepointpath, BatchFileName))
                                                        if (CheckFileExistinNewCSSite(NewCSSharepointFolderPath, BatchFileName))
                                                        {
                                                            if (strExistingFiles == "")
                                                            {
                                                                //if (billingFlag)
                                                                //    strExistingFiles = "<br/>" + NewCSSharepointFolderPath + "/ComplianceDocuments/" + BatchFileName;
                                                                //else
                                                                strExistingFiles = "<br/>" + NewCSSharepointFolderPath + "/PublishedDocuments/" + BatchFileName;
                                                            }
                                                            else
                                                            {
                                                                //if (billingFlag)
                                                                //    strExistingFiles = "<br/>" + NewCSSharepointFolderPath + "/ComplianceDocuments/" + BatchFileName;
                                                                //else
                                                                strExistingFiles = strExistingFiles + ",<br/>" + NewCSSharepointFolderPath + "/PublishedDocuments/" + BatchFileName;
                                                            }
                                                        }
                                                        string BatchType1 = row.Cells[37].Text;
                                                        //Commented New Sharepoint Client Services site Changes 
                                                        //  CopyFilenew(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);
                                                        //ClientServicesLink = CopyFilenewCS(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, year, BatchType1, Quarter, LegalEntityUUID, billingFlag);

                                                        ClientServicesLink = CopyFilenewCS(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, year, BatchType1, Quarter, ssi_billingid);


                                                        if (ClientServicesLink != "")

                                                        {
                                                            // ClientServicesLink = "https://greshampartners.sharepoint.com" + ClientServicesLink;
                                                            ClientServicesLink = AppLogic.GetParam(AppLogic.ConfigParam.SharepointURL) + ClientServicesLink;

                                                            FinalClientServiceLink = ClientServicesLink.Replace(BatchFileName, "");
                                                            SucessClientName.Add(Batch_Name1);
                                                            SucessFileName.Add(FinalClientServiceLink);

                                                            FinalClientServiceLink = FinalClientServiceLink.Remove(FinalClientServiceLink.Length - 1);// remove last "/" from URL
                                                            FinalClientServiceLink = HttpUtility.UrlPathEncode(FinalClientServiceLink); // Encode String to URL

                                                        }
                                                    }
                                                    catch (Exception exc)
                                                    {
                                                        // strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                                        Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                               " to " + NewCSSharepointFolderPath + "<br/>" + exc.Message + ", " + exc.StackTrace);

                                                        //FailClientName.Add();
                                                        //FailFileName.Add

                                                    }

                                                }
                                                else
                                                {
                                                    //strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + NewCSSharepointFolderPath;

                                                    if (ssi_SPSiteType == "100000000")
                                                    {
                                                        strFolderExists = strFolderExists + "<br/>File Not saved to Client Services <br/>Client Services Site not found for the CS household of the below batch <br/> Batch Name: " + Batch_Name1;
                                                    }
                                                    else
                                                    {
                                                        strFolderExists = strFolderExists + "<br/>File Not saved to Client Services <br/>Client Services Site not found for the CS LegalEntity of the below batch <br/> Batch Name: " + Batch_Name1;
                                                    }


                                                }
                                            }



                                        }
                                        else
                                        {
                                            //Commented New Sharepoint Client Services site Changes 
                                            //if (SharepointFolder.ToLower().Contains("clientserv"))
                                            //{
                                            //    string newsharepointPath = sharepointFolderPath(SharepointFolder);
                                            //    SharepointFolder = "https://greshampartners.sharepoint.com/clientserv/" + newsharepointPath;
                                            //}

                                            if (strFolderExists == "")
                                            {
                                                //   strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + NewCSSharepointFolderPath;
                                                if (ssi_SPSiteType == "100000000")//CS Household
                                                {

                                                    strFolderExists = strFolderExists + "<br/>File Not saved to Client Services<br/>CS Household is Empty for the below batch,SP Site type = Household <br/>Batch Name: " + Batch_Name1;
                                                }
                                                else
                                                {
                                                    strFolderExists = strFolderExists + "<br/>File Not saved to Client Services<br/>CS LegalEntity is Empty for the below batch,SP Site type = LegalEntity  <br/>Batch Name: " + Batch_Name1;
                                                }


                                            }
                                            else
                                            {
                                                strFolderExists = strFolderExists + "<br/>Batch Name: " + Batch_Name1 + "<br/>Sharepoint folder path: " + NewCSSharepointFolderPath;
                                            }
                                        }
                                    }
                                    else
                                        NewCSSharepointFolderPath = "";
                                }
                                else if (ssi_CSCiteUUID != "")
                                {
                                    BatchName = (string)row.Cells[4].Text;
                                }

                                //---------------------- Clientfolder Region
                                if ((BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                {

                                    if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && ClientFolderFilePath != "&nbsp;" && ClientFolderFilePath != "")
                                    {

                                        if (ClientFolder.ToLower().Contains("clientportal"))
                                        {
                                            string HouseHoldName = row.Cells[15].Text;
                                            //   string HouseHold = row.Cells[15].Text;
                                            HouseHoldName = HouseHoldName.Replace(" Family", "");
                                            HouseHoldName = HouseHoldName.Replace(" family", "");
                                            HouseHoldName = HouseHoldName.Replace(" FAMILY", "");

                                            //var Name = HouseHold.Split(' ');
                                            //string HouseHoldName = Name[0];

                                            string BatchType1 = row.Cells[37].Text;
                                            //string Batch_Name1 = row.Cells[38].Text;

                                            // CopyFile(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);
                                            newsClientFolder = "Documents taxonomy";
                                            string newClientsharepointPath = newsClientFolder;
                                           // newsClientFolder = "https://greshampartners.sharepoint.com/ClientPortal/" + newsClientFolder;
                                           
                                            newsClientFolder = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL) +"/"+ newsClientFolder;
                                            try
                                            {
                                                //Commented New Sharepoint Client Services site Changes 
                                                // string FileName = Path.GetFileName(SharepointFolderFilePath);
                                                string FileName = BatchFileName;

                                                //bool result=  CopyFiletoSharepoint(ClientFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, HouseHoldName, year, SharepointFolderFilePath);\
                                                //    bool result = CopyFiletoSharepoint(SharepointFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, ssi_clientportalname, year, SharepointFolderFilePath);

                                                //Commented New Sharepoint Client Services site Changes 
                                                // bool result = CopyFiletoSharepoint(SharepointFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, HouseHoldName, year, SharepointFolderFilePath, ssi_clientportalname);
                                                bResultClientPortal = CopyFiletoSharepoint("", FileName, BatchFilePath, BatchType1, Batch_Name1, HouseHoldName, year, "", ssi_clientportalname, ssi_billingid);
                                                if (bResultClientPortal)
                                                {
                                                    // strFolderExists = strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsClientFolder;
                                                    SucessClientName.Add(Batch_Name1);
                                                    SucessFileName.Add(newsClientFolder);
                                                }
                                                else
                                                {
                                                    // strFolderExists = strFolderExists + "<br/>Can not File in Client Portal Tag Missing, Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsClientFolder;
                                                    //  FailClientName.Add(BatchName);
                                                    //FailClientName.Add(ssi_clientportalname + "," + BatchName);
                                                    //FailFileName.Add(newsClientFolder);
                                                    // FailClientName.Add(ssi_clientportalname + "," + HouseHoldName);
                                                    string strFailClientName = "";
                                                    if (ssi_clientportalname != "")
                                                    {
                                                        strFailClientName = ssi_clientportalname;
                                                    }
                                                    else
                                                    {
                                                        strFailClientName = HouseHoldName;
                                                    }
                                                    FailClientName.Add(strFailClientName);
                                                    FailFileName.Add(newsClientFolder);
                                                    ListBatchName.Add(Batch_Name1);
                                                }

                                            }
                                            catch (Exception exce)
                                            {
                                                //  strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsClientFolder;
                                                Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                        " to " + newsClientFolder + "/" + BatchFileName + "<br/>" + exce.Message + ", " + exce.StackTrace);
                                            }



                                        }
                                        else
                                        {
                                            try
                                            {
                                                if (System.IO.File.Exists(newsClientFolder + "/" + BatchFileName))
                                                {
                                                    if (strExistingFiles == "")
                                                        strExistingFiles = "<br/>" + newsClientFolder + "/" + BatchFileName;
                                                    else
                                                        strExistingFiles = strExistingFiles + ",<br/>" + newsClientFolder + "/" + BatchFileName;


                                                    // ExistingFileName.Add(SharepointFolderFilePath + "/" + BatchFileName);
                                                    System.IO.File.Copy(BatchFilePath, newsClientFolder + "/" + BatchFileName, true);
                                                    strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Client folder path: " + newsClientFolder;
                                                }

                                                //System.IO.File.Copy(BatchFilePath, ClientFolderFilePath, true);
                                                //strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Client folder path: " + ClientFolder;
                                            }
                                            catch (Exception ex)
                                            {
                                                Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                                         " to " + newsClientFolder + "/" + BatchFileName + "<br/>" + ex.Message + ", " + ex.StackTrace);

                                            }
                                        }
                                        //SetFileReadAccess(ClientFolderFilePath, true);
                                        //Response.Write("<br/>File Copied: " + ClientFolder + "<br/>");
                                    }
                                    else
                                        ClientFolderFilePath = "";
                                }
                                else if (ClientFolder != "")
                                {

                                    BatchName = (string)row.Cells[4].Text;

                                    if (strFolderExists == "")
                                    {
                                        strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Client folder path: " + newsClientFolder;
                                    }
                                    else
                                        strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Client folder path: " + newsClientFolder;
                                }
                                if (bResultClientPortal || ClientServicesLink != "")
                                {
                                    Count++;
                                    #region Update Batch Status to 'Final Report Sent'

                                    //ssi_batch objBatch = new ssi_batch();
                                    Entity objBatch = new Entity("ssi_batch");

                                    //objBatch.ssi_batchid = new Key();
                                    //objBatch.ssi_batchid.Value = new Guid(ssi_batchid);
                                    objBatch["ssi_batchid"] = new Guid(ssi_batchid);

                                    if (bResultClientPortal && ClientServicesLink != "")
                                    {
                                        //objBatch.ssi_sharepointreportfolderfinal = FinalSharepointFolder.Replace(" Family", "").Replace(",", "%2C").Replace(" ", "%20").Replace("'", "%27").Replace("&#39;", "'").ToString();
                                        // objBatch["ssi_sharepointreportfolderfinal"] = FinalSharepointFolder.Replace(" Family", "").Replace(",", "%2C").Replace(" ", "%20").Replace("'", "%27").Replace("&#39;", "'").ToString();

                                        objBatch["ssi_sharepointreportfolderfinal"] = HttpUtility.UrlPathEncode(FinalClientServiceLink);// FinalClientServiceLink;
                                    }
                                    else if (bResultClientPortal)
                                    {
                                        objBatch["ssi_sharepointreportfolderfinal"] = HttpUtility.UrlPathEncode(newsClientFolder); // newsClientFolder;//+ BatchFileName;
                                    }
                                    else if (ClientServicesLink != "")
                                    {
                                        objBatch["ssi_sharepointreportfolderfinal"] = HttpUtility.UrlPathEncode(FinalClientServiceLink);// FinalClientServiceLink;
                                    }

                                    if (row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace("&nbsp;", "") != "")
                                    {
                                        //objBatch.ssi_sharepointemaillink = (string)row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                        // objBatch["ssi_sharepointemaillink"] = (string)row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                        objBatch["ssi_sharepointemaillink"] = BatchFileName.Trim().Replace("ssi_spvfilename", "").Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                    }
                                    else
                                    {
                                        //objBatch.ssi_sharepointemaillink = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(" ", "%20");
                                        objBatch["ssi_sharepointemaillink"] = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(" ", "%20");
                                    }

                                    if (MailPref.ToUpper().Contains("EMAIL"))
                                    {
                                        //objBatch.ssi_finalreportcreatedflag = new CrmBoolean();
                                        //objBatch.ssi_finalreportcreatedflag.Value = true;
                                        objBatch["ssi_finalreportcreatedflag"] = true;

                                        updateSentData(ssi_secondaryownerid, ssi_mailrecordsId);
                                    }
                                    else
                                    {
                                        //objBatch.ssi_sendemailib = new CrmBoolean();
                                        //objBatch.ssi_sendemailib.Value = true;
                                    }

                                    //objBatch.ssi_sendemailib = new CrmBoolean();
                                    //objBatch.ssi_sendemailib.Value = true;
                                    objBatch["ssi_sendemailib"] = true;


                                    if (MailPref.ToUpper().Contains("EMAIL") && ssi_ReviewReqByid != "")
                                    {
                                        // Send Email to Review Required By 
                                        //objBatch.ssi_reviewrequiredbyid = new Lookup();
                                        //objBatch.ssi_reviewrequiredbyid.type = EntityName.systemuser.ToString();
                                        //objBatch.ssi_reviewrequiredbyid.Value = new Guid(ssi_ReviewReqByid);
                                        objBatch["ssi_reviewrequiredbyid"] = new EntityReference("systemuser", new Guid(ssi_ReviewReqByid));

                                        //objBatch.ssi_sendrrbymail = new CrmBoolean();
                                        //objBatch.ssi_sendrrbymail.Value = true;
                                        objBatch["ssi_sendrrbymail"] = true;

                                        //objBatch.ssi_reporttrackerstatus = new Picklist();
                                        //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Sent'
                                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                        // intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                        intStatus = 4;



                                        //BatchReportStatus(intStatus, UserId, ssi_batchid, BillingHandedOff);
                                        BillingHandedOff = "true";
                                    }
                                    else if (MailPref.ToUpper().Contains("EMAIL") && ReviewRequiredBy == "")
                                    {
                                        // Send Email to Associate
                                        //objBatch.ssi_reporttrackerstatus = new Picklist();
                                        //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Sent'
                                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                        //objBatch.ssi_sendmailassociate = new CrmBoolean();
                                        //objBatch.ssi_sendmailassociate.Value = true;
                                        objBatch["ssi_sendmailassociate"] = true;

                                        //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                        intStatus = 4;
                                        //BatchReportStatus(intStatus, UserId, ssi_batchid, BillingHandedOff);
                                        BillingHandedOff = "true";
                                    }
                                    else
                                    {
                                        int trackerstatus;
                                        if (ssi_internalbillingcontactid != "")
                                        {
                                            //objBatch.ssi_billinghandedoff = new CrmBoolean();
                                            //objBatch.ssi_billinghandedoff.Value = true;
                                            objBatch["ssi_billinghandedoff"] = true;
                                            BillingHandedOff = "true";
                                        }
                                        else
                                        {
                                            //objBatch.ssi_billinghandedoff = new CrmBoolean();
                                            //objBatch.ssi_billinghandedoff.Value = false;
                                            objBatch["ssi_billinghandedoff"] = false;
                                            BillingHandedOff = "false";
                                        }


                                        if (MailPref.ToUpper() == "Client Portal".ToUpper())
                                        {
                                            //objBatch.ssi_reporttrackerstatus = new Picklist();
                                            //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Final Report Sent'
                                            objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                            trackerstatus = 4;

                                            //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                            updateSentData(ssi_ReviewReqByid, ssi_mailrecordsId);
                                        }
                                        else
                                        {
                                            //objBatch.ssi_reporttrackerstatus = new Picklist();
                                            //objBatch.ssi_reporttrackerstatus.Value = 9;// Batch status 'Final Report Sent'
                                            objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(9);
                                            trackerstatus = 9;
                                        }

                                        //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                        intStatus = trackerstatus;

                                    }

                                    if (ssi_ReviewReqByid != "" && MailPref.ToUpper().Contains("EMAIL"))
                                    {
                                        lblError.Text = "Some of the reports you selected need to be given to the person designated in Review required by.";
                                    }

                                    service.Update(objBatch);

                                    bool billingHandedOff = Convert.ToBoolean(BillingHandedOff);
                                    BatchReportStatus(intStatus, UserId, ssi_batchid, billingHandedOff);

                                    #endregion



                                    #region update Billing Invoice


                                    if (ssi_billinginvoiceid != "")
                                    {
                                        Entity objbillingInvoice = new Entity("ssi_billinginvoice");
                                        //objbillingInvoice.ssi_billinginvoiceid = new Key();
                                        //objbillingInvoice.ssi_billinginvoiceid.Value = new Guid(Convert.ToString(dsUpdateInvoice.Tables[0].Rows[j]["Ssi_billinginvoiceId"]));
                                        objbillingInvoice["ssi_billinginvoiceid"] = new Guid(ssi_billinginvoiceid);

                                        objbillingInvoice["ssi_invoicedate"] = DateTime.Now;

                                        service.Update(objbillingInvoice);

                                    }

                                    //if (Convert.ToString(dsUpdateInvoice.Tables[0].Rows[j]["Ssi_InvoiceDate"]) != "")
                                    //{
                                    //objbillingInvoice.ssi_invoicedate = new CrmDateTime();
                                    //objbillingInvoice.ssi_invoicedate.Value = Convert.ToString(dsUpdateInvoice.Tables[0].Rows[j]["Ssi_InvoiceDate"]);


                                    //}


                                    // intResult++;

                                    #endregion

                                    #region Update Mail status to 'Created'
                                    if (MailPref.ToUpper().Contains("EMAIL") && ssi_ReviewReqByid != "")
                                    {
                                        //objMailRecords.ssi_mailstatus = new Picklist();
                                        //objMailRecords.ssi_mailstatus.Value = 3;//mail status 'Sent to FINAL Reviewer'
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(3);
                                    }
                                    else if ((MailPref.ToUpper().Contains("EMAIL") || MailPref.ToUpper() == "Client Portal".ToUpper()) && ssi_ReviewReqByid == "")
                                    {
                                        //objMailRecords.ssi_mailstatus = new Picklist();
                                        //objMailRecords.ssi_mailstatus.Value = 4;//mail status 'Sent to FINAL Reviewer'
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                    }
                                    else if (MailPref.ToUpper() == "Client Portal".ToUpper() && ssi_ReviewReqByid != "")
                                    {
                                        //objMailRecords.ssi_mailstatus = new Picklist();
                                        //objMailRecords.ssi_mailstatus.Value = 4;//mail status 'Sent to FINAL Reviewer'
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                    }
                                    else
                                    {
                                        //objMailRecords.ssi_mailstatus = new Picklist();
                                        //objMailRecords.ssi_mailstatus.Value = 8;//mail status 'Created'
                                        //4-sent
                                        //  if (billingFlag)
                                        if (ssi_billingid != "")
                                            objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                        else
                                            objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(8);

                                    }



                                    service.Update(objMailRecords);
                                    #endregion


                                }

                                //---------------------- SubFolder Region
                                if (Directory.Exists(SubFolderFilePath) && (BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                {

                                    if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && SubFolderFilePath != "&nbsp;" && SubFolderFilePath != "")
                                    {
                                        if (System.IO.File.Exists(SubFolderFilePath))
                                        {
                                            if (strExistingFiles == "")
                                                strExistingFiles = "<br/>" + strSubFolderPath;
                                            else
                                                strExistingFiles = strExistingFiles + ",<br/>" + strSubFolderPath;

                                            //ExistingFileName.Add(SharepointFolderFilePath + "/" + BatchFileName);
                                        }

                                        try
                                        {
                                            System.IO.File.Copy(BatchFilePath, strSubFolderPath, true);
                                        }
                                        catch (Exception ex)
                                        { }
                                        //SetFileReadAccess(ClientFolderFilePath, true);
                                        //Response.Write("<br/>File Copied: " + ClientFolder + "<br/>");
                                    }

                                }
                                else if (SubFolder != "")
                                {
                                    Directory.CreateDirectory(SubFolderFilePath);
                                    if (Directory.Exists(SubFolderFilePath) && (BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                    {
                                        if (System.IO.File.Exists(strSubFolderPath))
                                        {
                                            if (strExistingFiles == "")
                                                strExistingFiles = "<br/>" + strSubFolderPath;
                                            else
                                                strExistingFiles = strExistingFiles + ",<br/>" + strSubFolderPath;

                                            //ExistingFileName.Add(SharepointFolderFilePath + "/" + BatchFileName);
                                        }

                                        try
                                        {
                                            System.IO.File.Copy(BatchFilePath, strSubFolderPath, true);
                                        }
                                        catch (Exception ex)
                                        { }

                                    }
                                }

                            }

                            //OPS Approved

                            else if (ddlAction.SelectedValue == "11" && (BatchStatus == "8" || BatchStatus == "9" || BatchStatus == "4"))// Save Reports to SharePoint and Client Portal //&& HoldReport == ""
                            {
                                // lg.AddinLogFile(LogFileName, "Save Reports to SharePoint only ");
                                #region OLD Sharepoint Code commented after New client Sevices Site 
                                //#region Update Batch Status to 'Final Report Sent'
                                ////ssi_batch objBatch = new ssi_batch();
                                //Entity objBatch = new Entity("ssi_batch");

                                ////objBatch.ssi_batchid = new Key();
                                ////objBatch.ssi_batchid.Value = new Guid(ssi_batchid);
                                //objBatch["ssi_batchid"] = new Guid(ssi_batchid);

                                //if (FinalSharepointFolder != "")
                                //{
                                //    //objBatch.ssi_sharepointreportfolderfinal = FinalSharepointFolder.Replace(" Family", "").Replace(",", "%2C").Replace(" ", "%20").Replace("'", "%27").Replace("&#39;", "'").ToString();
                                //    objBatch["ssi_sharepointreportfolderfinal"] = FinalSharepointFolder.Replace(" Family", "").Replace(",", "%2C").Replace(" ", "%20").Replace("'", "%27").Replace("&#39;", "'").ToString();
                                //}
                                //if (row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace("&nbsp;", "") != "")
                                //{
                                //    //objBatch.ssi_sharepointemaillink = (string)row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                //    objBatch["ssi_sharepointemaillink"] = (string)row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                //}
                                //else
                                //{
                                //    //objBatch.ssi_sharepointemaillink = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(" ", "%20");
                                //    objBatch["ssi_sharepointemaillink"] = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(" ", "%20");
                                //}

                                //if (MailPref.ToUpper().Contains("EMAIL"))
                                //{
                                //    //objBatch.ssi_finalreportcreatedflag = new CrmBoolean();
                                //    //objBatch.ssi_finalreportcreatedflag.Value = true;
                                //    objBatch["ssi_finalreportcreatedflag"] = true;

                                //    updateSentData(ssi_secondaryownerid, ssi_mailrecordsId);

                                //}
                                //else
                                //{
                                //    //objBatch.ssi_sendemailib = new CrmBoolean();
                                //    //objBatch.ssi_sendemailib.Value = true;
                                //}

                                ////objBatch.ssi_sendemailib = new CrmBoolean();
                                ////objBatch.ssi_sendemailib.Value = true;
                                //objBatch["ssi_sendemailib"] = true;

                                //if (MailPref.ToUpper().Contains("EMAIL") && ssi_ReviewReqByid != "")
                                //{
                                //    // Send Email to Review Required By 
                                //    //objBatch.ssi_reviewrequiredbyid = new Lookup();
                                //    //objBatch.ssi_reviewrequiredbyid.type = EntityName.systemuser.ToString();
                                //    //objBatch.ssi_reviewrequiredbyid.Value = new Guid(ssi_ReviewReqByid);
                                //    objBatch["ssi_reviewrequiredbyid"] = new EntityReference("systemuser", new Guid(ssi_ReviewReqByid));

                                //    //objBatch.ssi_sendrrbymail = new CrmBoolean();
                                //    //objBatch.ssi_sendrrbymail.Value = true;
                                //    objBatch["ssi_sendrrbymail"] = true;

                                //    //objBatch.ssi_reporttrackerstatus = new Picklist();
                                //    //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Sent'
                                //    objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                //    //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                //    intStatus = 4;
                                //    //BatchReportStatus(intStatus, UserId, ssi_batchid, BillingHandedOff);
                                //    BillingHandedOff = "true";
                                //}
                                //else if (MailPref.ToUpper().Contains("EMAIL") && ReviewRequiredBy == "")
                                //{
                                //    // Send Email to Associate
                                //    //objBatch.ssi_reporttrackerstatus = new Picklist();
                                //    //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Sent'
                                //    objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                //    //objBatch.ssi_sendmailassociate = new CrmBoolean();
                                //    //objBatch.ssi_sendmailassociate.Value = true;
                                //    objBatch["ssi_sendmailassociate"] = true;

                                //    intStatus = 4;
                                //    //BatchReportStatus(intStatus, UserId, ssi_batchid, BillingHandedOff);
                                //    BillingHandedOff = "true";
                                //}
                                //else
                                //{
                                //    if (ssi_internalbillingcontactid != "")
                                //    {
                                //        //objBatch.ssi_billinghandedoff = new CrmBoolean();
                                //        //objBatch.ssi_billinghandedoff.Value = true;
                                //        objBatch["ssi_billinghandedoff"] = true;
                                //        BillingHandedOff = "true";

                                //    }
                                //    else
                                //    {
                                //        //objBatch.ssi_billinghandedoff = new CrmBoolean();
                                //        //objBatch.ssi_billinghandedoff.Value = false;
                                //        objBatch["ssi_billinghandedoff"] = false;
                                //        BillingHandedOff = "false";
                                //    }


                                //    if (MailPref.ToUpper() == "Client Portal".ToUpper())
                                //    {
                                //        //objBatch.ssi_reporttrackerstatus = new Picklist();
                                //        //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Final Report Sent'
                                //        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                //        intStatus = 4;
                                //        updateSentData(ssi_ReviewReqByid, ssi_mailrecordsId);
                                //        //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                //    }
                                //    else
                                //    {
                                //        //objBatch.ssi_reporttrackerstatus = new Picklist();
                                //        //objBatch.ssi_reporttrackerstatus.Value = 9;// Batch status 'Final Report Sent'
                                //        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(9);
                                //        intStatus = 9;
                                //    }

                                //    // intStatus = objBatch.ssi_reporttrackerstatus.Value;

                                //}

                                //if (ssi_ReviewReqByid != "" && MailPref.ToUpper().Contains("EMAIL"))
                                //{
                                //    lblError.Text = "Some of the reports you selected need to be given to the person designated in Review required by.";
                                //}

                                //service.Update(objBatch);

                                //bool billingHandedOff = Convert.ToBoolean(BillingHandedOff);
                                //BatchReportStatus(intStatus, UserId, ssi_batchid, billingHandedOff);

                                //#endregion

                                //#region Update Mail status to 'Created'
                                //if (MailPref.ToUpper().Contains("EMAIL") && ssi_ReviewReqByid != "")
                                //{
                                //    //objMailRecords.ssi_mailstatus = new Picklist();
                                //    //objMailRecords.ssi_mailstatus.Value = 3;//mail status 'Sent to FINAL Reviewer'
                                //    objBatch["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(3);
                                //}
                                //else if ((MailPref.ToUpper().Contains("EMAIL") || MailPref.ToUpper() == "Client Portal".ToUpper()) && ssi_ReviewReqByid == "")
                                //{
                                //    //objMailRecords.ssi_mailstatus = new Picklist();
                                //    //objMailRecords.ssi_mailstatus.Value = 4;//mail status 'Sent to FINAL Reviewer'
                                //    objBatch["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                //}
                                //else if (MailPref.ToUpper() == "Client Portal".ToUpper() && ssi_ReviewReqByid != "")
                                //{
                                //    //objMailRecords.ssi_mailstatus = new Picklist();
                                //    //objMailRecords.ssi_mailstatus.Value = 4;//mail status 'Sent to FINAL Reviewer'
                                //    objBatch["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                //}
                                //else
                                //{
                                //    //objMailRecords.ssi_mailstatus = new Picklist();
                                //    //objMailRecords.ssi_mailstatus.Value = 8;//mail status 'Created'
                                //    objBatch["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(8);

                                //}

                                //#region Update Address not used
                                ////if (MailAddress != ContactAddress)
                                ////{
                                ////    string[] NewAddress = ContactAddress.Split('\n');
                                ////    //Address1
                                ////    if (NewAddress[0] != "")
                                ////    {
                                ////        objMailRecords.ssi_addressline1_mail = NewAddress[0];
                                ////    }
                                ////    //Address2
                                ////    if (NewAddress[1] != "")
                                ////    {
                                ////        objMailRecords.ssi_addressline2_mail = NewAddress[1];
                                ////    }
                                ////    //Address3
                                ////    if (NewAddress[2] != "")
                                ////    {
                                ////        objMailRecords.ssi_addressline3_mail = NewAddress[2];
                                ////    }
                                ////    //City
                                ////    if (NewAddress[3] != "")
                                ////    {
                                ////        objMailRecords.ssi_city_mail = NewAddress[3];
                                ////    }

                                ////    // State Province
                                ////    if (NewAddress[4] != "")
                                ////    {
                                ////        objMailRecords.ssi_stateprovince_mail = NewAddress[4];
                                ////    }

                                ////    // Zip Code
                                ////    if (NewAddress[5] != "")
                                ////    {
                                ////        objMailRecords.ssi_zipcode_mail = NewAddress[5];
                                ////    }

                                ////}
                                //#endregion

                                //service.Update(objMailRecords);
                                //#endregion

                                //string BatchName = string.Empty;
                                //string SubFolderName = string.Empty;
                                //string DestinationFileName = string.Empty;
                                //string SubFolder = string.Empty;

                                //BatchName = (string)row.Cells[4].Text;
                                //SubFolderName = (string)row.Cells[32].Text;

                                ////\\GRPAO1-VWFS01\shared$\Mail Merge\Completed Mailings
                                ////\\GRPAO1-VWFS01\shared$\Mail Merge\Completed Mailings
                                //SubFolder = "\\\\GRPAO1-VWFS01\\shared$\\Mail Merge\\Completed Mailings\\" + SubFolderName;
                                ////SubFolder = "\\\\GRPAO1-VWFS01\\opsreports$\\Mail Merge\\Completed Mailings\\" + SubFolderName;

                                //string BatchFilePath = (string)row.Cells[18].Text;
                                //string AsOfDate = (string)row.Cells[3].Text;
                                //string[] date = AsOfDate.Split(new char[] { '/' });
                                //string year = date[2];
                                ////string BatchFileName = (string)row.Cells[21].Text;  // BatchFilePath.Substring(BatchFilePath.LastIndexOf("\\") + 1); // 
                                //string SPVFileName = row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace("&nbsp;", "");

                                ////BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "");
                                //if (chkTest.Checked == true)
                                //{
                                //    if (SPVFileName != "")
                                //    {
                                //        BatchFileName = SPVFileName.Replace(".pdf", "_Test.pdf");
                                //    }
                                //    else
                                //    {
                                //        BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(".pdf", "_Test.pdf");
                                //    }
                                //}
                                //else
                                //{
                                //    if (SPVFileName != "")
                                //    {
                                //        BatchFileName = SPVFileName;
                                //    }
                                //    else
                                //    {
                                //        BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "");
                                //    }
                                //}


                                //string ClientFolder = (string)row.Cells[19].Text == "&nbsp;" || (string)row.Cells[19].Text == "" ? "" : (string)row.Cells[19].Text;
                                //string SharepointFolder = (string)row.Cells[20].Text == "&nbsp;" || (string)row.Cells[20].Text == "" ? "" : (string)row.Cells[20].Text;

                                //ClientFolder = ClientFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();
                                //SharepointFolder = SharepointFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();


                                //string ClientFolderFilePath = ClientFolder + "\\" + BatchFileName;
                                ////string SharepointFolderFilePath = SharepointFolder + "\\" + BatchFileName;

                                ////string ClientFolderFilePath = ClientFolder + "//" + BatchFileName;
                                //string SharepointFolderFilePath = SharepointFolder + "//" + BatchFileName;

                                //string SubFolderFilePath = SubFolder;

                                ////string SharepointFolder = "C:\\Reports" + "\\" + BatchFileName;

                                //ClientFolderFilePath = ClientFolderFilePath.Replace("%20", " ");
                                //SharepointFolderFilePath = SharepointFolderFilePath.Replace("%20", " ");

                                //string strSubFolderPath = SubFolderFilePath + "\\" + BatchFileName;
                                //strSubFolderPath = strSubFolderPath.Replace("%20", " ").Replace("&#39;", "'").ToString();
                                //string Batch_Name1 = row.Cells[38].Text;

                                //#region Not in Use

                                ////"\\\\sp02\\DavWWWRoot\\ClientServ\\Documents\\Clients\\Active\\Anathan\\Correspondence\\Quarterly AXYS Reports\\" + BatchFileName;

                                ////string SharepointFolder = "\\\\GRPAO1-VWFS01\\_ops_C_I_R_group\\AdventReport\\" + "Anathan Family-Anathan G_2011-0930_2011_11_25_07_48.pdf";

                                ////try
                                ////{

                                ////if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && SharepointFolder != "" && SharepointFolder != "&nbsp;")
                                ////{
                                ////    if (File.Exists(SharepointFolder))
                                ////    {
                                ////        if (strExistingFiles == "")
                                ////            strExistingFiles = "<br/> <a href='" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + " </a>";
                                ////        else
                                ////            strExistingFiles = strExistingFiles + ",<br/><a href='" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + " </a>";
                                ////    }
                                ////    else
                                ////    {
                                ////        if (strNewFiles == "")
                                ////            strNewFiles = "<br/> <a href='" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "</a>";
                                ////        else
                                ////            strNewFiles = strNewFiles + ",<br/><a href='" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//") + "'>" + "http" + "://" + SharepointFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "</a>";
                                ////    }

                                ////   //File.Copy(BatchFilePath, SharepointFolder, true);
                                ////   // Response.Write("<br/>File Copied: " + SharepointFolder +"<br/>");
                                ////}
                                ////else
                                ////    SharepointFolder = "";

                                ////if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && ClientFolder != "&nbsp;" && ClientFolder != "")
                                ////{
                                ////    if (File.Exists(ClientFolder))
                                ////    {
                                ////        if (strExistingFiles == "")
                                ////            strExistingFiles = "<br/><a href='" + "http" + ":" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + " </a>";
                                ////        else
                                ////            strExistingFiles = strExistingFiles + ",<br/> <a href='" + "http" + "://" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "</a>";
                                ////    }
                                ////    else
                                ////    {
                                ////        if (strNewFiles == "")
                                ////            strNewFiles = "<br/> <a href='" + "http" + "://" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "'>" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "</a>";
                                ////        else
                                ////            strNewFiles = strNewFiles + ",<br/><a href='" + "http" + "://" + ClientFolder.Replace(" ", "%20").Replace("\\", "//") + "'>" + ClientFolder.Replace(" ", "%20").Replace("\\", "//").Substring(4) + "</a>";
                                ////    }
                                ////    //File.Copy(BatchFilePath, ClientFolder, true);
                                ////   // Response.Write("<br/>File Copied: " + SharepointFolder+"<br/>");
                                ////}
                                ////else
                                ////    ClientFolder = "";

                                ////if (strExistingFiles != "")
                                ////    lblError.Text = lblError.Text + "<br/>Batch Report Saved to Sharepoint folder, " + strExistingFiles + "<br/> above files are overwritten.";
                                ////else if (strNewFiles != "")
                                ////    lblError.Text = lblError.Text + "<br/>New Batch Report added to Sharepoint folder, " + strNewFiles;
                                ////else
                                ////    lblError.Text = lblError.Text + "<br/>Batch Report Saved to Sharepoint folder";

                                //#endregion

                                //if ((BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                //{
                                //    if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && SharepointFolderFilePath != "" && SharepointFolderFilePath != "&nbsp;")
                                //    {
                                //        BatchName = (string)row.Cells[4].Text;
                                //        if (SharepointFolder != "")
                                //        {


                                //            if (SharepointFolder.ToLower().Contains("clientserv"))
                                //            {


                                //                string newsharepointPath = sharepointFolderPath(SharepointFolder);
                                //                string sharepointpath = newsharepointPath;
                                //                newsharepointPath = "https://greshampartners.sharepoint.com/clientserv/" + newsharepointPath;
                                //                // CopyFile(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);

                                //                bool IsSharepointFolderExists = CheckFolderPathExists(SharepointFolder);
                                //                if (IsSharepointFolderExists)
                                //                {
                                //                    try
                                //                    {

                                //                        if (checkSharepouintFileExist(sharepointpath, BatchFileName))
                                //                        {
                                //                            if (strExistingFiles == "")
                                //                                strExistingFiles = "<br/>" + newsharepointPath + "/" + BatchFileName;
                                //                            else
                                //                                strExistingFiles = strExistingFiles + ",<br/>" + newsharepointPath + "/" + BatchFileName;

                                //                            //ExistingFileName.Add(newsharepointPath + "/" + BatchFileName);

                                //                        }


                                //                        CopyFilenew(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);

                                //                        // strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                        SucessClientName.Add(Batch_Name1);
                                //                        SucessFileName.Add(newsharepointPath);
                                //                    }
                                //                    catch (Exception exc)
                                //                    {
                                //                        // strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                        Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                //               " to " + SharepointFolderFilePath + "<br/>" + exc.Message + ", " + exc.StackTrace);
                                //                    }

                                //                }
                                //                else
                                //                {
                                //                    strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                }
                                //            }
                                //            else if (SharepointFolder.ToLower().Contains("clientportal"))
                                //            {
                                //                string HouseHoldName = row.Cells[15].Text;
                                //                //string HouseHold = row.Cells[15].Text;
                                //                HouseHoldName = HouseHoldName.Replace(" Family", "");
                                //                HouseHoldName = HouseHoldName.Replace(" family", "");
                                //                HouseHoldName = HouseHoldName.Replace(" FAMILY", "");
                                //                //var Name = HouseHold.Split(' ');
                                //                //string HouseHoldName = Name[0];

                                //                string BatchType1 = row.Cells[37].Text;


                                //                //string newsharepointPath =sharepointFolderPath(SharepointFolder);
                                //                string newsharepointPath = "Documents taxonomy";
                                //                //string sharepointpath = newsharepointPath;
                                //                newsharepointPath = "https://greshampartners.sharepoint.com/ClientPortal/" + newsharepointPath;

                                //                //string BatchType = row.Cells[37].Text;
                                //                //string Batch_Name = row.Cells[38].Text;
                                //                // CopyFile(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);



                                //                try
                                //                {

                                //                    string FileName = Path.GetFileName(SharepointFolderFilePath);
                                //                    //bool result= CopyFiletoSharepoint(SharepointFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, HouseHoldName, year, SharepointFolderFilePath);
                                //                    // bool result = CopyFiletoSharepoint(SharepointFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, ssi_clientportalname, year, SharepointFolderFilePath);
                                //                    bool result = CopyFiletoSharepoint(SharepointFolder, FileName, BatchFilePath, BatchType1, Batch_Name1, HouseHoldName, year, SharepointFolderFilePath, ssi_clientportalname);
                                //                    if (result)
                                //                    {
                                //                        // strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                        SucessClientName.Add(Batch_Name1);
                                //                        SucessFileName.Add(newsharepointPath);
                                //                    }
                                //                    else
                                //                    {
                                //                        //  strFolderExists = strFolderExists + "<br/>Can not File in Client Portal Tag Missing, Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                        //FailClientName.Add(BatchName);
                                //                        //FailClientName.Add(ssi_clientportalname + "," + BatchName);
                                //                        //FailFileName.Add(newsharepointPath);

                                //                        string strFailClientName = "";
                                //                        if (ssi_clientportalname != "")
                                //                        {
                                //                            strFailClientName = ssi_clientportalname;
                                //                        }
                                //                        else
                                //                        {
                                //                            strFailClientName = HouseHoldName;
                                //                        }
                                //                        FailClientName.Add(strFailClientName);
                                //                        FailFileName.Add(newsharepointPath);
                                //                        ListBatchName.Add(Batch_Name1);
                                //                    }

                                //                    //  strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                }
                                //                catch (Exception exc)
                                //                {
                                //                    // strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                //                    Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                //           " to " + SharepointFolderFilePath + "<br/>" + exc.Message + ", " + exc.StackTrace);
                                //                }


                                //            }
                                //            else
                                //            {
                                //                if (System.IO.File.Exists(SharepointFolderFilePath))
                                //                {
                                //                    if (strExistingFiles == "")
                                //                        strExistingFiles = "<br/>" + SharepointFolderFilePath;
                                //                    else
                                //                        strExistingFiles = strExistingFiles + ",<br/>" + SharepointFolderFilePath;
                                //                    // ExistingFileName.Add(SharepointFolderFilePath );

                                //                    try
                                //                    {
                                //                        System.IO.File.Copy(BatchFilePath, SharepointFolderFilePath, true);
                                //                        strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + SharepointFolder;
                                //                    }
                                //                    catch (Exception ex)
                                //                    {
                                //                        Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                //                            " to " + SharepointFolderFilePath + "<br/>" + ex.Message + ", " + ex.StackTrace);
                                //                    }
                                //                }
                                //                //SetFileReadAccess(SharepointFolderFilePath, true);
                                //                //Response.Write("<br/>File Copied: " + SharepointFolder +"<br/>");
                                //            }
                                //        }
                                //        else
                                //        {
                                //            if (SharepointFolder.ToLower().Contains("clientserv"))
                                //            {
                                //                string newsharepointPath = sharepointFolderPath(SharepointFolder);
                                //                SharepointFolder = "https://greshampartners.sharepoint.com/clientserv/" + newsharepointPath;
                                //            }

                                //            if (strFolderExists == "")
                                //                strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + SharepointFolder;
                                //            else
                                //            {
                                //                strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + SharepointFolder;
                                //            }
                                //        }
                                //    }
                                //    else
                                //        SharepointFolderFilePath = "";
                                //}
                                //else if (SharepointFolder != "")
                                //{
                                //    BatchName = (string)row.Cells[4].Text;

                                //    //if (strFolderExists == "")
                                //    //    strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + SharepointFolder;
                                //    //else
                                //    //    strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + SharepointFolder;
                                //}


                                //if (Directory.Exists(ClientFolder) && (BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                //{

                                //    if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && ClientFolderFilePath != "&nbsp;" && ClientFolderFilePath != "")
                                //    {
                                //        if (System.IO.File.Exists(ClientFolderFilePath))
                                //        {
                                //            if (strExistingFiles == "")
                                //                strExistingFiles = "<br/>" + ClientFolderFilePath;
                                //            else
                                //                strExistingFiles = strExistingFiles + ",<br/>" + ClientFolderFilePath;

                                //            // ExistingFileName.Add(ClientFolderFilePath);
                                //        }

                                //        /** Commented
                                //        //try
                                //        //{
                                //        //    File.Copy(BatchFilePath, ClientFolderFilePath, true);
                                //        //}
                                //        //catch (Exception ex)
                                //        //{ }
                                //        ****/

                                //        //SetFileReadAccess(ClientFolderFilePath, true);
                                //        //Response.Write("<br/>File Copied: " + ClientFolder + "<br/>");
                                //    }
                                //    else
                                //        ClientFolderFilePath = "";
                                //}
                                //else if (ClientFolder != "")
                                //{

                                //    BatchName = (string)row.Cells[4].Text;

                                //    if (strFolderExists == "")
                                //    {
                                //        strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Client folder path: " + ClientFolder;
                                //    }
                                //    else
                                //        strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Client folder path: " + ClientFolder;
                                //}

                                //// SubFolder Region
                                //if (Directory.Exists(SubFolderFilePath) && (BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                //{

                                //    if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && SubFolderFilePath != "&nbsp;" && SubFolderFilePath != "")
                                //    {
                                //        if (System.IO.File.Exists(SubFolderFilePath))
                                //        {
                                //            if (strExistingFiles == "")
                                //                strExistingFiles = "<br/>" + strSubFolderPath;
                                //            else
                                //                strExistingFiles = strExistingFiles + ",<br/>" + strSubFolderPath;

                                //            //ExistingFileName.Add("");
                                //        }

                                //        try
                                //        {
                                //            System.IO.File.Copy(BatchFilePath, strSubFolderPath, true);
                                //        }
                                //        catch (Exception ex)
                                //        { }
                                //        //SetFileReadAccess(ClientFolderFilePath, true);
                                //        //Response.Write("<br/>File Copied: " + ClientFolder + "<br/>");
                                //    }

                                //}
                                //else if (SubFolder != "")
                                //{
                                //    Directory.CreateDirectory(SubFolderFilePath);
                                //    if (Directory.Exists(SubFolderFilePath) && (BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                //    {
                                //        if (System.IO.File.Exists(strSubFolderPath))
                                //        {
                                //            if (strExistingFiles == "")
                                //                strExistingFiles = "<br/>" + strSubFolderPath;
                                //            else
                                //                strExistingFiles = strExistingFiles + ",<br/>" + strSubFolderPath;
                                //        }

                                //        try
                                //        {
                                //            System.IO.File.Copy(BatchFilePath, strSubFolderPath, true);
                                //        }
                                //        catch (Exception ex)
                                //        { }

                                //    }
                                //}

                                #endregion


                                //  bool billingflag = Convert.ToBoolean(ViewState["BillingFlg"]);
                                string BatchName = string.Empty;
                                string SubFolderName = string.Empty;
                                string DestinationFileName = string.Empty;
                                string SubFolder = string.Empty;

                                BatchName = (string)row.Cells[4].Text;
                                SubFolderName = (string)row.Cells[32].Text;
                                string CompletedMailings = AppLogic.GetParam(AppLogic.ConfigParam.CompletedMailings);
                                //\\GRPAO1-VWFS01\shared$\Mail Merge\Completed Mailings
                                //\\GRPAO1-VWFS01\shared$\Mail Merge\Completed Mailings
                                // SubFolder = "\\\\GRPAO1-VWFS01\\shared$\\Mail Merge\\Completed Mailings\\" + SubFolderName;
                                SubFolder = CompletedMailings + SubFolderName; // shared drive changes - 7_4_2019
                                                                               //SubFolder = "\\\\GRPAO1-VWFS01\\opsreports$\\Mail Merge\\Completed Mailings\\" + SubFolderName;

                                string BatchFilePath = (string)row.Cells[18].Text;
                                string AsOfDate = (string)row.Cells[3].Text;
                                string[] date = AsOfDate.Split(new char[] { '/' });
                                string year = date[2];

                                // New Sharepoint Client Services site Changes 
                                DateTime dtAsOfDate = Convert.ToDateTime(AsOfDate);
                                string Quarter = GetQuarter(dtAsOfDate);


                                //string BatchFileName = (string)row.Cells[21].Text;  // BatchFilePath.Substring(BatchFilePath.LastIndexOf("\\") + 1); // 
                                string SPVFileName = row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace("&nbsp;", "");

                                //BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "");
                                if (chkTest.Checked == true)
                                {
                                    if (SPVFileName != "")
                                    {
                                        // New Sharepoint Client Services site Changes 
                                        // BatchFileName = SPVFileName.Replace(".pdf", "_Test.pdf");
                                        BatchFileName = "zzTest_" + SPVFileName;
                                    }
                                    else
                                    {
                                        // New Sharepoint Client Services site Changes 
                                        // BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(".pdf", "_Test.pdf");
                                        BatchFileName = "zzTest_" + BatchFileName.Replace(" Family", "").Replace(",", "");
                                    }
                                }
                                else
                                {
                                    if (SPVFileName != "")
                                    {
                                        BatchFileName = SPVFileName;
                                    }
                                    else
                                    {
                                        BatchFileName = BatchFileName.Replace(" Family", "").Replace(",", "");
                                    }
                                }

                                string ClientFolder = (string)row.Cells[19].Text == "&nbsp;" || (string)row.Cells[19].Text == "" ? "" : (string)row.Cells[19].Text;
                                //  string SharepointFolder = (string)row.Cells[20].Text == "&nbsp;" || (string)row.Cells[20].Text == "" ? "" : (string)row.Cells[20].Text;

                                ClientFolder = ClientFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();
                                // SharepointFolder = SharepointFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();


                                string ClientFolderFilePath = ClientFolder + "\\" + BatchFileName;
                                //string SharepointFolderFilePath = SharepointFolder + "\\" + BatchFileName;

                                //string ClientFolderFilePath = ClientFolder + "//" + BatchFileName;
                                //string SharepointFolderFilePath = SharepointFolder + "//" + BatchFileName;

                                string SubFolderFilePath = SubFolder;

                                //string SharepointFolder = "C:\\Reports" + "\\" + BatchFileName;

                                ClientFolderFilePath = ClientFolderFilePath.Replace("%20", " ");
                                // SharepointFolderFilePath = SharepointFolderFilePath.Replace("%20", " ");

                                string strSubFolderPath = SubFolderFilePath + "\\" + BatchFileName;
                                strSubFolderPath = strSubFolderPath.Replace("%20", " ").Replace("&#39;", "'").ToString();
                                string Batch_Name1 = row.Cells[38].Text;

                                DataSet dtDocTax = null;
                                DataTable dtHouseholdUUID = null;
                                DataTable dtLegalEntityUUID = null;
                                DataTable dtActiveClient = null;
                                string ClientServicesLink = string.Empty;
                                string NewCSSharepointFolderPath = string.Empty;
                                string FinalClientServiceLink = string.Empty;
                                bool bLegalEntityFolder = false;
                                bool IsSharepointSiteExists = false;

                                if ((BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                {
                                    //Commented New Sharepoint Client Services site Changes 
                                    //  if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && SharepointFolderFilePath != "" && SharepointFolderFilePath != "&nbsp;")
                                    if (BatchFilePath != "&nbsp;" && BatchFilePath != "")// && ssi_CSCiteUUID != "" && ssi_CSCiteUUID != "&nbsp;")
                                    {
                                        BatchName = (string)row.Cells[4].Text;
                                        //if (SharepointFolder != "")
                                        //f (ssi_CSCiteUUID != "")
                                        if (ssi_CSCiteUUID != "" && ssi_CSCiteUUID != "&nbsp;")
                                        {

                                            if (ssi_SPSiteType == "100000000" && ssi_SPLEFolder != "")//CS Household with LegalEntity Folder
                                            {
                                                bLegalEntityFolder = true;
                                            }

                                            if (bLegalEntityFolder) // For LegalEntity Folder inside Household Library
                                            {
                                                dtActiveClient = (DataTable)ViewState["dtActiveClientList"];
                                                //dtDocTax = (DataSet)ViewState["dtDocumentTaxonomy"];

                                                //dtHouseholdUUID = dtDocTax.Tables[0];
                                                //dtLegalEntityUUID = dtDocTax.Tables[1];


                                                NewCSSharepointFolderPath = sp.FetchNewSpURL(dtActiveClient, ssi_CSCiteUUID);

                                                // lg.AddinLogFile(LogFileName, "ssi_CSCiteUUID " + ssi_CSCiteUUID);
                                                // lg.AddinLogFile(LogFileName, "NewCSSharepointFolderPath " + NewCSSharepointFolderPath);
                                                // lg.AddinLogFile(LogFileName, "ssi_SPLEFolder " + ssi_SPLEFolder);
                                                IsSharepointSiteExists = CheckLEgalEntityFolderExist(NewCSSharepointFolderPath, "PublishedDocuments", ssi_SPLEFolder);
                                                //  lg.AddinLogFile(LogFileName, "IsSharepointSiteExists " + IsSharepointSiteExists);
                                                if (IsSharepointSiteExists)
                                                {
                                                    try
                                                    {

                                                        // if (checkSharepouintFileExist(sharepointpath, BatchFileName))
                                                        if (CheckFileExistinLegalEntity(NewCSSharepointFolderPath, "PublishedDocuments", ssi_SPLEFolder, BatchFileName))
                                                        {
                                                            // lg.AddinLogFile(LogFileName, "CheckFileExistinLegalEntity nside");
                                                            if (strExistingFiles == "")
                                                                strExistingFiles = "<br/>" + NewCSSharepointFolderPath + "/PublishedDocuments/" + ssi_SPLEFolder + "/" + BatchFileName;
                                                            else
                                                                strExistingFiles = strExistingFiles + ",<br/>" + NewCSSharepointFolderPath + "/PublishedDocuments/" + ssi_SPLEFolder + "/" + BatchFileName;

                                                            //ExistingFileName.Add(newsharepointPath + "/" + BatchFileName);

                                                        }

                                                        // New Sharepoint Client Services site Changes 
                                                        //  CopyFilenew(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);
                                                        string BatchType1 = row.Cells[37].Text;
                                                        // lg.AddinLogFile(LogFileName, "CopyFileinLegalEntityFolder before");
                                                        // ClientServicesLink = CopyFilenewCS(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, year, BatchType1, Quarter, LegalEntityUUID, dtLegalEntityUUID);
                                                        ClientServicesLink = CopyFileinLegalEntityFolder(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, "PublishedDocuments/" + ssi_SPLEFolder, year, BatchType1, Quarter, ssi_billingid);

                                                        // ClientServicesLink = CopyFileinLegalEntityFolder(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, "PublishedDocuments/" + ssi_SPLEFolder, year, BatchType1, Quarter, LegalEntityUUID, billingFlag);
                                                        //  lg.AddinLogFile(LogFileName, "ClientServicesLink " + ClientServicesLink);
                                                        if (ClientServicesLink != "")
                                                        {
                                                            // ClientServicesLink = "https://greshampartners.sharepoint.com" + ClientServicesLink;
                                                            ClientServicesLink = AppLogic.GetParam(AppLogic.ConfigParam.SharepointURL) + ClientServicesLink;
                                                            FinalClientServiceLink = ClientServicesLink.Replace(BatchFileName, "");
                                                            SucessClientName.Add(Batch_Name1);
                                                            SucessFileName.Add(FinalClientServiceLink);

                                                            FinalClientServiceLink = FinalClientServiceLink.Remove(FinalClientServiceLink.Length - 1);// remove last "/" from URL
                                                            FinalClientServiceLink = HttpUtility.UrlPathEncode(FinalClientServiceLink); // Encode String to URL
                                                        }
                                                    }
                                                    catch (Exception exc)
                                                    {
                                                        // lg.AddinLogFile(LogFileName, "DErrorr " + exc.Message.ToString());
                                                        // strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                                        Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                               " to " + NewCSSharepointFolderPath + "/PublishedDocuments/" + ssi_SPLEFolder + "<br/>" + exc.Message + ", " + exc.StackTrace);
                                                    }
                                                }
                                                else
                                                {
                                                    string BatchType1 = row.Cells[37].Text;



                                                    // ClientServicesLink = CopyFilenewCS(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, year, BatchType1, Quarter, LegalEntityUUID, billingflag);

                                                    ClientServicesLink = CopyFilenewCS(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, year, BatchType1, Quarter, ssi_billingid);

                                                    // strFolderExists = strFolderExists + "<br/>File Not saved to Client Services <br/>Legal Entity Folder not found of the below batch<br/> Batch Name: " + Batch_Name1;
                                                    
                                                    //string ClientServicesFinalLink = "https://greshampartners.sharepoint.com" + ClientServicesLink.Trim().Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                                    string ClientServicesFinalLink = AppLogic.GetParam(AppLogic.ConfigParam.SharepointURL) + ClientServicesLink.Trim().Replace(" ", "%20").Replace("&#39;", "'").ToString();

                                                    //  Response.Write("Before" + ClientServicesFinalLink);
                                                    //FinalClientServiceLink = "https://greshampartners.sharepoint.com" + ClientServicesLink.Replace(BatchFileName, "");



                                                    //FinalClientServiceLink = FinalClientServiceLink.Remove(FinalClientServiceLink.Length - 1);// remove last "/" from URL
                                                    //FinalClientServiceLink = HttpUtility.UrlPathEncode(FinalClientServiceLink); // Encode String to URL

                                                    //strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + ClientServicesFinalLink.Trim().Replace(" ", "%20").Replace("&#39;", "'").ToString(); ;

                                                    //if (ClientServicesLink != "")
                                                    //{
                                                    //    ClientServicesLink = "https://greshampartners.sharepoint.com" + ClientServicesLink;

                                                    //    FinalClientServiceLink = ClientServicesLink.Replace(BatchFileName, "");
                                                    //    SucessClientName.Add(Batch_Name1);
                                                    //    SucessFileName.Add(FinalClientServiceLink);

                                                    //    FinalClientServiceLink = FinalClientServiceLink.Remove(FinalClientServiceLink.Length - 1);// remove last "/" from URL
                                                    //    FinalClientServiceLink = HttpUtility.UrlPathEncode(FinalClientServiceLink); // Encode String to URL
                                                    //}
                                                    string link = NewCSSharepointFolderPath + "/" + "PublishedDocuments";

                                                    // strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + NewCSSharepointFolderPath + "/" + "PublishedDocuments";



                                                    //strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + "<a href='" + link + "'>clickme</a>";

                                                    strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + "<a href='" + link + "' target=_blank >" + link + "</a>";

                                                    //  strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + NewCSSharepointFolderPath + "/" + "PublishedDocuments";

                                                    // Response.Write(FinalClientServiceLink);

                                                    //  strFolderExists = strFolderExists + "<br/>The Legal Entity folder not found. The file is saved at the below path:" + "<br/>" + FinalClientServiceLink;

                                                    SendEmail(ssi_SPLEFolder, ClientServicesFinalLink);

                                                    // Response.Write("after" + ClientServicesFinalLink);
                                                }
                                            }
                                            else
                                            {

                                                //  lg.AddinLogFile(LogFileName, "Omnly Household " );
                                                dtActiveClient = (DataTable)ViewState["dtActiveClientList"];
                                                //dtDocTax = (DataSet)ViewState["dtDocumentTaxonomy"];

                                                //dtHouseholdUUID = dtDocTax.Tables[0];
                                                //dtLegalEntityUUID = dtDocTax.Tables[1];

                                                // NewCSSharepointFolderPath = sp.FetchSharepointLink(dtHouseholdUUID, dtActiveClient, ssi_CSCiteUUID);
                                                NewCSSharepointFolderPath = sp.FetchNewSpURL(dtActiveClient, ssi_CSCiteUUID);

                                                // bool IsSharepointFolderExists = CheckFolderPathExists(SharepointFolder);
                                                IsSharepointSiteExists = CheckNewCSSiteExists(NewCSSharepointFolderPath, ssi_billingid);
                                                if (IsSharepointSiteExists)
                                                {
                                                    try
                                                    {

                                                        // if (checkSharepouintFileExist(sharepointpath, BatchFileName))
                                                        if (CheckFileExistinNewCSSite(NewCSSharepointFolderPath, BatchFileName))
                                                        {
                                                            if (strExistingFiles == "")
                                                                strExistingFiles = "<br/>" + NewCSSharepointFolderPath + "/PublishedDocuments/" + BatchFileName;
                                                            else
                                                                strExistingFiles = strExistingFiles + ",<br/>" + NewCSSharepointFolderPath + "/PublishedDocuments/" + BatchFileName;

                                                            //ExistingFileName.Add(newsharepointPath + "/" + BatchFileName);

                                                        }

                                                        // New Sharepoint Client Services site Changes 
                                                        //  CopyFilenew(SharepointFolder, Path.GetFileName(SharepointFolderFilePath), BatchFilePath);
                                                        string BatchType1 = row.Cells[37].Text;
                                                        // ClientServicesLink = CopyFilenewCS(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, year, BatchType1, Quarter, LegalEntityUUID, billingflag);

                                                        ClientServicesLink = CopyFilenewCS(NewCSSharepointFolderPath, BatchFileName, BatchFilePath, year, BatchType1, Quarter, ssi_billingid);


                                                        if (ClientServicesLink != "")
                                                        {
                                                            //   ClientServicesLink = "https://greshampartners.sharepoint.com" + ClientServicesLink;
                                                            ClientServicesLink = AppLogic.GetParam(AppLogic.ConfigParam.SharepointURL) + ClientServicesLink;

                                                            FinalClientServiceLink = ClientServicesLink.Replace(BatchFileName, "");
                                                            SucessClientName.Add(Batch_Name1);
                                                            SucessFileName.Add(FinalClientServiceLink);

                                                            FinalClientServiceLink = FinalClientServiceLink.Remove(FinalClientServiceLink.Length - 1);// remove last "/" from URL
                                                            FinalClientServiceLink = HttpUtility.UrlPathEncode(FinalClientServiceLink); // Encode String to URL
                                                        }
                                                    }
                                                    catch (Exception exc)
                                                    {
                                                        // strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + newsharepointPath;
                                                        Response.Write("<br/>Error Occured when trying to copy file from: " + BatchFilePath +
                                               " to " + NewCSSharepointFolderPath + "<br/>" + exc.Message + ", " + exc.StackTrace);
                                                    }

                                                }
                                                else
                                                {
                                                    // strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + NewCSSharepointFolderPath;
                                                    if (ssi_SPSiteType == "100000000")
                                                    {
                                                        strFolderExists = strFolderExists + "<br/>File Not saved to Client Services <br/>Client Services Site not found for the CS household of the below batch <br/> Batch Name: " + Batch_Name1;
                                                    }
                                                    else
                                                    {
                                                        strFolderExists = strFolderExists + "<br/>File Not saved to Client Services <br/>Client Services Site not found for the CS LegalEntity of the below batch <br/> Batch Name: " + Batch_Name1;
                                                    }
                                                }
                                            }

                                        }
                                        else
                                        {
                                            // lblError2.Text = lblError2.Text + " " + "Eror Fetching URL for Client Services";
                                            if (strFolderExists == "")
                                            {
                                                //   strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Sharepoint folder path: " + NewCSSharepointFolderPath;
                                                if (ssi_SPSiteType == "100000000")//CS Household
                                                {

                                                    strFolderExists = strFolderExists + "<br/>File Not saved to Client Services<br/>CS Household is Empty for the below batch,SP Site type = Household <br/>Batch Name: " + Batch_Name1;
                                                }
                                                else
                                                {
                                                    strFolderExists = strFolderExists + "<br/>File Not saved to Client Services<br/>CS LegalEntity is Empty for the below batch,SP Site type = LegalEntity <br/>Batch Name: " + Batch_Name1;
                                                }

                                            }
                                            else
                                            {
                                                strFolderExists = strFolderExists + "<br/>Batch Name: " + Batch_Name1 + "<br/>Sharepoint folder path: " + NewCSSharepointFolderPath;
                                            }
                                        }
                                    }
                                    else
                                        NewCSSharepointFolderPath = "";
                                }
                                else if (ssi_CSCiteUUID != "")
                                {
                                    BatchName = (string)row.Cells[4].Text;

                                }
                                if (ClientServicesLink != "")
                                {
                                    Count++;
                                    #region Update Batch Status to 'Final Report Sent'

                                    //ssi_batch objBatch = new ssi_batch();
                                    Entity objBatch = new Entity("ssi_batch");

                                    //objBatch.ssi_batchid = new Key();
                                    //objBatch.ssi_batchid.Value = new Guid(ssi_batchid);
                                    objBatch["ssi_batchid"] = new Guid(ssi_batchid);

                                    if (ClientServicesLink != "")
                                    {
                                        //objBatch.ssi_sharepointreportfolderfinal = FinalSharepointFolder.Replace(" Family", "").Replace(",", "%2C").Replace(" ", "%20").Replace("'", "%27").Replace("&#39;", "'").ToString();
                                        // objBatch["ssi_sharepointreportfolderfinal"] = FinalSharepointFolder.Replace(" Family", "").Replace(",", "%2C").Replace(" ", "%20").Replace("'", "%27").Replace("&#39;", "'").ToString();

                                        objBatch["ssi_sharepointreportfolderfinal"] = HttpUtility.UrlPathEncode(FinalClientServiceLink); //FinalClientServiceLink;
                                    }


                                    if (row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace("&nbsp;", "") != "")
                                    {
                                        //objBatch.ssi_sharepointemaillink = (string)row.Cells[33].Text.Trim().Replace("ssi_spvfilename", "").Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                        objBatch["ssi_sharepointemaillink"] = BatchFileName.Trim().Replace("ssi_spvfilename", "").Replace(" ", "%20").Replace("&#39;", "'").ToString();
                                    }
                                    else
                                    {
                                        //objBatch.ssi_sharepointemaillink = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(" ", "%20");
                                        objBatch["ssi_sharepointemaillink"] = BatchFileName.Replace(" Family", "").Replace(",", "").Replace(" ", "%20");
                                    }

                                    if (MailPref.ToUpper().Contains("EMAIL"))
                                    {
                                        //objBatch.ssi_finalreportcreatedflag = new CrmBoolean();
                                        //objBatch.ssi_finalreportcreatedflag.Value = true;
                                        objBatch["ssi_finalreportcreatedflag"] = true;

                                        updateSentData(ssi_secondaryownerid, ssi_mailrecordsId);
                                    }
                                    else
                                    {
                                        //objBatch.ssi_sendemailib = new CrmBoolean();
                                        //objBatch.ssi_sendemailib.Value = true;
                                    }

                                    //objBatch.ssi_sendemailib = new CrmBoolean();
                                    //objBatch.ssi_sendemailib.Value = true;
                                    objBatch["ssi_sendemailib"] = true;


                                    if (MailPref.ToUpper().Contains("EMAIL") && ssi_ReviewReqByid != "")
                                    {
                                        // Send Email to Review Required By 
                                        //objBatch.ssi_reviewrequiredbyid = new Lookup();
                                        //objBatch.ssi_reviewrequiredbyid.type = EntityName.systemuser.ToString();
                                        //objBatch.ssi_reviewrequiredbyid.Value = new Guid(ssi_ReviewReqByid);
                                        objBatch["ssi_reviewrequiredbyid"] = new EntityReference("systemuser", new Guid(ssi_ReviewReqByid));

                                        //objBatch.ssi_sendrrbymail = new CrmBoolean();
                                        //objBatch.ssi_sendrrbymail.Value = true;
                                        objBatch["ssi_sendrrbymail"] = true;

                                        //objBatch.ssi_reporttrackerstatus = new Picklist();
                                        //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Sent'
                                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                        // intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                        intStatus = 4;



                                        //BatchReportStatus(intStatus, UserId, ssi_batchid, BillingHandedOff);
                                        BillingHandedOff = "true";
                                    }
                                    else if (MailPref.ToUpper().Contains("EMAIL") && ReviewRequiredBy == "")
                                    {
                                        // Send Email to Associate
                                        //objBatch.ssi_reporttrackerstatus = new Picklist();
                                        //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Sent'
                                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                        //objBatch.ssi_sendmailassociate = new CrmBoolean();
                                        //objBatch.ssi_sendmailassociate.Value = true;
                                        objBatch["ssi_sendmailassociate"] = true;

                                        //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                        intStatus = 4;
                                        //BatchReportStatus(intStatus, UserId, ssi_batchid, BillingHandedOff);
                                        BillingHandedOff = "true";
                                    }
                                    else
                                    {
                                        int trackerstatus;
                                        if (ssi_internalbillingcontactid != "")
                                        {
                                            //objBatch.ssi_billinghandedoff = new CrmBoolean();
                                            //objBatch.ssi_billinghandedoff.Value = true;
                                            objBatch["ssi_billinghandedoff"] = true;
                                            BillingHandedOff = "true";
                                        }
                                        else
                                        {
                                            //objBatch.ssi_billinghandedoff = new CrmBoolean();
                                            //objBatch.ssi_billinghandedoff.Value = false;
                                            objBatch["ssi_billinghandedoff"] = false;
                                            BillingHandedOff = "false";
                                        }


                                        if (MailPref.ToUpper() == "Client Portal".ToUpper())
                                        {
                                            //objBatch.ssi_reporttrackerstatus = new Picklist();
                                            //objBatch.ssi_reporttrackerstatus.Value = 4;// Batch status 'Final Report Sent'
                                            objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                            trackerstatus = 4;

                                            //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                            updateSentData(ssi_ReviewReqByid, ssi_mailrecordsId);
                                        }
                                        else
                                        {
                                            //objBatch.ssi_reporttrackerstatus = new Picklist();
                                            //objBatch.ssi_reporttrackerstatus.Value = 9;// Batch status 'Final Report Sent'
                                            objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(9);
                                            trackerstatus = 9;
                                        }

                                        //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                        intStatus = trackerstatus;

                                    }

                                    if (ssi_ReviewReqByid != "" && MailPref.ToUpper().Contains("EMAIL"))
                                    {
                                        lblError.Text = "Some of the reports you selected need to be given to the person designated in Review required by.";
                                    }

                                    service.Update(objBatch);

                                    bool billingHandedOff = Convert.ToBoolean(BillingHandedOff);
                                    BatchReportStatus(intStatus, UserId, ssi_batchid, billingHandedOff);

                                    #endregion

                                    #region Update Mail status to 'Created'
                                    if (MailPref.ToUpper().Contains("EMAIL") && ssi_ReviewReqByid != "")
                                    {
                                        //objMailRecords.ssi_mailstatus = new Picklist();
                                        //objMailRecords.ssi_mailstatus.Value = 3;//mail status 'Sent to FINAL Reviewer'
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(3);
                                    }
                                    else if ((MailPref.ToUpper().Contains("EMAIL") || MailPref.ToUpper() == "Client Portal".ToUpper()) && ssi_ReviewReqByid == "")
                                    {
                                        //objMailRecords.ssi_mailstatus = new Picklist();
                                        //objMailRecords.ssi_mailstatus.Value = 4;//mail status 'Sent to FINAL Reviewer'
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                    }
                                    else if (MailPref.ToUpper() == "Client Portal".ToUpper() && ssi_ReviewReqByid != "")
                                    {
                                        //objMailRecords.ssi_mailstatus = new Picklist();
                                        //objMailRecords.ssi_mailstatus.Value = 4;//mail status 'Sent to FINAL Reviewer'
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                    }
                                    else
                                    {
                                        //objMailRecords.ssi_mailstatus = new Picklist();
                                        //objMailRecords.ssi_mailstatus.Value = 8;//mail status 'Created'
                                        if (ssi_billingid != "")
                                            objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                        else
                                            objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(8);

                                    }



                                    service.Update(objMailRecords);
                                    #endregion


                                }
                                #region Commennted New Sharepoint Client Services site Changes 
                                //if (Directory.Exists(ClientFolder) && (BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                //{

                                //    if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && ClientFolderFilePath != "&nbsp;" && ClientFolderFilePath != "")
                                //    {
                                //        if (System.IO.File.Exists(ClientFolderFilePath))
                                //        {
                                //            if (strExistingFiles == "")
                                //                strExistingFiles = "<br/>" + ClientFolderFilePath;
                                //            else
                                //                strExistingFiles = strExistingFiles + ",<br/>" + ClientFolderFilePath;

                                //            // ExistingFileName.Add(ClientFolderFilePath);
                                //        }


                                //    }
                                //    else
                                //        ClientFolderFilePath = "";
                                //}
                                //else if (ClientFolder != "")
                                //{

                                //    BatchName = (string)row.Cells[4].Text;

                                //    if (strFolderExists == "")
                                //    {
                                //        strFolderExists = strFolderExists + "<br/>Below folder path not found for related batch <br/>Batch Name: " + BatchName + "<br/>Client folder path: " + ClientFolder;
                                //    }
                                //    else
                                //        strFolderExists = strFolderExists + "<br/>Batch Name: " + BatchName + "<br/>Client folder path: " + ClientFolder;
                                //}
                                #endregion
                                // SubFolder Region
                                if (Directory.Exists(SubFolderFilePath) && (BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                {

                                    if (BatchFilePath != "&nbsp;" && BatchFilePath != "" && SubFolderFilePath != "&nbsp;" && SubFolderFilePath != "")
                                    {
                                        if (System.IO.File.Exists(SubFolderFilePath))
                                        {
                                            if (strExistingFiles == "")
                                                strExistingFiles = "<br/>" + strSubFolderPath;
                                            else
                                                strExistingFiles = strExistingFiles + ",<br/>" + strSubFolderPath;

                                            //ExistingFileName.Add("");
                                        }

                                        try
                                        {
                                            System.IO.File.Copy(BatchFilePath, strSubFolderPath, true);
                                        }
                                        catch (Exception ex)
                                        { }
                                        //SetFileReadAccess(ClientFolderFilePath, true);
                                        //Response.Write("<br/>File Copied: " + ClientFolder + "<br/>");
                                    }

                                }
                                else if (SubFolder != "")
                                {
                                    Directory.CreateDirectory(SubFolderFilePath);
                                    if (Directory.Exists(SubFolderFilePath) && (BatchFilePath != "&nbsp;" && BatchFilePath != ""))
                                    {
                                        if (System.IO.File.Exists(strSubFolderPath))
                                        {
                                            if (strExistingFiles == "")
                                                strExistingFiles = "<br/>" + strSubFolderPath;
                                            else
                                                strExistingFiles = strExistingFiles + ",<br/>" + strSubFolderPath;
                                        }

                                        try
                                        {
                                            System.IO.File.Copy(BatchFilePath, strSubFolderPath, true);
                                        }
                                        catch (Exception ex)
                                        { }

                                    }
                                }



                            }

                            else if (ddlAction.SelectedValue == "5")//Mark all records Canceled 
                            {
                                if (ssi_mailrecordsId != "")// && ssi_batchid == "")
                                {
                                    //objMailRecords.ssi_mailrecordsid = new Key();
                                    //objMailRecords.ssi_mailrecordsid.Value = new Guid(ssi_mailrecordsId);
                                    objMailRecords["ssi_mailrecordsid"] = new Guid(ssi_mailrecordsId);

                                    //objMailRecords.ssi_mailstatus = new Picklist();
                                    //objMailRecords.ssi_mailstatus.Value = 6;//Canceled
                                    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(6);

                                    service.Update(objMailRecords);
                                }

                                #region Update Batch

                                //ssi_batch objBatch = new ssi_batch();
                                Entity objBatch = new Entity("ssi_batch");
                                if (ssi_batchid != "")
                                {
                                    //objBatch.ssi_batchid = new Key();
                                    //objBatch.ssi_batchid.Value = new Guid(ssi_batchid);
                                    objBatch["ssi_batchid"] = new Guid(ssi_batchid);


                                    //objBatch.ssi_reporttrackerstatus = new Picklist();
                                    //objBatch.ssi_reporttrackerstatus.Value = 6; //Pend Approval
                                    objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(6);
                                    intStatus = 6;
                                    if (BatchStatus == "8" || BatchStatus == "1" || BatchStatus == "9" || BatchStatus == "4" || BatchStatus == "5")
                                    {
                                        //objBatch.ssi_billingcontactchange = new CrmBoolean();
                                        //objBatch.ssi_billingcontactchange.Value = true;//To check email notification when status is changed to 'Pend Approval'
                                        objBatch["ssi_billingcontactchange"] = true;
                                    }
                                    else if (BatchStatus == "7")//OPS Change Requested
                                    {
                                        //objBatch.ssi_opschangecomplete = new CrmBoolean();
                                        //objBatch.ssi_opschangecomplete.Value = true;//Snd Email When Status is changed from 'OPS Change Requested ' to 'Pend Approval'
                                        objBatch["ssi_opschangecomplete"] = true;
                                    }

                                    //objBatch.ssi_batchdisplayfilename = "";
                                    objBatch["ssi_batchdisplayfilename"] = "";

                                    //objBatch.ssi_batchfilename = "";
                                    objBatch["ssi_batchfilename"] = "";


                                    //SecurityPrincipal assignee = new SecurityPrincipal();
                                    //assignee.PrincipalId = new Guid(ssi_secondaryownerid);///HouseHold Secondary Owner ID

                                    //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();
                                    //targetAssign.EntityId = new Guid(ssi_batchid);
                                    //targetAssign.EntityName = EntityName.ssi_batch.ToString();

                                    //AssignRequest assign = new AssignRequest();
                                    //assign.Assignee = assignee;
                                    //assign.Target = targetAssign;

                                    //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);

                                    //  service.Update(objBatch);



                                    AssignRequest assignRequest = new AssignRequest
                                    {
                                        Assignee = new EntityReference("systemuser",
                                            new Guid(ssi_secondaryownerid)),
                                        Target = new EntityReference("ssi_batch",
                                            new Guid(ssi_batchid))
                                    };


                                    service.Execute(assignRequest);
                                    service.Update(objBatch);






                                    // intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                    BatchReportStatus(intStatus, UserId, ssi_batchid, Convert.ToBoolean(BillingHandedOff));
                                }

                                #endregion

                                lblError.Text = "Records Canceled Successfully";
                            }
                            else if (ddlAction.SelectedValue == "6")//Mark All Records Sent
                            {
                                int intResult = 0;
                                if (ssi_mailrecordsId != "")// && ssi_batchid == "")
                                {
                                    //objMailRecords.ssi_mailrecordsid = new Key();
                                    //objMailRecords.ssi_mailrecordsid.Value = new Guid(ssi_mailrecordsId);
                                    objMailRecords["ssi_mailrecordsid"] = new Guid(ssi_mailrecordsId);

                                    //objMailRecords.ssi_mailstatus = new Picklist();
                                    if (ssi_ReviewReqByid != "")
                                    {
                                        //objMailRecords.ssi_mailstatus.Value = 3;// Sent to Final Reviewer
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(3);

                                    }
                                    else
                                    {
                                        //objMailRecords.ssi_mailstatus.Value = 4;//Sent
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);
                                    }

                                    service.Update(objMailRecords);
                                    intResult++;

                                    //ssi_batch objBatch = new ssi_batch();
                                    Entity objBatch = new Entity("ssi_batch");

                                    if (ssi_batchid != "")
                                    {
                                        //objBatch.ssi_batchid = new Key();
                                        //objBatch.ssi_batchid.Value = new Guid(ssi_batchid);
                                        objBatch["ssi_batchid"] = new Guid(ssi_batchid);

                                        //objBatch.ssi_reporttrackerstatus = new Picklist();
                                        //objBatch.ssi_reporttrackerstatus.Value = 4; //Sent
                                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                        service.Update(objBatch);
                                        intResult++;

                                        //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                        intStatus = 4;
                                        BatchReportStatus(intStatus, UserId, ssi_batchid, Convert.ToBoolean(BillingHandedOff));
                                    }


                                    lblError.Text = "All Records Sent Successfully";
                                }

                                if (intResult > 0)
                                {
                                    bMarkAllRecordsSent = true;
                                }

                            }//                    if (MailType.ToUpper() != "EMAIL" || MailType.ToUpper() != "REGULAR MAIL AND EMAIL" || MailType.ToUpper() != "FEDEX - 2DAY AND EMAIL" || MailType != "FEDEX - OVERNIGHT AND EMAIL")
                            else if (ddlAction.SelectedValue == "7" && (MailPref.ToUpper().Contains("EMAIL") || MailType.ToUpper() == "REGULAR MAIL AND EMAIL" || MailType.ToUpper() == "FEDEX - 2DAY AND EMAIL" || MailType.ToUpper() != "FEDEX - OVERNIGHT AND EMAIL")) //Send Email to Associate and Mark Sent 
                            {
                                if (ssi_batchid != "")
                                {
                                    #region Update Batch Status Send Email Flag to TRUE
                                    //ssi_batch objBatch = new ssi_batch();
                                    Entity objBatch = new Entity("ssi_batch");

                                    //objBatch.ssi_batchid = new Key();
                                    //objBatch.ssi_batchid.Value = new Guid(ssi_batchid);
                                    objBatch["ssi_batchid"] = new Guid(ssi_batchid);

                                    //objBatch.ssi_reporttrackerstatus = new Picklist();
                                    //objBatch.ssi_reporttrackerstatus.Value = 4;//batch report tracker status 'Sent'
                                    objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                    //objBatch.ssi_sendemail = new CrmBoolean();
                                    //objBatch.ssi_sendemail.Value = true;
                                    objBatch["ssi_sendemail"] = true;

                                    service.Update(objBatch);

                                    //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                    intStatus = 4;

                                    BatchReportStatus(intStatus, UserId, ssi_batchid, Convert.ToBoolean(BillingHandedOff));

                                    //Response.Write("Batch Send flag updated");
                                    #endregion
                                }

                                if (ssi_mailrecordsId != "")
                                {
                                    //objMailRecords.ssi_mailrecordsid = new Key();
                                    //objMailRecords.ssi_mailrecordsid.Value = new Guid(ssi_mailrecordsId);
                                    objMailRecords["ssi_mailrecordsid"] = new Guid(ssi_mailrecordsId); ;

                                    //objMailRecords.ssi_mailstatus = new Picklist();
                                    //objMailRecords.ssi_mailstatus.Value = 4;//Sent
                                    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);


                                    //Response.Write("Mail record Send flag updated");
                                    service.Update(objMailRecords);
                                }

                                lblError.Text = "Send Email to Associate and Mark Sent action completed.";
                            }
                            else if (ddlAction.SelectedValue == "7")//Send Email to Associate and Mark Sent
                            {
                                if (ssi_batchid != "")
                                {
                                    #region Update Batch Status Send Email Flag to TRUE
                                    //ssi_batch objBatch = new ssi_batch();
                                    Entity objBatch = new Entity("ssi_batch");

                                    //objBatch.ssi_batchid = new Key();
                                    //objBatch.ssi_batchid.Value = new Guid(ssi_batchid);
                                    objBatch["ssi_batchid"] = new Guid(ssi_batchid);


                                    //objBatch.ssi_reporttrackerstatus = new Picklist();
                                    //objBatch.ssi_reporttrackerstatus.Value = 4;//batch report tracker status 'Sent'
                                    objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                    //objBatch.ssi_sendemail = new CrmBoolean();
                                    //objBatch.ssi_sendemail.Value = true;
                                    objBatch["ssi_sendemail"] = true;

                                    //objBatch.ssi_sendemailib = new CrmBoolean(); // send email to 'Internal Billing' 
                                    //objBatch.ssi_sendemailib.Value = true;
                                    objBatch["ssi_sendemailib"] = true;


                                    service.Update(objBatch);

                                    //intStatus = objBatch.ssi_reporttrackerstatus.Value;
                                    intStatus = 4;

                                    BatchReportStatus(intStatus, UserId, ssi_batchid, Convert.ToBoolean(BillingHandedOff));
                                    //Response.Write("Batch Send flag updated");
                                    #endregion
                                }

                                if (ssi_mailrecordsId != "")
                                {
                                    //objMailRecords.ssi_mailrecordsid = new Key();
                                    //objMailRecords.ssi_mailrecordsid.Value = new Guid(ssi_mailrecordsId);
                                    objMailRecords["ssi_mailrecordsid"] = new Guid(ssi_mailrecordsId);

                                    //objMailRecords.ssi_mailstatus = new Picklist();
                                    //objMailRecords.ssi_mailstatus.Value = 4;//Sent
                                    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(4);

                                    //Response.Write("Mail record Send flag updated");
                                    service.Update(objMailRecords);
                                }
                            }
                            else if (ddlAction.SelectedValue == "8")//Mark Printed
                            {

                                if (MailType.ToUpper() != "QUARTERLY STATEMENT") //Quarterly Statement
                                {
                                    if (ssi_mailrecordsId != "")
                                    {
                                        //objMailRecords.ssi_mailrecordsid = new Key();
                                        //objMailRecords.ssi_mailrecordsid.Value = ssi_mailrecordsid
                                        objMailRecords["ssi_mailrecordsid"] = new Guid(ssi_mailrecordsId);

                                        //objMailRecords.ssi_mailstatus = new Picklist();
                                        //objMailRecords.ssi_mailstatus.Value = 2;
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(2);

                                        //objMailRecords.ssi_printdate = new CrmDateTime();
                                        //objMailRecords.ssi_printdate.Value = DateTime.Now.ToString();
                                        objMailRecords["ssi_printdate"] = DateTime.Now;

                                        service.Update(objMailRecords);
                                    }
                                }
                            }
                        }
                    }
                    catch (System.Web.Services.Protocols.SoapException exc)
                    {
                        bProceed = false;
                        strDescription = "Exception occurred, Error Detail: " + exc.Detail.InnerText + ", " + exc.StackTrace;
                        lblError.Text = strDescription;
                    }
                    catch (Exception exc)
                    {
                        bProceed = false;
                        strDescription = "Exception occurred, Error Detail: " + exc.Message + ", " + exc.StackTrace;
                        lblError.Text = strDescription;
                    }
                }
                string NewFolderWarningMsg1 = "";
                if (ddlAction.SelectedValue == "1" || ddlAction.SelectedValue == "11")
                {
                    if (strExistingFiles != "")
                    {
                        // lblError.Text = lblError.Text + "<br/>Make sure to filter on the Mail Status = \"Created\" to mark them sent and see who they should be given to.<br/>Batch Report Saved to Sharepoint folder, " + strExistingFiles + "<br/> above files are overwritten.";
                        NewFolderWarningMsg = "<br/>Make sure to filter on the Mail Status = \"Created\" to mark them sent and see who they should be given to.<br/>";
                        NewFolderWarningMsg1 = "Batch Report Saved to Sharepoint folder, " + strExistingFiles + "<br/> above files are overwritten.";
                    }
                    else if (Count > 0)
                    {
                        //lblError.Text = lblError.Text + "<br/>Make sure to filter on the Mail Status = \"Created\" to mark them sent and see who they should be given to.<br/>Batch Report Saved to Sharepoint folder.";
                        NewFolderWarningMsg = "<br/>Make sure to filter on the Mail Status = \"Created\" to mark them sent and see who they should be given to.<br/>Batch Report Saved to Sharepoint folder.";
                    }

                    if (strFolderExists != "")
                    {
                        lblError.Text = lblError.Text + "<br/> " + strFolderExists;
                        //System.Web.UI.WebControls.HyperLink Link = new System.Web.UI.WebControls.HyperLink();
                        //Link.ID = "link" + i.ToString();
                        //Link.NavigateUrl = SucessFileName[i];
                        //Link.Text = SucessFileName[i];
                        //Link.Target = "_blank";// added 3_5_2019- New CS  sharepoint changes

                        //// divControlContainer.Controls.Add(lbl);
                        //divControlContainer.Controls.Add(Link);
                    }


                    string strSucess = "";
                    if (SucessClientName.Count > 0)
                    {
                        for (int i = 0; i < SucessClientName.Count; i++)
                        {
                            System.Web.UI.WebControls.Label lbl = new System.Web.UI.WebControls.Label();
                            lbl.ID = "Label" + i.ToString();
                            lbl.Text = "<br />  Batch Name: <b>" + SucessClientName[i].ToString() + "</b> <br />Sharepoint folder path:";

                            System.Web.UI.WebControls.HyperLink Link = new System.Web.UI.WebControls.HyperLink();
                            Link.ID = "link" + i.ToString();
                            Link.NavigateUrl = SucessFileName[i];
                            Link.Text = SucessFileName[i];
                            Link.Target = "_blank";// added 3_5_2019- New CS  sharepoint changes

                            divControlContainer.Controls.Add(lbl);
                            divControlContainer.Controls.Add(Link);


                            strSucess = "Batch Name: " + SucessClientName[i].ToString() + "<br/>Sharepoint folder path: " + SucessFileName[i];
                        }
                    }

                    if (FailClientName.Count > 0)
                    {
                        for (int i = 0; i < FailClientName.Count; i++)
                        {
                            int id = SucessClientName.Count + 1 + i;
                            System.Web.UI.WebControls.Label lbl = new System.Web.UI.WebControls.Label();
                            lbl.ID = "Label" + id.ToString();
                            //  lbl.Text = "<br/>Error filing Batch to Client portal, client tag not found Batch Name <b>: " + FailClientName[i].ToString() + "</b> <br/>Sharepoint folder path : "; // +FailFileName[i].ToString();
                            lbl.Text = "<br/>Error filing Batch to Client portal, Client tag <b>" + FailClientName[i].ToString() + "</b> not found, Batch Name: <b>" + ListBatchName[i].ToString() + "</b> <br/>Sharepoint folder path : "; // +FailFileName[i].ToString();
                            lbl.ForeColor = System.Drawing.Color.Red;
                            // not found Batch Name <b>:

                            System.Web.UI.WebControls.HyperLink Link = new System.Web.UI.WebControls.HyperLink();
                            Link.ID = "link" + id.ToString();
                            Link.NavigateUrl = FailFileName[i];
                            Link.Text = FailFileName[i];
                            Link.Target = "_blank";// added 3_5_2019- New CS  sharepoint changes

                            divControlContainer.Controls.Add(lbl);
                            divControlContainer.Controls.Add(Link);


                        }
                    }
                    if (NewFolderWarningMsg != "")
                    {
                        lblError2.Text = NewFolderWarningMsg;
                        NewFolderWarningMsg = "";
                        if (NewFolderWarningMsg1 != "")
                        {
                            System.Web.UI.WebControls.Label lbl = new System.Web.UI.WebControls.Label();
                            lbl.ID = "LabelErrormsg2";
                            lbl.Text = NewFolderWarningMsg1;
                            lbl.ForeColor = System.Drawing.Color.Black;
                            divControlContainer2.Controls.Add(lbl);
                        }
                        //for (int i = 0; i < ExistingFileName.Count; i++)
                        //{
                        //    System.Web.UI.WebControls.HyperLink Link = new System.Web.UI.WebControls.HyperLink();
                        //    Link.ID = "linkExist" + i.ToString();
                        //    Link.NavigateUrl = ExistingFileName[i];
                        //    Link.Text = ExistingFileName[i];
                        //    divControlContainer2.Controls.Add(Link);
                        //}


                    }




                }
                else if (ddlAction.SelectedValue == "8")
                {
                    if (Message != "")
                    {
                        lblError.Text = Message;
                    }
                    else
                    {
                        lblError.Text = "Your request has been completed.";
                    }

                }
                else if (ddlAction.SelectedValue == "6")
                {
                    if (bMarkAllRecordsSent == true)
                    {
                        ActionMarkAllSent();
                    }
                }


                #region Export CSV
                if (ddlAction.SelectedValue == "4")//Create FED-EX CSV
                {
                    MailIdNmbListTxt = "";
                    fileName = "FedEx.csv";

                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                        CheckBox CheckAll = (CheckBox)GridView1.HeaderRow.FindControl("chkBoxAll");

                        if (chkSelectNC.Checked == true)
                        {
                            if (MailIdNmbListTxt == "")
                            {
                                MailIdNmbListTxt = row.Cells[17].Text;
                            }
                            else
                            {
                                MailIdNmbListTxt = MailIdNmbListTxt + "," + row.Cells[17].Text;
                            }
                        }
                    }

                    sqlstr = "   @MailIdNmbList='" + MailIdNmbListTxt + "'";
                    ds = clsDB.getDataSet(sqlstr);

                    RKLib.ExportData.Export objExport = new RKLib.ExportData.Export("Web");

                    //string filePath = (Server.MapPath("") + @"\ExcelTemplate\CSVOutput\" + fileName);

                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        try
                        {
                            string csv = GetCSV(ds.Tables[0]);
                            ExportCSV(csv, fileName);
                        }
                        catch (System.Threading.ThreadAbortException)
                        {
                            //Thrown when calling Response.End in ExportCSV
                        }
                        catch (Exception ex)
                        {
                            //lblMessage.Text = string.Concat("An error occurred: ", ex.Message);
                        }
                    }

                    //ExporttoCSV();
                }
                else if (ddlAction.SelectedValue == "2")//Create Mailing CSV for Merge
                {

                    //Response.Write(lstMailType.SelectedValue);
                    MailIdNmbListTxt = "";
                    string MailType = string.Empty;
                    string MailTypeId = string.Empty;//added 11_9_2019 -- change Maling Names to ID
                    bool bProceed = true;


                    foreach (GridViewRow row in GridView1.Rows)
                    {
                        CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

                        if (chkSelectNC.Checked == true)
                        {
                            if (MailIdNmbListTxt == "")
                            {
                                MailType = row.Cells[2].Text;
                                MailTypeId = row.Cells[46].Text;
                                MailIdNmbListTxt = row.Cells[17].Text;
                            }
                            else
                            {
                                //if (MailType != row.Cells[2].Text)
                                //{
                                //    bProceed = false;
                                //}
                                if (MailTypeId != row.Cells[46].Text)
                                {
                                    bProceed = false;
                                }
                                MailIdNmbListTxt = MailIdNmbListTxt + "," + row.Cells[17].Text;

                            }
                        }
                    }

                    if (!bProceed)
                    {
                        lblError.Text = "Please select single Mail Type records to output merge csv file.";
                        return;
                    }

                    //if (MailType.ToUpper() == "BILLING") //Billing
                    if (MailTypeId.ToUpper() == "3FB190D9-B2CD-E011-A19B-0019B9E7EE05") //Billing//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_BILLING_EXPORT";
                        fileName = "Billing.csv";
                    }
                    //else if (MailType.ToUpper() == "CLIENT MAILING") //Client Mailing
                    else if (MailTypeId.ToUpper() == "3BD7D776-E1D3-E011-A19B-0019B9E7EE05") //Client Mailing//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_CLIENT_MAILING_EXPORT";
                        fileName = "ClientMailing.csv";
                    }
                    // else if (MailType.ToUpper() == "FUND CAPITAL CALL LETTER") //Fund Capital Call Letter
                    else if (MailTypeId.ToUpper() == "A1A079A4-D7BE-E011-A19B-0019B9E7EE05") //Fund Capital Call Letter//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_CAPITAL_CALL_LETTER_EXPORT";
                        fileName = "Fund Capital Call Letter.csv";
                    }
                    //else if (MailType.ToUpper() == "FUND CAPITAL CALL WIRE INSTRUCTIONS") //Fund Capital Call Wire Instructions
                    else if (MailTypeId.ToUpper() == "81091A9B-2AE9-E011-9141-0019B9E7EE05") //Fund Capital Call Wire Instructions//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_CAPITAL_CALL_WIRE_EXPORT";
                        fileName = "Fund Capital Call Wire Instructions.csv";
                    }
                    //else if (MailType.ToUpper() == "FUND DISTRIBUTION") //Fund Distribution
                    else if (MailTypeId.ToUpper() == "78612B2B-5ADD-E011-AD4D-0019B9E7EE05") //Fund Distribution//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_FUND_DISTRIBUTION_EXPORT";
                        fileName = "Fund Distribution.csv";
                    }
                    // else if (MailType.ToUpper() == "Fund Info (Contacts for Confirmed Recommendations)".ToUpper()) //Fund Info (Contacts for Confirmed Recommendations)
                    else if (MailTypeId.ToUpper() == "B46939F9-59DD-E011-AD4D-0019B9E7EE05".ToUpper()) //Fund Info (Contacts for Confirmed Recommendations)//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_FUND_INFO_RECOMMENDATIONS_EXPORT";
                        fileName = "Fund Info Contacts for Confirmed Recommendations.csv";
                    }
                    //else if (MailType.ToUpper() == "Fund Info (Current Holders)".ToUpper()) //Fund Info (Current Holders)
                    else if (MailTypeId.ToUpper() == "10089FA9-59DD-E011-AD4D-0019B9E7EE05".ToUpper()) //Fund Info (Current Holders)//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_FUND_INFO_CURRENT_HOUSEHOLD_EXPORT";
                        fileName = "Fund Info Current Holders.csv";
                    }
                    //else if (MailType.ToUpper() == "FUND MAILING (SIGNATURE REQUIRED)") //Fund Mailing
                    else if (MailTypeId.ToUpper() == "3CBAF86D-5EDD-E011-AD4D-0019B9E7EE05") //Fund Mailing//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_FUND_MAILING_EXPORT";
                        fileName = "Fund Mailing.csv";
                    }
                    //  else if (MailType.ToUpper() == "GENERAL MAILING") //General Mailing
                    else if (MailTypeId.ToUpper() == "99B74584-E2D3-E011-A19B-0019B9E7EE05") //General Mailing //changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_GENERAL_SMART_PROSPECT_MAILING_EXPORT";
                        fileName = "General Mailing.csv";
                    }
                    // else if (MailType.ToUpper() == "OWNER MAILING (VIA MAILING DESIGNATION)") //Owner Mailing (via Mailing Designation)
                    else if (MailTypeId.ToUpper() == "5A79AB69-E60B-E111-B3CD-0019B9E7EE05") //Owner Mailing (via Mailing Designation)//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_GENERAL_SMART_PROSPECT_MAILING_EXPORT";
                        fileName = "Owner Mailing.csv";
                    }
                    //  else if (MailType.ToUpper() == "PROSPECT MAILING") //Prospect Mailing
                    else if (MailTypeId.ToUpper() == "C71108DA-E1D3-E011-A19B-0019B9E7EE05") //Prospect Mailing //changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_GENERAL_SMART_PROSPECT_MAILING_EXPORT";
                        fileName = "Prospect Mailing.csv";
                    }
                    // else if (MailType.ToUpper() == "QUARTERLY STATEMENT") //Quarterly Statement
                    else if (MailTypeId.ToUpper() == "EB776A64-CDBE-E011-A19B-0019B9E7EE05") //Quarterly Statement//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_QUARTERLY_STATEMENT_EXPORT";
                        fileName = "Quarterly Statement.csv";
                    }
                    // else if (MailType.ToUpper() == "QUARTERLY/ANNUAL REVIEW - MANAGER SPECIFIC") //Quarterly/Annual Review - Manager Specific
                    else if (MailTypeId.ToUpper() == "0F4C85F4-D0BE-E011-A19B-0019B9E7EE05") //Quarterly/Annual Review - Manager Specific//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_QUARTERLY_ANNUAL_REVIEW_EXPORT";
                        fileName = "QuarterlyAnnual Review - Manager Specific.csv";
                    }
                    // else if (MailType.ToUpper() == "SMART MAILING") //Smart Mailing
                    else if (MailTypeId.ToUpper() == "C10BA3B7-E1D3-E011-A19B-0019B9E7EE05") //Smart Mailing//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_GENERAL_SMART_PROSPECT_MAILING_EXPORT";
                        fileName = "Smart Mailing.csv";
                    }
                    //else if (MailType.ToUpper() == "CLIENT PORTAL") //Smart Mailing
                    else if (MailTypeId.ToUpper() == "2357D455-F762-E111-BD8F-0019B9E7EE05") //Smart Mailing//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_CLIENT_PORTAL_EXPORT";
                        fileName = "Client Portal.csv";
                    }
                    //else if (MailType.ToUpper() == "FUND DISTRIBUTION LETTER") //Smart Mailing
                    else if (MailTypeId.ToUpper() == "6D7545DA-8164-E111-BD8F-0019B9E7EE05") //Smart Mailing //changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_FUND_DISTRIBUTION_LETTER_EXPORT";
                        fileName = "Fund Distribution Letter.csv";
                    }
                    // else if (MailType.ToUpper() == "CLIENT EVENT MAILING") //Client Mailing
                    else if (MailTypeId.ToUpper() == "005403C5-D5D1-E111-A4D8-0019B9E7EE05") //Client Mailing //changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_Client_Event_Mailing_EXPORT";
                        fileName = "ClientEventMailing.csv";
                    }
                    // else if (MailType.ToUpper() == "ON DEMAND MAILING") //Client Mailing
                    else if (MailTypeId.ToUpper() == "01AB40EB-DDD1-E111-A4D8-0019B9E7EE05") //Client Mailing //changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_On_Demand_Mailing_EXPORT";
                        fileName = "OnDemandMailing.csv";
                    }
                    // else if (MailType.ToUpper() == "ONE TIME MAILING") //One Time Mailing -sasmit(6_30_2017)
                    else if (MailTypeId.ToUpper() == "18DFB14E-B156-E711-9422-005056A0567E") //One Time Mailing -sasmit(6_30_2017)//changed 11_9_2019 -- change Mailing Names to ID
                    {
                        sqlstr = "SP_S_On_Demand_Mailing_EXPORT @OneTimeMailFlg = 1";
                        fileName = "OnTimeMailing.csv";
                    }

                    //ssi_name	ssi_mailid
                    //Billing	3FB190D9-B2CD-E011-A19B-0019B9E7EE05
                    //Client Mailing	3BD7D776-E1D3-E011-A19B-0019B9E7EE05
                    //Fund Capital Call Letter	A1A079A4-D7BE-E011-A19B-0019B9E7EE05
                    //Fund Capital Call Wire Instructions	81091A9B-2AE9-E011-9141-0019B9E7EE05
                    //Fund Distribution	78612B2B-5ADD-E011-AD4D-0019B9E7EE05
                    //Fund Info (Contacts for Confirmed Recommendations)	B46939F9-59DD-E011-AD4D-0019B9E7EE05
                    //Fund Info (Current Holders)	10089FA9-59DD-E011-AD4D-0019B9E7EE05
                    //Fund Mailing (Signature Required)	3CBAF86D-5EDD-E011-AD4D-0019B9E7EE05
                    //General Mailing	99B74584-E2D3-E011-A19B-0019B9E7EE05
                    //Owner Mailing (via Mailing Designation)	5A79AB69-E60B-E111-B3CD-0019B9E7EE05
                    //Prospect Mailing	C71108DA-E1D3-E011-A19B-0019B9E7EE05
                    //Quarterly Statement	EB776A64-CDBE-E011-A19B-0019B9E7EE05
                    //Quarterly/Annual Review - Manager Specific	0F4C85F4-D0BE-E011-A19B-0019B9E7EE05
                    //Smart Mailing	C10BA3B7-E1D3-E011-A19B-0019B9E7EE05

                    //if (MailType.ToUpper() == "Fund Info (Contacts for Confirmed Recommendations)".ToUpper() || MailType.ToUpper() == "Fund Info (Current Holders)".ToUpper())
                    //{
                    //    //sqlstr;//Not to Attach MailIdNmb for these two mail Types
                    //}
                    //else 

                    //  if (MailType.ToUpper() == "SMART MAILING".ToUpper() || MailType.ToUpper() == "PROSPECT MAILING".ToUpper() || MailType.ToUpper() == "OWNER MAILING (VIA MAILING DESIGNATION)".ToUpper() || MailType.ToUpper() == "GENERAL MAILING".ToUpper())
                    if (MailTypeId.ToUpper() == "C10BA3B7-E1D3-E011-A19B-0019B9E7EE05".ToUpper() || MailTypeId.ToUpper() == "C71108DA-E1D3-E011-A19B-0019B9E7EE05".ToUpper() || MailTypeId.ToUpper() == "5A79AB69-E60B-E111-B3CD-0019B9E7EE05".ToUpper() || MailTypeId.ToUpper() == "99B74584-E2D3-E011-A19B-0019B9E7EE05".ToUpper())
                    {
                        if (hdCheckAll.Value == "1")
                        {
                            sqlstr = sqlstr + " @MailType='" + MailType + "',@MailIdNmbList='" + MailIdNmbListTxt + "'";
                            ds = clsDB.getDataSet(sqlstr);
                        }
                        else if (hdCheckAll.Value == "0" && sqlstr != "")
                        {
                            sqlstr = sqlstr + " @MailIdNmbList='" + MailIdNmbListTxt + "'";
                            ds = clsDB.getDataSet(sqlstr);
                        }

                    }
                    //else if (sqlstr != "")-added by Sasmit(6_30_2017)
                    //{
                    //    sqlstr = sqlstr + " @MailIdNmbList='" + MailIdNmbListTxt + "'";
                    //    ds = clsDB.getDataSet(sqlstr);
                    //}
                    else if (sqlstr != "")
                    {
                        // if (MailType.ToUpper() == "ONE TIME MAILING")
                        if (MailTypeId.ToUpper() == "18DFB14E-B156-E711-9422-005056A0567E")
                        {
                            sqlstr = sqlstr + " ,@MailIdNmbList='" + MailIdNmbListTxt + "'";
                            ds = clsDB.getDataSet(sqlstr);
                        }
                        else
                        {
                            sqlstr = sqlstr + " @MailIdNmbList='" + MailIdNmbListTxt + "'";
                            ds = clsDB.getDataSet(sqlstr);
                        }
                    }

                    hdCheckAll.Value = "0";
                    RKLib.ExportData.Export objExport = new RKLib.ExportData.Export("Web");

                    //string filePath = (Server.MapPath("") + @"\ExcelTemplate\CSVOutput\" + fileName);
                    if (ds != null)
                    {

                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            try
                            {
                                string csv = GetCSV(ds.Tables[0]);
                                ExportCSV(csv, fileName);
                            }
                            catch (System.Threading.ThreadAbortException)
                            {
                                //Thrown when calling Response.End in ExportCSV
                            }
                            catch (Exception ex)
                            {
                                //lblMessage.Text = string.Concat("An error occurred: ", ex.Message);
                            }
                            //objExport.ExportDetails(ds.Tables[0], RKLib.ExportData.Export.ExportFormat.CSV, fileName);
                        }
                    }

                    //ExportGridToCSV();
                }

                #endregion

                BindGridView();
            }
            else
            {
                // lblError.Text = "Error Fetching Sharepoint List, Contact Administrator";
                lblError.Text = "Unable to connect to SharePoint , Please try again after sometime";
            }
        }
        catch (Exception ex)
        {
            lblError.Text = "Error Occured : " + ex.Message.ToString();
        }
        finally
        {
            //delete tempfolder creted at local Directory
            if (Directory.Exists(TempFolderPath))
            {
                Directory.Delete(TempFolderPath, true);
            }
        }


        if (sw != null)
        {
            sw.Flush();
            sw.Close();
        }
    }


    // private string ConvertTextUrlToLink(string url)
    // {
    ////     string regex = @"((www\.|(http|https|ftp|news|file)+\:\/\/)[_.a-z0-9-]+\.
    ////[a-z0-9\/_:@=.+?,##%&~-]*[^.|\'|\# |!|\(|?|,| |>|<|;|\)])";
    ////     Regex r = new Regex(regex, RegexOptions.IgnoreCase);
    ////     return r.Replace(url, "a href=\"$1\" title=\"Click here to open in a new window or tab\" 

    ////    // target =\"_blank\">$1</a>").Replace("href=\"www", "href=\"http://www");
    // }


    public void SendEmail(string LEFolderName, string CLientServiceLink)
    {
        try
        {
            string mailmessage = string.Empty;
            MailMessage myMessage = new MailMessage();
            // SmtpClient SMTPSERVER = new SmtpClient();

            string EmailID = AppLogic.GetParam(AppLogic.ConfigParam.EmailId);
            string Password = AppLogic.GetParam(AppLogic.ConfigParam.Password);
            string SMTPHost = AppLogic.GetParam(AppLogic.ConfigParam.SMTPHost);
            string ToEmailIDs2 = AppLogic.GetParam(AppLogic.ConfigParam.ToEmailIDs2);

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

            myMessage.Bcc.Add("auto-emails@infograte.com");


            // string str = "" + LEFolderName + "";

            string Server = AppLogic.GetParam(AppLogic.ConfigParam.Server);


            if (Server.ToLower() == "test")
                myMessage.Subject = '"' + LEFolderName + '"' + " " + " folder missing - Test CRM";
            else if (Server.ToLower() == "prod")
                myMessage.Subject = '"' + LEFolderName + '"' + " " + " folder missing - Prod CRM";

            myMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
            myMessage.Body = LEFolderName + " folder is missing from the SharePoint Client Service site." + "<br/>The file has been saved to the Published document at the below given link" + "<br/>" + CLientServiceLink;


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
        }
        catch (Exception ex)
        {
            string strDescription = "Error sending Mail :" + ex.Message.ToString();
            //commented on 12_4_2018 Jscalise nolonger in process
            lblError.Text = lblError.Text + "  ," + "Send Error" + ex.Message.ToString();
            // lblEmail.Text = "Send Error" + ex.Message.ToString();

            //LogMessage(sw, strDescription);
        }

    }

    private void GroupBy(string TempFolderPath)
    {
        string FirstName = string.Empty;
        string ComparingFitstName = string.Empty;
        int groupCount = 0;

        for (int i = 0; i < GridView1.Rows.Count; i++)
        {
            CheckBox chkSelectNC = (CheckBox)GridView1.Rows[i].Cells[0].FindControl("chkSelectNC");

            if (chkSelectNC.Checked == true)
            {
                FirstName = GridView1.Rows[i].Cells[35].Text;
                //if (i == 0)
                //    ComparingFitstName = FirstName;

                for (int j = 0; j < GridView1.Rows.Count; j++)
                {

                    CheckBox chkSelectrr = (CheckBox)GridView1.Rows[j].Cells[0].FindControl("chkSelectNC");
                    if (chkSelectrr.Checked == true)
                    {
                        if (i != j)
                        {
                            ComparingFitstName = GridView1.Rows[j].Cells[35].Text;
                            if (FirstName == ComparingFitstName)
                            {
                                groupCount++;
                                i++;
                            }
                        }
                    }
                }
                bool ret = GenerateMergeTypeConsolidatedPDF(FirstName, TempFolderPath);
            }

        }
    }

    public string GetCSV(DataTable dt)
    {
        StringBuilder sb = new StringBuilder();

        //Line for column names
        for (int i = 0; i < dt.Columns.Count; i++)
        {
            sb.Append(dt.Columns[i]);

            if (i < dt.Columns.Count - 1)
            {
                sb.Append(",");
            }
        }

        sb.AppendLine();

        //Loop through table and create a line for each row
        foreach (DataRow dr in dt.Rows)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (!Convert.IsDBNull(dr[i]))
                {
                    string value = dr[i].ToString();

                    //Check if the value contans a comma and place it in quotes if so
                    if (value.Contains(","))
                    {
                        value = string.Concat("\"", value, "\"");
                    }

                    //Replace any \r or \n special characters from a new line with a space
                    if (value.Contains("\r"))
                    {
                        value = value.Replace("\r", " ");
                    }
                    if (value.Contains("\n"))
                    {
                        value = value.Replace("\n", " ");
                    }

                    sb.Append(value);
                }

                if (i < dt.Columns.Count - 1)
                {
                    sb.Append(",");
                }
            }

            sb.AppendLine();
        }

        return sb.ToString();
    }

    private void ExportCSV(string csv, string filename)
    {
        Response.Clear();
        Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", filename));
        Response.Charset = "";
        Response.ContentType = "text/csv";
        Response.ContentEncoding = System.Text.Encoding.Default;// GetEncoding("UTF-8");
        //Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());
        Response.AddHeader("Pragma", "public");
        Response.Write(csv);
        Response.End();
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
        sqlstr = "";
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

    public void BatchReportStatus(int Status, string UpdatedBy, string BatchId, bool BillingHandedOff)
    {
        string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        string orgName = "GreshamPartners";
        //string orgName = "Webdev";
        IOrganizationService service = null;


        //lblError.Text = "";
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
            lblError.Text = strDescription;
        }
        catch (Exception exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
            lblError.Text = strDescription;
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

            if (Status != 0)
            {
                //objReportStatusLog.ssi_status = new Picklist();
                //objReportStatusLog.ssi_status.Value = Status;
                objReportStatusLog["ssi_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(Status);
            }

            //if (BillingHandedOff.ToUpper() == "TRUE")
            //{

            //objReportStatusLog.ssi_billinghandedoff = new CrmBoolean();
            //objReportStatusLog.ssi_billinghandedoff.Value = BillingHandedOff;
            objReportStatusLog["ssi_billinghandedoff"] = BillingHandedOff;

            //}
            //else
            //{
            //    objReportStatusLog.ssi_billinghandedoff = new CrmBoolean();
            //    objReportStatusLog.ssi_billinghandedoff.Value = false;
            //}

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

            service.Create(objReportStatusLog);

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
    }

    public bool GenerateConsolidatedPDF(string TempFolderPath)
    {
        try
        {
            string[] SourceFileName = new string[0];
            if (!chkGroupRecandSpouse.Checked)
            {
                int NoOfBatches = 0;
                int checkBoxChecked = 0;
                //this loop will get the number of files to merge according to conditions.
                for (int j = 0; j < GridView1.Rows.Count; j++)
                {
                    CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");
                    string ReportPath = (string)GridView1.Rows[j].Cells[18].Text;

                    if (chkBox.Checked && ReportPath.Replace("&nbsp;", "") != "")
                    {
                        NoOfBatches++;//
                        checkBoxChecked = NoOfBatches;
                        if (!chkMailingSheets.Checked)
                        {
                            string IncludeSalutation = (string)GridView1.Rows[j].Cells[8].Text;
                            if (IncludeSalutation.Trim().ToLower() == "regular mail" || IncludeSalutation.Trim().ToLower() == "regular mail and email" || IncludeSalutation.Trim().ToLower() == "regular mail and client portal")
                            {
                                NoOfBatches = NoOfBatches + 1;
                            }
                        }
                        if (!chkReportSeperator.Checked)
                        {
                            NoOfBatches = NoOfBatches + 1;
                        }
                    }
                }


                string FileName = string.Empty;

                int NoofFiles = 0;

                NoofFiles = NoOfBatches;

                SourceFileName = new string[NoofFiles];
                if (SourceFileName.Length < 1)
                {
                    lblError.Text = "No Report found for selected records";
                    return false;
                }
                int FileNo = 0;
                //this loop will get the paths of files to merge according to conditions.
                for (int j = 0; j < GridView1.Rows.Count; j++)
                {
                    CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");
                    string ReportPath1 = (string)GridView1.Rows[j].Cells[18].Text;
                    if (chkBox.Checked && ReportPath1.Replace("&nbsp;", "") != "")
                    {
                        if (!chkMailingSheets.Checked)
                        {
                            string IncludeSalutation = (string)GridView1.Rows[j].Cells[8].Text;
                            if (IncludeSalutation.Trim().ToLower() == "regular mail" || IncludeSalutation.Trim().ToLower() == "regular mail and email" || IncludeSalutation.Trim().ToLower() == "regular mail and client portal")
                            {
                                string Name = (string)GridView1.Rows[j].Cells[6].Text;
                                string MailingAddress = (string)GridView1.Rows[j].Cells[7].Text;
                                string Salutation = (string)GridView1.Rows[j].Cells[34].Text;

                                SourceFileName[FileNo] = GenerateSalutaionPage(Name, MailingAddress, Salutation, TempFolderPath);
                                FileNo++;
                            }
                        }
                        string ReportName = (string)GridView1.Rows[j].Cells[16].Text;
                        string ReportPath = (string)GridView1.Rows[j].Cells[18].Text;

                        SourceFileName[FileNo] = ReportPath.Replace("&nbsp;", "");
                        FileNo++;
                        if (!chkReportSeperator.Checked)
                        {
                            SourceFileName[FileNo] = Server.MapPath("") + "/ExcelTemplate/Template/EndReport.pdf";
                            FileNo++;
                        }
                    }
                }
            }
            else
            {
                // to group the recipient and spouse data or arrange the recipient and spouse
                string[] src = GroupRecipientandSpouse(TempFolderPath);
                int arrCount = src.Length;
                SourceFileName = new string[arrCount];
                src.CopyTo(SourceFileName, 0);
                if (SourceFileName.Length < 1)
                {
                    lblError.Text = "No Report found for selected records";
                    return false;
                }
            }
            string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
            string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
            string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
            string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();

            string ConsolidatedPDFFileName = "ConsolidatedPDF_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond;

            string DestinationFileName = string.Empty;

            string CombinedPdfs = AppLogic.GetParam(AppLogic.ConfigParam.CombinedPdfs);

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                DestinationFileName = Request.MapPath("\\Advent Report\\ExcelTemplate\\BATCH REPORTS\\" + ConsolidatedPDFFileName + ".pdf"); //Server.MapPath("") + "\\ExcelTemplate\\" + ConsolidatedPDFFileName + ".pdf";
            }
            else
            {
                //   DestinationFileName = "\\\\GRPAO1-VWFS01\\opsreports$\\Combined PDFs\\" + ConsolidatedPDFFileName + ".pdf";
                DestinationFileName = CombinedPdfs + ConsolidatedPDFFileName + ".pdf"; // shared drive changes - 7_4_2019
            }
            //DestinationFileName = "\\\\GRPAO1-VWFS01\\shared$\\OPS REPORTS\\" + ConsolidatedPDFFileName + ".pdf";

            //DestinationFileName = "\\\\GRPAO1-VWFS01\\opsreports$\\" + ConsolidatedPDFFileName + ".pdf";


            // SourceFileName = GetDistinctValues<string>(SourceFileName);

            //string DestinationFileName = "D:\\Gresham\\TestMerge.pdf";
            PDFMerge PDF = new PDFMerge();
            PDF.MergeFiles(DestinationFileName, SourceFileName);
            ViewState["ConsolidatedSinglePDF"] = DestinationFileName;
            return true;
        }
        catch (Exception ex)
        {
            lblError.Text = ex.ToString();
            return false;
        }

    }

    private string[] GroupRecipientandSpouse(string TempFolderPath)
    {
        string strID = string.Empty;
        for (int j = 0; j < GridView1.Rows.Count; j++)
        {
            CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");
            if (chkBox.Checked)
            {
                string BatchId = (string)GridView1.Rows[j].Cells[13].Text;
                if (strID == "")
                {
                    strID = BatchId;
                }
                else
                {
                    strID = strID + "," + BatchId;
                }
            }
        }
        string sqlstr = "SP_S_REPORT_MAIL_QUEUE @ssi_mailrecordsId='" + strID + "'";
        DataSet DS = clsDB.getDataSet(sqlstr);
        DataTable dt = DS.Tables[0];
        int NoOfBatches = 0;
        int checkBoxChecked = 0;
        //this loop will get the number of files to merge according to conditions.
        for (int j = 0; j < dt.Rows.Count; j++)
        {
            string ReportPath = Convert.ToString(dt.Rows[j]["ssi_batchfilename"]);
            if (ReportPath.Replace("&nbsp;", "") != "")
            {
                NoOfBatches++;
                checkBoxChecked = NoOfBatches;
                if (!chkMailingSheets.Checked)
                {
                    string IncludeSalutation = Convert.ToString(dt.Rows[j]["Mail Preference"]);
                    if (IncludeSalutation.Trim().ToLower() == "regular mail")
                    {
                        NoOfBatches = NoOfBatches + 1;
                    }
                }
                if (!chkReportSeperator.Checked)
                {
                    NoOfBatches = NoOfBatches + 1;
                }
            }
        }

        string[] SourceFileName = new string[0];
        string FileName = string.Empty;

        int NoofFiles = 0;

        NoofFiles = NoOfBatches;

        SourceFileName = new string[NoofFiles];
        if (SourceFileName.Length < 1)
        {
            lblError.Text = "No Report found for selected records";
            return SourceFileName;
        }
        int FileNo = 0;
        //this loop will get the paths of files to merge according to conditions.
        for (int j = 0; j < dt.Rows.Count; j++)
        {
            string ReportPath1 = Convert.ToString(dt.Rows[j]["ssi_batchfilename"]);
            if (ReportPath1.Replace("&nbsp;", "") != "")
            {
                if (!chkMailingSheets.Checked)
                {
                    string IncludeSalutation = Convert.ToString(dt.Rows[j]["Mail Preference"]);
                    if (IncludeSalutation.Trim().ToLower() == "regular mail")
                    {
                        string Name = Convert.ToString(dt.Rows[j]["Receipent"]);
                        string MailingAddress = Convert.ToString(dt.Rows[j]["Mailing Address/ Email"]);
                        string Salutation = Convert.ToString(dt.Rows[j]["ssi_salutation_mail"]);

                        SourceFileName[FileNo] = GenerateSalutaionPage(Name, MailingAddress, Salutation, TempFolderPath);
                        FileNo++;
                    }
                }

                string ReportPath = Convert.ToString(dt.Rows[j]["ssi_batchfilename"]);
                SourceFileName[FileNo] = ReportPath.Replace("&nbsp;", "");
                FileNo++;
                if (!chkReportSeperator.Checked)
                {
                    SourceFileName[FileNo] = Server.MapPath("") + "/ExcelTemplate/Template/EndReport.pdf";
                    FileNo++;
                }
            }
        }
        return SourceFileName;
    }

    private string GenerateSalutaionPage(string Name, string MailingAddress, string Salutation, string TempFolderPath)
    {
        char lf = (char)10;
        MailingAddress = MailingAddress.Replace(lf.ToString(), "*");
        string[] Data = MailingAddress.Split('*');
        string Salutaion = Data.Length > 0 ? Data[0] : "";
        string AddressLine1 = Data.Length > 1 ? Data[1] : "";
        string AddressLine2 = Data.Length > 2 ? Data[2] : "";
        string AddressLine3 = Data.Length > 3 ? Data[3] : "";
        string City = Data.Length > 4 ? Data[4] : "";
        string StateProvinence = Data.Length > 5 ? Data[5] : "";
        string ZipCode = Data.Length > 6 ? Data[6] : "";
        string CountryRegion = Data.Length > 7 ? Data[7] : "";

        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = String.Empty;
        string strRndNumber = clsGM.CreateRandomNumber(4);
        Guid id = new Guid();
        id = Guid.NewGuid();
        string strGUID = id.ToString();
        //string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + strRndNumber; 

        //String fsFinalLocation = String.Empty;
        // fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\Slatn" + strGUID + ".xls";

        //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 137, 30, 144, 8);//10,10
        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 63, 30, 122, 8);//10,10


        // String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        String fsFinalLocation = TempFolderPath + "//" + Guid.NewGuid().ToString() + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));
        document.Open();

        iTextSharp.text.Table loTable = new iTextSharp.text.Table(2);
        loTable.Width = 100;
        int[] headerwidths = { 39, 45 }; //{ 47, 35 }
        loTable.SetWidths(headerwidths);
        loTable.Border = 0;

        iTextSharp.text.Cell loCell = new Cell();
        Chunk loChunk = new Chunk();


        string FormatName = Name + "\n" + AddressLine1 + "\n" + AddressLine2 + "\n" + AddressLine3 + "\n" + City + "," + StateProvinence + " " + ZipCode;

        if (Salutation.Length > 31)
            Name = Name.Replace("and", "\n" + "and");

        loChunk = new Chunk("dev", Font8Whitecheck("test"));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.HorizontalAlignment = 0;
        loCell.Border = 0;
        loTable.AddCell(loCell);

        loCell = new Cell();
        loChunk = new Chunk(Name, setFontsAll(11, 0, 0));//setFontsAll(18, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(AddressLine1, setFontsAll(11, 0, 0));//setFontsAll(18, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(AddressLine2, setFontsAll(11, 0, 0));//setFontsAll(26, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(AddressLine3, setFontsAll(11, 0, 0));//setFontsAll(26, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(City + "," + StateProvinence + " " + ZipCode, setFontsAll(11, 0, 0));//setFontsAll(26, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);


        //loCell = new Cell();
        //loChunk = new Chunk(StateProvinence, setFontsAll(18, 0, 0));//setFontsAll(26, 0, 0));
        //loCell.Add(loChunk);
        //loCell.Colspan = 2;
        //loCell.Border = 0;
        //loCell.HorizontalAlignment = 0;
        //loCell.Leading = 25f;
        //loTable.AddCell(loCell);


        //loCell = new Cell();
        //loChunk = new Chunk(ZipCode, setFontsAll(18, 0, 0));//setFontsAll(26, 0, 0));
        //loCell.Add(loChunk);
        //loCell.Colspan = 2;
        //loCell.Border = 0;
        //loCell.HorizontalAlignment = 0;
        //loCell.Leading = 25f;
        //loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(CountryRegion, setFontsAll(11, 0, 0));//setFontsAll(26, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        //document.Add(png);

        document.Add(loTable);

        document.Close();

        //FileInfo loFile = new FileInfo(ls);
        //loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        ////Response.Write("<br/>"+fsFinalLocation.Replace(".xls", ".pdf"));
        return fsFinalLocation.Replace(".xls", ".pdf");

    }


    private string GenerateSalutaionPageForMailingSheet(string Name, string MailingAddress, string Salutation, string TempFolderPath)
    {
        char lf = (char)10;
        MailingAddress = MailingAddress.Replace(lf.ToString(), "*");
        string[] Data = MailingAddress.Split('*');
        string Salutaion = Data.Length > 0 ? Data[0] : "";
        string AddressLine1 = Data.Length > 1 ? Data[1] : "";
        string AddressLine2 = Data.Length > 2 ? Data[2] : "";
        string AddressLine3 = Data.Length > 3 ? Data[3] : "";
        string City = Data.Length > 4 ? Data[4] : "";
        string StateProvinence = Data.Length > 5 ? Data[5] : "";
        string ZipCode = Data.Length > 6 ? Data[6] : "";
        string CountryRegion = Data.Length > 7 ? Data[7] : "";

        liPageSize = 29;
        DataSet newdataset;
        DB clsDB = new DB();
        newdataset = null;
        String lsFooterTxt = String.Empty;
        string strRndNumber = clsGM.CreateRandomNumber(4);
        Guid id = new Guid();
        id = Guid.NewGuid();
        string strGUID = id.ToString();
        //string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss") + strRndNumber; 

        String fsFinalLocation = String.Empty;
        //   fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\Slatn" + strGUID + ".xls";

        //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 137, 30, 144, 8);//10,10
        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 63, 30, 122, 8);//10,10


        //  String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
        // String ls = TempFolderPath + "//" + Guid.NewGuid().ToString() + ".pdf";
        fsFinalLocation = TempFolderPath + "//" + Guid.NewGuid().ToString() + ".pdf";
        // iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(fsFinalLocation, FileMode.Create));
        document.Open();

        iTextSharp.text.Table loTable = new iTextSharp.text.Table(2);
        loTable.Width = 100;
        int[] headerwidths = { 39, 45 }; //{ 47, 35 }
        loTable.SetWidths(headerwidths);
        loTable.Border = 0;

        iTextSharp.text.Cell loCell = new Cell();
        Chunk loChunk = new Chunk();


        string FormatName = Name + "\n" + AddressLine1 + "\n" + AddressLine2 + "\n" + AddressLine3 + "\n" + City + ", " + StateProvinence + " " + ZipCode;

        if (Salutation.Length > 31)
            Name = Name.Replace("and", "\n" + "and");

        loChunk = new Chunk("dev", Font8Whitecheck("test"));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.HorizontalAlignment = 0;
        loCell.Border = 0;
        loTable.AddCell(loCell);

        loCell = new Cell();
        loChunk = new Chunk(Name, setFontsAll(11, 0, 0));//setFontsAll(18, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(AddressLine1, setFontsAll(11, 0, 0));//setFontsAll(18, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(AddressLine2, setFontsAll(11, 0, 0));//setFontsAll(26, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(AddressLine3, setFontsAll(11, 0, 0));//setFontsAll(26, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(City + ", " + StateProvinence + " " + ZipCode, setFontsAll(11, 0, 0));//setFontsAll(26, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);


        //loCell = new Cell();
        //loChunk = new Chunk(StateProvinence, setFontsAll(18, 0, 0));//setFontsAll(26, 0, 0));
        //loCell.Add(loChunk);
        //loCell.Colspan = 2;
        //loCell.Border = 0;
        //loCell.HorizontalAlignment = 0;
        //loCell.Leading = 25f;
        //loTable.AddCell(loCell);


        //loCell = new Cell();
        //loChunk = new Chunk(ZipCode, setFontsAll(18, 0, 0));//setFontsAll(26, 0, 0));
        //loCell.Add(loChunk);
        //loCell.Colspan = 2;
        //loCell.Border = 0;
        //loCell.HorizontalAlignment = 0;
        //loCell.Leading = 25f;
        //loTable.AddCell(loCell);


        loCell = new Cell();
        loChunk = new Chunk(CountryRegion, setFontsAll(11, 0, 0));//setFontsAll(26, 0, 0));
        loCell.Add(loChunk);
        loCell.Colspan = 2;
        loCell.Border = 0;
        loCell.HorizontalAlignment = 0;
        loCell.Leading = 15f;
        loTable.AddCell(loCell);

        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(Server.MapPath("") + @"\images\Gresham_Logo.png");
        png.SetAbsolutePosition(45, 557);//540
        //png.ScaleToFit(288f, 42f);
        png.ScalePercent(10);
        //document.Add(png);

        document.Add(loTable);

        document.Close();

        //FileInfo loFile = new FileInfo(ls);
        //loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        ////Response.Write("<br/>"+fsFinalLocation.Replace(".xls", ".pdf"));
        return fsFinalLocation.Replace(".xls", ".pdf");

    }

    private void ExportCSV(DataTable dt, string filename)
    {

    }

    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (ViewState["checkedRowsList"] != null)
        {
            ArrayList checkedRowsList = (ArrayList)ViewState["checkedRowsList"];
            GridViewRow gvRow = e.Row;
            if (gvRow.RowType == DataControlRowType.DataRow)
            {
                CheckBox chkSelect = (CheckBox)gvRow.FindControl("chkSelectNC");
                string rowIndex = Convert.ToString(GridView1.DataKeys[gvRow.RowIndex]["ssi_mailrecordsId"]);
                //int rowIndex = Convert.ToInt32(gvRow.RowIndex) + 

                Convert.ToInt32(GridView1.PageIndex);
                if (checkedRowsList.Contains(rowIndex))
                {
                    chkSelect.Checked = true;
                }
            }
        }
    }

    protected void ddlAssociate_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
    }
    protected void ddlAdvisor_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
    }
    protected void lstMailId_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
    }
    protected void lstMailPreference_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
    }
    protected void lstMailStatus_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
    }
    protected void ddlHousehold_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
    }
    protected void ddlAsofDate_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
        //ClearControls();
    }

    private void ClearControls()
    {
        lstMailId.ClearSelection();
        lstMailType.ClearSelection();
        ddlHousehold.ClearSelection();
        ddlAssociate.ClearSelection();
        ddlAdvisor.ClearSelection();
        lstMailPreference.ClearSelection();
        lstMailStatus.ClearSelection();
        lstMailPreference.ClearSelection();
        ddlSalutationPref.ClearSelection();
        ddlCreatedBy.ClearSelection();
        txtCreatedOn.Text = "";
        lblError.Text = "";
    }

    private void GenerateReport()
    {
        try
        {
            lblError.Text = "";
            string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
            //string crmServerURL = "http://server:5555/";

            string orgName = "GreshamPartners";
            string currentuser = null;
            //string orgName = "Webdev";
            IOrganizationService service = null;
            Boolean checkrunreport = false;
            String DestinationPath = string.Empty;
            string ConsolidatePdfFileName = string.Empty;
            string ReportOpFolder = string.Empty;
            //  string ApprovedReports = "\\\\GRPAO1-VWFS01\\opsreports$\\Approved Reports\\";
            string ApprovedReports = AppLogic.GetParam(AppLogic.ConfigParam.ApprovedReports);// shared drive changes - 7_4_2019


            try
            {
                service = GM.GetCrmService();
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

            //string CurrentDateTime = DateTime.Now.ToShortDateString() + " " + " " + strHour + "-" + strMinute + "-" + strSecond;

            string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

            //  string strUserName = HttpContext.Current.User.Identity.Name.ToString();
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

            ViewState["ParentFolder"] = strUserName + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond;

            //string ReportOpFolder = "\\\\GRPAO1-VWFS01\\_ops_C_I_R_group\\Quarterly_Reports\\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();


            //ReportOpFolder = Request.MapPath("ExcelTemplate\\BATCH REPORTS\\") + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            //if (Request.Url.AbsoluteUri.Contains("localhost"))
            //{
            //    ReportOpFolder = Request.MapPath("ExcelTemplate\\BATCH REPORTS\\") + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            //}
            //else
            //    ReportOpFolder = Request.MapPath("ExcelTemplate\\BATCH REPORTS\\") + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            // ReportOpFolder = "\\\\GRPAO1-VWFS01\\opsreports$";//"\\\\GRPAO1-VWFS01\\shared$\\OPS REPORTS\\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports);// shared drive changes - 7_4_2019

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                ReportOpFolder = @"C:\Reports\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            }
            else
            {
                //   ReportOpFolder = "\\\\GRPAO1-VWFS01\\opsreports$";//"\\\\GRPAO1-VWFS01\\shared$\\OPS REPORTS\\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
                ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports);
            }
            string ContactFolderName = string.Empty;
            FileInfo loCoversheetCheck;
            String ReportOpFolder1 = String.Empty;

            /*****Start :  Array declaration for PDF merge **************/
            PDFMerge PDF = new PDFMerge();
            int sourcefilecount = 0;//= dtBatch.Rows.Count + 1;
            string[] SourceFileArray;
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

                string CurrentTimeStamp = FileYear + "_" + FileMonth + "_" + FileDay + "_" + FileHour + "_" + FileMinute;

                if (chkBox.Checked)
                {
                    checkrunreport = true;
                    String BatchIdListTxt = Convert.ToString(GridView1.Rows[j].Cells[12].Text);
                    dtBatch = GetDataTable(BatchIdListTxt);

                    //String TempName =  GridView1.Rows[j].Cells[6].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();

                    //String HHName = GridView1.Rows[j].Cells[6].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                    //string ssi_batchid = GridView1.Rows[j].Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                    String HHName = "";
                    string OldHHName = GridView1.Rows[j].Cells[16].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                    OldHHName = OldHHName.Replace("/", "");
                    //string TempName = HttpContext.Current.User.Identity.Name.ToString() + "_" + 

                    sourcefilecount = dtBatch.Rows.Count + 1;
                    SourceFileArray = new string[sourcefilecount];

                    for (int i = 0; i < dtBatch.Rows.Count; i++)
                    {
                        if (Convert.ToString(dtBatch.Rows[i]["ssi_spvfilename"]) != "")
                        {
                            HHName = Convert.ToString(dtBatch.Rows[i]["ssi_spvfilename"]);
                            HHName = HHName.Replace("/", "");
                        }
                        else
                        {
                            HHName = GridView1.Rows[j].Cells[16].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                            HHName = HHName.Replace("/", "");
                        }

                        ContactFolderName = GridView1.Rows[j].Cells[14].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                        //ContactFolderName = Convert.ToString(dtBatch.Rows[i]["Ssi_ContactIdName"]).Replace(",", "");
                        bool isExist = System.IO.Directory.Exists(ReportOpFolder + "\\" + ContactFolderName);

                        if (!isExist)
                        {
                            //Response.Write("Folder: " + ReportOpFolder + "\\" + ContactFolderName);
                            System.IO.Directory.CreateDirectory(ReportOpFolder + "\\" + ContactFolderName);
                        }

                        ViewState["AsOfDate"] = Convert.ToString(dtBatch.Rows[i]["Ssi_EndAsOfDate2"]);
                        // ViewState["PdfFileName"] = HHName = Convert.ToString(dtBatch.Rows[i]["PdfFileName"]);

                        String fsAllocationGroup = Convert.ToString(dtBatch.Rows[i]["Ssi_AllocationGroup"]);
                        String fsHouseholdName = Convert.ToString(dtBatch.Rows[i]["Ssi_HouseholdIdName"]);
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
                        String lsFinalTitleAfterChange = String.Empty;


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

                        string strGUID = Guid.NewGuid().ToString();
                        strGUID = strGUID.Substring(0, 5);
                        //String lsExcleSavePath = ReportOpFolder + "\\" + ContactFolderName + "\\" + fsHouseholdName.Replace(",", "") + "_" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + strGUID + ".xls";
                        String lsExcleSavePath = ReportOpFolder + "\\" + ContactFolderName + "\\" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + lsFinalTitleAfterChange.Replace(",", "").Replace("/", "").Replace("\\", "") + "_" + Convert.ToDateTime(fsAsofDate).ToString("yyyyMMdd") + "_" + strGUID + ".xls";
                        //String lsSavePathCombReport  = ReportOpFolder + "\\" + ContactFolderName + "\\" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + lsFinalTitleAfterChange.Replace(",", "").Replace("/", "").Replace("\\", "") + "_" + Convert.ToDateTime(fsAsofDate).ToString("yyyyMMdd") + "_" + strGUID + "_Combined.pdf"; 
                        String lsCoversheet = ReportOpFolder + "\\" + ContactFolderName + "\\Coversheet.xls";
                        //String fsHouseHoldReportTitle = "";

                        // Generate report on excel and pdf

                        bool CombinedFileName = false;
                        if (fsGreshReportIdName != "Asset Distribution" && fsGreshReportIdName != "Asset Distribution Comparison")
                        {
                            CombinedFileName = generateCombinedPDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath.Replace(".xls", ".pdf"), fsFooterTxt, fsGreshReportIdName, LegalEntity, FundID);
                        }
                        else
                        {
                            SetValuesToVariable(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt);
                            // generatesExcelsheets(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt);
                            generatePDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, fsFooterTxt);
                            CombinedFileName = true;
                        }

                        loCoversheetCheck = new FileInfo(lsCoversheet);
                        if (!loCoversheetCheck.Exists)
                        {
                            generateCoversheetPDF(fsAsofDate, lsCoversheet, fsAllocationGroup, fsHouseholdName, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange);
                            generatesCoverExcel(fsAsofDate, fsHouseholdName, fsAllocationGroup, lsCoversheet, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange);
                        }

                        /* Array fill with the PATH + Fullname of PDF*/

                        if (i == 0)
                        {
                            SourceFileArray[i] = lsCoversheet.Replace(".xls", ".pdf");
                            if (CombinedFileName == true)
                                SourceFileArray[i + 1] = lsExcleSavePath.Replace(".xls", ".pdf");
                        }
                        else
                        {
                            if (CombinedFileName == true)
                                SourceFileArray[i + 1] = lsExcleSavePath.Replace(".xls", ".pdf");

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

                    string DisplayFileName = HHName + "_" + strYear + "-" + strMonth + strDay + ".pdf";
                    DisplayFileName = GeneralMethods.RemoveSpecialCharacters(DisplayFileName);

                    string OldDisplayFileName = OldHHName + " " + strYear + "-" + strMonth + strDay + ".pdf";



                    if (!System.IO.File.Exists(ReportOpFolder + "\\" + ConsolidatePdfFileName))
                        System.IO.File.Copy(ReportOpFolder + "\\" + ContactFolderName + "\\Coversheet.pdf", ReportOpFolder + "\\" + ConsolidatePdfFileName);

                    DestinationPath = ReportOpFolder + "\\" + GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);

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


                    if (ContactFolderName.Contains("MTGBK")) //generate without coversheet
                    {
                        string[] target = new string[sourcefilecount - 1];
                        Array.Copy(SourceFileArray, 1, target, 0, sourcefilecount - 1);
                        PDF.MergeFiles(DestinationPath, target);
                    }
                    else //generate with coversheet
                    {
                        PDF.MergeFiles(DestinationPath, SourceFileArray);
                    }

                    System.IO.Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
                }

            }


            System.IO.File.Copy(DestinationPath, ApprovedReports + ConsolidatePdfFileName);
            ////////////////////////////////////
            if (NoOfBatches == 1)
            {
                string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + ConsolidatePdfFileName);

                System.IO.File.Copy(DestinationPath, strDirectory, true);
                //Directory.Delete(ReportOpFolder, true);

                try
                {
                    //loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

                    Response.Write("<script>");
                    string lsFileNamforFinal = "./ExcelTemplate/" + ConsolidatePdfFileName;
                    //Response.Write("window.open('" + lsFileNamforFinal + "', 'mywindow')");
                    Response.Write("window.open('ViewReport.aspx?" + ConsolidatePdfFileName + "', 'mywindow')");

                    Response.Write("</script>");

                }
                catch (Exception exc)
                {
                    Response.Write(exc.Message);
                }
            }

        }
        catch (System.Web.Services.Protocols.SoapException exc)
        {
            bProceed = false;
            strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
            lblError.Text = strDescription;
        }
        catch (Exception ex)
        {
            lblError.Text = "Error Generating Report " + ex.ToString();
        }
    }

    public bool generateCombinedPDF(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFooterTxt, String ReportName, String LegalEntityId, String FundId)
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

        string filepdfname = objCombinedReports.MergeReports(fsFinalLocation, ReportName);

        if (filepdfname == "")
        {
            return false;
        }
        else
            return true;

    }

    public void generatePDF(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFooterTxt)
    {

        liPageSize = 29;
        DataSet lodataset; DB clsDB = new DB();
        lodataset = null;

        String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2);
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
        String ls = Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
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
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty));
                    document.NewPage();
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
                        loCell.BackgroundColor = new iTextSharp.text.Color(216, 216, 216);
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
                document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));
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

    public void SetValuesToVariable(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFinalReportTitle, String lsFooterTxt)
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
            lsDistributionName = "ASSET DISTRIBUTION COMPARISON";
        else
            lsDistributionName = "ASSET DISTRIBUTION";

        lsFamiliesName = lsfamilyName;
        lsDateName = Convert.ToDateTime(fsAsofDate).ToString("MMMM dd, yyyy") + "";
    }

    public void generateCoversheetPDF(String lsDateString, String fsFinalLocation, String fsAllocationGroup, String fsHouseholdName, String fsContactId, DataTable foTable, String fsKeyContactID, String fsHouseHoldTitle, String fsContactFullname, String fsDisplayContactName, String lsFinalReportTitle)
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
                UpperspaceCount = 12;
                RptTitleCount = 11;
            }

        }
        else if (TotalReportCount >= 6 && TotalReportCount < 9)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 7;
                RptTitleCount = 13;
            }
            else
            {
                UpperspaceCount = 9;
                RptTitleCount = 14;
            }
        }
        else if (TotalReportCount >= 9 && TotalReportCount < 11)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 5;
                RptTitleCount = 12;
            }
            else
            {
                UpperspaceCount = 7;
                RptTitleCount = 13;
            }
        }
        else if (TotalReportCount >= 11 && TotalReportCount < 13)
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 4;
                RptTitleCount = 16;
            }
            else
            {
                UpperspaceCount = 6;
                RptTitleCount = 17;
            }
        }
        else
        {
            if (MainTitleLengthCount >= 54)
            {
                UpperspaceCount = 1;
                RptTitleCount = 16;
            }
            else
            {
                UpperspaceCount = 2;
                RptTitleCount = 16;
            }
        }

        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4.Rotate(), 80, 80, 31, 5);
        String ls = Server.MapPath("") + "/a" + System.DateTime.Now.ToString("MMddyyhhmmss") + ".pdf";
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
        for (int liCounter = 0; liCounter < UpperspaceCount; liCounter++)//13//7
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
        for (int liCounter = 0; liCounter < RptTitleCount; liCounter++)
        {
            rowdiff = RptTitleCount - rowcount;
            if (liCounter >= rowdiff)
            {
                if (fsContactId == Convert.ToString(foTable.Rows[j]["ssi_ContactID"]).Replace(",", ""))
                {
                    loCell = new Cell();
                    loChunk = new Chunk("dev", Font8Whitecheck("test"));
                    loCell.Add(loChunk);
                    loCell.Colspan = 0;
                    loCell.HorizontalAlignment = 0;
                    loCell.Leading = 0.3f;//0.7f
                    loCell.Border = 1;
                    loTable.AddCell(loCell);

                    loCell = new Cell();
                    String lsAllocationGroupNEW = Convert.ToString(foTable.Rows[j]["Ssi_AllocationGroup"]);

                    String lsFinalTitleAfterChange = String.Empty;
                    if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["HouseHoldReportTitle"])))
                        lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["HouseHoldReportTitle"]);

                    if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["AllocationGroupReportTitle"])))
                        lsFinalTitleAfterChange = Convert.ToString(foTable.Rows[j]["AllocationGroupReportTitle"]);

                    loChunk = new Chunk(Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]) + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));//setFontsAll(10, 0, 1));

                    loCell.Add(loChunk);
                    loCell.Colspan = 1;
                    loCell.Border = 0;
                    loCell.Width = 45;//20                    
                    loCell.HorizontalAlignment = 0;
                    loTable.AddCell(loCell);
                    j++;
                }
            }
            else
            {
                if (liCounter == rowdiff - 1)
                {
                    loCell = new Cell();
                    loChunk = new Chunk("dev", Font8Whitecheck("test"));
                    loCell.Add(loChunk);
                    loCell.Colspan = 0;
                    loCell.Leading = 1f;
                    loCell.HorizontalAlignment = 0;
                    loCell.Border = 1;
                    loTable.AddCell(loCell);

                    loCell = new Cell();
                    loChunk = new Chunk("Reports included:", setFontsAll(10, 0, 1));
                    loCell.Add(loChunk);
                    loCell.Colspan = 1;
                    loCell.Border = 0;
                    loCell.HorizontalAlignment = 0;
                    loTable.AddCell(loCell);
                }
                else
                {
                    loCell = new Cell();
                    loChunk = new Chunk("dev", Font8Whitecheck("test"));
                    loCell.Add(loChunk);
                    loCell.Colspan = 2;
                    loCell.HorizontalAlignment = 1;
                    loCell.Border = 0;
                    loTable.AddCell(loCell);
                }
            }

        }

        for (int liCounter1 = 0; liCounter1 < 2; liCounter1++)
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
        loTable.AddCell(loCell);

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

    public void generatesCoverExcel(String fsAsofDate, String fsHouseholdName, String fsAllocationGroup, String fsFinalLocation, String fsContactID, DataTable foTable, String fsKeyContactID, String fsHouseHoldTitle, String fsContactFullname, String fsDisplayContactName, String lsFinalReportTitle)
    {

        String lsFileNamforFinalXls = System.DateTime.Now.ToString("MMddyyhhmmss") + ".xls";
        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\coversheet.xls");
        string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls);


        string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls.Replace("xls", "xml"));


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

    public string getFinalSp(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2)
    {
        String lsSQL = "";
        if (!String.IsNullOrEmpty(fsAllocationGroup))
        {
            lsSQL = "SP_R_Advent_Report_Allocation @AllocationGroupNameTxt='" + fsAllocationGroup.Replace("'", "''") + "', ";
        }
        else
        {
            lsSQL = "SP_R_Advent_Report_Other";
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


        lsSQL += "@LookThruDetailTxt = '" + fsLookthrogh.Replace("'", "''") + "'," +
                "@ContactFullNameTxt = '" + fsContactFullname.Replace("'", "''") + "'," +
                "@VersionTxt = '" + fsVersion.Replace("'", "''") + "'," +
                 "@summaryflgtxt = '" + fsSummaryFlag + "'," +
                   "@ReportType = '" + fsAllignment + "'," +
                "@ReportGroupFlg = " + fsReportGroupflag +
                ",@Report2GroupFlg = " + fsReportgroupflag2;

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


        lochunk = new Chunk("\n" + lsDistributionName, setFontsAll(12, 0, 0));
        loCell.Add(lochunk);

        lochunk = new Chunk("\n" + lsDateName, setFontsAll(10, 0, 1));
        loCell.Add(lochunk);
        loCell.Border = 0;
        //   loCell.Add(loParagraph);
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
                foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216);
            }
        }
        catch { }
    }

    public void setGreyBorder(Cell foCell)
    {

        foCell.BorderWidthBottom = 0.1F;
        //foCell.BorderColorBottom = new iTextSharp.text.Color(242, 242, 242);
        foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216);

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

        if (GridView1.Rows[rowIndex].Cells[17].Text != "")
        {
            string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + GridView1.Rows[rowIndex].Cells[18].Text);

            System.IO.File.Copy(GridView1.Rows[rowIndex].Cells[17].Text, strDirectory, true);
            //Directory.Delete(ReportOpFolder, true);

            try
            {
                Response.Write("<script>");
                string lsFileNamforFinal = "./ExcelTemplate/TempFolder/" + GridView1.Rows[rowIndex].Cells[18].Text;
                Response.Write("window.open('ViewReport.aspx?" + GridView1.Rows[rowIndex].Cells[18].Text + "', 'mywindow')");
                Response.Write("</script>");

            }
            catch (Exception exc)
            {
                Response.Write(exc.Message);
            }
        }
    }
    protected void ddlSalutationPref_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
    }
    protected void txtCreatedOn_TextChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
    }
    protected void ddlCreatedBy_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        Session.RemoveAll();
        if (ViewState["BatchIdListTxt"] != null)
        {
            lstMailStatus.SelectedValue = "0";
        }
        ViewState["BatchIdListTxt"] = null;
        BindGridView();
    }
    protected void ddlAction_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlAction.SelectedValue != "3")
        {
            chkMailingSheets.Enabled = false;
            chkReportSeperator.Enabled = false;
        }
        else
        {
            chkMailingSheets.Enabled = true;
            chkReportSeperator.Enabled = true;
        }
    }
    protected void btnRefresh_Click(object sender, EventArgs e)
    {
        lblError.Text = "";
        BindMailId(lstMailId);
        BindGridView();
    }

    // Sets the read-only value of a file.
    public static void SetFileReadAccess(string FileName, bool SetReadOnly)
    {
        // Create a new FileInfo object.
        FileInfo fInfo = new FileInfo(FileName);

        // Set the IsReadOnly property.
        fInfo.IsReadOnly = SetReadOnly;

    }

    private void ActionMarkAllSent()
    {
        int Count = 0;
        foreach (GridViewRow row in GridView1.Rows)
        {
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
            string ssi_mailrecordsId = row.Cells[13].Text.Trim().Replace("ssi_mailrecordsId", "").Replace("&nbsp;", "");
            string MailType = row.Cells[2].Text.Trim().Replace("Mail Type", "").Replace("&nbsp;", "");
            string ssi_reviewreqdbyid = row.Cells[23].Text.Trim().Replace("Ssi_ReviewReqdById", "").Replace("&nbsp;", "");

            if (chkSelectNC.Checked)
            {

                //if (MailType.ToUpper() != "Quarterly Statement".ToUpper() && MailType.ToUpper() != "Quarterly Statement".ToUpper() && MailType.ToUpper() != "Quarterly Statement".ToUpper()) //Commented --- Report Date and report by should update for all mail type.
                //{

                if (ddlType.SelectedValue != "5")
                {
                    Count++;
                    if (MailRecordsIdListTxt == "")
                        MailRecordsIdListTxt = ssi_mailrecordsId;
                    else
                        MailRecordsIdListTxt = MailRecordsIdListTxt + "," + ssi_mailrecordsId;


                }
                else
                    updateSentData(ssi_reviewreqdbyid, ssi_mailrecordsId);
                // }
            }
        }

        string[] Test = MailRecordsIdListTxt.Split(',');

        Session["MailRecordsIdList"] = MailRecordsIdListTxt;

        if (Count > 0)
        {
            string csname2 = "ClientScript";
            System.Text.StringBuilder cstext2 = new System.Text.StringBuilder();
            cstext2.Append("<script type=\"text/javascript\"> ");
            cstext2.Append("var myObject = window.open('SenderPopUp.aspx','win2','toolbar=0,status=no,resizable=yes,menubar=0,scrollbars=1,width=700,height=250,TOP=150,left=100');");
            cstext2.Append(" myObject.focus();");
            cstext2.Append("</script>");
            RegisterClientScriptBlock(csname2, cstext2.ToString());
        }

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

        //ssi_mailrecords objMailRecords = null;
        Entity objMailRecords = null;

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
            objMailRecords = new Entity("ssi_mailrecords");

            //objMailRecords.ssi_mailrecordsid = new Key();
            //objMailRecords.ssi_mailrecordsid.Value = new Guid(ssi_mailrecordsId);
            objMailRecords["ssi_mailrecordsid"] = new Guid(ssi_mailrecordsId);

            //objMailRecords.ssi_sentbyid = new Lookup();
            //objMailRecords.ssi_sentbyid.Value = new Guid(UserId);
            objMailRecords["ssi_sentbyid"] = new EntityReference("systemuser", new Guid(UserId));

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
    protected void ddlType_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlType.SelectedValue != "")
        {
            BindGridView();
        }
    }

    public bool GenerateMergeTypeConsolidatedPDF(string fName, string TempFolderPath)
    {
        try
        {
            string[] SourceFileName = new string[0];
            string BatchTypeId = string.Empty;

            string MailType = string.Empty; //2
            string AsOfDate = string.Empty; //3
            string HouseholdNameTxt = string.Empty; //15

            string FirstNameSort = string.Empty; //35
            string LastNameSort = string.Empty; //36

            if (!chkGroupRecandSpouse.Checked)
            {
                int NoOfBatches = 0;
                int checkBoxChecked = 0;
                //this loop will get the number of files to merge according to conditions.
                for (int j = 0; j < GridView1.Rows.Count; j++)
                {
                    CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");
                    string ReportPath = (string)GridView1.Rows[j].Cells[18].Text;
                    string ComparingFitstName = GridView1.Rows[j].Cells[35].Text;

                    if (chkBox.Checked && ReportPath.Replace("&nbsp;", "") != "")
                    {
                        if (fName == ComparingFitstName) //or create single PDF
                        {
                            NoOfBatches++;
                            checkBoxChecked = NoOfBatches;
                            if (!chkMailingSheets.Checked)
                            {
                                string IncludeSalutation = (string)GridView1.Rows[j].Cells[8].Text;
                                NoOfBatches = NoOfBatches + 1;
                                //if (IncludeSalutation.Trim().ToLower() == "regular mail")
                                //{
                                //    NoOfBatches = NoOfBatches + 1;
                                //}
                            }
                            if (!chkReportSeperator.Checked)
                            {
                                NoOfBatches = NoOfBatches + 1;
                            }

                            MailType = GridView1.Rows[j].Cells[2].Text;  //2
                            AsOfDate = GridView1.Rows[j].Cells[3].Text; //3
                            HouseholdNameTxt = GridView1.Rows[j].Cells[15].Text;//15

                            FirstNameSort = GridView1.Rows[j].Cells[35].Text; //35
                            LastNameSort = GridView1.Rows[j].Cells[36].Text;//36
                        }
                    }
                }


                string FileName = string.Empty;
                int NoofFiles = 0;
                NoofFiles = NoOfBatches;

                SourceFileName = new string[NoofFiles];
                if (SourceFileName.Length < 1)
                {
                    lblError.Text = "No Report found for selected records";
                    return false;
                }
                int FileNo = 0;
                //this loop will get the paths of files to merge according to conditions.
                for (int j = 0; j < GridView1.Rows.Count; j++)
                {
                    CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");
                    string ReportPath1 = (string)GridView1.Rows[j].Cells[18].Text;
                    if (chkBox.Checked && ReportPath1.Replace("&nbsp;", "") != "")
                    {
                        string ComparingFitstName = GridView1.Rows[j].Cells[35].Text;
                        if (fName == ComparingFitstName)
                        {
                            if (!chkMailingSheets.Checked)
                            {
                                string IncludeSalutation = (string)GridView1.Rows[j].Cells[8].Text;
                                string Name = (string)GridView1.Rows[j].Cells[6].Text;
                                string MailingAddress = (string)GridView1.Rows[j].Cells[7].Text;
                                string Salutation = (string)GridView1.Rows[j].Cells[34].Text;

                                SourceFileName[FileNo] = GenerateSalutaionPage(Name, MailingAddress, Salutation, TempFolderPath);
                                FileNo++;
                                //if (IncludeSalutation.Trim().ToLower() == "regular mail")
                                //{
                                //    string Name = (string)GridView1.Rows[j].Cells[6].Text;
                                //    string MailingAddress = (string)GridView1.Rows[j].Cells[7].Text;

                                //    SourceFileName[FileNo] = GenerateSalutaionPage(Name, MailingAddress);
                                //    FileNo++;
                                //}
                            }
                            string ReportName = (string)GridView1.Rows[j].Cells[16].Text;
                            string ReportPath = (string)GridView1.Rows[j].Cells[18].Text;

                            SourceFileName[FileNo] = ReportPath.Replace("&nbsp;", "");
                            FileNo++;
                            if (!chkReportSeperator.Checked)
                            {
                                SourceFileName[FileNo] = Server.MapPath("") + "/ExcelTemplate/Template/EndReport.pdf";
                                FileNo++;
                            }
                        }
                    }
                }
            }
            else
            {
                // to group the recipient and spouse data or arrange the recipient and spouse
                string[] src = GroupRecipientandSpouse(TempFolderPath);
                int arrCount = src.Length;
                SourceFileName = new string[arrCount];
                src.CopyTo(SourceFileName, 0);
                if (SourceFileName.Length < 1)
                {
                    lblError.Text = "No Report found for selected records";
                    return false;
                }
            }
            string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
            string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
            string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
            string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();
            string strMilliSecond = DateTime.Now.Millisecond.ToString().Length < 2 ? "0" + DateTime.Now.Millisecond.ToString() : DateTime.Now.Millisecond.ToString();

            string ConsolidatedPDFFileName = "ConsolidatedPDF_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond + strMilliSecond;
            ConsolidatedPDFFileName = GetPDFFileName(HouseholdNameTxt, AsOfDate, MailType, FirstNameSort, LastNameSort);
            string DestinationFileName = string.Empty;

            string CombinedPdfs = AppLogic.GetParam(AppLogic.ConfigParam.CombinedPdfs);

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                DestinationFileName = Request.MapPath("\\Advent Report\\ExcelTemplate\\BATCH REPORTS\\" + ConsolidatedPDFFileName + ".pdf"); //Server.MapPath("") + "\\ExcelTemplate\\" + ConsolidatedPDFFileName + ".pdf";
            }
            else
            {   // DestinationFileName = "\\\\GRPAO1-VWFS01\\opsreports$\\Combined PDFs\\" + ConsolidatedPDFFileName + "_TEST.pdf";
                DestinationFileName = CombinedPdfs + ConsolidatedPDFFileName + "_TEST.pdf"; // shared drive Changes- 7_4_2019
            }

            PDFMerge PDF = new PDFMerge();
            PDF.MergeFiles(DestinationFileName, SourceFileName);
            return true;
        }
        catch (Exception ex)
        {
            lblError.Text = ex.ToString();
            return false;
        }

    }

    private string GetPDFFileName(string HH, string asofDate, string MailType, string FirstName, string LastName)
    {
        string fileName = string.Empty;
        HH = HH.Replace("Family", "");
        string[] dt = asofDate.Split('/');
        if (MailType == "Fund Capital Call Letter")
        {
            fileName = HH + " - Capital Call " + dt[2] + "-" + dt[0] + dt[1] + " - " + FirstName.Substring(0, 1) + " " + LastName;
        }
        else if (MailType == "Fund Distribution Letter")
        {
            fileName = HH + " - Distribution Letter " + dt[2] + "-" + dt[0] + dt[1] + " - " + FirstName.Substring(0, 1) + " " + LastName;
        }
        else
        {
            fileName = HH + " - " + MailType + " " + dt[2] + "-" + dt[0] + dt[1] + " - " + FirstName.Substring(0, 1) + " " + LastName;
        }
        return fileName;
    }

    public bool InsertMailingsheet(string TempFolderPath)
    {
        try
        {
            string[] SourceFileName = new string[GridView1.Rows.Count];
            if (!chkGroupRecandSpouse.Checked)
            {
                int NoOfBatches = 0;
                int checkBoxChecked = 0;
                //this loop will get the number of files to merge according to conditions.
                for (int j = 0; j < GridView1.Rows.Count; j++)
                {
                    CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");
                    string ReportPath = (string)GridView1.Rows[j].Cells[18].Text;

                    if (chkBox.Checked == true)
                    {
                        NoOfBatches++;//
                        checkBoxChecked = NoOfBatches;
                        if (!chkMailingSheets.Checked)
                        {
                            string IncludeSalutation = (string)GridView1.Rows[j].Cells[8].Text;
                            NoOfBatches = NoOfBatches + 1;
                        }
                        if (!chkReportSeperator.Checked)
                        {
                            NoOfBatches = NoOfBatches + 1;
                        }
                    }
                }


                string FileName = string.Empty;

                int NoofFiles = 0;

                NoofFiles = NoOfBatches;

                SourceFileName = new string[NoofFiles];
                if (SourceFileName.Length < 1)
                {
                    lblError.Text = "No Report found for selected records";
                    return false;
                }
                int FileNo = 0;
                //this loop will get the paths of files to merge according to conditions.
                for (int j = 0; j < GridView1.Rows.Count; j++)
                {
                    CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");
                    string ReportPath1 = (string)GridView1.Rows[j].Cells[18].Text;
                    if (chkBox.Checked == true)
                    {
                        if (!chkMailingSheets.Checked)
                        {
                            string IncludeSalutation = (string)GridView1.Rows[j].Cells[8].Text;
                            string Name = (string)GridView1.Rows[j].Cells[6].Text.Replace("&amp;", "&").Replace("&#39;", "'");
                            string MailingAddress = (string)GridView1.Rows[j].Cells[7].Text.Replace("&amp;", "&").Replace("&#39;", "'");
                            string Salutation = (string)GridView1.Rows[j].Cells[34].Text;

                            SourceFileName[FileNo] = GenerateSalutaionPageForMailingSheet(Name, MailingAddress, Salutation, TempFolderPath);
                            FileNo++;
                        }
                        string ReportName = (string)GridView1.Rows[j].Cells[16].Text;
                        string ReportPath = (string)GridView1.Rows[j].Cells[18].Text;

                        //SourceFileName[FileNo] = ReportPath.Replace("&nbsp;", "");
                        //FileNo++;
                        //if (!chkReportSeperator.Checked)
                        //{
                        //SourceFileName[FileNo] = Server.MapPath("") + "/ExcelTemplate/Template/EndReport.pdf";
                        //FileNo++;
                        //}
                    }
                }
            }
            else
            {
                // to group the recipient and spouse data or arrange the recipient and spouse
                string[] src = GroupRecipientandSpouse(TempFolderPath);
                int arrCount = src.Length;
                SourceFileName = new string[arrCount];
                src.CopyTo(SourceFileName, 0);
                if (SourceFileName.Length < 1)
                {
                    lblError.Text = "No Report found for selected records";
                    return false;
                }
            }
            string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
            string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
            string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
            string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();
            string strMilliSecond = DateTime.Now.Millisecond.ToString().Length < 2 ? "0" + DateTime.Now.Millisecond.ToString() : DateTime.Now.Millisecond.ToString();
            string ConsolidatedPDFFileName = "ConsolidatedPDF_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond + strMilliSecond;

            string DestinationFileName = string.Empty;

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                DestinationFileName = Request.MapPath("\\Advent Report\\ExcelTemplate\\BATCH REPORTS\\" + ConsolidatedPDFFileName + ".pdf"); //Server.MapPath("") + "\\ExcelTemplate\\" + ConsolidatedPDFFileName + ".pdf";
            }
            else
            {
                //   DestinationFileName = Server.MapPath("../ExcelTemplate/pdfOutput/" + ConsolidatedPDFFileName + ".pdf");//"\\\\GRPAO1-VWFS01\\opsreports$\\Combined PDFs\\" + ConsolidatedPDFFileName + ".pdf";
                DestinationFileName = TempFolderPath + "//" + ConsolidatedPDFFileName + ".pdf";//"\\\\GRPAO1-VWFS01\\opsreports$\\Combined PDFs\\" + ConsolidatedPDFFileName + ".pdf";
            }
            //DestinationFileName = "\\\\GRPAO1-VWFS01\\shared$\\OPS REPORTS\\" + ConsolidatedPDFFileName + ".pdf";

            //DestinationFileName = "\\\\GRPAO1-VWFS01\\opsreports$\\" + ConsolidatedPDFFileName + ".pdf";


            // SourceFileName = GetDistinctValues<string>(SourceFileName);

            //string DestinationFileName = "D:\\Gresham\\TestMerge.pdf";
            PDFMerge PDF = new PDFMerge();
            PDF.MergeFiles(DestinationFileName, SourceFileName);


            string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ConsolidatedPDFFileName + ".pdf");

            //Response.Write(strDirectory);

            System.IO.File.Copy(DestinationFileName, strDirectory, true);

            string filesToDelete = "*.pdf";
            string DestPath = "C:\\AdventReport\\ExcelTemplate\\pdfOutput";
            string DestPath2 = "C:\\AdventReport\\BatchReport\\ExcelTemplate\\pdfOutput";
            if (DestPath != "")
            {
                if (Directory.Exists(DestPath))
                {
                    string[] fileList = System.IO.Directory.GetFiles(DestPath, filesToDelete);

                    foreach (string file in fileList)
                    {
                        try
                        {
                            System.IO.File.Delete(file);
                        }
                        catch
                        { }

                        //sResult += "\n" + file + "\n";
                    }
                }
            }

            if (DestPath2 != "")
            {
                if (Directory.Exists(DestPath2))
                {
                    string[] fileList = System.IO.Directory.GetFiles(DestPath2, filesToDelete);

                    foreach (string file in fileList)
                    {
                        try
                        {
                            System.IO.File.Delete(file);
                        }
                        catch
                        { }

                        //sResult += "\n" + file + "\n";
                    }
                }
            }

            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            Type tp = this.GetType();
            sb.Append("\n<script type=text/javascript>\n");
            sb.Append("\nwindow.open('ViewReport.aspx?" + ConsolidatedPDFFileName + ".pdf', 'mywindow');");
            sb.Append("</script>");
            ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());

            return true;
        }
        catch (Exception ex)
        {
            lblError.Text = ex.ToString();
            return false;
        }

    }

    #region Sharepoint
    public bool CheckFolderPathExists(String folderPath)
    {
        string siteUrl = "https://greshampartners.sharepoint.com/clientserv";
        string filename = @"E:\devlopment\GP\SharepointCode\DemoTest.txt";
        ClientContext context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

        string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);

        Web site = context.Web;
        try
        {
            //Get the required RootFolder
            //string barRootFolderRelativeUrl = "Shared Documents/test 2/";
            //  string barRootFolderRelativeUrl = folderPath;

            //folderPath = folderPath.Replace("\\", "/");
            //folderPath = "Documents/" + folderPath;
            //Folder barFolder = site.GetFolderByServerRelativeUrl(folderPath);

            int len = folderPath.Length;
            int indexlen = folderPath.IndexOf("Documents");
            indexlen = indexlen + 10;
            int cnt = len - indexlen;
            string vNewSharePointReportFolder = folderPath.Substring(indexlen, cnt);
            vNewSharePointReportFolder = vNewSharePointReportFolder.Replace("\\", "/").Replace(@"\", "/");




            vNewSharePointReportFolder = vNewSharePointReportFolder.Replace("\\", "/");
            vNewSharePointReportFolder = "Documents/" + vNewSharePointReportFolder;
            Folder barFolder = site.GetFolderByServerRelativeUrl(vNewSharePointReportFolder);

            // context.Load(barFolder);
            context.ExecuteQuery();

            return true;
        }
        catch
        {
            return false;
        }

        //  return true;
    }

    public void CopyFile(string Ssi_SharePointReportFolder, string destFilename, string vSourcrFile)  // string vSourcefile, string vDestinationFile
    {


        //string Ssi_ClientPortalFolder = @"\\sp02\\Client%20Portal\Documents\Scalise%20Test\Gresham%20Statements\2016";

        // Ssi_SharePointReportFolder = @"\\sp02\ClientServ\Documents\Test%20JMASA\";

        //string Filename = "test_Masa T 2016-0630.pdf";

        //string filename = @"E:\devlopment\GP\SharepointCode\DemoTest.txt";

        Ssi_SharePointReportFolder = Ssi_SharePointReportFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();
        //  Ssi_ClientPortalFolder = Ssi_ClientPortalFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();

        int len = Ssi_SharePointReportFolder.Length;
        int indexlen = Ssi_SharePointReportFolder.IndexOf("Documents");
        indexlen = indexlen + 10;
        int cnt = len - indexlen;
        string vNewSharePointReportFolder = Ssi_SharePointReportFolder.Substring(indexlen, cnt);
        vNewSharePointReportFolder = vNewSharePointReportFolder.Replace("\\", "/").Replace(@"\", "/");


        //len = Ssi_ClientPortalFolder.Length;
        //indexlen = Ssi_ClientPortalFolder.IndexOf("Documents");
        //indexlen = indexlen + 10;
        //cnt = len - indexlen;
        //string vNewClientFolderPath = Ssi_ClientPortalFolder.Substring(indexlen, cnt);

        //  string FilePath = vNewSharePointReportFolder + Filename;

        // FilePath = FilePath.Replace("\\", "/");

        vNewSharePointReportFolder = "Documents/" + vNewSharePointReportFolder;
        // Response.Write(vNewSharePointReportFolder);

        string siteUrl = "https://greshampartners.sharepoint.com/clientserv";
        ClientContext context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

        string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);

        Web site = context.Web;

        Folder currentRunFolder = site.GetFolderByServerRelativeUrl(vNewSharePointReportFolder);
        FileCreationInformation newFile = new FileCreationInformation { Content = System.IO.File.ReadAllBytes(vSourcrFile), Url = Path.GetFileName(destFilename), Overwrite = true };
        currentRunFolder.Files.Add(newFile);

        currentRunFolder.Update();

        context.ExecuteQuery();

        //  bool result= CheckFolderPathExis(vNewSharePointReportFolder);



    }

    public void CopyFilenew(string Ssi_SharePointReportFolder, string destFilename, string vSourcrFile)  // string vSourcefile, string vDestinationFile
    {

        #region not used
        //string Ssi_ClientPortalFolder = @"\\sp02\\Client%20Portal\Documents\Scalise%20Test\Gresham%20Statements\2016";

        // Ssi_SharePointReportFolder = @"\\sp02\ClientServ\Documents\Test%20JMASA\";

        //string Filename = "test_Masa T 2016-0630.pdf";

        //string filename = @"E:\devlopment\GP\SharepointCode\DemoTest.txt";

        //Ssi_SharePointReportFolder = Ssi_SharePointReportFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();
        ////  Ssi_ClientPortalFolder = Ssi_ClientPortalFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();

        //int len = Ssi_SharePointReportFolder.Length;
        //int indexlen = Ssi_SharePointReportFolder.IndexOf("Documents");
        //indexlen = indexlen + 10;
        //int cnt = len - indexlen;
        //string vNewSharePointReportFolder = Ssi_SharePointReportFolder.Substring(indexlen, cnt);
        //vNewSharePointReportFolder = vNewSharePointReportFolder.Replace("\\", "/").Replace(@"\", "/");

        //len = Ssi_ClientPortalFolder.Length;
        //indexlen = Ssi_ClientPortalFolder.IndexOf("Documents");
        //indexlen = indexlen + 10;
        //cnt = len - indexlen;
        //string vNewClientFolderPath = Ssi_ClientPortalFolder.Substring(indexlen, cnt);

        //  string FilePath = vNewSharePointReportFolder + Filename;

        // FilePath = FilePath.Replace("\\", "/");

        //   vNewSharePointReportFolder = "Documents/" + vNewSharePointReportFolder;
        //  Response.Write(vNewSharePointReportFolder);
        #endregion


        string vNewSharePointReportFolder = sharepointFolderPath(Ssi_SharePointReportFolder);
        string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientservURL);
        //string siteUrl = "https://greshampartners.sharepoint.com/clientserv";
        ClientContext context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        ////foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        ////context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

        string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);

        Web site = context.Web;

        byte[] bytes = System.IO.File.ReadAllBytes(vSourcrFile);
        System.IO.Stream stream = new System.IO.MemoryStream(bytes);

        Folder currentRunFolder = site.GetFolderByServerRelativeUrl(vNewSharePointReportFolder);
        FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(destFilename), Overwrite = true };
        currentRunFolder.Files.Add(newFile);

        currentRunFolder.Update();

        context.ExecuteQuery();

        //  bool result= CheckFolderPathExis(vNewSharePointReportFolder);



    }
    public bool LoadList()
    {
        bool bProceed = false;
        try
        {
            // DataTable FolderData;     // clientPortal folderPath Datatable\
            DataSet dsTaxonomyclientPortal1;   // clientPortal taxonomy data
            DataTable dtSiteClientList;

            DataTable dtLEClientList;

            DataTable dtCorrespondenceType;
            DataTable dsDocumentTaxonomy;
            DataTable dtActiveClientList;

            //   FolderData = sp.getSPList();
            dsTaxonomyclientPortal1 = sp.getTaxonomyClientPortal();
            dtSiteClientList = sp.getSiteClientList();

            //if (billingflg)
            //{
            //    dsDocumentTaxonomy = sp.getTaxonomyClientService();
            //    ViewState["dtDocumentTaxonomy"] = dsDocumentTaxonomy;
            //}

            //   ViewState["dtFolderData"] = FolderData;
            ViewState["dsTaxonomyclientPortal"] = dsTaxonomyclientPortal1;
            ViewState["dtSiteClientList"] = dtSiteClientList;




            #region New Client Services Shaepoint
            //dsDocumentTaxonomy = sp.getTaxonomyClientService();
            dtActiveClientList = sp.getActiveClientList();
            dtCorrespondenceType = sp.getTaxonomyCorrespondenceType();

            //  ViewState["dtDocumentTaxonomy"] = dsDocumentTaxonomy;
            ViewState["dtActiveClientList"] = dtActiveClientList;
            ViewState["dtCorrespondenceType"] = dtCorrespondenceType;



            #endregion
            //if (FolderData != null && dsTaxonomyclientPortal1 != null && dtSiteClientList != null && dsDocumentTaxonomy != null && dtActiveClientList != null && dtCorrespondenceType != null)
            if (dsTaxonomyclientPortal1 != null && dtSiteClientList != null && dtActiveClientList != null && dtCorrespondenceType != null)
            {
                bProceed = true;
            }
            else
            {
                bProceed = false;
                lblError2.Visible = true;
                // lblError2.Text = "Error Fetching Taxonomy List";
                lblError2.Text = "Unable to Connect to SharePoint , Please try again after sometime";
            }
        }
        catch (Exception Ex)
        {
            lblError2.Visible = true;
            // lblError2.Text = "Error Fetching Taxonomy List";
            lblError2.Text = "Unable to Connect to SharePoint , Please try again after sometime";
            bProceed = false;
        }
        return bProceed;
    }

    public bool CopyFiletoSharepoint(string sharepointfolderpath, string destFilename, string vSourcrFile, string BatchType, string BatchName, string HouseHoldName, string year, string SharepointFolderFilePath, string ClientPortalName, string billinginvoiceid)
    {
        //DataTable dtFolderData = (DataTable)ViewState["dtFolderData"];
        DataSet dsTaxonomyclientPortal = (DataSet)ViewState["dsTaxonomyclientPortal"];
        DataTable dtDocumentType = dsTaxonomyclientPortal.Tables[1];
        DataTable dtClientSite = dsTaxonomyclientPortal.Tables[0];
        DataTable dtYear = dsTaxonomyclientPortal.Tables[2];
        DataTable dtClient = (DataTable)ViewState["dtSiteClientList"];

        string taggingClientID = string.Empty;
        string taggingClientName = string.Empty;

        //vSourcrFile = @"D:\Back Data 19-08-2016\D Data\Practice\11_4_2016\Eterovic Family_2016-0930_2016_11_04_08_58_37.pdf";
        //destFilename = "Eterovic Family_2016-0930_2016_11_04_08_58_37.pdf";
        HouseHoldName = HouseHoldName.Replace(" Family", "");
        HouseHoldName = HouseHoldName.Replace(" family", "");
        HouseHoldName = HouseHoldName.Replace(" FAMILY", "");

        ClientPortalName = ClientPortalName.Replace(" Family", "");
        ClientPortalName = ClientPortalName.Replace(" family", "");
        ClientPortalName = ClientPortalName.Replace(" FAMILY", "");
        string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientportalURL);
       // string siteUrl = "https://greshampartners.sharepoint.com/ClientPortal";
        ClientContext context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

        string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);

        Web site = context.Web;
        //    ClientPortalName;


        foreach (DataRow rw in dtClientSite.Rows)
        {
            if (ClientPortalName.ToLower() == rw["ClientName"].ToString().ToLower())
            {
                //onPortal = rw["OnPortal"].ToString();
                taggingClientID = rw["iID"].ToString();
                break;
            }
        }
        taggingClientName = ClientPortalName;
        //if (taggingClientID == "")
        //{
        //    foreach (DataRow rw in dtClientSite.Rows)
        //    {
        //        if (HouseHoldName.ToLower() == rw["ClientName"].ToString().ToLower())
        //        {
        //            //onPortal = rw["OnPortal"].ToString();
        //            taggingClientID = rw["iID"].ToString();
        //            break;
        //        }
        //    }
        //}


        if (taggingClientID != "")
        {

            //string exte = System.IO.Path.GetExtension(destFilename);
            //string Filenames = destFilename;  

            //string filenameWithoutext = Filenames.Substring(0, Filenames.LastIndexOf("."));

            //string iClientID = string.Empty;
            //foreach (DataRow rw in dtClient.Rows)
            //{
            //    if (HouseHoldName == rw["ClientName"].ToString())
            //    {
            //        iClientID = rw["iID"].ToString();
            //        break;
            //    }
            //}
            //Filenames = filenameWithoutext + "_" + iClientID + exte;


            byte[] bytes = System.IO.File.ReadAllBytes(vSourcrFile);
            System.IO.Stream stream = new System.IO.MemoryStream(bytes);

            Folder currentRunFolder = site.GetFolderByServerRelativeUrl("Documents taxonomy");
            // Folder currentRunFolder = site.GetFolderByServerRelativeUrl("docTaxTest");
            FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = destFilename, Overwrite = true };
            //currentRunFolder.Files.Add(newFile);


            Microsoft.SharePoint.Client.List docs = context.Web.Lists.GetByTitle("Documents taxonomy");
            // Microsoft.SharePoint.Client.List docs = context.Web.Lists.GetByTitle("docTaxTest");

            Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile);

            context.Load(uploadFile);
            context.Load(docs);
            context.ExecuteQuery();

            context.Load(uploadFile.ListItemAllFields);
            context.ExecuteQuery();

            // Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

            Microsoft.SharePoint.Client.ListItem item = docs.GetItemById(uploadFile.ListItemAllFields.Id);
            context.Load(item);
            context.ExecuteQuery();

            string PathTaggingName = string.Empty;
            string PathTaggingID = string.Empty;



            //foreach (DataRow rw in dtFolderData.Rows)
            //{
            //    if (sharepointfolderpath.Contains(rw["FolderPath"].ToString()))
            //    {
            //        PathTaggingName = rw["Tag"].ToString();
            //        //vIsYear = rw["OnPortal"].ToString();
            //        break;
            //    }

            //}


            sw.WriteLine("Batchtype:" + BatchType);
            if (BatchType.ToUpper() == "Q" || BatchType.ToUpper() == "M")
            {

                PathTaggingName = "Gresham Statements";
            }
            else if (BatchType.ToUpper() == "MERGE")// && (BatchName.ToLower().Contains("cap") || BatchName.ToLower().Contains("dist")))
            {
                PathTaggingName = "NonMarketable";
            }

            string Taggingyear = string.Empty;
            string TaggingYearID = string.Empty;

            string iClientID = string.Empty;
            string onPortal = string.Empty;

            Taggingyear = year;
            foreach (DataRow rw in dtYear.Rows)
            {
                if (rw["Year"].ToString() == Taggingyear)
                    TaggingYearID = rw["iID"].ToString();

            }

            TaxonomyFieldValue taxonomyFieldValueClient = new TaxonomyFieldValue();
            TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
            TaxonomyFieldValue taxonomyFieldValueYear = new TaxonomyFieldValue();

            if (BatchType.ToUpper() == "Q" || BatchType.ToUpper() == "M")
            {
                // taxonomyFieldValuePath.TermGuid = PathTaggingID;
                taxonomyFieldValuePath.TermGuid = "953cf71a-90c8-42e8-8393-f71adf3df1f2";
                taxonomyFieldValuePath.Label = "Gresham Statements";

                taxonomyFieldValueClient.TermGuid = taggingClientID;
                taxonomyFieldValueClient.Label = taggingClientName;


            }
            else if (BatchType.ToUpper() == "MERGE")///&& (BatchName.ToLower().Contains("capital call") || BatchName.ToLower().Contains("distribution")))
            {
                sw.WriteLine("Batchtype:" + BatchType + "under merge ");
                //taxonomyFieldValuePath.TermGuid = PathTaggingID;
                sw.WriteLine("billinginvoiceid:" + billinginvoiceid);
                //    Response.Write("billinginvoiceid:" + billinginvoiceid);
                if (billinginvoiceid != "")
                {
                    taxonomyFieldValuePath.TermGuid = "ba310341-29be-4754-9077-40c81f676f7b";
                    //  taxonomyFieldValuePath.Label = PathTaggingName;
                    taxonomyFieldValuePath.Label = "Billing";
                }
                else
                {
                    sw.WriteLine("billinginvoiceid:" + billinginvoiceid + "under capitacl call");
                    taxonomyFieldValuePath.TermGuid = "60418785-db1d-4558-bb0b-630efc86ecbb";
                    //  taxonomyFieldValuePath.Label = PathTaggingName;
                    taxonomyFieldValuePath.Label = "NonMarketable";
                }
                taxonomyFieldValueClient.TermGuid = taggingClientID;
                taxonomyFieldValueClient.Label = taggingClientName;
            }

            // taxonomyFieldValueClient.TermGuid = taggingClientID;


            taxonomyFieldValueYear.TermGuid = TaggingYearID;
            taxonomyFieldValueYear.Label = Taggingyear;

            //else
            //{

            //    item["p9cafa43d635492cb87a8a60d0ebb191"] = "";
            //}

            //item["d19c761c862c4a1d960e584c607dfa04"] = taxonomyFieldValueClient;
            item["g6508b71d21947cdacac1f29db22f573"] = taxonomyFieldValuePath;
            item["d19c761c862c4a1d960e584c607dfa04"] = taxonomyFieldValueClient;
            if (billinginvoiceid == "")
                item["p9cafa43d635492cb87a8a60d0ebb191"] = taxonomyFieldValueYear;

            item.Update();
            docs.Update();
            context.ExecuteQuery();

            sw.WriteLine("Sucess flow");
            return true;

        }
        else
        {
            //Response.Write("<br/>Error Occured when trying to copy file from: " + vSourcrFile +
            //                      " to " + SharepointFolderFilePath + "<br/> Cient Name not found: " +HouseHoldName );
            sw.WriteLine("Sucess flow return false");
            return false;

        }
    }

    public string sharepointFolderPath(string Ssi_SharePointReportFolder)
    {
        Ssi_SharePointReportFolder = Ssi_SharePointReportFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();
        //  Ssi_ClientPortalFolder = Ssi_ClientPortalFolder.Replace("%20", " ").Replace("&#39;", "'").ToString();

        int len = Ssi_SharePointReportFolder.Length;
        int indexlen = Ssi_SharePointReportFolder.IndexOf("Documents");
        indexlen = indexlen + 10;
        int cnt = len - indexlen;
        string vNewSharePointReportFolder = Ssi_SharePointReportFolder.Substring(indexlen, cnt);
        vNewSharePointReportFolder = vNewSharePointReportFolder.Replace("\\", "/").Replace(@"\", "/");

        vNewSharePointReportFolder = "Documents/" + vNewSharePointReportFolder;
        return vNewSharePointReportFolder;
    }

    public bool checkSharepouintFileExist(string FilePath, string filename)
    {

        string filePath = "/clientserv/" + FilePath + "/" + filename;


        string siteUrl = "https://greshampartners.sharepoint.com";
        ClientContext clientContext = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        // clientContext.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

        string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        clientContext.Credentials = new SharePointOnlineCredentials(user, passWord);

        // Web site = context.Web;

        Web web = clientContext.Web;
        Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl(filePath);
        bool bExists = false;
        try
        {
            clientContext.Load(file);
            clientContext.ExecuteQuery();
            bExists = file.Exists;
        }
        catch
        {
            bExists = false;
        }



        return bExists;

    }

    public bool CheckNewCSSiteExists(String URL, string billingInvoiceID)
    {
        bool bProceed = false;
        try
        {

            #region Commented 
            //if (billingInvoiceID != "")
            //{
            //    string siteUrl = URL;
            //    ClientContext context = new ClientContext(siteUrl);
            //    SecureString passWord = new SecureString();
            //    //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //    //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            //    string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId).ToString();
            //    string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
            //    foreach (var c in Pass) passWord.AppendChar(c);
            //    context.Credentials = new SharePointOnlineCredentials(user, passWord);

            //    Web site = context.Web;
            //    List list = context.Web.Lists.GetByTitle("Compliance Documents");

            //    context.Load(list);
            //    context.ExecuteQuery();

            //    bProceed = true;

            //}

            //else
            //{
            //    string siteUrl = URL;
            //    ClientContext context = new ClientContext(siteUrl);
            //    SecureString passWord = new SecureString();
            //    //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //    //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            //    string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId).ToString();
            //    string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
            //    foreach (var c in Pass) passWord.AppendChar(c);
            //    context.Credentials = new SharePointOnlineCredentials(user, passWord);

            //    Web site = context.Web;
            //    List list = context.Web.Lists.GetByTitle("Published Documents");

            //    context.Load(list);
            //    context.ExecuteQuery();

            //    bProceed = true;
            //}

            #endregion


            string siteUrl = URL;
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;
            List list = context.Web.Lists.GetByTitle("Published Documents");

            context.Load(list);
            context.ExecuteQuery();

            bProceed = true;

        }
        catch (Exception Ex)
        {
            return false;
        }

        return bProceed;
    }
    public bool CheckFileExistinNewCSSite(String URL, string FileName)
    {
        bool Proceed = false;
        try
        {

            string siteUrl = URL;
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;
            List list = context.Web.Lists.GetByTitle("Published Documents");

            context.Load(list);
            context.ExecuteQuery();

            CamlQuery camlQuery = new CamlQuery();

            //WORKING
            camlQuery.ViewXml = @"<View Scope='RecursiveAll'>
                   <Query>
   <Where>
      <Eq>
         <FieldRef Name='FileLeafRef' />
         <Value Type='File'>" + FileName + "</Value></Eq> </Where></Query> <RowLimit>4990</RowLimit> </View>";

            Microsoft.SharePoint.Client.ListItemCollection listItems = list.GetItems(camlQuery);
            context.Load(listItems);
            context.ExecuteQuery();
            int FileCount = listItems.Count;

            if (FileCount > 0)
            {
                Proceed = true;
            }
        }
        catch (Exception Ex)
        {
            return false;
        }
        return Proceed;

    }
    public string CopyFilenewCS(string Ssi_SharePointReportFolder, string destFilename, string vSourcrFile, string year, string BatchType, string Quarter, string BillingInvoiceId)  // string vSourcefile, string vDestinationFile
    {
        string FileLink = string.Empty;
        string siteUrl = string.Empty;
        try
        {

            DataSet dsTaxonomyclientPortal = (DataSet)ViewState["dsTaxonomyclientPortal"];
            DataTable dtYear = dsTaxonomyclientPortal.Tables[2];
            DataTable dtCorrespondenceType = (DataTable)ViewState["dtCorrespondenceType"];

            // DataTable dtLEList = (DataTable)ViewState["dtDocumentTaxonomy"];

            //   bool billingflag = Convert.ToBoolean(ViewState["BillingFlg"]);


            siteUrl = Ssi_SharePointReportFolder;
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;


            // vSourcrFile = @"D:\Infograte\Site\TEST Report Output\OPS REPORTS\zzTest_Cap Call - Masa TEST Legal Entity 2 2019-0731.pdf";

            byte[] bytes = System.IO.File.ReadAllBytes(vSourcrFile);
            System.IO.Stream stream = new System.IO.MemoryStream(bytes);

            FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = destFilename, Overwrite = true };

            //if (billingflag)
            //{
            #region billing

            //DataTable dtchoice = getListChoiceField(Ssi_SharePointReportFolder);




            //Microsoft.SharePoint.Client.List docs = context.Web.Lists.GetByTitle("Compliance Documents");
            //context.ExecuteQuery();



            //// List list = context.Web.Lists.GetByTitle("Compliance Documents");

            //// fldChoice.


            //Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile); // NEW File to be created and uploaded 

            //context.Load(uploadFile);
            //context.Load(docs);
            //context.ExecuteQuery();


            //context.Load(uploadFile.ListItemAllFields);
            //context.ExecuteQuery();

            //// Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

            //Microsoft.SharePoint.Client.ListItem item = docs.GetItemById(uploadFile.ListItemAllFields.Id); // Fetch the Uploaded File to Tag
            //context.Load(item);
            //context.ExecuteQuery();

            //TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
            //TaxonomyFieldValue taxonomyFieldValueLegalEntity = new TaxonomyFieldValue();

            //FileLink = item["FileRef"].ToString();
            //string docType = string.Empty;
            //string iID = string.Empty;

            //string LEName = string.Empty;
            //string LEID = string.Empty;

            //string chfieldname = "billing";
            //string chfieldid = string.Empty;
            //int cnt = 0;


            //Field choice = docs.Fields.GetByInternalNameOrTitle("Compliance Document Type");

            ////FieldChoice fldChoice = context.CastTo<FieldChoice>(choice);
            //context.Load(choice);
            //context.Load(item);
            //context.ExecuteQuery();

            ////var values = fldChoice.Choices;

            //var fieldStatus = context.CastTo<FieldChoice>(choice);
            //var values = fieldStatus.Choices;

            //foreach (DataRow rw in dtchoice.Rows)
            //{
            //    cnt++;

            //    if (chfieldname == rw["Choicefield"].ToString().ToLower())
            //    {
            //        chfieldid = Convert.ToString(cnt);
            //        chfieldname = rw["Choicefield"].ToString();
            //        break;
            //    }
            //}


            //if (LegalEntityUUID != "" && LegalEntityUUID != null && LegalEntityUUID != "&nbsp;")
            //{
            //    foreach (DataRow rw in dtLEList.Rows)
            //    {


            //        if (LegalEntityUUID == rw["TaxonomyValue"].ToString().ToLower())
            //        {
            //            LEID = LegalEntityUUID;
            //            LEName = rw["TaxonomyName"].ToString();
            //            break;
            //        }
            //    }
            //}


            //if (LegalEntityUUID != "" && LegalEntityUUID != null && LegalEntityUUID != "&nbsp;")
            //{
            //    taxonomyFieldValueLegalEntity.TermGuid = LEID; // LegalEntity ID
            //    taxonomyFieldValueLegalEntity.Label = LEName;//LegalEntityUUID Name
            //}
            ////taxonomyFieldValuePath.TermGuid = iID;// Id from The Correspondence Taxonomy
            ////taxonomyFieldValuePath.Label = docType;// "Correspondence"  OR "Quarterly Report";
            //try
            //{
            //    if (LegalEntityUUID != "" && LegalEntityUUID != null && LegalEntityUUID != "&nbsp;")
            //    {
            //        item["ce9c7b54e9364a2f9eff5e92f7d79d65"] = taxonomyFieldValueLegalEntity;
            //    }
            //    else
            //    {
            //        item["ce9c7b54e9364a2f9eff5e92f7d79d65"] = "";
            //    }
            //    //item["Quarter"] = Quarter; // Quarter Tag
            //    //item["ocac1b27043549bf95a6b3be20a5e5ea"] = taxonomyFieldValuePath; // Correspondence Type
            //    //item["Year"] = year; // Year Field
            //    //  item["ComplianceDocumentType"] = chfieldname;

            //    item["ComplianceDocumentType"] = chfieldname;



            //    item.Update();
            //    docs.Update();
            //    context.ExecuteQuery();


            //    // Response.Write(chfieldname);
            //}
            //catch (Exception Ex)
            //{
            //    Response.Write("<br/>Error Occured while tagging the File: " + vSourcrFile +
            //                         " to " + Ssi_SharePointReportFolder + "Compliance Documents" + "<br/>" + Ex.Message + ", " + Ex.StackTrace);
            //}


            #endregion
            //}

            //else
            //{

            #region non billing



            Microsoft.SharePoint.Client.List docs = context.Web.Lists.GetByTitle("Published Documents");
            context.ExecuteQuery();

            Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile); // NEW File to be created and uploaded 

            context.Load(uploadFile);
            context.Load(docs);
            context.ExecuteQuery();

            context.Load(uploadFile.ListItemAllFields);
            context.ExecuteQuery();

            // Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

            Microsoft.SharePoint.Client.ListItem item = docs.GetItemById(uploadFile.ListItemAllFields.Id); // Fetch the Uploaded File to Tag
            context.Load(item);
            context.ExecuteQuery();

            TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
            TaxonomyFieldValue taxonomyFieldValueLegalEntity = new TaxonomyFieldValue();

            FileLink = item["FileRef"].ToString();
            string docType = string.Empty;
            string iID = string.Empty;

            string LEName = string.Empty;
            string LEID = string.Empty;

            //if (LegalEntityUUID != "" && LegalEntityUUID != null && LegalEntityUUID != "&nbsp;")
            //{
            //    foreach (DataRow rw in dtLegalEntityUUID.Rows)
            //    {


            //        if (LegalEntityUUID == rw["TaxonomyValue"].ToString().ToLower())
            //        {
            //            LEID = LegalEntityUUID;
            //            LEName = rw["TaxonomyName"].ToString();
            //            break;
            //        }
            //    }
            //}

            if (BatchType.ToUpper() == "Q" || BatchType.ToUpper() == "M")
            {
                docType = "Quarterly Report";
            }
            else if (BatchType.ToUpper() == "MERGE" && BillingInvoiceId == "")
            {
                //  docType = "Correspondence";
                docType = "Cap Call/Distribution";//added on 07/26/2019
            }
            else if (BatchType.ToUpper() == "MERGE" && BillingInvoiceId != "")
            {
                // docType = "Correspondence";
                docType = "Billing";  // changed to billing 19_9_2019 - sasmit(Basceamp request)
            }



            //  Response.Write("biilingflag=" + billingflag);

            foreach (DataRow rw in dtCorrespondenceType.Rows)
            {
                if (docType.ToLower() == rw["DocumentType"].ToString().ToLower())
                {
                    iID = rw["iID"].ToString();
                    break;
                }
            }
            //if (LegalEntityUUID != "" && LegalEntityUUID != null && LegalEntityUUID != "&nbsp;")
            //{
            //    taxonomyFieldValueLegalEntity.TermGuid = LEID; // LegalEntity ID
            //    taxonomyFieldValueLegalEntity.Label = LEName;//LegalEntityUUID Name
            //}
            taxonomyFieldValuePath.TermGuid = iID;// Id from The Correspondence Taxonomy
            taxonomyFieldValuePath.Label = docType;// "Correspondence"  OR "Quarterly Report";
            try
            {
                //if (LegalEntityUUID != "" && LegalEntityUUID != null && LegalEntityUUID != "&nbsp;")
                //{
                //    item["ce9c7b54e9364a2f9eff5e92f7d79d65"] = taxonomyFieldValueLegalEntity;
                //}
                //else
                //{
                //    item["ce9c7b54e9364a2f9eff5e92f7d79d65"] = "";
                //}
                item["Quarter"] = Quarter; // Quarter Tag
                item["ocac1b27043549bf95a6b3be20a5e5ea"] = taxonomyFieldValuePath; // Correspondence Type
                item["Year"] = year; // Year Field
                item.Update();
                docs.Update();
                context.ExecuteQuery();
            }
            catch (Exception Ex)
            {
                Response.Write("<br/>Error Occured while tagging the File: " + vSourcrFile +
                                     " to " + Ssi_SharePointReportFolder + "Published Documents" + "<br/>" + Ex.Message + ", " + Ex.StackTrace);
            }


            #endregion
            //}
        }
        catch (Exception Exx)
        {
            FileLink = "";
            Response.Write("<br/>Error Occured when trying to copy file from: " + vSourcrFile +
                                      " to " + Ssi_SharePointReportFolder + "Published Documents" + "<br/>" + Exx.Message + ", " + Exx.StackTrace);
            return "";

        }
        return FileLink;

    }



    public DataTable getListChoiceField(string Ssi_SharePointReportFolder)
    {

        string siteUrl = Ssi_SharePointReportFolder;
        ClientContext context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();
        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

        string user = AppLogic.GetParam(AppLogic.ConfigParam.EmailId).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.CRMPassword).ToString();
        foreach (var c in Pass) passWord.AppendChar(c);
        context.Credentials = new SharePointOnlineCredentials(user, passWord);

        Web site = context.Web;

        DataTable dt = new DataTable();
        dt.Columns.Add("Choicefield");

        // ClientContext context = new ClientContext(source_site_url);
        List list = context.Web.Lists.GetByTitle("Compliance Documents");
        Field choice = list.Fields.GetByInternalNameOrTitle("Compliance Document Type");

        FieldChoice fldChoice = context.CastTo<FieldChoice>(choice);
        context.Load(fldChoice, f => f.Choices);
        context.ExecuteQuery();
        foreach (string item in fldChoice.Choices)
        {
            DataRow row = dt.NewRow();
            row["Choicefield"] = item.ToString();
            // row["iID"] = ts.Id.ToString();
            dt.Rows.Add(row);

            //add choices to dropdown list
        }

        return dt;
    }
    public string GetQuarter(DateTime date)
    {
        if (date.Month >= 1 && date.Month <= 3)
            return "1";
        else if (date.Month >= 4 && date.Month <= 6)
            return "2";
        else if (date.Month >= 7 && date.Month <= 9)
            return "3";
        else
            return "4";
    }
    public bool CheckFileExistinLegalEntity(string URL, string FolderNAme, string SubFolder, string FileName)
    {
        bool bProceed = false;

        try
        {
            string siteUrl = URL;
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            ////foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            ////context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;


            Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderNAme);//PublishedDocuments
                                                                                    // int count = currentRunFolder.Files.Count;
            context.Load(currentRunFolder);
            context.ExecuteQuery();

            Folder subRunFolder = currentRunFolder.Folders.GetByUrl(SubFolder); // LegalEntity
            context.Load(subRunFolder);
            context.ExecuteQuery();

            Microsoft.SharePoint.Client.File file = subRunFolder.Files.GetByUrl(FileName);
            context.Load(file);
            context.ExecuteQuery();


            bProceed = true;
        }
        catch (Exception ex)
        {
            return false;
        }
        return bProceed;
    }

    public bool CheckLEgalEntityFolderExist(string URL, string FolderNAme, string SubFolder)
    {
        bool bProceed = false;

        try
        {
            string siteUrl = URL;
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            ////foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            ////context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;


            Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderNAme);//PublishedDocuments
                                                                                    // int count = currentRunFolder.Files.Count;
            context.Load(currentRunFolder);
            context.ExecuteQuery();

            Folder subRunFolder = currentRunFolder.Folders.GetByUrl(SubFolder); // LegalEntity
            context.Load(subRunFolder);
            context.ExecuteQuery();

            bProceed = true;
        }
        catch (Exception ex)
        {
            //lg.AddinLogFile(LogFileName, "IsSharepointSiteExists Error" + ex.Message.ToString());
            return false;

        }
        return bProceed;
    }
    public string CopyFileinLegalEntityFolder(string URL, string destFilename, string vSourcrFile, string FolderNAme, string year, string BatchType, string Quarter, string BillingInvoiceId)
    {
        string FileLink = string.Empty;
        try
        {
            DataTable dtCorrespondenceType = (DataTable)ViewState["dtCorrespondenceType"];

            //  DataTable dtLEList = (DataTable)ViewState["dtDocumentTaxonomy"];

            //   bool billingflag = Convert.ToBoolean(ViewState["BillingFlg"]);

            string siteUrl = URL;
            ClientContext context = new ClientContext(siteUrl);
            SecureString passWord = new SecureString();
            ////foreach (var c in "w!ldWind36") passWord.AppendChar(c);
            ////context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);

            string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
            string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
            foreach (var c in Pass) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(user, passWord);

            Web site = context.Web;

            byte[] bytes = System.IO.File.ReadAllBytes(vSourcrFile);
            System.IO.Stream stream = new System.IO.MemoryStream(bytes);

            Folder currentRunFolder = site.GetFolderByServerRelativeUrl(FolderNAme);

            FileCreationInformation newFile = new FileCreationInformation { ContentStream = stream, Url = Path.GetFileName(destFilename), Overwrite = true };

            //if (billingflag)
            //{
            //    #region billing

            //    DataTable dtchoice = getListChoiceField(URL);




            //    Microsoft.SharePoint.Client.List docs = context.Web.Lists.GetByTitle("Compliance Documents");
            //    context.ExecuteQuery();



            //    // List list = context.Web.Lists.GetByTitle("Compliance Documents");

            //    // fldChoice.


            //    Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile); // NEW File to be created and uploaded 

            //    context.Load(uploadFile);
            //    context.Load(docs);
            //    context.ExecuteQuery();


            //    context.Load(uploadFile.ListItemAllFields);
            //    context.ExecuteQuery();

            //    // Microsoft.SharePoint.Client.ListItem item2 = uploadFile.ListItemAllFields;

            //    Microsoft.SharePoint.Client.ListItem item = docs.GetItemById(uploadFile.ListItemAllFields.Id); // Fetch the Uploaded File to Tag
            //    context.Load(item);
            //    context.ExecuteQuery();

            //    TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
            //    TaxonomyFieldValue taxonomyFieldValueLegalEntity = new TaxonomyFieldValue();

            //    FileLink = item["FileRef"].ToString();
            //    string docType = string.Empty;
            //    string iID = string.Empty;

            //    string LEName = string.Empty;
            //    string LEID = string.Empty;

            //    string chfieldname = "billing";
            //    string chfieldid = string.Empty;
            //    int cnt = 0;


            //    Field choice = docs.Fields.GetByInternalNameOrTitle("Compliance Document Type");

            //    //FieldChoice fldChoice = context.CastTo<FieldChoice>(choice);
            //    context.Load(choice);
            //    context.Load(item);
            //    context.ExecuteQuery();

            //    //var values = fldChoice.Choices;

            //    var fieldStatus = context.CastTo<FieldChoice>(choice);
            //    var values = fieldStatus.Choices;

            //    foreach (DataRow rw in dtchoice.Rows)
            //    {
            //        cnt++;

            //        if (chfieldname == rw["Choicefield"].ToString().ToLower())
            //        {
            //            chfieldid = Convert.ToString(cnt);
            //            chfieldname = rw["Choicefield"].ToString();
            //            break;
            //        }
            //    }


            //    if (LegalEntityUUID != "" && LegalEntityUUID != null && LegalEntityUUID != "&nbsp;")
            //    {
            //        foreach (DataRow rw in dtLEList.Rows)
            //        {


            //            if (LegalEntityUUID == rw["TaxonomyValue"].ToString().ToLower())
            //            {
            //                LEID = LegalEntityUUID;
            //                LEName = rw["TaxonomyName"].ToString();
            //                break;
            //            }
            //        }
            //    }


            //    if (LegalEntityUUID != "" && LegalEntityUUID != null && LegalEntityUUID != "&nbsp;")
            //    {
            //        taxonomyFieldValueLegalEntity.TermGuid = LEID; // LegalEntity ID
            //        taxonomyFieldValueLegalEntity.Label = LEName;//LegalEntityUUID Name
            //    }
            //    //taxonomyFieldValuePath.TermGuid = iID;// Id from The Correspondence Taxonomy
            //    //taxonomyFieldValuePath.Label = docType;// "Correspondence"  OR "Quarterly Report";
            //    try
            //    {
            //        if (LegalEntityUUID != "" && LegalEntityUUID != null && LegalEntityUUID != "&nbsp;")
            //        {
            //            item["ce9c7b54e9364a2f9eff5e92f7d79d65"] = taxonomyFieldValueLegalEntity;
            //        }
            //        else
            //        {
            //            item["ce9c7b54e9364a2f9eff5e92f7d79d65"] = "";
            //        }
            //        //item["Quarter"] = Quarter; // Quarter Tag
            //        //item["ocac1b27043549bf95a6b3be20a5e5ea"] = taxonomyFieldValuePath; // Correspondence Type
            //        //item["Year"] = year; // Year Field
            //        //  item["ComplianceDocumentType"] = chfieldname;

            //        item["ComplianceDocumentType"] = chfieldname;



            //        item.Update();
            //        docs.Update();
            //        context.ExecuteQuery();
            //    }
            //    catch (Exception Ex)
            //    {
            //        Response.Write("<br/>Error Occured while tagging the File: " + vSourcrFile +
            //                             " to " + URL + "Compliance Documents" + "<br/>" + Ex.Message + ", " + Ex.StackTrace);
            //    }


            //    #endregion
            //}

            //else
            //{
            #region non billing

            currentRunFolder.Files.Add(newFile);

            int count = currentRunFolder.Files.Count;

            currentRunFolder.Update();

            context.ExecuteQuery();

            Microsoft.SharePoint.Client.File upload = currentRunFolder.Files.GetByUrl(newFile.Url);
            context.Load(upload);

            context.Load(upload.ListItemAllFields);
            context.ExecuteQuery();


            Microsoft.SharePoint.Client.ListItem item = upload.ListItemAllFields;
            context.Load(item);
            context.ExecuteQuery();

            FileLink = item["FileRef"].ToString();
            string docType = string.Empty;
            string iID = string.Empty;


            TaxonomyFieldValue taxonomyFieldValuePath = new TaxonomyFieldValue();
            if (BatchType.ToUpper() == "Q" || BatchType.ToUpper() == "M")
            {
                docType = "Quarterly Report";
            }
            else if (BatchType.ToUpper() == "MERGE" && BillingInvoiceId == "")
            {
                //  docType = "Correspondence";
                docType = "Cap Call/Distribution";//added on 07/26/2019
            }

            else if (BatchType.ToUpper() == "MERGE" && BillingInvoiceId != "")
            {
                // docType = "Correspondence";
                docType = "Billing"; // changed to billing 19_9_2019 - sasmit(Basceamp request)
            }


            //Response.Write("biilingflag:" + billingflag);

            foreach (DataRow rw in dtCorrespondenceType.Rows)
            {
                if (docType.ToLower() == rw["DocumentType"].ToString().ToLower())
                {
                    iID = rw["iID"].ToString();
                    break;
                }
            }
            taxonomyFieldValuePath.TermGuid = iID;// Id from The Correspondence Taxonomy
            taxonomyFieldValuePath.Label = docType;// "Correspondence"  OR "Quarterly Report";


            try
            {
                item["ocac1b27043549bf95a6b3be20a5e5ea"] = taxonomyFieldValuePath; // Correspondence Type
                item["Year"] = year;
                item["Quarter"] = Quarter;

                item.Update();
                currentRunFolder.Update();
                context.ExecuteQuery();
            }
            catch (Exception Ex)
            {
                Response.Write("<br/>Error Occured while tagging the File: " + vSourcrFile +
                                     " to " + URL + "Published Documents/" + FolderNAme + "<br/>" + Ex.Message + ", " + Ex.StackTrace);
            }

            #endregion

            //}


        }
        catch (Exception Exx)
        {
            FileLink = "";
            Response.Write("<br/>Error Occured when trying to copy file from: " + vSourcrFile +
                                      " to " + URL + "Published Documents/" + FolderNAme + "<br/>" + Exx.Message + ", " + Exx.StackTrace);
            return "";
        }

        return FileLink;
    }





    #endregion
}