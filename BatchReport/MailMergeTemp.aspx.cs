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
using Spire.Xls;
using System.Drawing;
using System.Data.Common;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.pdf;
//using CrmSdk;
using System.Data;
using System.Data.SqlClient;
using System.Text;

using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Smo;
using Microsoft.SqlServer.Management.Smo.Agent;

using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System.IO.Compression;
using System.Collections.Generic;
using System.Threading;
using Microsoft.IdentityModel.Claims;


public partial class _MailMergeTemp : System.Web.UI.Page
{
    // List<string> lstFile = new List<string>();
    bool bUnify = false;
    int Success = 0;
    Dictionary<string, string> lstFile = new Dictionary<string, string>();
    Dictionary<string, string> lstSuccessFundName = new Dictionary<string, string>();
    string strErrorOccured = string.Empty;
    bool bProceed = true;
    int intResult = 0;
    string strDescription;
    GeneralMethods clsGM = new GeneralMethods();
    DB clsDB = null;
    SqlConnection cn = null;
    public String _dbErrorMsg;
    public string _strTemplateId = string.Empty;
    string MailId = string.Empty;
    string Mailing_Id = string.Empty;
    //ssi_mailrecordstemp objMailRecordsTemp = null;
    //ssi_mailinglist objMailingList = null;
    //ssi_billinginvoice objbillingInvoice = null;

    // old string con = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TransactionLoad_DB;Data Source=SQL01";
    string con = AppLogic.GetParam(AppLogic.ConfigParam.DBTransactions);//"Data Source=sql01;User ID=MPIUser;Initial Catalog=TransactionLoad_DB;Persist Security Info=True;Password=slater6;";
    string DTSFilePath = AppLogic.GetParam(AppLogic.ConfigParam.DTSFilePath);

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            clsDB = new DB();
            mvShowReport.ActiveViewIndex = 0;
            MailType();
            BindFund();
            BindTemplates();
            BindMailingId();
            BindLegalEntity();
            trLegalentity.Style.Add("display", "none");
            DataSet ds = clsDB.getDataSet("SP_S_FUND_LKUP");
            if (ds.Tables.Count > 0)
            {
                DataTable dt = ds.Tables[0];
                if (ViewState["dtFund"] == null)
                {
                    ViewState["dtFund"] = dt;
                }
            }
        }

        lblError.Text = "";
        lblErrortxt.Text = "";
        lblSuccess.Text = "";
    }

    public void BindLegalEntity()
    {
        lstLegalEntity.Items.Clear();

        //string strType = lstType.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstType) + "'"; 
        string sqlstr = "SP_S_LEGAL_ENTITY_LIST";
        clsGM.getListForBindListBox(lstLegalEntity, sqlstr, "LegalEntityName", "LegalEntityNameId");

        lstLegalEntity.Items.Insert(0, "All");
        lstLegalEntity.Items[0].Value = "0";
        lstLegalEntity.SelectedIndex = 0;
    }

    public void MailType()
    {
        //DB clsDB = new DB(); /* Contact Specific mail details */
        //DataSet loDataset = clsDB.getDataSet("SP_S_CONTACT_SPECIFIC");

        string sql = "SP_S_MAIL_LKUP";
        clsGM.getBindDDL(ddlMailType, sql, "ssi_name", "ssi_mailid");
    }

    public void BindFund()
    {
        string sql = "SP_S_FUND_LKUP";
        clsGM.getListForBindListBox(lstFund, sql, "ssi_name", "ssi_FundId");
        //clsGM.getBindDDL(drpHouseHoldReportTitle, sql, "ssi_name", "ssi_FundId");
    }


    public void BindTemplates()
    {
        if (ddlMailId.SelectedValue != "" && ddlMailId.SelectedValue != "0")
        {
            string strMailId = ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0" ? "null" : "'" + ddlMailId.SelectedValue + "'";
            string sql = "SP_S_Template_MailidTemp @ssi_mailidtemp=" + strMailId;//"Select ssi_templateid from ssi_MailRecordsTemp Where Deletionstatecode=0 AND ssi_mailidtemp=" + strMailId;
            DataSet loDataset = clsDB.getDataSet(sql);

            if (loDataset.Tables[0].Rows.Count > 0)
            {
                ddlTemplates.SelectedValue = Convert.ToString(loDataset.Tables[0].Rows[0]["ssi_templateid"]);
                _strTemplateId = ddlTemplates.SelectedValue;
                ddlTemplates.Enabled = false;
                txtWireAsofDate.Enabled = false;
                txtLetterDate.Enabled = false;
                img1.Style.Add("display", "none");
                img2.Style.Add("display", "none");
            }
        }
        else
        {
            string sql = "SP_S_Template";//"Select ssi_name,ssi_templateid from ssi_template Where Deletionstatecode=0";
            clsGM.getBindDDL(ddlTemplates, sql, "ssi_name", "ssi_templateid");
            //ddlTemplates.Items.Insert(0, "All");
            //ddlTemplates.Items[0].Value = "0";
            //ddlTemplates.SelectedIndex = 0;
        }
    }


    public void BindMailingId()
    {
        string sql = "SP_S_MailIdemp_Lkup";//"Select distinct Isnull(ssi_mailidtemp,0) as ssi_mailidtemp   from ssi_MailRecordsTemp Where Deletionstatecode=0 AND ssi_Unifiedflg=0 AND ssi_ApprovedFlg=0";
        clsGM.getListForBindDDL(ddlMailId, sql, "ssi_mailidtemp", "ssi_mailidtemp");
        ddlMailId.Items.Insert(0, "Select");
        ddlMailId.Items[0].Value = "0";
        ddlMailId.SelectedIndex = 0;
    }



    #region OLD CODE COMMENTED CRM2016 UPGRADE
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
    #endregion
    protected void Button1_Click(object sender, EventArgs e)
    {
        lbtnExceptionReport.Visible = false;
        DB clsDB = new DB();
        //string str = Server.MapPath("..//BatchReport//" + DateTime.Now.ToString("dd_MMM_yyyy_hh_ss") + ".pdf");
        //string str1 = Server.MapPath("..//BatchReport//022412015148.pdf");
        //string str2 = Server.MapPath("..//BatchReport//a030212045741.pdf");
        //string[] str3 = new string[2];
        //str3[0] = str1;
        //str3[1] = str2;
        //PDFMerge pdfMerge = new PDFMerge();
        //pdfMerge.MergeFiles(str, str3);

        //string test = ddlMailType.SelectedValue;
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
        //string crmServerURL = "http://server:5555/";
        lblmailid.Visible = false;
        ddlMailId.Visible = true;
        //string orgName = "GreshamPartners";
        //string orgName = "Webdev";

        //CrmService service = null;
        IOrganizationService service = null;
        string test = ddlMailType.SelectedValue;

        lblError.Text = "";
        lblErrortxt.Text = "";
        lblSuccess.Text = "";
        DataSet loInvoiceData = null;
        try
        {
            //  service = GetCrmService(crmServerUrl, orgName);
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

        #region OLD CODE CRM2016 UPGRADE -Commented
        //try
        //{
        //    //Response.Write(service.Url);
        //    service.PreAuthenticate = true;
        //    service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        //}
        //catch (NullReferenceException ne)
        //{
        //    //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        //}
        #endregion

        #region not in used


        //ssi_mailrecords objMailRecords = null;
        //invoice objInvoice = null;


        //        #region Mail Records Temp Update

        //        string query = @"select b.Ssi_billingId, ms.Ssi_mailinglistId
        //from Ssi_billing b
        //join Ssi_mailinglist ms on b.Ssi_contactid = ms.Ssi_MailingContactsId
        //where ms.Ssi_MailPieceNameIdName = 'Billing'";

        //        DataSet dsupdate = clsDB.getDataSet(query);


        //        for (int j = 0; j < dsupdate.Tables[0].Rows.Count; j++)
        //        {

        //            try
        //            {
        //                //objMailRecordsTemp = new ssi_mailrecordstemp();
        //                Entity objMailingList = new Entity("ssi_mailinglist");

        //                if (Convert.ToString(dsupdate.Tables[0].Rows[j]["Ssi_mailinglistId"]) != "")
        //                {
        //                    //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
        //                    //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(MailRecordsDataset1.Tables[0].Rows[j]["ssi_mailrecordstempid"]));
        //                    objMailingList["ssi_mailinglistid"] = new Guid(Convert.ToString(dsupdate.Tables[0].Rows[j]["Ssi_mailinglistId"]));

        //                    //objMailRecordsTemp.ssi_batchstatus = new Picklist();
        //                    //objMailRecordsTemp.ssi_batchstatus.Value = 1;//Batched


        //                    objMailingList["ssi_billingid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_billing", new Guid(Convert.ToString(dsupdate.Tables[0].Rows[j]["Ssi_billingId"])));



        //                    //objMailRecordsTemp.ssi_batchidtxt = Convert.ToString(DSBatch.Tables[0].Rows[0]["ssi_BatchID"]);

        //                    service.Update(objMailingList);
        //                }

        //            }


        //            catch (Exception ex)
        //            {

        //            }
        //        }

        //        #endregion


        #endregion

        try
        {
            string Mail_type = null;
            clsDB = new DB();
            DataSet MailType = new DataSet();
            object ssi_mailid = ddlMailType.SelectedValue == "" ? "null" : "'" + ddlMailType.SelectedValue + "'";
            MailType = clsDB.getDataSet("SP_S_MAILTYPE @ssi_mailid=" + ssi_mailid);


            for (int i = 0; i < MailType.Tables[0].Rows.Count; i++)
            {
                Mail_type = Convert.ToString(MailType.Tables[0].Rows[i]["ssi_mailtype"]);
            }


            #region Unification Logic Here


            if (chkUnify.Checked == true)
            {
                if (ddlMailId.SelectedValue != "" && ddlMailId.SelectedValue != "0")
                {
                    MailId = ddlMailId.SelectedValue;
                }
                else
                {
                    string strsql = "SP_S_MailIdemp_Max";//" Select max(IsNull(ssi_mailidtemp,0)) + 1  as ssi_mailidtemp  from ssi_MailRecordsTemp Where Deletionstatecode=0 AND ssi_Unifiedflg=0";
                    clsDB = new DB();
                    DataSet loDataset = clsDB.getDataSet(strsql);
                    if ((ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0") && Convert.ToString(loDataset.Tables[0].Rows[0]["ssi_mailidtemp"]) == "")
                    {
                        MailId = "1";
                    }
                    else
                    {
                        MailId = Convert.ToString(loDataset.Tables[0].Rows[0]["ssi_mailidtemp"]);
                    }
                }
            }
            else if (chkUnify.Checked == false)
            {
                if (ddlMailId.SelectedValue != "" && ddlMailId.SelectedValue != "0")
                {
                    MailId = ddlMailId.SelectedValue;
                }
                else
                {
                    string strsql = "SP_S_MailIdemp_Max";//" Select max(IsNull(ssi_mailidtemp,0)) + 1  as ssi_mailidtemp  from ssi_MailRecordsTemp Where Deletionstatecode=0 AND ssi_Unifiedflg=0";
                    clsDB = new DB();
                    DataSet loDataset = clsDB.getDataSet(strsql);
                    if ((ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0") && Convert.ToString(loDataset.Tables[0].Rows[0]["ssi_mailidtemp"]) == "")
                    {
                        MailId = "1";
                    }
                    else
                    {
                        MailId = Convert.ToString(loDataset.Tables[0].Rows[0]["ssi_mailidtemp"]);
                    }
                }
            }


            #endregion

            if (ddlMailType.SelectedValue == "3fb190d9-b2cd-e011-a19b-0019b9e7ee05")//3FB190D9-B2CD-E011-A19B-0019B9E7EE05 --Billing 
            {
                #region Old Logic Commented

                //#region Billing Specific

                //DataSet BillingData = LoadDataSet("SP_S_BILLING_LOG");

                //for (int m = 0; m < BillingData.Tables[0].Rows.Count; m++)
                //{
                //    ssi_billing objBilling = new ssi_billing();

                //    objBilling.ssi_billingid = new Key();
                //    objBilling.ssi_billingid.Value = Guid.NewGuid();

                //    if (Convert.ToString(BillingData.Tables[0].Rows[m]["BillingId"]) != "")
                //    {
                //        objBilling.ssi_billingid_billing = new CrmNumber();
                //        objBilling.ssi_billingid_billing.Value = Convert.ToInt32(BillingData.Tables[0].Rows[m]["BillingId"]);
                //    }

                //    if (Convert.ToString(BillingData.Tables[0].Rows[m]["Contactid"]) != "")
                //    {
                //        objBilling.ssi_contactid = new Lookup();
                //        objBilling.ssi_contactid.type = EntityName.contact.ToString();
                //        objBilling.ssi_contactid.Value = new Guid(Convert.ToString(BillingData.Tables[0].Rows[m]["Contactid"]));
                //    }

                //    if (Convert.ToString(BillingData.Tables[0].Rows[m]["AccountId"]) != "")
                //    {
                //        objBilling.ssi_householdid = new Lookup();
                //        objBilling.ssi_householdid.type = EntityName.account.ToString();
                //        objBilling.ssi_householdid.Value = new Guid(Convert.ToString(BillingData.Tables[0].Rows[m]["AccountId"]));
                //    }

                //    if (Convert.ToString(BillingData.Tables[0].Rows[m]["ssi_LegalEntityId"]) != "")
                //    {
                //        objBilling.ssi_legalentityid = new Lookup();
                //        objBilling.ssi_legalentityid.type = EntityName.ssi_legalentity.ToString();
                //        objBilling.ssi_legalentityid.Value = new Guid(Convert.ToString(BillingData.Tables[0].Rows[m]["ssi_LegalEntityId"]));
                //    }

                //    if (Convert.ToString(BillingData.Tables[0].Rows[m]["ss_Accountid"]) != "")
                //    {
                //        objBilling.ssi_billingaccountid = new Lookup();
                //        objBilling.ssi_billingaccountid.type = EntityName.ssi_account.ToString();
                //        objBilling.ssi_billingaccountid.Value = new Guid(Convert.ToString(BillingData.Tables[0].Rows[m]["ss_Accountid"]));
                //    }

                //    if (Convert.ToString(BillingData.Tables[0].Rows[m]["Fee Type"]) != "")
                //    {
                //        objBilling.ssi_feetype_billing = new Picklist();
                //        objBilling.ssi_feetype_billing.Value = Convert.ToInt32(BillingData.Tables[0].Rows[m]["Fee Type"]);
                //    }

                //    if (Convert.ToString(BillingData.Tables[0].Rows[m]["Billing Note"]) != "")
                //    {
                //        objBilling.ssi_billingnote_billing = Convert.ToString(BillingData.Tables[0].Rows[m]["Billing Note"]);
                //    }

                //    if (Convert.ToString(BillingData.Tables[0].Rows[m]["Custodian Account"]) != "")
                //    {
                //        objBilling.ssi_custodianaccount = Convert.ToString(BillingData.Tables[0].Rows[m]["Custodian Account"]);
                //    }

                //    if (Convert.ToString(BillingData.Tables[0].Rows[m]["Billing Method"]) != "")
                //    {
                //        objBilling.ssi_billingmethod_billing = new Picklist();
                //        objBilling.ssi_billingmethod_billing.Value = Convert.ToInt32(BillingData.Tables[0].Rows[m]["Billing Method"]);
                //    }

                //    service.Create(objBilling);
                //}

                //#endregion

                //#region Add Invoice
                //if (FileUpload1.HasFile == true)
                //{
                //    if (System.IO.Path.GetExtension(FileUpload1.FileName) == ".xls")
                //    {

                //        if (Request.Url.AbsoluteUri.Contains("localhost"))
                //        {
                //            FileUpload1.PostedFile.SaveAs(@"C:\\Reports\\" + FileUpload1.FileName);

                //            if (File.Exists(@"C:\\Reports\\" + FileUpload1.FileName))
                //            {
                //                File.Delete(@"C:\\Reports\\Invoice.xls");
                //                FileUpload1.PostedFile.SaveAs(@"C:\\Reports\\" + FileUpload1.FileName);
                //                File.Move(@"C:\\Reports\\" + FileUpload1.FileName, @"C:\\Reports\\Invoice.xls");
                //            }
                //        }
                //        else
                //        {
                //            string extension = System.IO.Path.GetExtension(FileUpload1.FileName);
                //            //Response.Write("FileUpload1.FileName:" + FileUpload1.FileName + "<br/><br/><br/>");
                //            string strFileName = "Invoice" + extension;
                //            //Response.Write("New FileName:" + strFileName + "<br/><br/><br/>");
                //            FileUpload1.PostedFile.SaveAs("\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + strFileName);

                //            if (File.Exists("\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + strFileName))
                //            {
                //                File.Delete("\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\Invoice.xls");
                //                FileUpload1.PostedFile.SaveAs("\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + strFileName);
                //                File.Move("\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + strFileName, "\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\Invoice.xls");
                //            }
                //            //FileUpload1.PostedFile.SaveAs("\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + FileUpload1.FileName);

                //            //if (File.Exists("\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + FileUpload1.FileName))
                //            //{
                //            //    File.Delete("\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\Invoice.xls");

                //            //    FileUpload1.PostedFile.SaveAs("\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + FileUpload1.FileName);
                //            //    File.Move("\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + FileUpload1.FileName, "\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\Invoice.xls");
                //            //}
                //        }
                //    }
                //    //"Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=TransactionLoad_DB;Data Source=sql01";
                //    SqlConnection sqlconn = null;
                //    try
                //    {
                //        int retVal = 1;
                //        try
                //        {
                //            using (SqlConnection connection3 = new SqlConnection(con))
                //            {
                //                DateTime time;
                //                ServerConnection serverConnection = new ServerConnection(connection3);
                //                Server server = new Server(serverConnection);
                //                Job job = server.JobServer.Jobs["InvoiceUpload"];
                //                JobHistoryFilter filter = new JobHistoryFilter();
                //                filter.JobName = "InvoiceUpload";
                //                time = time = job.LastRunDate;
                //                job.Start();
                //                while (time == job.LastRunDate)
                //                {
                //                    job.Refresh();
                //                }
                //                if (job.LastRunOutcome == CompletionResult.Succeeded)
                //                {
                //                    retVal = 0;
                //                }
                //                else
                //                {
                //                    retVal = 1;
                //                }
                //            }
                //        }
                //        catch (Exception exception3)
                //        {
                //            lblError.Text = "InvoiceUpload Load Job Failed to Execute." + exception3.Message;
                //        }
                //        /*
                //        //string con = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TransactionLoad_DB;Data Source=SQL01";
                //        sqlconn = new SqlConnection(con);

                //        string strsql = "SP_DTS_RunPackage_InvoiceUpload";
                //        SqlCommand cmd = new SqlCommand();

                //        SqlParameter TypeId = cmd.Parameters.Add("@TypeId", SqlDbType.Int);
                //        TypeId.Value = 1;

                //        SqlParameter returncode = cmd.Parameters.Add("@returncode", SqlDbType.Int);
                //        returncode.Direction = ParameterDirection.Output;

                //        //cmd.Parameters["@returncode"].Direction = ParameterDirection.Output;
                //        cmd.CommandText = strsql;
                //        cmd.Connection = sqlconn;
                //        cmd.CommandType = CommandType.StoredProcedure;

                //        sqlconn.Open();
                //        //Response.Write(sqlconn.State + "<br/><br/><br/>");
                //        int result = cmd.ExecuteNonQuery();
                //        System.Threading.Thread.Sleep(1000);

                //        int retVal = (int)cmd.Parameters["@returncode"].Value;

                //        //Response.Write("ret value:" + retVal.ToString());
                //        */
                //        if (retVal == 0)
                //        {
                //            //Response.Write(retVal);
                //            string sql = "SP_S_INVOICE @Date='" + txtAsofdate.Text + "'";
                //            loInvoiceData = LoadDataSet(sql);
                //            //Response.Write("data count:" + loInvoiceData.Tables[0].Rows.Count.ToString());

                //            #region Load Invoice Data

                //            string InvoiceLoadId = Guid.NewGuid().ToString();

                //            for (int i = 0; i < loInvoiceData.Tables[0].Rows.Count; i++)
                //            {
                //                ssi_billinginvoice objBillingInvoice = new ssi_billinginvoice();

                //                objBillingInvoice.ssi_invoiceloadid = InvoiceLoadId;

                //                objBillingInvoice.ssi_billinginvoiceid = new Key();
                //                objBillingInvoice.ssi_billinginvoiceid.Value = Guid.NewGuid();

                //                //objBillingInvoice.ssi_name = "test";

                //                //ssi_billingid
                //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["ssi_billingId"]) != "")
                //                {
                //                    objBillingInvoice.ssi_billingprimaryid = new Lookup();
                //                    objBillingInvoice.ssi_billingprimaryid.type = EntityName.ssi_billing.ToString();
                //                    objBillingInvoice.ssi_billingprimaryid.Value = new Guid(Convert.ToString(loInvoiceData.Tables[0].Rows[i]["ssi_billingId"]));
                //                }

                //                //InvoiceId
                //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["InvoiceId"]) != "")
                //                {
                //                    objBillingInvoice.ssi_invoiceid = Convert.ToString(loInvoiceData.Tables[0].Rows[i]["InvoiceId"]);
                //                }


                //                //BillingId
                //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Billing Id"]) != "")
                //                {
                //                    objBillingInvoice.ssi_billingid = new CrmNumber();
                //                    objBillingInvoice.ssi_billingid.Value = Convert.ToInt32(loInvoiceData.Tables[0].Rows[i]["Billing Id"]);
                //                    //Response.Write(objBillingInvoice.ssi_billingid.Value + "<br/><br/><br/>");
                //                }



                //                //Fee Rate
                //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Fee Rate"]) != "")
                //                {
                //                    objBillingInvoice.ssi_feerate = new CrmDecimal();
                //                    objBillingInvoice.ssi_feerate.Value = Convert.ToDecimal(Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Fee Rate"]));
                //                }

                //                //AUM
                //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["AUM"]) != "")
                //                {
                //                    objBillingInvoice.ssi_aum = new CrmMoney();
                //                    objBillingInvoice.ssi_aum.Value = Convert.ToDecimal(Convert.ToString(loInvoiceData.Tables[0].Rows[i]["AUM"]));
                //                }


                //                //Annual Fee
                //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Annual"]) != "")
                //                {
                //                    objBillingInvoice.ssi_annualfee = new CrmMoney();
                //                    objBillingInvoice.ssi_annualfee.Value = Convert.ToDecimal(Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Annual"]));
                //                }

                //                //Quarterly Fee
                //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Quarterly Fee"]) != "")
                //                {
                //                    objBillingInvoice.ssi_quarterlyfee = new CrmMoney();
                //                    objBillingInvoice.ssi_quarterlyfee.Value = Convert.ToDecimal(Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Quarterly Fee"]));
                //                }


                //                System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");
                //                int Month;
                //                if (ddlMonths.SelectedValue != "")
                //                {
                //                    if (ddlMonths.SelectedValue == "1")
                //                    {
                //                        Month = 2;//Starting from february

                //                        //Month 1
                //                        String MonthName1 = enUS.DateTimeFormat.GetMonthName(Month);
                //                        objBillingInvoice.ssi_month1 = MonthName1;

                //                        //Month 2
                //                        String MonthName2 = enUS.DateTimeFormat.GetMonthName(Month + 1);
                //                        objBillingInvoice.ssi_month2 = MonthName2;

                //                        //Month 3
                //                        String MonthName3 = enUS.DateTimeFormat.GetMonthName(Month + 2);
                //                        objBillingInvoice.ssi_month3 = MonthName3;
                //                    }
                //                    else if (ddlMonths.SelectedValue == "2")
                //                    {
                //                        Month = 5;//Starting from may

                //                        //Month 1
                //                        String MonthName1 = enUS.DateTimeFormat.GetMonthName(Month);
                //                        objBillingInvoice.ssi_month1 = MonthName1;


                //                        //Month 2
                //                        String MonthName2 = enUS.DateTimeFormat.GetMonthName(Month + 1);
                //                        objBillingInvoice.ssi_month2 = MonthName2;

                //                        //Month 3
                //                        String MonthName3 = enUS.DateTimeFormat.GetMonthName(Month + 2);
                //                        objBillingInvoice.ssi_month3 = MonthName3;
                //                    }
                //                    else if (ddlMonths.SelectedValue == "3")
                //                    {
                //                        Month = 8;//Starting from August

                //                        //Month 1
                //                        String MonthName1 = enUS.DateTimeFormat.GetMonthName(Month);
                //                        objBillingInvoice.ssi_month1 = MonthName1;


                //                        //Month 2
                //                        String MonthName2 = enUS.DateTimeFormat.GetMonthName(Month + 1);
                //                        objBillingInvoice.ssi_month2 = MonthName2;

                //                        //Month 3
                //                        String MonthName3 = enUS.DateTimeFormat.GetMonthName(Month + 2);
                //                        objBillingInvoice.ssi_month3 = MonthName3;
                //                    }
                //                    else if (ddlMonths.SelectedValue == "4")
                //                    {
                //                        Month = 11;//Starting from November

                //                        //Month 1
                //                        String MonthName1 = enUS.DateTimeFormat.GetMonthName(Month);
                //                        objBillingInvoice.ssi_month1 = MonthName1;


                //                        //Month 2
                //                        String MonthName2 = enUS.DateTimeFormat.GetMonthName(Month + 1);
                //                        objBillingInvoice.ssi_month2 = MonthName2;

                //                        //Month 3
                //                        String MonthName3 = enUS.DateTimeFormat.GetMonthName(Month + 2);
                //                        objBillingInvoice.ssi_month3 = MonthName3;
                //                    }
                //                }


                //                //Month1 Fee
                //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Month1 Fee"]) != "")
                //                {
                //                    objBillingInvoice.ssi_month1fee = new CrmMoney();
                //                    objBillingInvoice.ssi_month1fee.Value = Convert.ToDecimal(Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Month1 Fee"]));
                //                }

                //                //Month2 Fee
                //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Month2 Fee"]) != "")
                //                {
                //                    objBillingInvoice.ssi_month2fee = new CrmMoney();
                //                    objBillingInvoice.ssi_month2fee.Value = Convert.ToDecimal(Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Month2 Fee"]));
                //                }



                //                //Month3 Fee
                //                if (Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Month3 Fee"]) != "")
                //                {
                //                    objBillingInvoice.ssi_month3fee = new CrmMoney();
                //                    objBillingInvoice.ssi_month3fee.Value = Convert.ToDecimal(Convert.ToString(loInvoiceData.Tables[0].Rows[i]["Month3 Fee"]));
                //                }
                //                //*/

                //                if (txtAsofdate.Text != "")
                //                {
                //                    objBillingInvoice.ssi_invoicedate = new CrmDateTime();
                //                    objBillingInvoice.ssi_invoicedate.Value = txtAsofdate.Text;
                //                }



                //                service.Create(objBillingInvoice);
                //                intResult++;
                //            }
                //            #endregion

                //            //Response.Write(loInvoiceData.Tables[0].Rows.Count);

                //            #region Contact Specific

                //            clsDB = new DB();
                //            DataSet loDataset;
                //            /* get invoice and billing related specific information */
                //            if (chkEmailRecipients.Checked == true)
                //            {
                //                loDataset = clsDB.getDataSet("SP_S_CONTACT_SPECIFIC @IncEmailRecipients=1");
                //            }
                //            else
                //            {
                //                loDataset = clsDB.getDataSet("SP_S_CONTACT_SPECIFIC @IncEmailRecipients=0");
                //            }
                //            //Response.Write("data count:" + loDataset.Tables[0].Rows.Count.ToString() + "<br/>");
                //            for (int i = 0; i < loDataset.Tables[0].Rows.Count; i++)
                //            {
                //                objMailRecords = new ssi_mailrecords();

                //                //Mail Type
                //                objMailRecords.ssi_mailtypeid = new Lookup();
                //                objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();//3FB190D9-B2CD-E011-A19B-0019B9E7EE05
                //                objMailRecords.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue); //new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_mailid"]));

                //                //[Spouse Name]
                //                //First Name
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]) != "")
                //                {
                //                    objMailRecords.ssi_spousepart_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]);
                //                }


                //                //First Name
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["FirstName"]) != "")
                //                {
                //                    objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["FirstName"]);
                //                }


                //                //Last Name
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["LastName"]) != "")
                //                {
                //                    objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["LastName"]);
                //                }

                //                //House Hold
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["HouseHold"]) != "")
                //                {
                //                    objMailRecords.ssi_hholdinst_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["HouseHold"]);
                //                }


                //                //Contact
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Contact"]) != "")
                //                {
                //                    objMailRecords.ssi_fullname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact"]);
                //                }

                //                //Address Line 1
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["AddressLine1"]) != "")
                //                {
                //                    objMailRecords.ssi_addressline1_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["AddressLine1"]);
                //                }
                //                else
                //                {
                //                    objMailRecords.ssi_addressline1_mail = "";
                //                }

                //                //Address Line 2
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["AddressLine2"]) != "")
                //                {
                //                    objMailRecords.ssi_addressline2_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["AddressLine2"]);
                //                }
                //                else
                //                {
                //                    objMailRecords.ssi_addressline2_mail = "";
                //                }

                //                //Address Line 3
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["AddressLine3"]) != "")
                //                {
                //                    objMailRecords.ssi_addressline3_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["AddressLine3"]);
                //                }
                //                else
                //                {
                //                    objMailRecords.ssi_addressline3_mail = "";
                //                }

                //                //City
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["City"]) != "")
                //                {
                //                    objMailRecords.ssi_city_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["City"]);
                //                }

                //                //State Or Province
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["State Or Province"]) != "")
                //                {
                //                    objMailRecords.ssi_stateprovince_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["State Or Province"]);
                //                }


                //                //Zip Code
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Zip Code"]) != "")
                //                {
                //                    objMailRecords.ssi_zipcode_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Zip Code"]);
                //                }

                //                //Country Or Region
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Country Or Region"]) != "")
                //                {
                //                    objMailRecords.ssi_countryregion_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Country Or Region"]);
                //                }

                //                //Dear
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]) != "")
                //                {
                //                    objMailRecords.ssi_dear_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]);
                //                }

                //                //Mail Preference
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]) != "")
                //                {
                //                    objMailRecords.ssi_mailpreference_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]);
                //                }

                //                //Salutation
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]) != "")
                //                {
                //                    objMailRecords.ssi_salutation_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]);
                //                }


                //                if (txtAsofdate.Text != "")
                //                {
                //                    objMailRecords.ssi_asofdate = new CrmDateTime();
                //                    objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                //                }


                //                if (txtAsofdate.Text != "")
                //                {
                //                    //Month1
                //                    DateTime AsOfDate = Convert.ToDateTime(txtAsofdate.Text);

                //                    int Month1 = AsOfDate.Month;
                //                    System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");
                //                    String MonthName1 = enUS.DateTimeFormat.GetMonthName(Month1);
                //                    objMailRecords.ssi_month1 = MonthName1;

                //                    int Month2 = AsOfDate.AddMonths(1).Month;

                //                    String MonthName2 = enUS.DateTimeFormat.GetMonthName(Month2);
                //                    objMailRecords.ssi_month2 = MonthName2;

                //                    int Month3 = AsOfDate.AddMonths(2).Month;

                //                    String MonthName3 = enUS.DateTimeFormat.GetMonthName(Month3);
                //                    objMailRecords.ssi_month3 = MonthName3;

                //                }

                //                //Month1Fee
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Month1Fee"]) != "")
                //                {
                //                    objMailRecords.ssi_month1fee = Convert.ToString(loDataset.Tables[0].Rows[i]["Month1Fee"]);
                //                }

                //                //Month2Fee
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Month2Fee"]) != "")
                //                {
                //                    objMailRecords.ssi_month2fee = Convert.ToString(loDataset.Tables[0].Rows[i]["Month2Fee"]);
                //                }

                //                //Month3Fee
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Month3Fee"]) != "")
                //                {
                //                    objMailRecords.ssi_month3fee = Convert.ToString(loDataset.Tables[0].Rows[i]["Month3Fee"]);
                //                }

                //                //Mailing ID
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_MailingID"]) != "")
                //                {
                //                    objMailRecords.ssi_mailingid = new CrmNumber();
                //                    objMailRecords.ssi_mailingid.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_MailingID"]);
                //                }

                //                ////Full Name
                //                //if (Convert.ToString(loDataset.Tables[0].Rows[i]["FullName"]) != "")
                //                //{
                //                //    objMailRecords.ssi_fullname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["FullName"]);
                //                //}

                //                //Invoice ID
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_InvoiceId"]) != "")
                //                {
                //                    objMailRecords.ssi_invoiceid = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_InvoiceId"]);
                //                }

                //                //Billing Method
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Billing Method"]) != "")
                //                {
                //                    objMailRecords.ssi_billingmethod = Convert.ToString(loDataset.Tables[0].Rows[i]["Billing Method"]);
                //                }

                //                //Bank
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Bank"]) != "")
                //                {
                //                    objMailRecords.ssi_bankid = new Lookup();
                //                    objMailRecords.ssi_bankid.type = EntityName.ssi_account.ToString();
                //                    objMailRecords.ssi_bankid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["Bank"]));
                //                }

                //                //Custodian Account
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account"]) != "")
                //                {
                //                    objMailRecords.ssi_custodianaccount = Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account"]);
                //                }

                //                //Invoiced On
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Invoiced On"]) != "")
                //                {
                //                    objMailRecords.ssi_invoice1 = Convert.ToString(loDataset.Tables[0].Rows[i]["Invoiced On"]);
                //                }

                //                //Mail Type
                //                if (ddlMailType.SelectedValue != "")
                //                {
                //                    objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                //                }

                //                //Billing ID
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Billing ID"]) != "")
                //                {
                //                    objMailRecords.ssi_billingid = Convert.ToString(loDataset.Tables[0].Rows[i]["Billing ID"]);
                //                }

                //                //Name
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]) != "")
                //                {
                //                    if (txtAsofdate.Text != "")
                //                    {
                //                        objMailRecords.ssi_name = txtAsofdate.Text + "-" + Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);
                //                    }
                //                    else
                //                    {
                //                        objMailRecords.ssi_name = Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);
                //                    }
                //                }

                //                //HouseHold lookup
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"]) != "")
                //                {
                //                    objMailRecords.ssi_accountid = new Lookup();
                //                    objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"]));
                //                }

                //                //Contact lookup
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"]) != "")
                //                {
                //                    objMailRecords.ssi_contactfullnameid = new Lookup();
                //                    objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"]));
                //                }

                //                //ssi_LegalEntityId lookup
                //                if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
                //                {
                //                    //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                //                    objMailRecords.ssi_legalentitynameid = new Lookup();
                //                    objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                //                }

                //                // Response.Write("<br/> inserted:" + i.ToString());
                //                // CreatedByCustomid Field 
                //                //Rohit Pawar
                //                string Userid = GetcurrentUser();

                //                if (Userid != "")
                //                {
                //                    objMailRecords.ssi_createdbycustomid = new Lookup();
                //                    objMailRecords.ssi_createdbycustomid.Value = new Guid(Userid);
                //                }

                //                service.Create(objMailRecords);
                //                intResult++;
                //            }

                //            #endregion
                //        }
                //        else if (retVal == -1)
                //        {
                //            lblError.Text = "File Upload failed, Error detail: DTS Failed";
                //        }

                //    }
                //    catch (System.Web.Services.Protocols.SoapException exc1)
                //    {

                //        Response.Write("<br/>Exception: " + exc1.Detail.InnerText);

                //    }
                //    catch (Exception exc)
                //    {
                //        Response.Write(exc.Message + exc.StackTrace);
                //    }
                //    finally
                //    {
                //        if (sqlconn != null)
                //            if (sqlconn.State != System.Data.ConnectionState.Open)
                //                sqlconn.Close();
                //    }

                //}

                //#endregion

                #endregion

                #region update invoice

                clsDB = new DB();

                DataSet dsUpdateInvoice = clsDB.getDataSet("SP_U_BILLINGINVOICEDATE @AsOfDate='" + txtAsofdate.Text + "',@LetterDate='" + txtLetterDate.Text + "' ");


                for (int j = 0; j < dsUpdateInvoice.Tables[0].Rows.Count; j++)
                {
                    //  objbillingInvoice = new ssi_billinginvoice();
                    Entity objbillingInvoice = new Entity("ssi_billinginvoice");
                    if (Convert.ToString(dsUpdateInvoice.Tables[0].Rows[j]["Ssi_billinginvoiceId"]) != "")
                    {
                        //objbillingInvoice.ssi_billinginvoiceid = new Key();
                        //objbillingInvoice.ssi_billinginvoiceid.Value = new Guid(Convert.ToString(dsUpdateInvoice.Tables[0].Rows[j]["Ssi_billinginvoiceId"]));
                        objbillingInvoice["ssi_billinginvoiceid"] = new Guid(Convert.ToString(dsUpdateInvoice.Tables[0].Rows[j]["Ssi_billinginvoiceId"]));

                    }

                    if (Convert.ToString(dsUpdateInvoice.Tables[0].Rows[j]["Ssi_InvoiceDate"]) != "")
                    {
                        //objbillingInvoice.ssi_invoicedate = new CrmDateTime();
                        //objbillingInvoice.ssi_invoicedate.Value = Convert.ToString(dsUpdateInvoice.Tables[0].Rows[j]["Ssi_InvoiceDate"]);
                        objbillingInvoice["ssi_invoicedate"] = Convert.ToDateTime(dsUpdateInvoice.Tables[0].Rows[j]["Ssi_InvoiceDate"]);

                    }

                    //  service.Update(objbillingInvoice);
                   //commented 4_24_2020 Exception Report
                    // intResult++;
                }


                #endregion

                #region Billing Specific


                string strsql = "SP_S_MailIdemp_Max";//" Select max(IsNull(ssi_mailidtemp,0)) + 1  as ssi_mailidtemp  from ssi_MailRecordsTemp Where Deletionstatecode=0 AND ssi_Unifiedflg=0";

                DataSet loDataset = clsDB.getDataSet(strsql);
                if ((ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0") && Convert.ToString(loDataset.Tables[0].Rows[0]["ssi_mailidtemp"]) == "")
                {
                    MailId = "1";
                }
                else
                {
                    MailId = Convert.ToString(loDataset.Tables[0].Rows[0]["ssi_mailidtemp"]);
                }


                clsDB = new DB();

                DataSet dsBilling = clsDB.getDataSet("SP_S_BILLING @AsOfDate='" + txtAsofdate.Text + "',@LetterDate='" + txtLetterDate.Text + "' ");
                for (int i = 0; i < dsBilling.Tables[0].Rows.Count; i++)
                {
                    //objMailRecordsTemp = new ssi_mailrecordstemp();
                    Entity objMailRecordsTemp = new Entity("ssi_mailrecordstemp");
                    //Mail Type
                    //objMailRecordsTemp.ssi_mailtypeid = new Lookup();
                    //objMailRecordsTemp.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();//3FB190D9-B2CD-E011-A19B-0019B9E7EE05
                    //objMailRecordsTemp.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue); //new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_mailid"]));
                    objMailRecordsTemp["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));


                    //objMailRecordsTemp.ssi_templateid = new Lookup();
                    //objMailRecordsTemp.ssi_templateid.type = EntityName.ssi_template.ToString();
                    ////objMailRecordsTemp.ssi_templateid.Value = new Guid("73709E5B-849A-E511-9416-005056A0099E"); //billing
                    //objMailRecordsTemp.ssi_templateid.Value = new Guid("CD75AE88-D5A4-E511-9418-005056A0567E"); //billing
                    objMailRecordsTemp["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(Convert.ToString("CD75AE88-D5A4-E511-9418-005056A0567E")));


                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_AUMasofdate"]) != "")
                    {
                        //objMailRecordsTemp.ssi_aumasofdate = new CrmDateTime();
                        //objMailRecordsTemp.ssi_aumasofdate.Value = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_AUMasofdate"]);
                        objMailRecordsTemp["ssi_aumasofdate"] = Convert.ToDateTime(dsBilling.Tables[0].Rows[i]["ssi_AUMasofdate"]);

                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_AUMasofdate"]) != "")
                    {
                        //objMailRecordsTemp.ssi_asofdate = new CrmDateTime();
                        //objMailRecordsTemp.ssi_asofdate.Value = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_AUMasofdate"]);
                        objMailRecordsTemp["ssi_asofdate"] = Convert.ToDateTime(dsBilling.Tables[0].Rows[i]["ssi_AUMasofdate"]);

                    }


                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_AUM"]) != "")
                    {
                        // objMailRecordsTemp.ssi_aum = Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_AUM"]);
                        objMailRecordsTemp["ssi_aum"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_AUM"]);
                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_AnnualFee"]) != "")
                    {
                        //objMailRecordsTemp.ssi_annualfeebilling = new CrmMoney();
                        //objMailRecordsTemp.ssi_annualfeebilling.Value = Convert.ToDecimal(Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_AnnualFee"]));
                        objMailRecordsTemp["ssi_annualfeebilling"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["Ssi_AnnualFee"]));

                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_InvoiceDate"]) != "")
                    {
                        //objMailRecordsTemp.ssi_invoicedatebilling = new CrmDateTime();
                        //objMailRecordsTemp.ssi_invoicedatebilling.Value = Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_InvoiceDate"]);
                        objMailRecordsTemp["ssi_invoicedatebilling"] = Convert.ToDateTime(dsBilling.Tables[0].Rows[i]["Ssi_InvoiceDate"]);

                    }


                    //commented on 08/27/2019 as letter date not populated when we do billing
                    //if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_InvoiceDate"]) != "")
                    //{
                    //    //objMailRecordsTemp.ssi_letterdate = new CrmDateTime();
                    //    //objMailRecordsTemp.ssi_letterdate.Value = Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_InvoiceDate"]);
                    //    objMailRecordsTemp["ssi_letterdate"] = Convert.ToDateTime(dsBilling.Tables[0].Rows[i]["Ssi_InvoiceDate"]);

                    //}
                    //added on 04/16/2020 as discussed & confirm by sudeep 
                    if (txtLetterDate.Text != "")
                    {
                        //objMailRecordsTemp.ssi_asofdate = new CrmDateTime();
                        //objMailRecordsTemp.ssi_asofdate.Value = txtLetterDate.Text;
                        objMailRecordsTemp["ssi_letterdate"] = Convert.ToDateTime(txtLetterDate.Text);

                    }
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_FeeRate"]) != "")
                    {
                        //objMailRecordsTemp.ssi_feeratebilling = new CrmFloat();
                        //objMailRecordsTemp.ssi_feeratebilling.Value = Convert.ToDouble(Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_FeeRate"]));
                        objMailRecordsTemp["ssi_feeratebilling"] = Convert.ToDouble(Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_FeeRate"]));
                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_BillingID"]) != "")
                    {
                        //objMailRecordsTemp.ssi_billingidbilling = new CrmNumber();
                        //objMailRecordsTemp.ssi_billingidbilling.Value = Convert.ToInt32(dsBilling.Tables[0].Rows[i]["Ssi_BillingID"].ToString());
                        objMailRecordsTemp["ssi_billingidbilling"] = Convert.ToInt32(dsBilling.Tables[0].Rows[i]["Ssi_BillingID"]);

                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_quarterlyfee"]) != "")
                    {
                        //objMailRecordsTemp.ssi_quarterlyfeebilling = new CrmMoney();
                        //objMailRecordsTemp.ssi_quarterlyfeebilling.Value = Convert.ToDecimal(Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_quarterlyfee"]));
                        objMailRecordsTemp["ssi_quarterlyfeebilling"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_quarterlyfee"]));

                    }
                    //Contact lookup
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ContactId"]) != "")
                    {
                        //objMailRecordsTemp.ssi_contactfullnameid = new Lookup();
                        //objMailRecordsTemp.ssi_contactfullnameid.Value = new Guid(Convert.ToString(dsBilling.Tables[0].Rows[i]["ContactId"]));
                        objMailRecordsTemp["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(dsBilling.Tables[0].Rows[i]["ContactId"])));
                    }


                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ContactFullName"]) != "")
                    {
                        //objMailRecordsTemp.ssi_contactfullname = Convert.ToString(dsBilling.Tables[0].Rows[i]["ContactFullName"]);
                        objMailRecordsTemp["ssi_contactfullname"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ContactFullName"]);
                    }

                    //Address Line 1
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["AddressLine1"]) != "")
                    {
                        // objMailRecordsTemp.ssi_addressline1_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["AddressLine1"]);
                        objMailRecordsTemp["ssi_addressline1_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["AddressLine1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["AddressLine2"]) != "")
                    {
                        // objMailRecordsTemp.ssi_addressline2_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["AddressLine2"]);
                        objMailRecordsTemp["ssi_addressline2_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["AddressLine2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["AddressLine3"]) != "")
                    {
                        //objMailRecordsTemp.ssi_addressline3_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["AddressLine3"]);
                        objMailRecordsTemp["ssi_addressline3_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["AddressLine3"]);
                    }

                    //City
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["City"]) != "")
                    {
                        // objMailRecordsTemp.ssi_city_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["City"]);
                        objMailRecordsTemp["ssi_city_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["City"]);
                    }

                    //State or Province
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["StateOrProvince"]) != "")
                    {
                        //  objMailRecordsTemp.ssi_stateprovince_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["StateOrProvince"]);
                        objMailRecordsTemp["ssi_stateprovince_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["StateOrProvince"]);
                    }

                    //ZIP Code
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ZIPCode"]) != "")
                    {
                        //objMailRecordsTemp.ssi_zipcode_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["ZIPCode"]);
                        objMailRecordsTemp["ssi_zipcode_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ZIPCode"]);
                    }

                    //Country or Region
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["CountryOrRegion"]) != "")
                    {
                        //objMailRecordsTemp.ssi_countryregion_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["CountryOrRegion"]);
                        objMailRecordsTemp["ssi_countryregion_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["CountryOrRegion"]);
                    }

                    //Dear
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Dear"]) != "")
                    {
                        //objMailRecordsTemp.ssi_dear_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["Dear"]);
                        objMailRecordsTemp["ssi_dear_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["Dear"]);
                    }

                    //Salutation
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Salutation"]) != "")
                    {
                        // objMailRecordsTemp.ssi_salutation_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["Salutation"]);
                        objMailRecordsTemp["ssi_salutation_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["Salutation"]);
                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
                    {
                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecordsTemp.ssi_legalentitynameid = new Lookup();
                        //objMailRecordsTemp.ssi_legalentitynameid.Value = new Guid(Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                        objMailRecordsTemp["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_LegalEntityId"])));
                    }

                    //Legal Entity name 
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["LegalentityName"]) != "")
                    {
                        //  objMailRecordsTemp.ssi_legalentityname = Convert.ToString(dsBilling.Tables[0].Rows[i]["LegalentityName"]);
                        objMailRecordsTemp["ssi_legalentityname"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["LegalentityName"]);
                    }

                    //Owner First Name
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["OwnerFirstName"]) != "")
                    {
                        //  objMailRecordsTemp.ssi_ownerfirstname_hh_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["OwnerFirstName"]);
                        objMailRecordsTemp["ssi_ownerfirstname_hh_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["OwnerFirstName"]);
                    }

                    //Owner Last Name
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["OwnerLastName"]) != "")
                    {
                        // objMailRecordsTemp.ssi_ownerlname_hh_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["OwnerLastName"]);
                        objMailRecordsTemp["ssi_ownerlname_hh_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["OwnerLastName"]);
                    }
                    //Secondary Owner First Name
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["SecondaryOwnerFirstName"]) != "")
                    {
                        //  objMailRecordsTemp.ssi_secownerfname_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["SecondaryOwnerFirstName"]);
                        objMailRecordsTemp["ssi_secownerfname_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["SecondaryOwnerFirstName"]);
                    }

                    //Secondary Owner Last Name
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["SecondaryOwnerLastName"]) != "")
                    {
                        //objMailRecordsTemp.ssi_secownerlname_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["SecondaryOwnerLastName"]);
                        objMailRecordsTemp["ssi_secownerlname_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["SecondaryOwnerLastName"]);
                    }

                    //Mail Preference 
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["MailPreference"]) != "")
                    {
                        // objMailRecordsTemp.ssi_mailpreference_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["MailPreference"]);
                        objMailRecordsTemp["ssi_mailpreference_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["MailPreference"]);
                    }

                    //Name
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Name"]) != "")
                    {
                        //  objMailRecordsTemp.ssi_name = Convert.ToString(dsBilling.Tables[0].Rows[i]["Name"]);
                        objMailRecordsTemp["ssi_name"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["Name"]);
                    }


                    //ssi_mailStatus 
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_mailStatus"]) != "")
                    {
                        //objMailRecordsTemp.ssi_mailstatus = new Picklist();
                        //objMailRecordsTemp.ssi_mailstatus.Value = Convert.ToInt32(dsBilling.Tables[0].Rows[i]["ssi_mailStatus"]);
                        objMailRecordsTemp["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(dsBilling.Tables[0].Rows[i]["ssi_mailStatus"]));

                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_FileName"]) != "")
                    {
                        //objMailRecordsTemp.ssi_filename = Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_FileName"]);
                        objMailRecordsTemp["ssi_filename"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_FileName"]);
                    }


                    //MailingID
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["MailingID"]) != "")
                    {
                        //objMailRecordsTemp.ssi_mailingid = new CrmNumber();
                        //objMailRecordsTemp.ssi_mailingid.Value = Convert.ToInt32(dsBilling.Tables[0].Rows[i]["MailingID"]);
                        objMailRecordsTemp["ssi_mailingid"] = Convert.ToInt32(dsBilling.Tables[0].Rows[i]["MailingID"]);

                    }

                    if (MailId != "" && MailId != "0")
                    {
                        //objMailRecordsTemp.ssi_mailidtemp = new CrmNumber();
                        //objMailRecordsTemp.ssi_mailidtemp.Value = Convert.ToInt32(MailId);
                        objMailRecordsTemp["ssi_mailidtemp"] = Convert.ToInt32(MailId);

                    }

                    if (Convert.ToString(txtautoDebitDate.Text) != "")
                    {
                        //objMailRecordsTemp.ssi_autodebitdate = new CrmDateTime();
                        //objMailRecordsTemp.ssi_autodebitdate.Value = txtautoDebitDate.Text;
                        objMailRecordsTemp["ssi_autodebitdate"] = Convert.ToDateTime(txtautoDebitDate.Text);

                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_accountid"]) != "")
                    {
                        //objMailRecordsTemp.ssi_accountid = new Lookup();
                        //objMailRecordsTemp.ssi_accountid.Value = new Guid(Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_accountid"]));
                        objMailRecordsTemp["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_accountid"])));
                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_CustomBillingREF"]) != "")
                    {
                        // objMailRecordsTemp.ssi_custombillingref = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_CustomBillingREF"]);
                        objMailRecordsTemp["ssi_custombillingref"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_CustomBillingREF"]);
                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_billingname"]) != "")
                    {
                        //objMailRecordsTemp.ssi_billingname = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_billingname"]);
                        objMailRecordsTemp["ssi_billingname"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_billingname"]);
                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_AdjustedFee"]) != "")
                    {
                        //objMailRecordsTemp.ssi_adjustedfee = new CrmMoney();
                        //objMailRecordsTemp.ssi_adjustedfee.Value = Convert.ToDecimal(Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_AdjustedFee"]));
                        objMailRecordsTemp["ssi_adjustedfee"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_AdjustedFee"]));
                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_Adjustment"]) != "")
                    {
                        //objMailRecordsTemp.ssi_adjustment = new CrmMoney();
                        //objMailRecordsTemp.ssi_adjustment.Value = Convert.ToDecimal(Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_Adjustment"]));
                        objMailRecordsTemp["ssi_adjustment"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_Adjustment"]));
                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_AdjustmentReason"]) != "")
                    {
                        // objMailRecordsTemp.ssi_adjustmentreason = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_AdjustmentReason"]);
                        objMailRecordsTemp["ssi_adjustmentreason"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_AdjustmentReason"]);
                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_RelationshipFee"]) != "")
                    {
                        //objMailRecordsTemp.ssi_relationshipfee = new CrmMoney();
                        //objMailRecordsTemp.ssi_relationshipfee.Value = Convert.ToDecimal(Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_RelationshipFee"]));
                        objMailRecordsTemp["ssi_relationshipfee"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_RelationshipFee"]));
                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_CustomFee"]) != "")
                    {
                        //objMailRecordsTemp.ssi_customfee = new CrmMoney();
                        //objMailRecordsTemp.ssi_customfee.Value = Convert.ToDecimal(Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_CustomFee"]));
                        objMailRecordsTemp["ssi_customfee"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_CustomFee"]));
                    }

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_Discount"]) != "")
                    {
                        //objMailRecordsTemp.ssi_discount = new CrmDecimal();
                        //objMailRecordsTemp.ssi_discount.Value = Convert.ToDecimal(Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_Discount"]));
                        objMailRecordsTemp["ssi_discount"] = Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_Discount"]);

                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        // objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecordsTemp["ssi_mail"] = ddlMailType.SelectedItem.Text;
                    }



                    #region New Field added on 08/08/2019 for billing 


                    // ssi_billinginvoiceid
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_billinginvoiceId"]) != "")
                    {
                        objMailRecordsTemp["ssi_billinginvoiceid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_billinginvoice", new Guid(Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_billinginvoiceId"])));

                    }

                    // ssi_billingprimaryid
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_billingprimaryid"]) != "")
                    {
                        objMailRecordsTemp["ssi_billingprimaryid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_billing", new Guid(Convert.ToString(dsBilling.Tables[0].Rows[i]["Ssi_billingprimaryid"])));

                    }

                    //ssi_feeonfirst25mminbps

                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_feeonfirst25mminbps"]) != "")
                    {

                        objMailRecordsTemp["ssi_feeonfirst25mminbps"] = Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_feeonfirst25mminbps"]);

                    }

                    // ssi_maximumfeeasa
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_maximumfeeasa"]) != "")
                    {
                        objMailRecordsTemp["ssi_maximumfeeasa"] = Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_maximumfeeasa"]);
                    }

                    // ssi_minimumfeein
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_minimumfeein"]) != "")
                    {
                        objMailRecordsTemp["ssi_minimumfeein"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_minimumfeein"]));

                    }

                    //  ssi_totalbillableassets
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_totalbillableassets"]) != "")
                    {
                        objMailRecordsTemp["ssi_totalbillableassets"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_totalbillableassets"]));

                    }

                    //ssi_securityfeeaum
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_securityfeeaum"]) != "")
                    {
                        objMailRecordsTemp["ssi_securityfeeaum"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(dsBilling.Tables[0].Rows[i]["ssi_securityfeeaum"]));

                    }


                    #endregion


                    #region New Field Added on 05_28_2020
                    //[Spouse Name]
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_spousepart_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_spousepart_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["Spouse Name"]);
                        objMailRecordsTemp["ssi_spousepart_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_spousepart_mail"]);
                    }

                    //Ssi_AnzianoID
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_anzianoid"]) != "")
                    {
                        //objMailRecordsTemp.ssi_anzianoid = new CrmNumber();
                        //objMailRecordsTemp.ssi_anzianoid.Value = Convert.ToInt32(dsBilling.Tables[0].Rows[i]["Ssi_AnzianoID"]);
                        objMailRecordsTemp["ssi_anzianoid"] = Convert.ToInt32(dsBilling.Tables[0].Rows[i]["ssi_anzianoid"]);

                    }
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_tnrid_nv"]) != "")
                    {
                        // objMailRecordsTemp.ssi_tnrid_nv = Convert.ToString(dsBilling.Tables[0].Rows[i]["TNR ID"]);
                        objMailRecordsTemp["ssi_tnrid_nv"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_tnrid_nv"]);
                        //Response.Write(objMailRecords.ssi_tnrid_nv);
                    }

                    //Mailing Contact Type
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_mailingcontacttype"]) != "")
                    {
                        // objMailRecordsTemp.ssi_mailingcontacttype = Convert.ToString(dsBilling.Tables[0].Rows[i]["Mailing Contact Type"]);
                        objMailRecordsTemp["ssi_mailingcontacttype"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_mailingcontacttype"]);
                    }
                    //Contact Household
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_hholdinst_mail"]) != "")
                    {
                        // objMailRecordsTemp.ssi_hholdinst_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["Contact Household"]);
                        objMailRecordsTemp["ssi_hholdinst_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_hholdinst_mail"]);
                    }

                    //Contact Full Name
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_fullname_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_fullname_mail = Convert.ToString(dsBilling.Tables[0].Rows[i]["Contact Full Name"]);
                        objMailRecordsTemp["ssi_fullname_mail"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_fullname_mail"]);
                    }
                    //clientportalname 
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_clientportalname"]) != "")
                    {
                        //objMailRecordsTemp.ssi_clientportalname  = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_clientportalname"]);
                        objMailRecordsTemp["ssi_clientportalname"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_clientportalname"]);
                    }
                 
                    //clientreportfolder
                    if (Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_clientreportfolder"]) != "")
                    {
                        //objMailRecordsTemp.ssi_clientreportfolder = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_clientreportfolder"]);
                        objMailRecordsTemp["ssi_clientreportfolder"] = Convert.ToString(dsBilling.Tables[0].Rows[i]["ssi_clientreportfolder"]);
                    }


                    #endregion


                    string Userid = GetcurrentUser();

                    if (Userid != "")
                    {
                        //objMailRecordsTemp.ssi_createdbycustomid = new Lookup();
                        //objMailRecordsTemp.ssi_createdbycustomid.Value = new Guid(Userid);
                        objMailRecordsTemp["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
                    }

                    //objMailRecordsTemp.ssi_updateunifyflg = new CrmBoolean();
                    //objMailRecordsTemp.ssi_updateunifyflg.Value = true;
                    objMailRecordsTemp["ssi_updateunifyflg"] = true;

                    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
                    //objMailRecordsTemp.ssi_unifiedflg.Value = true;
                    objMailRecordsTemp["ssi_unifiedflg"] = true;

                    service.Create(objMailRecordsTemp);
                    intResult++;

                }
                lblmailid.Visible = true;

                if (dsBilling.Tables[0].Rows.Count > 0)
                {
                    lblmailid.Text = MailId;
                    ddlMailId.Visible = false;
                }
                else
                {
                    lblError.Text = "No Records Found";
                }

                #endregion

            }

            #region Quarterly/Annual Review

            /* get Quarterly and Annual related specific information */

            else if (ddlMailType.SelectedValue == "0f4c85f4-d0be-e011-a19b-0019b9e7ee05")
            {
                clsDB = new DB();
                DataSet QuarAnnuDataset = new DataSet();
                if (chkEmailRecipients.Checked == true)
                {

                    QuarAnnuDataset = clsDB.getDataSet("SP_S_QUARTERLY_ANNUAL_REVIEW @IncEmailRecipients=1");
                }
                else
                {
                    QuarAnnuDataset = clsDB.getDataSet("SP_S_QUARTERLY_ANNUAL_REVIEW @IncEmailRecipients=0");
                }



                for (int i = 0; i < QuarAnnuDataset.Tables[0].Rows.Count; i++)
                {

                    // objMailRecords = new ssi_mailrecords();
                    Entity objMailRecords = new Entity("ssi_mailrecords");
                    //Mail Type
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ssi_mailid"]) != "")
                    {
                        //objMailRecords.ssi_mailtypeid = new Lookup();
                        //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                        //objMailRecords.ssi_mailtypeid.Value = new Guid(Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ssi_mailid"]));
                        objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ssi_mailid"])));
                    }


                    //[Spouse Name]
                    //First Name
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Spouse Name"]) != "")
                    {
                        // objMailRecords.ssi_spousepart_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Spouse Name"]);
                        objMailRecords["ssi_spousepart_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Spouse Name"]);
                    }


                    //Name
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Name"]) != "")
                    {
                        if (txtAsofdate.Text != "")
                        {
                            //objMailRecords.ssi_name = txtAsofdate.Text + "-" + Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Name"]);
                            objMailRecords["ssi_name"] = txtAsofdate.Text + "-" + Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Name"]);
                        }
                        else
                        {
                            //objMailRecords.ssi_name = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Name"]);
                            objMailRecords["ssi_name"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Name"]);
                        }
                    }


                    //First Name
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["FirstName"]) != "")
                    {
                        //objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["FirstName"]);
                        objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["FirstName"]);
                    }

                    //Last Name
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["LastName"]) != "")
                    {
                        // objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["LastName"]);
                        objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["LastName"]);
                    }

                    //House Hold
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["HouseHold"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["HouseHold"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["HouseHold"]);
                    }

                    //Contact
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Contact"]) != "")
                    {
                        // objMailRecords.ssi_fullname_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Contact"]);
                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Contact"]);
                    }

                    //Address Line 1
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AddressLine1"]) != "")
                    {
                        //objMailRecords.ssi_addressline1_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AddressLine1"]);
                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AddressLine1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AddressLine2"]) != "")
                    {
                        // objMailRecords.ssi_addressline2_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AddressLine2"]);
                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AddressLine2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AddressLine3"]) != "")
                    {
                        // objMailRecords.ssi_addressline3_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AddressLine3"]);
                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AddressLine3"]);
                    }

                    //City
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["City"]) != "")
                    {
                        // objMailRecords.ssi_city_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["City"]);
                        objMailRecords["ssi_city_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["City"]);
                    }

                    //State Or Province
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["State Or Province"]) != "")
                    {
                        //objMailRecords.ssi_stateprovince_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["State Or Province"]);
                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["State Or Province"]);
                    }


                    //Zip Code
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Zip Code"]) != "")
                    {
                        // objMailRecords.ssi_zipcode_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Zip Code"]);
                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Zip Code"]);
                    }

                    //Country Or Region
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Country Or Region"]) != "")
                    {
                        //objMailRecords.ssi_countryregion_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Country Or Region"]);
                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Country Or Region"]);
                    }

                    //Dear
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Dear"]) != "")
                    {
                        //objMailRecords.ssi_dear_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Dear"]);
                        objMailRecords["ssi_dear_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Dear"]);
                    }

                    //Mail Preference
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Mail Preference"]) != "")
                    {
                        // objMailRecords.ssi_mailpreference_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Mail Preference"]);
                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Mail Preference"]);
                    }

                    //Salutation
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Salutation"]) != "")
                    {
                        // objMailRecords.ssi_salutation_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Salutation"]);
                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Salutation"]);
                    }


                    if (txtAsofdate.Text != "")
                    {
                        //objMailRecords.ssi_asofdate = new CrmDateTime();
                        //objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);

                    }

                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Ssi_MailingID"]) != "")
                    {
                        //objMailRecords.ssi_mailingid = new CrmNumber();
                        //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(QuarAnnuDataset.Tables[0].Rows[i]["Ssi_MailingID"]);
                        objMailRecords["ssi_mailingid"] = Convert.ToInt32(QuarAnnuDataset.Tables[0].Rows[i]["Ssi_MailingID"]);
                        Mailing_Id = Convert.ToInt32(QuarAnnuDataset.Tables[0].Rows[i]["Ssi_MailingID"]).ToString(); // added 6_11_2019
                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        // objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                    }
                    //if (txtAsofdate.Text != "")
                    //{
                    //    //Month1
                    //    DateTime InvoiceDate = Convert.ToDateTime(txtAsofdate.Text);
                    //    int Month1 = InvoiceDate.Month;
                    //    System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");
                    //    String MonthName1 = enUS.DateTimeFormat.GetMonthName(Month1);
                    //    objMailRecords.ssi_month1 = MonthName1;

                    //    //Month2
                    //    //DateTime InvoiceDate = Convert.ToDateTime(txtAsofdate.Text);
                    //    int Month2 = InvoiceDate.Month + 1;
                    //    //System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");
                    //    String MonthName2 = enUS.DateTimeFormat.GetMonthName(Month2);
                    //    objMailRecords.ssi_month2 = MonthName2;

                    //    //Month3
                    //    //DateTime InvoiceDate = Convert.ToDateTime(txtAsofdate.Text);
                    //    int Month3 = InvoiceDate.Month + 2;
                    //    //System.Globalization.CultureInfo enUS = new System.Globalization.CultureInfo("en-US");
                    //    String MonthName3 = enUS.DateTimeFormat.GetMonthName(Month3);
                    //    objMailRecords.ssi_month2 = MonthName3;
                    //}

                    //            //Month1Fee
                    //if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Month1Fee"]) != "")
                    //{
                    //    objMailRecords.ssi_month1fee = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Month1Fee"]);
                    //}

                    //            //Month2Fee
                    //if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Month2Fee"]) != "")
                    //{
                    //    objMailRecords.ssi_month2fee = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Month2Fee"]);
                    //}

                    //            //Month3Fee
                    //if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Month3Fee"]) != "")
                    //{
                    //    objMailRecords.ssi_month3fee = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Month3Fee"]);
                    //}



                    //if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["FullName"]) != "")
                    //{
                    //    objMailRecords.ssi_fullname_mail = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["FullName"]);
                    //}


                    //if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ssi_InvoiceId"]) != "")
                    //{
                    //    objMailRecords.ssi_invoiceid = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ssi_InvoiceId"]);
                    //}

                    //if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Billing Method"]) != "")
                    //{
                    //    objMailRecords.ssi_billingmethod = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Billing Method"]);
                    //}

                    //if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Bank"]) != "")
                    //{
                    //    objMailRecords.ssi_bankid = new Lookup();
                    //    objMailRecords.ssi_bankid.type = EntityName.ssi_account.ToString();
                    //    objMailRecords.ssi_bankid.Value = new Guid(Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Bank"]));
                    //}

                    //if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Custodian Account"]) != "")
                    //{
                    //    objMailRecords.ssi_custodianaccount = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Custodian Account"]);
                    //}

                    //if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Invoiced On"]) != "")
                    //{
                    //    objMailRecords.ssi_invoice1 = Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["Invoiced On"]);
                    //}

                    //HouseHold lookup
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AccountId"]) != "")
                    {
                        //objMailRecords.ssi_accountid = new Lookup();
                        //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AccountId"]));
                        objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["AccountId"])));
                    }

                    //Contact lookup
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ContactId"]) != "")
                    {
                        //objMailRecords.ssi_contactfullnameid = new Lookup();
                        //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ContactId"]));
                        objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ContactId"])));
                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
                    {
                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_legalentitynameid = new Lookup();
                        //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                        objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(QuarAnnuDataset.Tables[0].Rows[i]["ssi_LegalEntityId"])));
                    }

                    // CreatedByCustomid Field 
                    //Rohit Pawar
                    string Userid = GetcurrentUser();

                    if (Userid != "")
                    {
                        //objMailRecords.ssi_createdbycustomid = new Lookup();
                        //objMailRecords.ssi_createdbycustomid.Value = new Guid(Userid);
                        objMailRecords["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Convert.ToString(Userid)));
                    }

                    service.Create(objMailRecords);
                    intResult++;
                    trBrowsefiles.Style.Add("display", "none");
                    trUnify.Style.Add("display", "none");
                }
            }



            #endregion

            #region Client Mailing| ADV Part2

            /* get Client Mailing information */
            else if (ddlMailType.SelectedValue == "3bd7d776-e1d3-e011-a19b-0019b9e7ee05" || ddlMailType.SelectedValue.ToUpper() == "AA0B12FB-B357-E911-8106-000D3A1C025B")
            {
                clsDB = new DB();
                DataSet CMAILDataset = new DataSet(); ;

                if (chkEmailRecipients.Checked == true)
                {
                    if (ddlMailType.SelectedValue == "3bd7d776-e1d3-e011-a19b-0019b9e7ee05")
                    {
                        CMAILDataset = clsDB.getDataSet("SP_S_CLIENT_MAILING @IncEmailRecipients=1");
                    }
                    else if (ddlMailType.SelectedValue.ToUpper() == "AA0B12FB-B357-E911-8106-000D3A1C025B") //added 4_11_2019 ADV Part2
                    {
                        CMAILDataset = clsDB.getDataSet("SP_S_CLIENT_MAILING @IncEmailRecipients=1 , @MailingTypeNmb = 1");
                    }
                }
                else
                {
                    if (ddlMailType.SelectedValue == "3bd7d776-e1d3-e011-a19b-0019b9e7ee05")
                    {
                        CMAILDataset = clsDB.getDataSet("SP_S_CLIENT_MAILING @IncEmailRecipients=0");
                    }
                    else if (ddlMailType.SelectedValue.ToUpper() == "AA0B12FB-B357-E911-8106-000D3A1C025B")//added 4_11_2019 ADV Part2
                    {
                        CMAILDataset = clsDB.getDataSet("SP_S_CLIENT_MAILING @IncEmailRecipients=0 , @MailingTypeNmb = 1");
                    }
                }



                for (int j = 0; j < CMAILDataset.Tables[0].Rows.Count; j++)
                {

                    //objMailRecords = new ssi_mailrecords();
                    Entity objMailRecords = new Entity("ssi_mailrecords");
                    //Mail Type
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_mailid"]) != "")
                    {
                        //objMailRecords.ssi_mailtypeid = new Lookup();
                        //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                        //objMailRecords.ssi_mailtypeid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_mailid"]));
                        objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_mailid"])));
                    }

                    //Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]) != "")
                    {
                        if (txtAsofdate.Text != "")
                        {
                            // objMailRecords.ssi_name = txtAsofdate.Text + "-" + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                            objMailRecords["ssi_name"] = txtAsofdate.Text + "-" + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                        }
                        else
                        {
                            //objMailRecords.ssi_name = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                            objMailRecords["ssi_name"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                        }

                    }


                    //[Spouse Name]
                    //First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]) != "")
                    {
                        // objMailRecords.ssi_spousepart_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]);
                        objMailRecords["ssi_spousepart_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]);
                    }

                    //First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]) != "")
                    {
                        // objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]);
                        objMailRecords["ssi_ownerfname_cnt_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]);
                    }

                    //Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]) != "")
                    {
                        //objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]);
                        objMailRecords["ssi_ownerlname_cnt_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]);
                    }

                    //House Hold
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]);
                    }

                    //Contact
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]) != "")
                    {
                        //objMailRecords.ssi_fullname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]);
                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]);
                    }

                    //Address Line 1
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]) != "")
                    {
                        //objMailRecords.ssi_addressline1_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]);
                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]) != "")
                    {
                        // objMailRecords.ssi_addressline2_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]);
                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]) != "")
                    {
                        //  objMailRecords.ssi_addressline3_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]);
                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]);
                    }

                    //City
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]) != "")
                    {
                        //  objMailRecords.ssi_city_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]);
                        objMailRecords["ssi_city_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]);
                    }

                    //State Or Province
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]) != "")
                    {
                        // objMailRecords.ssi_stateprovince_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]);
                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]);
                    }


                    //Zip Code
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]) != "")
                    {
                        //objMailRecords.ssi_zipcode_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]);
                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]);

                    }

                    //Country Or Region
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]) != "")
                    {
                        // objMailRecords.ssi_countryregion_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]);
                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]);
                    }

                    //Dear
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]) != "")
                    {
                        //objMailRecords.ssi_dear_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]);
                        objMailRecords["ssi_dear_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]);
                    }

                    //Mail Preference
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]) != "")
                    {
                        // objMailRecords.ssi_mailpreference_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]);

                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]);
                    }

                    //Salutation
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]) != "")
                    {
                        // objMailRecords.ssi_salutation_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]);
                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]);
                    }

                    //As Of Date
                    if (txtAsofdate.Text != "")
                    {
                        //objMailRecords.ssi_asofdate = new CrmDateTime();
                        //objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);

                    }

                    //Mailing ID
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]) != "")
                    {
                        //objMailRecords.ssi_mailingid = new CrmNumber();
                        //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                        objMailRecords["ssi_mailingid"] = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                        Mailing_Id = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]).ToString(); // added 6_11_2019
                    }

                    //Household/Institution
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]);
                    }

                    //Household Owner First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]) != "")
                    {
                        //objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]);
                        objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]);
                    }

                    //Household Owner Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]) != "")
                    {
                        //objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]);
                        objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]);
                    }

                    //Household Secondary Owner First Name ssi_secownerfname_mail 
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]) != "")
                    {
                        //objMailRecords.ssi_secownerfname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]);
                        objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]);
                    }

                    //Household Secondary Owner Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]) != "")
                    {
                        // objMailRecords.ssi_secownerlname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]);
                        objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]);
                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        // objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                    }


                    //HouseHold lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"]) != "")
                    {
                        //objMailRecords.ssi_accountid = new Lookup();
                        //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"]));
                        objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"])));
                    }

                    //Contact lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"]) != "")
                    {
                        //objMailRecords.ssi_contactfullnameid = new Lookup();
                        //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"]));
                        objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"])));
                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]) != "")
                    {
                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_legalentitynameid = new Lookup();
                        //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]));
                        objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]

["ssi_LegalEntityId"])));
                    }

                    // CreatedByCustomid Field 
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
                    trBrowsefiles.Style.Add("display", "none");
                    trUnify.Style.Add("display", "none");
                }
            }



            #endregion

            #region General/Smart/Prospect Mailing

            /* get General/Smart/Prospect Mailing information*/

            else if (ddlMailType.SelectedValue == "99b74584-e2d3-e011-a19b-0019b9e7ee05" || ddlMailType.SelectedValue == "c10ba3b7-e1d3-e011-a19b-0019b9e7ee05" || ddlMailType.SelectedValue == "c71108da-e1d3-e011-a19b-0019b9e7ee05")
            {
                clsDB = new DB();
                DataSet GSPDataset; //= clsDB.getDataSet("SP_S_GENERAL_SMART_PROSPECT_MAILING @ssi_mailid='" + ddlMailType.SelectedValue + "', @IncEmailRecipients=1");

                if (chkEmailRecipients.Checked == true)
                {
                    GSPDataset = clsDB.getDataSet("SP_S_GENERAL_SMART_PROSPECT_MAILING @ssi_mailid='" + ddlMailType.SelectedValue + "', @IncEmailRecipients=1");
                }
                else
                {
                    GSPDataset = clsDB.getDataSet("SP_S_GENERAL_SMART_PROSPECT_MAILING @ssi_mailid='" + ddlMailType.SelectedValue + "', @IncEmailRecipients=0");
                }



                for (int l = 0; l < GSPDataset.Tables[0].Rows.Count; l++)
                {
                    //objMailRecords = new ssi_mailrecords();
                    Entity objMailRecords = new Entity("ssi_mailrecords");
                    //Mail Type
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["ssi_mailid"]) != "")
                    {
                        //objMailRecords.ssi_mailtypeid = new Lookup();
                        //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                        //objMailRecords.ssi_mailtypeid.Value = new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["ssi_mailid"]));
                        objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["ssi_mailid"])));
                    }

                    //Name
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Name"]) != "")
                    {
                        if (txtAsofdate.Text != "")
                        {
                            // objMailRecords.ssi_name = txtAsofdate.Text + "-" + Convert.ToString(GSPDataset.Tables[0].Rows[l]["Name"]);
                            objMailRecords["ssi_name"] = txtAsofdate.Text + "-" + Convert.ToString(GSPDataset.Tables[0].Rows[l]["Name"]);
                        }
                        else
                        {
                            // objMailRecords.ssi_name = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Name"]);
                            objMailRecords["ssi_name"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Name"]);
                        }

                    }

                    //[Spouse Name]

                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Spouse Name"]) != "")
                    {
                        // objMailRecords.ssi_spousepart_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Spouse Name"]);
                        objMailRecords["ssi_spousepart_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Spouse Name"]);
                    }

                    //First Name
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["FirstName"]) != "")
                    {
                        // objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["FirstName"]);
                        objMailRecords["ssi_ownerfname_cnt_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["FirstName"]);
                    }

                    //Last Name
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["LastName"]) != "")
                    {
                        // objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["LastName"]);
                        objMailRecords["ssi_ownerlname_cnt_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["LastName"]);
                    }


                    //Contact
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Contact"]) != "")
                    {
                        // objMailRecords.ssi_fullname_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Contact"]);
                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Contact"]);
                    }

                    //Address Line 1
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine1"]) != "")
                    {
                        //objMailRecords.ssi_addressline1_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine1"]);
                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine2"]) != "")
                    {
                        //objMailRecords.ssi_addressline2_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine2"]);
                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine3"]) != "")
                    {
                        // objMailRecords.ssi_addressline3_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine3"]);
                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine3"]);
                    }

                    //City
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["City"]) != "")
                    {
                        // objMailRecords.ssi_city_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["City"]);
                        objMailRecords["ssi_city_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["City"]);
                    }

                    //State Or Province
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["State Or Province"]) != "")
                    {
                        // objMailRecords.ssi_stateprovince_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["State Or Province"]);
                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["State Or Province"]);
                    }


                    //Zip Code
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Zip Code"]) != "")
                    {
                        // objMailRecords.ssi_zipcode_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Zip Code"]);
                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Zip Code"]);
                    }

                    //Country Or Region
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Country Or Region"]) != "")
                    {
                        //objMailRecords.ssi_countryregion_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Country Or Region"]);
                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Country Or Region"]);
                    }

                    //Dear
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Dear"]) != "")
                    {
                        //  objMailRecords.ssi_dear_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Dear"]);
                        objMailRecords["ssi_dear_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Dear"]);
                    }

                    //Mail Preference
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Mail Preference"]) != "")
                    {
                        //objMailRecords.ssi_mailpreference_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Mail Preference"]);
                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Mail Preference"]);
                    }


                    //Salutation
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Salutation"]) != "")
                    {
                        //objMailRecords.ssi_salutation_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Salutation"]);
                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Salutation"]);
                    }


                    if (txtAsofdate.Text != "")
                    {
                        //objMailRecords.ssi_asofdate = new CrmDateTime();
                        //objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);

                    }

                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Ssi_MailingID"]) != "")
                    {
                        //objMailRecords.ssi_mailingid = new CrmNumber();
                        //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(GSPDataset.Tables[0].Rows[l]["GSPDataset"]);
                        objMailRecords["ssi_mailingid"] = Convert.ToInt32(GSPDataset.Tables[0].Rows[l]["Ssi_MailingID"]);
                        Mailing_Id = Convert.ToInt32(GSPDataset.Tables[0].Rows[l]["Ssi_MailingID"]).ToString();//added 6_11_2019
                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        // objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                    }


                    //HouseHold lookup
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["AccountId"]) != "")
                    {
                        //objMailRecords.ssi_accountid = new Lookup();
                        //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["AccountId"]));
                        objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["AccountId"])));

                    }

                    //Contact lookup
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["ContactId"]) != "")
                    {
                        //objMailRecords.ssi_contactfullnameid = new Lookup();
                        //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["ContactId"]));
                        objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["ContactId"])));
                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["ssi_LegalEntityId"]) != "")
                    {
                        ////objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_legalentitynameid = new Lookup();
                        //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["ssi_LegalEntityId"]));
                        objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["ssi_LegalEntityId"])));
                    }


                    // CreatedByCustomid Field 
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
                    trBrowsefiles.Style.Add("display", "none");
                    trUnify.Style.Add("display", "none");
                }
            }



            #endregion

            #region Capital call Letter| Capital Call Letter Wire | Fund Distribution | Fund Distribution Letter


            else if (ddlMailType.SelectedValue == "a1a079a4-d7be-e011-a19b-0019b9e7ee05" || ddlMailType.SelectedValue == "81091a9b-2ae9-e011-9141-0019b9e7ee05" || ddlMailType.SelectedValue == "78612b2b-5add-e011-ad4d-0019b9e7ee05" || ddlMailType.SelectedValue == "6d7545da-8164-e111-bd8f-0019b9e7ee05")//Capital Call Letter| Capital Call Letter Wire | Fund Distribution | Fund Distribution Letter
            {
                if (ddlMailType.SelectedValue == "a1a079a4-d7be-e011-a19b-0019b9e7ee05" || ddlMailType.SelectedValue == "81091a9b-2ae9-e011-9141-0019b9e7ee05" || ddlMailType.SelectedValue == "78612b2b-5add-e011-ad4d-0019b9e7ee05" || ddlMailType.SelectedValue == "6d7545da-8164-e111-bd8f-0019b9e7ee05")//Capital Call Letter| Capital Call Letter Wire | Fund Distribution | Fund Distribution Letter
                {



                    if (FileUpload1.HasFile == true)
                    {
                        Success = 0;

                        #region OLD CODE 
                        //if (System.IO.Path.GetExtension(FileUpload1.FileName) == ".xls" || System.IO.Path.GetExtension(FileUpload1.FileName) == ".xlsx")
                        //{
                        //    if (Request.Url.AbsoluteUri.Contains("localhost"))
                        //    {
                        //        FileUpload1.PostedFile.SaveAs(@"C:\\Reports\\" + FileUpload1.FileName);

                        //        if (File.Exists(@"C:\\Reports\\" + FileUpload1.FileName))
                        //        {
                        //            File.Delete(@"C:\\Reports\\Capital Call Sample File Layout.xlsx");
                        //            FileUpload1.PostedFile.SaveAs(@"C:\\Reports\\" + FileUpload1.FileName);
                        //            File.Move(@"C:\\Reports\\" + FileUpload1.FileName, @"C:\\Reports\\Capital Call Sample File Layout.xlsx");
                        //        }
                        //    }
                        //    else
                        //    {
                        //        // string extension = System.IO.Path.GetExtension(FileUpload1.FileName);//----10_31_2017(sasmit)
                        //        //Response.Write("FileUpload1.FileName:" + FileUpload1.FileName + "<br/><br/><br/>");
                        //        //string strFileName = "Capital Call Sample File Layout" + extension; //-------10_31_2017(sasmit)

                        //        string strFileName = clsDB.FileName("capitalcall");//----------10_31_2017(sasmit)
                        //                                                           // string strFileName = "Capital Call Sample File Layout.xlsx";//----------10_31_2017(sasmit)
                        //                                                           //Response.Write("New FileName:" + strFileName + "<br/><br/><br/>");
                        //                                                           //  FileUpload1.PostedFile.SaveAs(@"\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + strFileName);//----------10_31_2017(sasmit)

                        //        //if (File.Exists(@"\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + strFileName))//----------10_31_2017(sasmit)
                        //        if (File.Exists(DTSFilePath + strFileName))
                        //        {
                        //            //File.Delete(@"\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\Capital Call Sample File Layout.xlsx");//----------10_31_2017(sasmit)

                        //            File.Delete(DTSFilePath + strFileName);//----------10_31_2017(sasmit)
                        //                                                   //FileUpload1.PostedFile.SaveAs(@"\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + strFileName);//----------10_31_2017(sasmit)
                        //            FileUpload1.PostedFile.SaveAs(DTSFilePath + strFileName);//----------10_31_2017(sasmit)
                        //                                                                     //File.Move(@"\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\" + strFileName, @"\\\\GRPAO1-VWFS01\\Shared$\\Invoice\\CRM2011\\Capital Call Sample File Layout.xlsx");
                        //            File.Move(DTSFilePath + strFileName, DTSFilePath + strFileName);//----------10_31_2017(sasmit)
                        //        }
                        //        else
                        //        {
                        //            FileUpload1.PostedFile.SaveAs(DTSFilePath + strFileName); //10_31_2017(sasmit)
                        //        }
                        //    }


                        //    //"Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=TransactionLoad_DB;Data Source=sql01";

                        //    try
                        //    {
                        //        using (SqlConnection connection3 = new SqlConnection(con))
                        //        {
                        //            DateTime time;
                        //            ServerConnection serverConnection = new ServerConnection(connection3);
                        //            Server server = new Server(serverConnection);
                        //            Job job = server.JobServer.Jobs["CapitalCall"];
                        //            JobHistoryFilter filter = new JobHistoryFilter();
                        //            filter.JobName = "CapitalCall";
                        //            time = time = job.LastRunDate;
                        //            job.Start();
                        //            while (time == job.LastRunDate)
                        //            {
                        //                job.Refresh();
                        //            }
                        //            if (job.LastRunOutcome == CompletionResult.Succeeded)
                        //            {
                        //                retVal = 0;
                        //            }
                        //            else
                        //            {
                        //                retVal = 1;
                        //            }
                        //        }
                        //    }
                        //    catch (Exception exception3)
                        //    {
                        //        lblError.Text = "CapitalCall Letter Load Job Failed to Execute." + exception3.Message;
                        //    }
                        //}

                        ///*
                        ////string con = "Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TransactionLoad_DB;Data Source=SQL01";
                        //sqlconn = new SqlConnection(con);

                        //string strsql = "SP_DTS_RunPackage_InvoiceUpload";
                        //SqlCommand cmd = new SqlCommand();
                        //SqlParameter TypeId = cmd.Parameters.Add("@TypeId", SqlDbType.Int);
                        //TypeId.Value = 2;
                        //SqlParameter returncode = cmd.Parameters.Add("@returncode", SqlDbType.Int);
                        //returncode.Direction = ParameterDirection.Output;

                        ////cmd.Parameters["@returncode"].Direction = ParameterDirection.Output;
                        //cmd.CommandText = strsql;
                        //cmd.Connection = sqlconn;
                        //cmd.CommandType = CommandType.StoredProcedure;

                        //sqlconn.Open();
                        ////Response.Write(sqlconn.State + "<br/><br/><br/>");
                        ////Response.Write(sqlconn.Database + "<br/><br/><br/>"); 
                        //int result = cmd.ExecuteNonQuery();
                        ////System.Threading.Thread.Sleep(1000);

                        //int retVal = (int)cmd.Parameters["@returncode"].Value;

                        ////Response.Write("ret value:" + retVal.ToString() + "<br/><br/><br/>");
                        //*/
                        #endregion
                        string TempPath = string.Empty;
                        int ZipFileCount = 0;
                        bool bProceed = false;
                        int countSelectedFund = 0;
                        int countSelectedLegalEntity = 0;
                        if (System.IO.Path.GetExtension(FileUpload1.FileName) == ".zip")
                        {
                            try
                            {
                                DateTime dt = DateTime.Now;

                                string strHour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
                                string strMinute = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
                                string strSecond = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();

                                string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
                                string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
                                string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();

                                // string strUserName = HttpContext.Current.User.Identity.Name.ToString();

                                string strUserName = string.Empty;

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

                                string dateTime = "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond;
                                string TempFolderName = strUserName + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond;

                                TempPath = Server.MapPath("") + @"\ExcelTemplate\Tempfolder\" + TempFolderName + "\\";


                                ZipFileCount = ProcessZip(TempPath, dateTime, ddlMailType.SelectedValue);

                                if (ZipFileCount > 0)
                                {
                                    foreach (System.Web.UI.WebControls.ListItem li in lstFund.Items)
                                    {
                                        if (li.Selected)
                                        {
                                            countSelectedFund++;
                                        }
                                    }
                                    foreach (System.Web.UI.WebControls.ListItem li in lstLegalEntity.Items)
                                    {
                                        if (li.Selected)
                                        {
                                            countSelectedLegalEntity++;
                                        }
                                    }

                                    //If Funds are not selected -Run for all files present in the Zip
                                    if (countSelectedFund == 0)
                                    {
                                        Proces_ALL_FUND(TempPath, service, ddlMailType.SelectedValue);
                                    }
                                    else // Run only for Selected Fund
                                    {
                                        Proces_Selected_FUND(TempPath, service, ddlMailType.SelectedValue);
                                    }

                                    if(countSelectedLegalEntity ==0) // No selected LegalEntity
                                    {
                                        //Check if the Files in Zip 
                                        if (countSelectedFund == 0 && Success == ZipFileCount)
                                        {
                                            //lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully" + ", MailId: " + MailId;
                                            lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully";
                                            bUnify = true;
                                        }
                                        else if (Success == countSelectedFund && countSelectedFund != 0)
                                        {
                                            // lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully" + ", MailId: " + MailId;
                                            lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully";
                                            bUnify = true;
                                        }
                                        else
                                        {
                                            bUnify = false;
                                        }
                                    }
                                    else if(countSelectedLegalEntity > 0)
                                    {
                                        if(intResult>0)
                                        {
                                            lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully";
                                            bUnify = true;
                                        }
                                    }
                                   
                                }
                                else
                                {
                                    lblError.Text = "Error in Loading Zip File";
                                }


                                //Delete TempoFolder
                                if (Directory.Exists(TempPath))
                                {
                                    Directory.Delete(TempPath, true);
                                }

                            }
                            catch (Exception exception3)
                            {
                                lblError.Text = "CapitalCall Letter Process Failed " + exception3.Message;
                            }
                        }
                        else
                        {
                            lblError.Text = "Please Upload Zip File";
                        }




                    }
                    else if (FileUpload1.HasFile == false)
                    {
                        if (MailId != "" && chkUnify.Checked == true)
                        {
                            UpdateMailRecords(MailId);
                        }
                        BindMailingId();

                        //lblError.Text = ddlMailType.SelectedItem.Text + " records Unified successfully";
                    }

                  

                }
                //else if (FileUpload1.HasFile == false)
                //{
                //    if (MailId != "" && chkUnify.Checked == true)
                //    {
                //        UpdateMailRecords(MailId);
                //    }
                //    BindMailingId();

                //    //lblError.Text = ddlMailType.SelectedItem.Text + " records Unified successfully";
                //}



            }

            #endregion


            #region Fund Mailing (Signature Required)



            else if (ddlMailType.SelectedValue == "3cbaf86d-5edd-e011-ad4d-0019b9e7ee05")//Fund Info (Contacts for Confirmed Recommendations)
            {



                clsDB = new DB();
                DataSet FundInfoContacts = new DataSet();


                object AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
                object FundId = lstFund.SelectedValue == "0" || lstFund.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFund) + "'";
                object LegalEntityId = lstLegalEntity.SelectedValue == "" || lstLegalEntity.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstLegalEntity) + "'";
                //  if (chkEmailRecipients.Checked == true)
                //  {
                // FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_CURRENT_HOUSEHOLD @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@TypeId='" + ddlMailType.SelectedValue + "',@IncEmailRecipients=1");
                // }
                // else
                //{
                //FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_CURRENT_HOUSEHOLD @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@TypeId='" + ddlMailType.SelectedValue + "',@IncEmailRecipients=0");
                // }

                if (chkEmailRecipients.Checked == true)
                {
                    FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_CURRENT_HOUSEHOLD_SLOA @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@LegalEntityNameID= " + LegalEntityId + ",@IncEmailRecipients=1");
                }
                else
                {
                    FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_CURRENT_HOUSEHOLD_SLOA @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@LegalEntityNameID= " + LegalEntityId + ",@IncEmailRecipients=0");
                }

                for (int i = 0; i < FundInfoContacts.Tables[0].Rows.Count; i++)
                {

                    // objMailRecordsTemp = new ssi_mailrecordstemp();
                    Entity objMailRecordsTemp = new Entity("ssi_mailrecordstemp");

                    //objMailRecordsTemp.ssi_mailtypeid = new Lookup();
                    //objMailRecordsTemp.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                    //objMailRecordsTemp.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue);
                    objMailRecordsTemp["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));




                    //[Spouse Name]
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]) != "")
                    {
                        // objMailRecordsTemp.ssi_spousepart_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]);
                        objMailRecordsTemp["ssi_spousepart_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]);
                    }

                    //MailingID
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["MailingID"]) != "")
                    {
                        //objMailRecordsTemp.ssi_mailingid = new CrmNumber();
                        //objMailRecordsTemp.ssi_mailingid.Value = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]);
                        objMailRecordsTemp["ssi_mailingid"] = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]);

                    }

                    //Mail
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Name"]) != "")
                    {
                        // objMailRecordsTemp.ssi_name = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Name"]);
                        objMailRecordsTemp["ssi_name"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Name"]);
                    }


                    //First Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]) != "")
                    {
                        // objMailRecordsTemp.ssi_ownerfirstname_hh_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]);
                        objMailRecordsTemp["ssi_ownerfirstname_hh_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]);
                    }

                    //Last Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]) != "")
                    {
                        //objMailRecordsTemp.ssi_ownerlname_hh_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]);
                        objMailRecordsTemp["ssi_ownerlname_hh_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]);
                    }

                    //House Hold
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]) != "")
                    {
                        // objMailRecordsTemp.ssi_hholdinst_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]);
                        objMailRecordsTemp["ssi_hholdinst_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]);
                    }

                    //Contact
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]) != "")
                    {
                        // objMailRecordsTemp.ssi_fullname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]);
                        objMailRecordsTemp["ssi_fullname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]);
                    }


                    //House Hold lookup
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"]) != "")
                    {
                        //objMailRecordsTemp.ssi_accountid = new Lookup();
                        //objMailRecordsTemp.ssi_accountid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"]));
                        objMailRecordsTemp["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"])));
                    }

                    //Contact lookup
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"]) != "")
                    {
                        //objMailRecordsTemp.ssi_contactfullnameid = new Lookup();
                        //objMailRecordsTemp.ssi_contactfullnameid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"]));
                        objMailRecordsTemp["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"])));
                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
                    {
                        //objMailRecordsTemp.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecordsTemp.ssi_legalentitynameid = new Lookup();
                        //objMailRecordsTemp.ssi_legalentitynameid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                        objMailRecordsTemp["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"])));
                    }

                    //Address Line 1
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]) != "")
                    {
                        // objMailRecordsTemp.ssi_addressline1_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]);
                        objMailRecordsTemp["ssi_addressline1_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]) != "")
                    {
                        //objMailRecordsTemp.ssi_addressline2_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]);
                        objMailRecordsTemp["ssi_addressline2_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]) != "")
                    {
                        // objMailRecordsTemp.ssi_addressline3_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]);
                        objMailRecordsTemp["ssi_addressline3_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]);
                    }

                    //City
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]) != "")
                    {
                        // objMailRecordsTemp.ssi_city_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]);
                        objMailRecordsTemp["ssi_city_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]);
                    }

                    //State Or Province
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]) != "")
                    {
                        // objMailRecordsTemp.ssi_stateprovince_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]);
                        objMailRecordsTemp["ssi_stateprovince_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]);
                    }


                    //Zip Code
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]) != "")
                    {
                        //   objMailRecordsTemp.ssi_zipcode_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]);
                        objMailRecordsTemp["ssi_zipcode_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]);
                    }

                    //Country Or Region
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]) != "")
                    {
                        // objMailRecordsTemp.ssi_countryregion_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]);
                        objMailRecordsTemp["ssi_countryregion_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]);
                    }

                    //Dear
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]) != "")
                    {
                        //objMailRecordsTemp.ssi_dear_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]);
                        objMailRecordsTemp["ssi_dear_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]);
                    }

                    //Mail Preference
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]) != "")
                    {
                        //objMailRecordsTemp.ssi_mailpreference_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]);
                        objMailRecordsTemp["ssi_mailpreference_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]);
                    }

                    //Salutation
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]) != "")
                    {
                        // objMailRecordsTemp.ssi_salutation_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]);
                        objMailRecordsTemp["ssi_salutation_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]);
                    }

                    //Only for SLOA  -- New Condition 
                    if (txtLetterDate.Text != "")
                    {
                        //objMailRecordsTemp.ssi_asofdate = new CrmDateTime();
                        //objMailRecordsTemp.ssi_asofdate.Value = txtLetterDate.Text;
                        objMailRecordsTemp["ssi_asofdate"] = Convert.ToDateTime(txtLetterDate.Text);

                    }
                    else
                    {
                        //ASOF DATE
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Ssi_AsofDate"]) != "")
                        {
                            //objMailRecordsTemp.ssi_asofdate = new CrmDateTime();
                            //objMailRecordsTemp.ssi_asofdate.Value = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Ssi_AsofDate"]);
                            objMailRecordsTemp["ssi_asofdate"] = Convert.ToDateTime(FundInfoContacts.Tables[0].Rows[i]["Ssi_AsofDate"]);

                        }
                    }

                    //Anziano ID
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]) != "")
                    {
                        //objMailRecordsTemp.ssi_anzianoid = new CrmNumber();
                        //objMailRecordsTemp.ssi_anzianoid.Value = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]);
                        objMailRecordsTemp["ssi_anzianoid"] = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]);

                    }

                    //TNR ID
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]) != "")
                    {
                        //  objMailRecordsTemp.ssi_tnrid_nv = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]);
                        objMailRecordsTemp["ssi_tnrid_nv"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]);
                    }



                    //Secondary Owner First Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
                    {
                        //objMailRecordsTemp.ssi_secownerfname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]);
                        objMailRecordsTemp["ssi_secownerfname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]);
                    }

                    //Secondary Owner Last Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
                    {
                        // objMailRecordsTemp.ssi_secownerlname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                        objMailRecordsTemp["ssi_secownerlname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                    }


                    //Mailing Contact Type
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]) != "")
                    {
                        //objMailRecordsTemp.ssi_legalentity = new  = new CrmNumber();
                        // objMailRecordsTemp.ssi_mailingcontacttype = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]);
                        objMailRecordsTemp["ssi_mailingcontacttype"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]);
                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        //objMailRecordsTemp.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecordsTemp["ssi_mail"] = ddlMailType.SelectedItem.Text;
                    }

                    //Fund Name
                    if (lstFund.SelectedValue != "")
                    {
                        // objMailRecordsTemp.ssi_fundname = lstFund.SelectedItem.Text; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                        objMailRecordsTemp["ssi_fundname"] = lstFund.SelectedItem.Text; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                    }
                    //added by sasmit 5_3_2017
                    //clientportalname 
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_clientportalname"]) != "")
                    {
                        // objMailRecordsTemp.ssi_clientportalname = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_clientportalname"]);
                        objMailRecordsTemp["ssi_clientportalname"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_clientportalname"]);
                    }
                    //added by sasmit 5_3_2017
                    //clientreportfolder
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_clientreportfolder"]) != "")
                    {
                        //objMailRecordsTemp.ssi_clientreportfolder = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_clientreportfolder"]);
                        objMailRecordsTemp["ssi_clientreportfolder"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_clientreportfolder"]);
                    }
                    // CreatedByCustomid Field 

                    string Userid = GetcurrentUser();

                    if (Userid != "")
                    {
                        //objMailRecordsTemp.ssi_createdbycustomid = new Lookup();
                        //objMailRecordsTemp.ssi_createdbycustomid.Value = new Guid(Userid);

                        objMailRecordsTemp["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
                    }

                    #region Logic for Capital Call Approval Process

                    // Logic for Capital Call Approval Process

                    if (chkUnify.Checked == true)
                    {
                        //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
                        //objMailRecordsTemp.ssi_unifiedflg.Value = true;
                        objMailRecordsTemp["ssi_unifiedflg"] = true;
                    }
                    else if (chkUnify.Checked == false)
                    {
                        //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
                        //objMailRecordsTemp.ssi_unifiedflg.Value = false;
                        objMailRecordsTemp["ssi_unifiedflg"] = false;
                    }

                    if (MailId != "" && MailId != "0")
                    {
                        //objMailRecordsTemp.ssi_mailidtemp = new CrmNumber();
                        //objMailRecordsTemp.ssi_mailidtemp.Value = Convert.ToInt32(MailId);
                        objMailRecordsTemp["ssi_mailidtemp"] = Convert.ToInt32(MailId);

                    }


                    if (ddlTemplates.SelectedValue != "" && ddlTemplates.SelectedValue != "0")
                    {
                        //objMailRecordsTemp.ssi_templateid = new Lookup();
                        //objMailRecordsTemp.ssi_templateid.Value = new Guid(ddlTemplates.SelectedValue);
                        objMailRecordsTemp["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(ddlTemplates.SelectedValue));
                    }

                    ////Wire AsofDate
                    //if (txtWireAsofDate.Text != "")
                    //{
                    //    objMailRecordsTemp.ssi_wireasofdate = new CrmDateTime();
                    //    objMailRecordsTemp.ssi_wireasofdate.Value = txtWireAsofDate.Text;
                    //}

                    //Letter AsofDate
                    if (txtLetterDate.Text != "")
                    {
                        //objMailRecordsTemp.ssi_letterdate = new CrmDateTime();
                        //objMailRecordsTemp.ssi_letterdate.Value = txtLetterDate.Text;
                        objMailRecordsTemp["ssi_letterdate"] = Convert.ToDateTime(txtLetterDate.Text);

                    }


                    #endregion


                    if (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0" && chkUnify.Checked == false)
                    {
                        service.Create(objMailRecordsTemp);
                        intResult++;
                    }
                    else if (ddlMailId.SelectedValue != "" && chkUnify.Checked == true)
                    {
                        service.Create(objMailRecordsTemp);
                        intResult++;
                    }
                    else if (ddlMailId.SelectedValue != "" && chkUnify.Checked == false)
                    {
                        service.Create(objMailRecordsTemp);
                        intResult++;
                    }


                    //Response.Write(intResult.ToString());
                }

                #region Update Mailing List

                for (int j = 0; j < FundInfoContacts.Tables[1].Rows.Count; j++)
                {
                    //objMailingList = new ssi_mailinglist();
                    Entity objMailingList = new Entity("ssi_mailinglist");
                    if (Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]) != "")
                    {
                        //objMailingList.ssi_mailinglistid = new Key();
                        //objMailingList.ssi_mailinglistid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]));
                        objMailingList["ssi_mailinglistid"] = new Guid(Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]));

                    }


                    //if (txtWireAsofDate.Text != "")
                    //{
                    //    objMailingList.ssi_capitalcalldate = new CrmDateTime();
                    //    objMailingList.ssi_capitalcalldate.Value = txtWireAsofDate.Text;
                    //}

                    service.Update(objMailingList);
                    intResult++;
                }

                #endregion


                if (intResult > 0)
                {
                    lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully";
                    bUnify = true;
                }

                if (FundInfoContacts.Tables[1].Rows.Count == 0 && FundInfoContacts.Tables[0].Rows.Count == 0 && chkUnify.Checked == false && (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0"))
                    lblError.Text = "No Records Found";
                else if (FundInfoContacts.Tables[1].Rows.Count == 0 && FundInfoContacts.Tables[0].Rows.Count == 0 && chkUnify.Checked == false && (ddlMailId.SelectedValue != "" || ddlMailId.SelectedValue != "0"))
                    lblError.Text = "No Records Found";
                else if (FundInfoContacts.Tables[1].Rows.Count == 0 && FundInfoContacts.Tables[0].Rows.Count == 0 && chkUnify.Checked == true && (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0"))
                    lblError.Text = "No Records Found";
                else if (FundInfoContacts.Tables[1].Rows.Count == 0 && FundInfoContacts.Tables[0].Rows.Count == 0 && chkUnify.Checked == true && (ddlMailId.SelectedValue != "" || ddlMailId.SelectedValue != "0"))
                {
                    UpdateMailRecords(MailId);
                    BindMailingId();
                    lblError.Text = "No records found to save but records Unified successfully";
                }
            }

            #endregion

            #region Fund Info (Current Holders)

            else if (ddlMailType.SelectedValue == "10089fa9-59dd-e011-ad4d-0019b9e7ee05")//Fund Info (Contacts for Confirmed Recommendations)
            {

                clsDB = new DB();
                DataSet FundInfoContacts = new DataSet();


                object AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
                object FundId = lstFund.SelectedValue == "0" || lstFund.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFund) + "'";

                if (chkEmailRecipients.Checked == true)
                {
                    FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_CURRENT_HOUSEHOLD @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@MailID='" + ddlMailType.SelectedValue + "',@IncEmailRecipients=1");
                }
                else
                {
                    FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_CURRENT_HOUSEHOLD @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@MailID='" + ddlMailType.SelectedValue + "',@IncEmailRecipients=0");
                }

                for (int i = 0; i < FundInfoContacts.Tables[0].Rows.Count; i++)
                {

                    //objMailRecords = new ssi_mailrecords();
                    Entity objMailRecords = new Entity("ssi_mailrecords");
                    //Mail Type
                    if (ddlMailType.SelectedValue != "")
                    {
                        //objMailRecords.ssi_mailtypeid = new Lookup();
                        //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                        //objMailRecords.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue);
                        objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));
                    }

                    //[Spouse Name]
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]) != "")
                    {
                        //objMailRecords.ssi_spousepart_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]);
                        objMailRecords["ssi_spousepart_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]);
                    }

                    //MailingID
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["MailingID"]) != "")
                    {
                        //objMailRecords.ssi_mailingid = new CrmNumber();
                        //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]);
                        objMailRecords["ssi_mailingid"] = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]);
                        Mailing_Id = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]).ToString(); // added 6_11_2019
                    }

                    //Mail
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Name"]) != "")
                    {
                        // objMailRecords.ssi_name = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Name"]);
                        objMailRecords["ssi_name"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Name"]);
                    }


                    //First Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]) != "")
                    {
                        // objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]);
                        objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]);
                    }

                    //Last Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]) != "")
                    {
                        // objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]);
                        objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]);
                    }

                    //House Hold
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]);
                    }

                    //Contact
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]) != "")
                    {
                        // objMailRecords.ssi_fullname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]);
                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]);
                    }


                    //House Hold lookup
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"]) != "")
                    {
                        //objMailRecords.ssi_accountid = new Lookup();
                        //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"]));
                        objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"])));
                    }

                    //Contact lookup
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"]) != "")
                    {
                        //objMailRecords.ssi_contactfullnameid = new Lookup();
                        //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"]));
                        objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"])));
                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
                    {
                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_legalentitynameid = new Lookup();
                        //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                        objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"])));
                    }

                    //Address Line 1
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]) != "")
                    {
                        // objMailRecords.ssi_addressline1_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]);
                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]) != "")
                    {
                        // objMailRecords.ssi_addressline2_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]);
                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]) != "")
                    {
                        //objMailRecords.ssi_addressline3_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]);
                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]);
                    }

                    //City
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]) != "")
                    {
                        // objMailRecords.ssi_city_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]);
                        objMailRecords["ssi_city_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]);
                    }

                    //State Or Province
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]) != "")
                    {
                        //objMailRecords.ssi_stateprovince_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]);
                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]);
                    }


                    //Zip Code
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]) != "")
                    {
                        //objMailRecords.ssi_zipcode_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]);
                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]);
                    }

                    //Country Or Region
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]) != "")
                    {
                        //objMailRecords.ssi_countryregion_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]);
                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]);
                    }

                    //Dear
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]) != "")
                    {
                        //objMailRecords.ssi_dear_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]);
                        objMailRecords["ssi_dear_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]);
                    }

                    //Mail Preference
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]) != "")
                    {
                        // objMailRecords.ssi_mailpreference_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]);
                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]);
                    }

                    //Salutation
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]) != "")
                    {
                        //objMailRecords.ssi_salutation_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]);
                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]);
                    }

                    //ASOF DATE
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Ssi_AsofDate"]) != "")
                    {
                        //objMailRecords.ssi_asofdate = new CrmDateTime();
                        //objMailRecords.ssi_asofdate.Value = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Ssi_AsofDate"]);
                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(FundInfoContacts.Tables[0].Rows[i]["Ssi_AsofDate"]);

                    }

                    //Anziano ID
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]) != "")
                    {
                        //objMailRecords.ssi_anzianoid = new CrmNumber();
                        //objMailRecords.ssi_anzianoid.Value = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]);
                        objMailRecords["ssi_anzianoid"] = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]);

                    }

                    //TNR ID
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]) != "")
                    {
                        // objMailRecords.ssi_tnrid_nv = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]);
                        objMailRecords["ssi_tnrid_nv"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]);
                    }



                    //Secondary Owner First Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
                    {
                        //objMailRecords.ssi_secownerfname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]);
                        objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]);
                    }

                    //Secondary Owner Last Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
                    {
                        // objMailRecords.ssi_secownerlname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                        objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                    }


                    //Mailing Contact Type
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]) != "")
                    {
                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        // objMailRecords.ssi_mailingcontacttype = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]);
                        objMailRecords["ssi_mailingcontacttype"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]);
                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        //objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                    }

                    //Fund Name
                    if (lstFund.SelectedValue != "")
                    {
                        // objMailRecords.ssi_fundname = lstFund.SelectedItem.Text; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                        objMailRecords["ssi_fundname"] = lstFund.SelectedItem.Text; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                    }

                    // CreatedByCustomid Field 

                    string Userid = GetcurrentUser();

                    if (Userid != "")
                    {
                        //objMailRecords.ssi_createdbycustomid = new Lookup();
                        //objMailRecords.ssi_createdbycustomid.Value = new Guid(Userid);

                        objMailRecords["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
                    }

                    service.Create(objMailRecords);
                    intResult++;
                    trBrowsefiles.Style.Add("display", "none");
                    trUnify.Style.Add("display", "none");
                }


                #region Update Mailing List

                for (int j = 0; j < FundInfoContacts.Tables[1].Rows.Count; j++)
                {
                    //objMailingList = new ssi_mailinglist();
                    Entity objMailingList = new Entity("ssi_mailinglist");
                    if (Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]) != "")
                    {
                        //objMailingList.ssi_mailinglistid = new Key();
                        //objMailingList.ssi_mailinglistid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]));
                        objMailingList["ssi_mailinglistid"] = new Guid(Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]));

                    }


                    if (txtLetterDate.Text != "")
                    {
                        //objMailingList.ssi_fundmailingdate = new CrmDateTime();
                        //objMailingList.ssi_fundmailingdate.Value = txtLetterDate.Text;
                        objMailingList["ssi_fundmailingdate"] = Convert.ToDateTime(txtLetterDate.Text);

                    }

                    service.Update(objMailingList);
                    intResult++;
                }

                #endregion


                trBrowsefiles.Style.Add("display", "none");
                trMonths.Style.Add("display", "none");
                //RKLib.ExportData.Export objExport = new RKLib.ExportData.Export("Web");

                //if (FundInfoContacts.Tables[0].Rows.Count > 0)
                //{
                //    objExport.ExportDetails(FundInfoContacts.Tables[0], RKLib.ExportData.Export.ExportFormat.CSV, ddlMailType.SelectedItem.Text + ".CSV");
                //}

            }

            #endregion

            #region  Fund Info (Contacts for Confirmed Recommendations)

            else if (ddlMailType.SelectedValue == "b46939f9-59dd-e011-ad4d-0019b9e7ee05")//Fund Info (Contacts for Confirmed Recommendations)
            {
                clsDB = new DB();
                DataSet FundInfoContacts = new DataSet();

                object AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
                object FundId = lstFund.SelectedValue == "0" || lstFund.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFund) + "'";
                if (chkEmailRecipients.Checked == true)
                {
                    FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_RECOMMENDATIONS @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@IncEmailRecipients=1");
                }
                else
                {
                    FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_RECOMMENDATIONS @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@IncEmailRecipients=0");
                }


                for (int i = 0; i < FundInfoContacts.Tables[0].Rows.Count; i++)
                {

                    //objMailRecords = new ssi_mailrecords();
                    Entity objMailRecords = new Entity("ssi_mailrecords");

                    //Mail Type
                    if (ddlMailType.SelectedValue != "")
                    {
                        //objMailRecords.ssi_mailtypeid = new Lookup();
                        //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                        //objMailRecords.ssi_mailtypeid.Value = new Guid("b46939f9-59dd-e011-ad4d-0019b9e7ee05");
                        objMailRecords["ssi_mailtypeid"] = new EntityReference("ssi_mail", new Guid("b46939f9-59dd-e011-ad4d-0019b9e7ee05"));
                    }

                    //[Spouse Name]
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]) != "")
                    {
                        //objMailRecords.ssi_spousepart_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]);
                        objMailRecords["ssi_spousepart_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]);

                    }

                    //MailingID
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["MailingID"]) != "")
                    {
                        //objMailRecords.ssi_mailingid = new CrmNumber();
                        //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]);
                        objMailRecords["ssi_mailingid"] = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]);
                        Mailing_Id = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]).ToString(); // added 6_11_2019
                    }

                    //Mail
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail"]) != "")
                    {
                        //objMailRecords.ssi_name = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail"]);
                        objMailRecords["ssi_name"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail"]);
                    }


                    //First Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]) != "")
                    {
                        //objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]);
                        objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]);
                    }

                    //Last Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]) != "")
                    {
                        //objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]);
                        objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]);
                    }

                    //House Hold
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]);
                    }

                    //Contact
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]) != "")
                    {
                        //objMailRecords.ssi_fullname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]);
                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]);
                    }

                    //Address Line 1
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]) != "")
                    {
                        //objMailRecords.ssi_addressline1_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]);
                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]) != "")
                    {
                        //objMailRecords.ssi_addressline2_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]);
                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]) != "")
                    {
                        //objMailRecords.ssi_addressline3_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]);
                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]);
                    }

                    //City
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]) != "")
                    {
                        //objMailRecords.ssi_city_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]);
                        objMailRecords["ssi_city_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]);
                    }

                    //State Or Province
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]) != "")
                    {
                        //objMailRecords.ssi_stateprovince_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]);
                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]);
                    }

                    //Zip Code
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]) != "")
                    {
                        //objMailRecords.ssi_zipcode_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]);
                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]);
                    }

                    //Country Or Region
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]) != "")
                    {
                        //objMailRecords.ssi_countryregion_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]);
                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]);
                    }

                    //Dear
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]) != "")
                    {
                        //objMailRecords.ssi_dear_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]);
                        objMailRecords["ssi_dear_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]);
                    }

                    //Mail Preference
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]) != "")
                    {
                        //objMailRecords.ssi_mailpreference_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]);
                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]);
                    }

                    //Salutation
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]) != "")
                    {
                        //objMailRecords.ssi_salutation_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]);
                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]);
                    }

                    //ASOF DATE
                    if (txtAsofdate.Text != "")
                    {
                        //objMailRecords.ssi_asofdate = new CrmDateTime();
                        //objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);
                    }

                    //Anziano ID
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]) != "")
                    {
                        //objMailRecords.ssi_anzianoid = new CrmNumber();
                        //objMailRecords.ssi_anzianoid.Value = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]);
                        objMailRecords["ssi_anzianoid"] = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]);
                    }

                    //TNR ID
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]) != "")
                    {
                        //objMailRecords.ssi_tnrid_nv = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]);
                        objMailRecords["ssi_tnrid_nv"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]);
                    }


                    //Legal Entity Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Legal Entity Name"]) != "")
                    {
                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_legalentity = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Legal Entity Name"]);
                        objMailRecords["ssi_legalentity"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Legal Entity Name"]);
                    }


                    //Mailing Contact Type
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]) != "")
                    {
                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_mailingcontacttype = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]);
                        objMailRecords["ssi_mailingcontacttype"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]);
                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        //objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                    }

                    //HouseHold lookup
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"]) != "")
                    {
                        //objMailRecords.ssi_accountid = new Lookup();
                        //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"]));
                        objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"])));
                    }

                    //Contact lookup
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"]) != "")
                    {
                        //objMailRecords.ssi_contactfullnameid = new Lookup();
                        //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"]));
                        objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"])));
                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
                    {
                        ////objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_legalentitynameid = new Lookup();
                        //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                        objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"])));
                    }

                    //Fund Name
                    if (lstFund.SelectedValue != "")
                    {
                        //objMailRecords.ssi_fundname = lstFund.SelectedItem.Text; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                        objMailRecords["ssi_fundname"] = lstFund.SelectedItem.Text;
                    }



                    //Secondary Owner First Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
                    {
                        //objMailRecords.ssi_secownerfname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]);
                        objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]);
                    }

                    //Secondary Owner Last Name
                    if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
                    {
                        //objMailRecords.ssi_secownerlname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                        objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                    }

                    // CreatedByCustomid Field 

                    string Userid = GetcurrentUser();

                    if (Userid != "")
                    {
                        //objMailRecords.ssi_createdbycustomid = new Lookup();
                        //objMailRecords.ssi_createdbycustomid.Value = new Guid(Userid);

                        objMailRecords["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
                    }

                    service.Create(objMailRecords);
                    intResult++;
                    trBrowsefiles.Style.Add("display", "none");
                }


                #region Update Mailing List

                for (int j = 0; j < FundInfoContacts.Tables[1].Rows.Count; j++)
                {
                    //objMailingList = new ssi_mailinglist();
                    Entity objMailingList = new Entity("ssi_mailinglist");

                    if (Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]) != "")
                    {
                        //objMailingList.ssi_mailinglistid = new Key();
                        //objMailingList.ssi_mailinglistid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]));
                        objMailingList["ssi_mailinglistid"] = new Guid(Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]));
                    }


                    if (txtLetterDate.Text != "")
                    {
                        //objMailingList.ssi_fundmailingdate = new CrmDateTime();
                        //objMailingList.ssi_fundmailingdate.Value = txtLetterDate.Text;
                        objMailingList["ssi_fundmailingdate"] = Convert.ToDateTime(txtLetterDate.Text);
                    }

                    service.Update(objMailingList);
                    intResult++;
                }

                #endregion



                trBrowsefiles.Style.Add("display", "none");
                trMonths.Style.Add("display", "none");
                trUnify.Style.Add("display", "none");
                //RKLib.ExportData.Export objExport = new RKLib.ExportData.Export("Web");

                //if (FundInfoContacts.Tables[0].Rows.Count > 0)
                //{
                //    objExport.ExportDetails(FundInfoContacts.Tables[0], RKLib.ExportData.Export.ExportFormat.CSV, ddlMailType.SelectedItem.Text + ".CSV");
                //}


            }

            #endregion

            #region Owner Mailing

            else if (ddlMailType.SelectedValue == "5a79ab69-e60b-e111-b3cd-0019b9e7ee05")
            {
                clsDB = new DB();
                DataSet GSPDataset; //= clsDB.getDataSet("SP_S_GENERAL_SMART_PROSPECT_MAILING @ssi_mailid='" + ddlMailType.SelectedValue + "', @IncEmailRecipients=1");

                if (chkEmailRecipients.Checked == true)
                {
                    GSPDataset = clsDB.getDataSet("SP_S_GENERAL_SMART_PROSPECT_MAILING @ssi_mailid='" + ddlMailType.SelectedValue + "', @IncEmailRecipients=1");
                }
                else
                {
                    GSPDataset = clsDB.getDataSet("SP_S_GENERAL_SMART_PROSPECT_MAILING @ssi_mailid='" + ddlMailType.SelectedValue + "', @IncEmailRecipients=0");
                }



                for (int l = 0; l < GSPDataset.Tables[0].Rows.Count; l++)
                {
                    //objMailRecords = new ssi_mailrecords();
                    Entity objMailRecords = new Entity("ssi_mailrecords");

                    //Mail Type
                    if (ddlMailType.SelectedValue != "")
                    {
                        //objMailRecords.ssi_mailtypeid = new Lookup();
                        //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                        //objMailRecords.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue);
                        objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));

                    }

                    //Name
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Name"]) != "")
                    {
                        if (txtAsofdate.Text != "")
                        {
                            //objMailRecords.ssi_name = txtAsofdate.Text + "-" + Convert.ToString(GSPDataset.Tables[0].Rows[l]["Name"]);
                            objMailRecords["ssi_name"] = txtAsofdate.Text + "-" + Convert.ToString(GSPDataset.Tables[0].Rows[l]["Name"]);
                        }
                        else
                        {
                            //objMailRecords.ssi_name = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Name"]);
                            objMailRecords["ssi_name"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Name"]);
                        }

                    }


                    //[Spouse Name]
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Spouse Name"]) != "")
                    {
                        //objMailRecords.ssi_spousepart_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Spouse Name"]);
                        objMailRecords["ssi_spousepart_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Spouse Name"]);
                    }


                    //First Name
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["FirstName"]) != "")
                    {
                        //objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["FirstName"]);
                        objMailRecords["ssi_ownerfname_cnt_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["FirstName"]);
                    }

                    //Last Name
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["LastName"]) != "")
                    {
                        //objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["LastName"]);
                        objMailRecords["ssi_ownerlname_cnt_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["LastName"]);
                    }


                    //Contact
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Contact"]) != "")
                    {
                        //objMailRecords.ssi_fullname_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Contact"]);
                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Contact"]);
                    }

                    //Address Line 1
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine1"]) != "")
                    {
                        //objMailRecords.ssi_addressline1_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine1"]);
                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine2"]) != "")
                    {
                        //objMailRecords.ssi_addressline2_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine2"]);
                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine3"]) != "")
                    {
                        //objMailRecords.ssi_addressline3_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine3"]);
                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["AddressLine3"]);
                    }

                    //City
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["City"]) != "")
                    {
                        //objMailRecords.ssi_city_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["City"]);
                        objMailRecords["ssi_city_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["City"]);
                    }

                    //State Or Province
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["State Or Province"]) != "")
                    {
                        //objMailRecords.ssi_stateprovince_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["State Or Province"]);
                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["State Or Province"]);
                    }


                    //Zip Code
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Zip Code"]) != "")
                    {
                        //objMailRecords.ssi_zipcode_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Zip Code"]);
                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Zip Code"]);
                    }

                    //Country Or Region
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Country Or Region"]) != "")
                    {
                        //objMailRecords.ssi_countryregion_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Country Or Region"]);
                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Country Or Region"]);
                    }

                    //Dear
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Dear"]) != "")
                    {
                        //objMailRecords.ssi_dear_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Dear"]);
                        objMailRecords["ssi_dear_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Dear"]);
                    }

                    //Mail Preference
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Mail Preference"]) != "")
                    {
                        //objMailRecords.ssi_mailpreference_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Mail Preference"]);
                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Mail Preference"]);
                    }


                    //Salutation
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Salutation"]) != "")
                    {
                        //objMailRecords.ssi_salutation_mail = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Salutation"]);
                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(GSPDataset.Tables[0].Rows[l]["Salutation"]);
                    }


                    if (txtAsofdate.Text != "")
                    {
                        //objMailRecords.ssi_asofdate = new CrmDateTime();
                        //objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);

                    }

                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["Ssi_MailingID"]) != "")
                    {
                        //objMailRecords.ssi_mailingid = new CrmNumber();
                        //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(GSPDataset.Tables[0].Rows[l]["Ssi_MailingID"]);
                        objMailRecords["ssi_mailingid"] = Convert.ToInt32(GSPDataset.Tables[0].Rows[l]["Ssi_MailingID"]);
                        Mailing_Id = Convert.ToInt32(GSPDataset.Tables[0].Rows[l]["Ssi_MailingID"]).ToString();
                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        //objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                    }


                    //HouseHold lookup
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["AccountId"]) != "")
                    {
                        //objMailRecords.ssi_accountid = new Lookup();
                        //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["AccountId"]));
                        objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["AccountId"])));

                    }

                    //Contact lookup
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["ContactId"]) != "")
                    {
                        //objMailRecords.ssi_contactfullnameid = new Lookup();
                        //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["ContactId"]));
                        objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["ContactId"])));
                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(GSPDataset.Tables[0].Rows[l]["ssi_LegalEntityId"]) != "")
                    {
                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_legalentitynameid = new Lookup();
                        //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["ssi_LegalEntityId"]));
                        objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(GSPDataset.Tables[0].Rows[l]["ssi_LegalEntityId"])));
                    }


                    // CreatedByCustomid Field 
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
                    trBrowsefiles.Style.Add("display", "none");
                    trUnify.Style.Add("display", "none");
                }
                trUnify.Style.Add("display", "none");
                trBrowsefiles.Style.Add("display", "none");
                trMonths.Style.Add("display", "none");
            }

            #endregion

            #region Client Portal
            /* get Client Portal information */
            else if (ddlMailType.SelectedValue == "2357d455-f762-e111-bd8f-0019b9e7ee05")
            {
                clsDB = new DB();
                DataSet CMAILDataset;

                if (chkEmailRecipients.Checked == true)
                {
                    CMAILDataset = clsDB.getDataSet("SP_S_CLIENT_Portal @IncEmailRecipients=1");
                }
                else
                {
                    CMAILDataset = clsDB.getDataSet("SP_S_CLIENT_Portal @IncEmailRecipients=0");
                }



                for (int j = 0; j < CMAILDataset.Tables[0].Rows.Count; j++)
                {

                    //objMailRecords = new ssi_mailrecords();
                    Entity objMailRecords = new Entity("ssi_mailrecords");

                    //Mail Type
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_mailid"]) != "")
                    {
                        //objMailRecords.ssi_mailtypeid = new Lookup();
                        //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                        //objMailRecords.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue);
                        objMailRecords["ssi_mailtypeid"] = new EntityReference("ssi_mail", new Guid(ddlMailType.SelectedValue));

                    }

                    //Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]) != "")
                    {
                        if (txtAsofdate.Text != "")
                        {
                            //objMailRecords.ssi_name = txtAsofdate.Text + "-" + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                            objMailRecords["ssi_name"] = txtAsofdate.Text + "-" + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                        }
                        else
                        {
                            //objMailRecords.ssi_name = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                            objMailRecords["ssi_name"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                        }
                    }


                    //[Spouse Name]
                    //First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]) != "")
                    {
                        //objMailRecords.ssi_spousepart_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]);
                        objMailRecords["ssi_spousepart_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]);

                    }

                    //First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]) != "")
                    {
                        //objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]);
                        objMailRecords["ssi_ownerfname_cnt_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]);
                    }

                    //Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]) != "")
                    {
                        //objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]);
                        objMailRecords["ssi_ownerlname_cnt_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]);
                    }

                    //House Hold
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]);
                    }

                    //Contact
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]) != "")
                    {
                        //objMailRecords.ssi_fullname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]);
                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]);
                    }

                    //Address Line 1
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]) != "")
                    {
                        //objMailRecords.ssi_addressline1_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]);
                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]) != "")
                    {
                        //objMailRecords.ssi_addressline2_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]);
                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]) != "")
                    {
                        //objMailRecords.ssi_addressline3_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]);
                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]);
                    }

                    //City
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]) != "")
                    {
                        //objMailRecords.ssi_city_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]);
                        objMailRecords["ssi_city_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]);
                    }

                    //State Or Province
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]) != "")
                    {
                        //objMailRecords.ssi_stateprovince_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]);
                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]);
                    }


                    //Zip Code
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]) != "")
                    {
                        //objMailRecords.ssi_zipcode_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]);
                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]);
                    }

                    //Country Or Region
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]) != "")
                    {
                        //objMailRecords.ssi_countryregion_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]);
                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]);
                    }

                    //Dear
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]) != "")
                    {
                        //objMailRecords.ssi_dear_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]);
                        objMailRecords["ssi_dear_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]);
                    }

                    //Mail Preference
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]) != "")
                    {
                        //objMailRecords.ssi_mailpreference_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]);
                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]);
                    }

                    //Salutation
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]) != "")
                    {
                        //objMailRecords.ssi_salutation_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]);
                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]);
                    }

                    //As Of Date
                    if (txtAsofdate.Text != "")
                    {
                        //objMailRecords.ssi_asofdate = new CrmDateTime();
                        //objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);
                    }

                    //Mailing ID
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]) != "")
                    {
                        //objMailRecords.ssi_mailingid = new CrmNumber();
                        //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                        objMailRecords["ssi_mailingid"] = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                        Mailing_Id = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]).ToString();//added 6_11_2019
                    }

                    //Household/Institution
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]);
                    }

                    //Household Owner First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]) != "")
                    {
                        //objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]);
                        objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]);
                    }

                    //Household Owner Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]) != "")
                    {
                        //objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]);
                        objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]);
                    }

                    //Household Secondary Owner First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]) != "")
                    {
                        //objMailRecords.ssi_secownerfname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]);
                        objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]);
                    }

                    //Household Secondary Owner Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]) != "")
                    {
                        //objMailRecords.ssi_secownerlname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]);
                        objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]);
                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        //objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;

                    }

                    //HouseHold lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"]) != "")
                    {
                        //objMailRecords.ssi_accountid = new Lookup();
                        //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"]));
                        objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"])));
                    }

                    //Contact lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"]) != "")
                    {
                        //objMailRecords.ssi_contactfullnameid = new Lookup();
                        //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"]));
                        objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"])));
                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]) != "")
                    {
                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_legalentitynameid = new Lookup();
                        //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]));
                        objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"])));
                    }


                    // CreatedByCustomid Field 
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
                    trBrowsefiles.Style.Add("display", "none");
                    trUnify.Style.Add("display", "none");
                }
            }

            #endregion

            #region Client Event Mailing

            /* get Client Mailing information */
            else if (ddlMailType.SelectedValue == "005403c5-d5d1-e111-a4d8-0019b9e7ee05")
            {
                clsDB = new DB();
                DataSet CMAILDataset;

                if (chkEmailRecipients.Checked == true)
                {
                    CMAILDataset = clsDB.getDataSet("SP_S_CLIENT_EVENT_MAILING @IncEmailRecipients=1");
                }
                else
                {
                    CMAILDataset = clsDB.getDataSet("SP_S_CLIENT_EVENT_MAILING @IncEmailRecipients=0");
                }


                for (int j = 0; j < CMAILDataset.Tables[0].Rows.Count; j++)
                {

                    //objMailRecords = new ssi_mailrecords();
                    Entity objMailRecords = new Entity("ssi_mailrecords");

                    //Mail Type
                    //objMailRecords.ssi_mailtypeid = new Lookup();
                    //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                    //objMailRecords.ssi_mailtypeid.Value = new Guid("005403C5-D5D1-E111-A4D8-0019B9E7EE05");
                    objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString("005403C5-D5D1-E111-A4D8-0019B9E7EE05")));

                    //Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]) != "")
                    {
                        if (txtAsofdate.Text != "")
                        {
                            //objMailRecords.ssi_name = txtAsofdate.Text + "-" + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                            objMailRecords["ssi_name"] = txtAsofdate.Text + "-" + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                        }
                        else
                        {
                            //objMailRecords.ssi_name = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                            objMailRecords["ssi_name"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                        }

                    }


                    //[Spouse Name]
                    //First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]) != "")
                    {
                        //objMailRecords.ssi_spousepart_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]);
                        objMailRecords["ssi_spousepart_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]);
                    }

                    //First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]) != "")
                    {
                        //objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]);
                        objMailRecords["ssi_ownerfname_cnt_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]);
                    }

                    //Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]) != "")
                    {
                        //objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]);
                        objMailRecords["ssi_ownerlname_cnt_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]);
                    }

                    //House Hold
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]);
                    }

                    //Contact
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]) != "")
                    {
                        //objMailRecords.ssi_fullname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]);
                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]);
                    }

                    //Address Line 1
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]) != "")
                    {
                        //objMailRecords.ssi_addressline1_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]);
                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]) != "")
                    {
                        //objMailRecords.ssi_addressline2_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]);
                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]) != "")
                    {
                        //objMailRecords.ssi_addressline3_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]);
                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]);
                    }

                    //City
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]) != "")
                    {

                        //objMailRecords.ssi_city_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]);
                        objMailRecords["ssi_city_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]);
                    }

                    //State Or Province
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]) != "")
                    {
                        //objMailRecords.ssi_stateprovince_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]);
                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]);
                    }


                    //Zip Code
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]) != "")
                    {
                        //objMailRecords.ssi_zipcode_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]);
                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]);
                    }

                    //Country Or Region
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]) != "")
                    {
                        //objMailRecords.ssi_countryregion_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]);
                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]);
                    }

                    //Dear
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]) != "")
                    {
                        //objMailRecords.ssi_dear_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]);
                        objMailRecords["ssi_dear_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]);
                    }

                    //Mail Preference
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]) != "")
                    {
                        //objMailRecords.ssi_mailpreference_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]);
                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]);
                    }

                    //Salutation
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]) != "")
                    {
                        //objMailRecords.ssi_salutation_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]);
                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]);
                    }

                    //As Of Date
                    if (txtAsofdate.Text != "")
                    {
                        //objMailRecords.ssi_asofdate = new CrmDateTime();
                        //objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);
                    }

                    //Mailing ID
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]) != "")
                    {
                        //objMailRecords.ssi_mailingid = new CrmNumber();
                        //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                        objMailRecords["ssi_mailingid"] = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                        Mailing_Id = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]).ToString(); // added 6_11_2019
                    }

                    //Household/Institution
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]);
                    }

                    //Household Owner First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]) != "")
                    {
                        //objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]);
                        objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]);
                    }

                    //Household Owner Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]) != "")
                    {
                        //objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]);
                        objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]);
                    }

                    //Household Secondary Owner First Name ssi_secownerfname_mail 
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]) != "")
                    {
                        //objMailRecords.ssi_secownerfname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]);
                        objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]);
                    }

                    //Household Secondary Owner Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]) != "")
                    {
                        //objMailRecords.ssi_secownerlname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]);
                        objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]);
                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        //objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                    }


                    //HouseHold lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"]) != "")
                    {
                        //objMailRecords.ssi_accountid = new Lookup();
                        //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"]));
                        objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"])));
                    }

                    //Contact lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"]) != "")
                    {
                        //objMailRecords.ssi_contactfullnameid = new Lookup();
                        //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["contact"]));
                        objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"])));

                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]) != "")
                    {
                        //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_legalentitynameid = new Lookup();
                        //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]));
                        objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"])));
                    }

                    // CreatedByCustomid Field 
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
                    trBrowsefiles.Style.Add("display", "none");
                    trMonths.Style.Add("display", "none");
                }
            }



            #endregion

            #region On Demand Mailing/One Time Mailing/Adhoc Mailing

            /* get Client Mailing information */
            else if (ddlMailType.SelectedValue == "01ab40eb-ddd1-e111-a4d8-0019b9e7ee05" || ddlMailType.SelectedValue == "18dfb14e-b156-e711-9422-005056a0567e" || ddlMailType.SelectedValue.ToUpper() == "D0AA2024-B457-E911-8106-000D3A1C025B")
            {
                clsDB = new DB();
                DataSet CMAILDataset = new DataSet();

                if (chkEmailRecipients.Checked == true)
                {
                    if (ddlMailType.SelectedValue == "01ab40eb-ddd1-e111-a4d8-0019b9e7ee05")
                    {
                        CMAILDataset = clsDB.getDataSet("SP_S_ON_DEMAND_MAILING @IncEmailRecipients=1");
                    }
                    else if (ddlMailType.SelectedValue == "18dfb14e-b156-e711-9422-005056a0567e")
                    {
                        CMAILDataset = clsDB.getDataSet("SP_S_On_Demand_Mailing @OneTimeMailFlg = 1 , @IncEmailRecipients=1");
                    }
                    else if (ddlMailType.SelectedValue.ToUpper() == "D0AA2024-B457-E911-8106-000D3A1C025B")// added 4_11_2019  - Adhoc Mailing
                    {
                        CMAILDataset = clsDB.getDataSet("SP_S_On_Demand_Mailing @OneTimeMailFlg = 2 , @IncEmailRecipients=1");
                    }
                    //CMAILDataset = clsDB.getDataSet("SP_S_ON_DEMAND_MAILING @IncEmailRecipients=1");
                }
                else
                {
                    if (ddlMailType.SelectedValue == "01ab40eb-ddd1-e111-a4d8-0019b9e7ee05")
                    {
                        CMAILDataset = clsDB.getDataSet("SP_S_ON_DEMAND_MAILING @IncEmailRecipients=0");
                    }
                    else if (ddlMailType.SelectedValue == "18dfb14e-b156-e711-9422-005056a0567e")//One Time Mailing
                    {
                        CMAILDataset = clsDB.getDataSet("SP_S_On_Demand_Mailing @OneTimeMailFlg = 1 , @IncEmailRecipients=0");
                    }
                    else if (ddlMailType.SelectedValue.ToUpper() == "D0AA2024-B457-E911-8106-000D3A1C025B")// added 4_11_2019  - Adhoc Mailing
                    {
                        CMAILDataset = clsDB.getDataSet("SP_S_On_Demand_Mailing @OneTimeMailFlg = 2 , @IncEmailRecipients=0");
                    }
                    // CMAILDataset = clsDB.getDataSet("SP_S_ON_DEMAND_MAILING @IncEmailRecipients=0");
                }


                for (int j = 0; j < CMAILDataset.Tables[0].Rows.Count; j++)
                {
                    //objMailRecords = new ssi_mailrecords();
                    Entity objMailRecords = new Entity("ssi_mailrecords");


                    //Mail Type
                    //objMailRecords.ssi_mailtypeid = new Lookup();
                    //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                    //objMailRecords.ssi_mailtypeid.Value = new Guid("01AB40EB-DDD1-E111-A4D8-0019B9E7EE05");
                    //  objMailRecords["ssi_mailtypeid"]=new EntityReference("ssi_mail", new Guid("01AB40EB-DDD1-E111-A4D8-0019B9E7EE05"));
                    //Mail Type
                    if (ddlMailType.SelectedValue == "01ab40eb-ddd1-e111-a4d8-0019b9e7ee05")
                    {
                        objMailRecords["ssi_mailtypeid"] = new EntityReference("ssi_mail", new Guid("01ab40eb-ddd1-e111-a4d8-0019b9e7ee05"));
                    }
                    else if (ddlMailType.SelectedValue == "18dfb14e-b156-e711-9422-005056a0567e")//One Time Mailing
                    {
                        objMailRecords["ssi_mailtypeid"] = new EntityReference("ssi_mail", new Guid("18dfb14e-b156-e711-9422-005056a0567e"));
                    }
                    else if (ddlMailType.SelectedValue.ToUpper() == "D0AA2024-B457-E911-8106-000D3A1C025B") // added 4_11_2019  - Adhoc Mailing
                    {
                        objMailRecords["ssi_mailtypeid"] = new EntityReference("ssi_mail", new Guid("D0AA2024-B457-E911-8106-000D3A1C025B"));
                    }


                    //Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]) != "")
                    {
                        if (txtAsofdate.Text != "")
                        {
                            //objMailRecords.ssi_name = txtAsofdate.Text + "-" + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                            objMailRecords["ssi_name"] = txtAsofdate.Text + "-" + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                        }
                        else
                        {
                            //objMailRecords.ssi_name = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                            objMailRecords["ssi_name"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                        }
                    }


                    //[Spouse Name]
                    //First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]) != "")
                    {
                        //objMailRecords.ssi_spousepart_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]);
                        objMailRecords["ssi_spousepart_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]);
                    }

                    //First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]) != "")
                    {
                        //objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]);
                        objMailRecords["ssi_ownerfname_cnt_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]);
                    }

                    //Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]) != "")
                    {
                        //objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]);
                        objMailRecords["ssi_ownerlname_cnt_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]);
                    }

                    //House Hold
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]);
                    }

                    //Contact
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]) != "")
                    {
                        //objMailRecords.ssi_fullname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]);
                        objMailRecords["ssi_fullname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]);
                    }

                    //Address Line 1
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]) != "")
                    {
                        //objMailRecords.ssi_addressline1_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]);
                        objMailRecords["ssi_addressline1_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]);
                    }

                    //Address Line 2
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]) != "")
                    {
                        //objMailRecords.ssi_addressline2_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]);
                        objMailRecords["ssi_addressline2_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]);
                    }

                    //Address Line 3
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]) != "")
                    {
                        //objMailRecords.ssi_addressline3_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]);
                        objMailRecords["ssi_addressline3_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]);
                    }

                    //City
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]) != "")
                    {
                        //objMailRecords.ssi_city_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]);
                        objMailRecords["ssi_city_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]);
                    }

                    //State Or Province
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]) != "")
                    {
                        //objMailRecords.ssi_stateprovince_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]);
                        objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]);
                    }


                    //Zip Code
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]) != "")
                    {
                        //objMailRecords.ssi_zipcode_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]);
                        objMailRecords["ssi_zipcode_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]);
                    }

                    //Country Or Region
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]) != "")
                    {
                        //objMailRecords.ssi_countryregion_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]);
                        objMailRecords["ssi_countryregion_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]);
                    }

                    //Dear
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]) != "")
                    {
                        //objMailRecords.ssi_dear_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]);
                        objMailRecords["ssi_dear_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]);
                    }

                    //Mail Preference
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]) != "")
                    {
                        //objMailRecords.ssi_mailpreference_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]);
                        objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]);
                    }

                    //Salutation
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]) != "")
                    {
                        //objMailRecords.ssi_salutation_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]);
                        objMailRecords["ssi_salutation_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]);
                    }

                    //As Of Date
                    if (txtAsofdate.Text != "")
                    {
                        //objMailRecords.ssi_asofdate = new CrmDateTime();
                        //objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                        objMailRecords["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);
                    }

                    //Mailing ID
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]) != "")
                    {
                        //objMailRecords.ssi_mailingid = new CrmNumber();
                        //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                        objMailRecords["ssi_mailingid"] = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                        Mailing_Id = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]).ToString(); // added 6_11_2019
                    }

                    //Household/Institution
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]) != "")
                    {
                        //objMailRecords.ssi_hholdinst_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]);
                        objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]);
                    }

                    //Household Owner First Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]) != "")
                    {
                        //objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]);
                        objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]);
                    }

                    //Household Owner Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]) != "")
                    {
                        //objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]);
                        objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]);
                    }

                    //Household Secondary Owner First Name ssi_secownerfname_mail 
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]) != "")
                    {
                        //objMailRecords.ssi_secownerfname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]);
                        objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]);
                    }

                    //Household Secondary Owner Last Name
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]) != "")
                    {
                        //objMailRecords.ssi_secownerlname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]);
                        objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]);
                    }

                    if (ddlMailType.SelectedValue != "")
                    {
                        //objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                        objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                    }


                    //HouseHold lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"]) != "")
                    {
                        //objMailRecords.ssi_accountid = new Lookup();
                        //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"]));
                        objMailRecords["ssi_accountid"] = new EntityReference("account", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"])));
                    }

                    //Contact lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"]) != "")
                    {
                        //objMailRecords.ssi_contactfullnameid = new Lookup();
                        //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"]));
                        objMailRecords["ssi_contactfullnameid"] = new EntityReference("contact", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"])));
                    }

                    //ssi_LegalEntityId lookup
                    if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]) != "")
                    {
                        ////objMailRecords.ssi_legalentity = new  = new CrmNumber();
                        //objMailRecords.ssi_legalentitynameid = new Lookup();
                        //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]));
                        objMailRecords["ssi_legalentitynameid"] = new EntityReference("ssi_legalentity", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"])));
                    }


                    // CreatedByCustomid Field 
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
                    trBrowsefiles.Style.Add("display", "none");
                    trMonths.Style.Add("display", "none");
                }
            }



            #endregion

            #region Other Mailing Type
            else
            {
                // if (Mail_type == "100000000")
                if (Mail_type == "1" || Mail_type == "5")// changed field(MailType to existing field Mailingtype) in CRM-6_25_2019
                {
                    #region Contact Specific|Contact Specific-Non client similar to Client Mailing| ADV Part2
                    clsDB = new DB();
                    DataSet CMAILDataset = new DataSet(); ;

                    if (chkEmailRecipients.Checked == true)
                    {

                        CMAILDataset = clsDB.getDataSet("SP_S_CLIENT_MAILING @IncEmailRecipients=1,@MailID='" + ddlMailType.SelectedValue + "'");

                    }
                    else
                    {

                        CMAILDataset = clsDB.getDataSet("SP_S_CLIENT_MAILING @IncEmailRecipients=0,@MailID='" + ddlMailType.SelectedValue + "'");

                    }



                    for (int j = 0; j < CMAILDataset.Tables[0].Rows.Count; j++)
                    {

                        //objMailRecords = new ssi_mailrecords();
                        Entity objMailRecords = new Entity("ssi_mailrecords");
                        //Mail Type
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_mailid"]) != "")
                        {
                            //objMailRecords.ssi_mailtypeid = new Lookup();
                            //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                            //objMailRecords.ssi_mailtypeid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_mailid"]));
                            objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));
                        }

                        //Name
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]) != "")
                        {
                            if (txtAsofdate.Text != "")
                            {
                                // objMailRecords.ssi_name = txtAsofdate.Text + "-" + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                                objMailRecords["ssi_name"] = txtAsofdate.Text + "-" + ddlMailType.SelectedItem.Text + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["MailName"]);
                            }
                            else
                            {
                                //objMailRecords.ssi_name = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Name"]);
                                objMailRecords["ssi_name"] = ddlMailType.SelectedItem.Text + Convert.ToString(CMAILDataset.Tables[0].Rows[j]["MailName"]);
                            }

                        }


                        //[Spouse Name]
                        //First Name
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]) != "")
                        {
                            // objMailRecords.ssi_spousepart_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]);
                            objMailRecords["ssi_spousepart_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Spouse Name"]);
                        }

                        //First Name
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]) != "")
                        {
                            // objMailRecords.ssi_ownerfname_cnt_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]);
                            objMailRecords["ssi_ownerfname_cnt_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["FirstName"]);
                        }

                        //Last Name
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]) != "")
                        {
                            //objMailRecords.ssi_ownerlname_cnt_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]);
                            objMailRecords["ssi_ownerlname_cnt_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["LastName"]);
                        }

                        //House Hold
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]) != "")
                        {
                            //objMailRecords.ssi_hholdinst_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]);
                            objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["HouseHold"]);
                        }

                        //Contact
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]) != "")
                        {
                            //objMailRecords.ssi_fullname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]);
                            objMailRecords["ssi_fullname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Contact"]);
                        }

                        //Address Line 1
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]) != "")
                        {
                            //objMailRecords.ssi_addressline1_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]);
                            objMailRecords["ssi_addressline1_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine1"]);
                        }

                        //Address Line 2
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]) != "")
                        {
                            // objMailRecords.ssi_addressline2_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]);
                            objMailRecords["ssi_addressline2_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine2"]);
                        }

                        //Address Line 3
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]) != "")
                        {
                            //  objMailRecords.ssi_addressline3_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]);
                            objMailRecords["ssi_addressline3_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AddressLine3"]);
                        }

                        //City
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]) != "")
                        {
                            //  objMailRecords.ssi_city_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]);
                            objMailRecords["ssi_city_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["City"]);
                        }

                        //State Or Province
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]) != "")
                        {
                            // objMailRecords.ssi_stateprovince_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]);
                            objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["State Or Province"]);
                        }


                        //Zip Code
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]) != "")
                        {
                            //objMailRecords.ssi_zipcode_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]);
                            objMailRecords["ssi_zipcode_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Zip Code"]);

                        }

                        //Country Or Region
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]) != "")
                        {
                            // objMailRecords.ssi_countryregion_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]);
                            objMailRecords["ssi_countryregion_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Country Or Region"]);
                        }

                        //Dear
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]) != "")
                        {
                            //objMailRecords.ssi_dear_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]);
                            objMailRecords["ssi_dear_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Dear"]);
                        }

                        //Mail Preference
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]) != "")
                        {
                            // objMailRecords.ssi_mailpreference_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]);

                            objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Mail Preference"]);
                        }

                        //Salutation
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]) != "")
                        {
                            // objMailRecords.ssi_salutation_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]);
                            objMailRecords["ssi_salutation_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Salutation"]);
                        }

                        //As Of Date
                        if (txtAsofdate.Text != "")
                        {
                            //objMailRecords.ssi_asofdate = new CrmDateTime();
                            //objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                            objMailRecords["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);

                        }

                        //Mailing ID
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]) != "")
                        {
                            //objMailRecords.ssi_mailingid = new CrmNumber();
                            //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                            objMailRecords["ssi_mailingid"] = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]);
                            Mailing_Id = Convert.ToInt32(CMAILDataset.Tables[0].Rows[j]["Ssi_MailingID"]).ToString(); // added 6_11_2019
                        }

                        //Household/Institution
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]) != "")
                        {
                            //objMailRecords.ssi_hholdinst_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]);
                            objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["Household"]);
                        }

                        //Household Owner First Name
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]) != "")
                        {
                            //objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]);
                            objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerFirstName"]);
                        }

                        //Household Owner Last Name
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]) != "")
                        {
                            //objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]);
                            objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["OwnerLastName"]);
                        }

                        //Household Secondary Owner First Name ssi_secownerfname_mail 
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]) != "")
                        {
                            //objMailRecords.ssi_secownerfname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]);
                            objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerFirstName"]);
                        }

                        //Household Secondary Owner Last Name
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]) != "")
                        {
                            // objMailRecords.ssi_secownerlname_mail = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]);
                            objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(CMAILDataset.Tables[0].Rows[j]["SecondaryOwnerLastName"]);
                        }

                        if (ddlMailType.SelectedValue != "")
                        {
                            // objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                            objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                        }


                        //HouseHold lookup
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"]) != "")
                        {
                            //objMailRecords.ssi_accountid = new Lookup();
                            //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"]));
                            objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["AccountId"])));
                        }

                        //Contact lookup
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"]) != "")
                        {
                            //objMailRecords.ssi_contactfullnameid = new Lookup();
                            //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"]));
                            objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ContactId"])));
                        }

                        //ssi_LegalEntityId lookup
                        if (Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]) != "")
                        {
                            //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                            //objMailRecords.ssi_legalentitynameid = new Lookup();
                            //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"]));
                            objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(CMAILDataset.Tables[0].Rows[j]["ssi_LegalEntityId"])));
                        }

                        // CreatedByCustomid Field 
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
                        trBrowsefiles.Style.Add("display", "none");
                        trUnify.Style.Add("display", "none");
                    }

                    // Response.Write("Contact Baswe Done" + DateTime.Now.ToString());
                    #endregion
                }
                //else if (Mail_type == "100000001")
                else if (Mail_type == "2")// changed field(MailType to existing field Mailingtype) in CRM-6_25_2019
                {
                    #region Position Based, Similar to Fund Info (Current Holders)
                    clsDB = new DB();
                    DataSet FundInfoContacts = new DataSet();


                    object AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
                    object FundId = lstFund.SelectedValue == "0" || lstFund.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFund) + "'";

                    if (chkEmailRecipients.Checked == true)
                    {
                        FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_CURRENT_HOUSEHOLD @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@MailID='" + ddlMailType.SelectedValue + "',@IncEmailRecipients=1");
                    }
                    else
                    {
                        FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_CURRENT_HOUSEHOLD @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@MailID='" + ddlMailType.SelectedValue + "',@IncEmailRecipients=0");
                    }

                    for (int i = 0; i < FundInfoContacts.Tables[0].Rows.Count; i++)
                    {

                        //objMailRecords = new ssi_mailrecords();
                        Entity objMailRecords = new Entity("ssi_mailrecords");
                        //Mail Type
                        if (ddlMailType.SelectedValue != "")
                        {
                            //objMailRecords.ssi_mailtypeid = new Lookup();
                            //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                            //objMailRecords.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue);
                            objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));
                        }

                        //[Spouse Name]
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]) != "")
                        {
                            //objMailRecords.ssi_spousepart_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]);
                            objMailRecords["ssi_spousepart_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]);
                        }

                        //MailingID
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["MailingID"]) != "")
                        {
                            //objMailRecords.ssi_mailingid = new CrmNumber();
                            //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]);
                            objMailRecords["ssi_mailingid"] = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]);
                            Mailing_Id = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]).ToString(); // added 6_11_2019
                        }

                        //Mail
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Name"]) != "")
                        {
                            // objMailRecords.ssi_name = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Name"]);
                            objMailRecords["ssi_name"] = ddlMailType.SelectedItem.Text + Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["MailName"]);
                        }


                        //First Name
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]) != "")
                        {
                            // objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]);
                            objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]);
                        }

                        //Last Name
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]) != "")
                        {
                            // objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]);
                            objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]);
                        }

                        //House Hold
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]) != "")
                        {
                            //objMailRecords.ssi_hholdinst_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]);
                            objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]);
                        }

                        //Contact
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]) != "")
                        {
                            // objMailRecords.ssi_fullname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]);
                            objMailRecords["ssi_fullname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]);
                        }


                        //House Hold lookup
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"]) != "")
                        {
                            //objMailRecords.ssi_accountid = new Lookup();
                            //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"]));
                            objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"])));
                        }

                        //Contact lookup
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"]) != "")
                        {
                            //objMailRecords.ssi_contactfullnameid = new Lookup();
                            //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"]));
                            objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"])));
                        }

                        //ssi_LegalEntityId lookup
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
                        {
                            //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                            //objMailRecords.ssi_legalentitynameid = new Lookup();
                            //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                            objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"])));
                        }

                        //Address Line 1
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]) != "")
                        {
                            // objMailRecords.ssi_addressline1_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]);
                            objMailRecords["ssi_addressline1_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]);
                        }

                        //Address Line 2
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]) != "")
                        {
                            // objMailRecords.ssi_addressline2_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]);
                            objMailRecords["ssi_addressline2_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]);
                        }

                        //Address Line 3
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]) != "")
                        {
                            //objMailRecords.ssi_addressline3_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]);
                            objMailRecords["ssi_addressline3_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]);
                        }

                        //City
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]) != "")
                        {
                            // objMailRecords.ssi_city_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]);
                            objMailRecords["ssi_city_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]);
                        }

                        //State Or Province
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]) != "")
                        {
                            //objMailRecords.ssi_stateprovince_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]);
                            objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]);
                        }


                        //Zip Code
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]) != "")
                        {
                            //objMailRecords.ssi_zipcode_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]);
                            objMailRecords["ssi_zipcode_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]);
                        }

                        //Country Or Region
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]) != "")
                        {
                            //objMailRecords.ssi_countryregion_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]);
                            objMailRecords["ssi_countryregion_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]);
                        }

                        //Dear
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]) != "")
                        {
                            //objMailRecords.ssi_dear_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]);
                            objMailRecords["ssi_dear_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]);
                        }

                        //Mail Preference
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]) != "")
                        {
                            // objMailRecords.ssi_mailpreference_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]);
                            objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]);
                        }

                        //Salutation
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]) != "")
                        {
                            //objMailRecords.ssi_salutation_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]);
                            objMailRecords["ssi_salutation_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]);
                        }

                        //ASOF DATE
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Ssi_AsofDate"]) != "")
                        {
                            //objMailRecords.ssi_asofdate = new CrmDateTime();
                            //objMailRecords.ssi_asofdate.Value = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Ssi_AsofDate"]);
                            objMailRecords["ssi_asofdate"] = Convert.ToDateTime(FundInfoContacts.Tables[0].Rows[i]["Ssi_AsofDate"]);

                        }

                        //Anziano ID
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]) != "")
                        {
                            //objMailRecords.ssi_anzianoid = new CrmNumber();
                            //objMailRecords.ssi_anzianoid.Value = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]);
                            objMailRecords["ssi_anzianoid"] = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]);

                        }

                        //TNR ID
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]) != "")
                        {
                            // objMailRecords.ssi_tnrid_nv = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]);
                            objMailRecords["ssi_tnrid_nv"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]);
                        }



                        //Secondary Owner First Name
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
                        {
                            //objMailRecords.ssi_secownerfname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]);
                            objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]);
                        }

                        //Secondary Owner Last Name
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
                        {
                            // objMailRecords.ssi_secownerlname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                            objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                        }


                        //Mailing Contact Type
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]) != "")
                        {
                            //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                            // objMailRecords.ssi_mailingcontacttype = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]);
                            objMailRecords["ssi_mailingcontacttype"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]);
                        }

                        if (ddlMailType.SelectedValue != "")
                        {
                            //objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                            objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                        }

                        //Fund Name
                        if (lstFund.SelectedValue != "")
                        {
                            // objMailRecords.ssi_fundname = lstFund.SelectedItem.Text; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                            objMailRecords["ssi_fundname"] = lstFund.SelectedItem.Text; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                        }

                        // CreatedByCustomid Field 

                        string Userid = GetcurrentUser();

                        if (Userid != "")
                        {
                            //objMailRecords.ssi_createdbycustomid = new Lookup();
                            //objMailRecords.ssi_createdbycustomid.Value = new Guid(Userid);

                            objMailRecords["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
                        }

                        service.Create(objMailRecords);
                        intResult++;
                        trBrowsefiles.Style.Add("display", "none");
                        trUnify.Style.Add("display", "none");
                    }


                    #region Update Mailing List

                    for (int j = 0; j < FundInfoContacts.Tables[1].Rows.Count; j++)
                    {
                        //objMailingList = new ssi_mailinglist();
                        Entity objMailingList = new Entity("ssi_mailinglist");
                        if (Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]) != "")
                        {
                            //objMailingList.ssi_mailinglistid = new Key();
                            //objMailingList.ssi_mailinglistid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]));
                            objMailingList["ssi_mailinglistid"] = new Guid(Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]));

                        }


                        if (txtLetterDate.Text != "")
                        {
                            //objMailingList.ssi_fundmailingdate = new CrmDateTime();
                            //objMailingList.ssi_fundmailingdate.Value = txtLetterDate.Text;
                            objMailingList["ssi_fundmailingdate"] = Convert.ToDateTime(txtLetterDate.Text);

                        }

                        service.Update(objMailingList);
                        intResult++;
                    }

                    #endregion


                    trBrowsefiles.Style.Add("display", "none");
                    trMonths.Style.Add("display", "none");
                    //RKLib.ExportData.Export objExport = new RKLib.ExportData.Export("Web");

                    //if (FundInfoContacts.Tables[0].Rows.Count > 0)
                    //{
                    //    objExport.ExportDetails(FundInfoContacts.Tables[0], RKLib.ExportData.Export.ExportFormat.CSV, ddlMailType.SelectedItem.Text + ".CSV");
                    //}
                    // Response.Write("Position Based Done" + DateTime.Now.ToString());
                    #endregion
                }
                // else if (Mail_type == "100000002")
                else if (Mail_type == "3")// changed field(MailType to existing field Mailingtype) in CRM-6_25_2019
                {
                    #region Recommendation  Based , Similar to Fund Info (Contacts for Confirmed Recommendations)
                    clsDB = new DB();
                    DataSet FundInfoContacts = new DataSet();

                    object AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
                    object FundId = lstFund.SelectedValue == "0" || lstFund.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFund) + "'";
                    if (chkEmailRecipients.Checked == true)
                    {
                        FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_RECOMMENDATIONS @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@IncEmailRecipients=1,@MailID='" + ddlMailType.SelectedValue + "'");
                    }
                    else
                    {
                        FundInfoContacts = clsDB.getDataSet("SP_S_FUND_INFO_RECOMMENDATIONS @AsOfDate=" + AsOfDate + ",@FundId=" + FundId + ",@IncEmailRecipients=0,@MailID='" + ddlMailType.SelectedValue + "'");
                    }


                    for (int i = 0; i < FundInfoContacts.Tables[0].Rows.Count; i++)
                    {

                        //objMailRecords = new ssi_mailrecords();
                        Entity objMailRecords = new Entity("ssi_mailrecords");

                        //Mail Type
                        if (ddlMailType.SelectedValue != "")
                        {
                            //objMailRecords.ssi_mailtypeid = new Lookup();
                            //objMailRecords.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();
                            //objMailRecords.ssi_mailtypeid.Value = new Guid("b46939f9-59dd-e011-ad4d-0019b9e7ee05");
                            objMailRecords["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));
                        }

                        //[Spouse Name]
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]) != "")
                        {
                            //objMailRecords.ssi_spousepart_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]);
                            objMailRecords["ssi_spousepart_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Spouse Name"]);

                        }

                        //MailingID
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["MailingID"]) != "")
                        {
                            //objMailRecords.ssi_mailingid = new CrmNumber();
                            //objMailRecords.ssi_mailingid.Value = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]);
                            objMailRecords["ssi_mailingid"] = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]);
                            Mailing_Id = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["MailingID"]).ToString();//added 6_11_2019
                        }

                        //Mail
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail"]) != "")
                        {
                            //objMailRecords.ssi_name = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail"]);
                            objMailRecords["ssi_name"] = ddlMailType.SelectedItem.Text;
                        }


                        //First Name
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]) != "")
                        {
                            //objMailRecords.ssi_ownerfirstname_hh_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]);
                            objMailRecords["ssi_ownerfirstname_hh_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner First Name"]);
                        }

                        //Last Name
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]) != "")
                        {
                            //objMailRecords.ssi_ownerlname_hh_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]);
                            objMailRecords["ssi_ownerlname_hh_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Owner Last Name"]);
                        }

                        //House Hold
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]) != "")
                        {
                            //objMailRecords.ssi_hholdinst_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]);
                            objMailRecords["ssi_hholdinst_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact HouseHold"]);
                        }

                        //Contact
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]) != "")
                        {
                            //objMailRecords.ssi_fullname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]);
                            objMailRecords["ssi_fullname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Contact Full Name"]);
                        }

                        //Address Line 1
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]) != "")
                        {
                            //objMailRecords.ssi_addressline1_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]);
                            objMailRecords["ssi_addressline1_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 1"]);
                        }

                        //Address Line 2
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]) != "")
                        {
                            //objMailRecords.ssi_addressline2_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]);
                            objMailRecords["ssi_addressline2_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 2"]);
                        }

                        //Address Line 3
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]) != "")
                        {
                            //objMailRecords.ssi_addressline3_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]);
                            objMailRecords["ssi_addressline3_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Address Line 3"]);
                        }

                        //City
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]) != "")
                        {
                            //objMailRecords.ssi_city_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]);
                            objMailRecords["ssi_city_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["City"]);
                        }

                        //State Or Province
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]) != "")
                        {
                            //objMailRecords.ssi_stateprovince_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]);
                            objMailRecords["ssi_stateprovince_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["State Or Province"]);
                        }

                        //Zip Code
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]) != "")
                        {
                            //objMailRecords.ssi_zipcode_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]);
                            objMailRecords["ssi_zipcode_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Zip Code"]);
                        }

                        //Country Or Region
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]) != "")
                        {
                            //objMailRecords.ssi_countryregion_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]);
                            objMailRecords["ssi_countryregion_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Country Or Region"]);
                        }

                        //Dear
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]) != "")
                        {
                            //objMailRecords.ssi_dear_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]);
                            objMailRecords["ssi_dear_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Dear"]);
                        }

                        //Mail Preference
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]) != "")
                        {
                            //objMailRecords.ssi_mailpreference_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]);
                            objMailRecords["ssi_mailpreference_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mail Preference"]);
                        }

                        //Salutation
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]) != "")
                        {
                            //objMailRecords.ssi_salutation_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]);
                            objMailRecords["ssi_salutation_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Salutation"]);
                        }

                        //ASOF DATE
                        if (txtAsofdate.Text != "")
                        {
                            //objMailRecords.ssi_asofdate = new CrmDateTime();
                            //objMailRecords.ssi_asofdate.Value = txtAsofdate.Text;
                            objMailRecords["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);
                        }

                        //Anziano ID
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]) != "")
                        {
                            //objMailRecords.ssi_anzianoid = new CrmNumber();
                            //objMailRecords.ssi_anzianoid.Value = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]);
                            objMailRecords["ssi_anzianoid"] = Convert.ToInt32(FundInfoContacts.Tables[0].Rows[i]["Anziano ID"]);
                        }

                        //TNR ID
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]) != "")
                        {
                            //objMailRecords.ssi_tnrid_nv = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]);
                            objMailRecords["ssi_tnrid_nv"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["TNR ID"]);
                        }


                        //Legal Entity Name
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Legal Entity Name"]) != "")
                        {
                            //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                            //objMailRecords.ssi_legalentity = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Legal Entity Name"]);
                            objMailRecords["ssi_legalentity"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Legal Entity Name"]);
                        }


                        //Mailing Contact Type
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]) != "")
                        {
                            //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                            //objMailRecords.ssi_mailingcontacttype = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]);
                            objMailRecords["ssi_mailingcontacttype"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Mailing Contact Type"]);
                        }

                        if (ddlMailType.SelectedValue != "")
                        {
                            //objMailRecords.ssi_mail = ddlMailType.SelectedItem.Text;
                            objMailRecords["ssi_mail"] = ddlMailType.SelectedItem.Text;
                        }

                        //HouseHold lookup
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"]) != "")
                        {
                            //objMailRecords.ssi_accountid = new Lookup();
                            //objMailRecords.ssi_accountid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"]));
                            objMailRecords["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["AccountId"])));
                        }

                        //Contact lookup
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"]) != "")
                        {
                            //objMailRecords.ssi_contactfullnameid = new Lookup();
                            //objMailRecords.ssi_contactfullnameid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"]));
                            objMailRecords["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ContactId"])));
                        }

                        //ssi_LegalEntityId lookup
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
                        {
                            ////objMailRecords.ssi_legalentity = new  = new CrmNumber();
                            //objMailRecords.ssi_legalentitynameid = new Lookup();
                            //objMailRecords.ssi_legalentitynameid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                            objMailRecords["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["ssi_LegalEntityId"])));
                        }

                        //Fund Name
                        if (lstFund.SelectedValue != "")
                        {
                            //objMailRecords.ssi_fundname = lstFund.SelectedItem.Text; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                            objMailRecords["ssi_fundname"] = lstFund.SelectedItem.Text;
                        }



                        //Secondary Owner First Name
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
                        {
                            //objMailRecords.ssi_secownerfname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]);
                            objMailRecords["ssi_secownerfname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner First Name"]);
                        }

                        //Secondary Owner Last Name
                        if (Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
                        {
                            //objMailRecords.ssi_secownerlname_mail = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                            objMailRecords["ssi_secownerlname_mail"] = Convert.ToString(FundInfoContacts.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                        }

                        // CreatedByCustomid Field 

                        string Userid = GetcurrentUser();

                        if (Userid != "")
                        {
                            //objMailRecords.ssi_createdbycustomid = new Lookup();
                            //objMailRecords.ssi_createdbycustomid.Value = new Guid(Userid);

                            objMailRecords["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
                        }

                        service.Create(objMailRecords);
                        intResult++;
                        trBrowsefiles.Style.Add("display", "none");
                    }


                    #region Update Mailing List

                    for (int j = 0; j < FundInfoContacts.Tables[1].Rows.Count; j++)
                    {
                        //objMailingList = new ssi_mailinglist();
                        Entity objMailingList = new Entity("ssi_mailinglist");

                        if (Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]) != "")
                        {
                            //objMailingList.ssi_mailinglistid = new Key();
                            //objMailingList.ssi_mailinglistid.Value = new Guid(Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]));
                            objMailingList["ssi_mailinglistid"] = new Guid(Convert.ToString(FundInfoContacts.Tables[1].Rows[j]["Ssi_MailingListID"]));
                        }


                        if (txtLetterDate.Text != "")
                        {
                            //objMailingList.ssi_fundmailingdate = new CrmDateTime();
                            //objMailingList.ssi_fundmailingdate.Value = txtLetterDate.Text;
                            objMailingList["ssi_fundmailingdate"] = Convert.ToDateTime(txtLetterDate.Text);
                        }

                        service.Update(objMailingList);
                        intResult++;
                    }

                    #endregion



                    trBrowsefiles.Style.Add("display", "none");
                    trMonths.Style.Add("display", "none");
                    trUnify.Style.Add("display", "none");
                    //RKLib.ExportData.Export objExport = new RKLib.ExportData.Export("Web");

                    //if (FundInfoContacts.Tables[0].Rows.Count > 0)
                    //{
                    //    objExport.ExportDetails(FundInfoContacts.Tables[0], RKLib.ExportData.Export.ExportFormat.CSV, ddlMailType.SelectedItem.Text + ".CSV");
                    //}
                    //   Response.Write("Recomandation Based Done" + DateTime.Now.ToString());
                    #endregion

                }
            }
            #endregion


            if (intResult > 0)
            {
                if (ddlMailType.SelectedValue == "a1a079a4-d7be-e011-a19b-0019b9e7ee05" || ddlMailType.SelectedValue == "6d7545da-8164-e111-bd8f-0019b9e7ee05" || ddlMailType.SelectedValue == "78612b2b-5add-e011-ad4d-0019b9e7ee05" || ddlMailType.SelectedValue == "81091a9b-2ae9-e011-9141-0019b9e7ee05" || ddlMailType.SelectedValue == "3cbaf86d-5edd-e011-ad4d-0019b9e7ee05")//Capital Call Letter
                {
                    if (MailId != "" && chkUnify.Checked == true)
                    {
                        if (bUnify)
                        {
                            UpdateMailRecords(MailId);
                            lblmailid.Visible = true;
                            lblmailid.Text = MailId;
                            ddlMailId.Visible = false;
                        }
                        else
                        {
                            lblmailid.Visible = true;
                            lblmailid.Text = MailId;
                        }
                    }
                    else
                    {
                        lblmailid.Visible = true;
                        lblmailid.Text = MailId;
                        //ddlMailId.Visible = false;
                    }
                }
                else
                {
                    if (ddlMailType.SelectedItem.Text != "")
                    {
                        if (Mailing_Id != "")
                        {
                            trMailID.Style.Add("display", "table-row");
                            lblmailid.Visible = true;
                            lblmailid.Text = Mailing_Id;
                            ddlMailId.Visible = false;
                        }
                        lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully.";
                        //lblError.Text = ddlMailType.SelectedItem.Text + " records Unified successfully";
                        
                        //added 4_23_2020 Exception Report
                        if (ddlMailType.SelectedValue == "3fb190d9-b2cd-e011-a19b-0019b9e7ee05") //Billing
                        {
                            string sql1 = "SP_S_BillingInvoiceException @AUMAsofDate='" + txtAsofdate.Text + "',@MailIDList=" + MailId;

                          //  Response.Write("SQL =" + sql1);
                            DataSet DSLegalEntitytemp1 = clsDB.getDataSet(sql1);
                            DataTable dtException = DSLegalEntitytemp1.Tables[0];
                            int rowCount = dtException.Rows.Count;
                            //Response.Write("rowCount =" + rowCount);
                            if (rowCount > 0)
                            {
                                string ExcelFilePath = GenerateExcel(DSLegalEntitytemp1);
                                if (ExcelFilePath != "")
                                {
                                    lbtnExceptionReport.Visible = true;
                                    ViewState["ExcetionReportPath"] = ExcelFilePath;
                                }


                            }
                        }
                    }
                }
                BindMailingId();

            }
            else
            {
                //added 4_23_2020 Exception Report
                if (ddlMailType.SelectedValue == "3fb190d9-b2cd-e011-a19b-0019b9e7ee05") //Billing
                {
                    string sql1 = "SP_S_BillingInvoiceException @AUMAsofDate='" + txtAsofdate.Text + "',@MailIDList=''";

                    //  Response.Write("SQL =" + sql1);
                    DataSet DSLegalEntitytemp1 = clsDB.getDataSet(sql1);
                    DataTable dtException = DSLegalEntitytemp1.Tables[0];
                    int rowCount = dtException.Rows.Count;
                    //Response.Write("rowCount =" + rowCount);
                    if (rowCount > 0)
                    {
                        string ExcelFilePath = GenerateExcel(DSLegalEntitytemp1);
                        if (ExcelFilePath != "")
                        {
                            lbtnExceptionReport.Visible = true;
                            ViewState["ExcetionReportPath"] = ExcelFilePath;
                        }


                    }
                }
            }

            if (lstSuccessFundName.Count > 0)
            {
                string isSuccess = string.Empty;
                string Success_FundShortName = string.Empty;
                lblSuccess.Text = "Files processed successfully for:";
                foreach (KeyValuePair<string, string> pair in lstSuccessFundName)
                {
                    Success_FundShortName = pair.Key.ToString(); // FilePath of Each File in Zip
                    isSuccess = pair.Value.ToString(); // fundShortName of each File Associated to
                    if (isSuccess.ToLower().ToString() == "success")
                    {
                        lblSuccess.Text = lblSuccess.Text + "<br/>" + Success_FundShortName;
                    }
                }
            }


            lblErrortxt.Text = strErrorOccured;
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

    private void Download_File(string FilePath, string FileName)
    {
        Response.ContentType = ContentType;
        Response.AppendHeader("Content-Disposition", "attachment; filename=" + FileName);
        Response.WriteFile(FilePath);
        Response.End();
    }
    public string GenerateExcel(DataSet ds)
    {
        string Server = AppLogic.GetParam(AppLogic.ConfigParam.Server);
        try
        {
            //string aod = txtAUMDate.Text;
            DateTime dAsofDate = DateTime.Now;



            if (!Directory.Exists(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder"))
            {
                Directory.CreateDirectory(HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder");
            }

            string strYear = DateTime.Now.Year.ToString().Length < 2 ? "0" + DateTime.Now.Year.ToString() : DateTime.Now.Year.ToString();
            string strMonth = DateTime.Now.Month.ToString().Length < 2 ? "0" + DateTime.Now.Month.ToString() : DateTime.Now.Month.ToString();
            string strDay = DateTime.Now.Day.ToString().Length < 2 ? "0" + DateTime.Now.Day.ToString() : DateTime.Now.Day.ToString();
            string strhour = DateTime.Now.Hour.ToString().Length < 2 ? "0" + DateTime.Now.Hour.ToString() : DateTime.Now.Hour.ToString();
            string strMin = DateTime.Now.Minute.ToString().Length < 2 ? "0" + DateTime.Now.Minute.ToString() : DateTime.Now.Minute.ToString();
            string strSec = DateTime.Now.Second.ToString().Length < 2 ? "0" + DateTime.Now.Second.ToString() : DateTime.Now.Second.ToString();
            string append_timestamp = strMonth + "_" + strDay + "_" + strYear + "_" + strhour + "_" + strMin + "_" + strSec;
            String lsFileNamforFinalXls = string.Empty;
            if (Server.ToLower() == "prod")
            {
                lsFileNamforFinalXls = "ExceptionReport" + "_" + append_timestamp;
            }
            else
            {
                lsFileNamforFinalXls = "ExceptionReport" + "_" + append_timestamp + "_test";
            }


            string ExcelFilePath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls + ".xlsx";

            if (System.IO.File.Exists(ExcelFilePath))
            {
                System.IO.File.Delete(ExcelFilePath);
            }


            if (ds.Tables.Count > 0)
            {

                #region Spire License Code
                string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
                Spire.License.LicenseProvider.SetLicenseKey(License);
                Spire.License.LicenseProvider.LoadLicense();
                #endregion

                // string SheetNme = ds.Tables[0].Rows[0][0].ToString();

                Workbook book = new Workbook();
                book.Version = ExcelVersion.Version2016;
                Worksheet sheet = book.Worksheets[0];
                //  sheet.Name = SheetNme;
                sheet.Range[1, 1, 1, ds.Tables[0].Columns.Count].Style.Font.IsBold = true;

                sheet.InsertDataTable(ds.Tables[0], true, 1, 1);

                sheet.Range[2, 12, ds.Tables[0].Rows.Count + 1, 12].NumberFormat = "$ #,##0.00_);($ #,##0.00)";
                sheet.Range[2, 13, ds.Tables[0].Rows.Count + 1, 13].NumberFormat = "$ #,##0.00_);($ #,##0.00)";
                sheet.Range[2, 14, ds.Tables[0].Rows.Count + 1, 14].NumberFormat = "$ #,##0.00_);($ #,##0.00)";

                sheet.Range[1, 1, ds.Tables[0].Rows.Count + 1, ds.Tables[0].Columns.Count].AutoFitColumns();
                sheet.Range[1, 1, ds.Tables[0].Rows.Count + 1, ds.Tables[0].Columns.Count].Style.HorizontalAlignment = HorizontalAlignType.Center;

                book.SaveToFile(ExcelFilePath);
                //string vContain = "Excel Report Generated Succesfully ";
            }
            return ExcelFilePath;

        }
        catch (Exception e)
        {

            //string vContain = "Excel Report Genration Fail,  Error " + e.ToString();

            return "";
        }
    }
    private SqlConnection OpenConnection()
    {
        try
        {
            //"Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";
            //"Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=TransactionLoad_DB;Data Source=SQL01";
            string ConnString = AppLogic.GetParam(AppLogic.ConfigParam.DBTransactions);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=TransactionLoad_DB;Data Source=sql01";
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



    public void updateMailingList(IOrganizationService service, DataSet ds)
    {

    }


    public void Proces_ALL_FUND(string TempPath, IOrganizationService service, string MailType)
    {
        clsDB = new DB();
        //foreach (System.Web.UI.WebControls.ListItem li in lstFund.Items)
        foreach (KeyValuePair<string, string> pair in lstFile)
        {
            int retVal = 1;
            string FundShortName = string.Empty;
            string FilePath = string.Empty;

            string SelectedFund = string.Empty;
            string SelectedFundValue = string.Empty;
            string strFileName = string.Empty;
            string SheetName = string.Empty;
            string TypeId = string.Empty;
            bool bProcess = false;

            if (MailType.ToLower() == "a1a079a4-d7be-e011-a19b-0019b9e7ee05" || MailType.ToLower() == "81091a9b-2ae9-e011-9141-0019b9e7ee05")//Cap Call Letter - Cap Call Wire
            {
                strFileName = clsDB.FileName("capitalcall");
                SheetName = "Capital Call";
                TypeId = "11";
            }
            else if (MailType.ToLower() == "6d7545da-8164-e111-bd8f-0019b9e7ee05" || MailType.ToLower() == "78612b2b-5add-e011-ad4d-0019b9e7ee05")//Fund Distribution Letter - Fund Distribution  
            {
                strFileName = clsDB.FileName("funddistribution");
                SheetName = "Distribution";
                TypeId = "10";
            }



            FilePath = pair.Key.ToString(); // FilePath of Each File in Zip
            FundShortName = pair.Value.ToString(); // fundShortName of each File Associated to


            string Ssi_ShortName = string.Empty;
            DataTable dtFund = (DataTable)ViewState["dtFund"];
            for (int i = 0; i < dtFund.Rows.Count; i++)
            {
                Ssi_ShortName = Convert.ToString(dtFund.Rows[i]["Ssi_ShortName"]);
                SelectedFund = Convert.ToString(dtFund.Rows[i]["ssi_name"]);
                SelectedFundValue = Convert.ToString(dtFund.Rows[i]["ssi_FundId"]);
                //if (SelectedFund == ssi_name)
                //{
                //    break;
                //}
                if (Ssi_ShortName == FundShortName)
                {
                    bProcess = true;
                    break;
                }
                else
                {
                    bProcess = false;
                }
            }
            try
            {


                if (bProcess)
                {
                    bool bCopy = ReadAlPsFile(FilePath, TempPath + strFileName, SheetName);//CleanUp the File from Alps
                    if (bCopy)
                    {


                        if (File.Exists(DTSFilePath + strFileName))
                        {
                            File.Delete(DTSFilePath + strFileName); // Delete File from DTS PAth

                            File.Move(TempPath + strFileName, DTSFilePath + strFileName);//Move File to DTS PAth
                        }
                        else
                        {
                            File.Move(TempPath + strFileName, DTSFilePath + strFileName);//Move File to DTS PAth
                        }

                        #region JOBCALL
                        try
                        {

                            SqlConnection Gresham_con = new SqlConnection(con);

                            SqlCommand cmd = new SqlCommand();
                            SqlDataAdapter dagersham = new SqlDataAdapter();

                            DataSet ds_gresham = new DataSet();
                            DataSet ds = new DataSet();

                            Gresham_con = new SqlConnection(con);
                            Gresham_con.Open();
                            string greshamquery = "SP_S_ExecuteJobs @TypeId = " + TypeId;
                            cmd = new SqlCommand();
                            cmd.Connection = Gresham_con;
                            cmd.CommandText = greshamquery;

                            cmd.ExecuteNonQuery();
                            retVal = 0;
                        }
                        catch (Exception exception3)
                        {
                            retVal = 1;
                            strErrorOccured = strErrorOccured + "<br/>Error Occurred in File for Fund :" + SelectedFund;// + "<br/>" + exception3.Message.ToString();
                        }

                        #endregion

                        if (retVal == 0)
                        {
                            if (MailType.ToLower() == "a1a079a4-d7be-e011-a19b-0019b9e7ee05")
                            {
                                Create_MailRecordsTemp_CapitalCallLetter(service, SelectedFund, SelectedFundValue);
                            }
                            else if (MailType.ToLower() == "81091a9b-2ae9-e011-9141-0019b9e7ee05")
                            {
                                Create_MailRecordsTEmp_CapCallWire(service, SelectedFund, SelectedFundValue);
                            }
                            else if (MailType.ToLower() == "78612b2b-5add-e011-ad4d-0019b9e7ee05")
                            {
                                Create_MailRecordsTemp_FundDistribution(service, SelectedFund, SelectedFundValue);
                            }
                            else if (MailType.ToLower() == "6d7545da-8164-e111-bd8f-0019b9e7ee05")
                            {
                                Create_MailRecordsTemp_FundDistributionLetter(service, SelectedFund, SelectedFundValue);
                            }
                            lstSuccessFundName.Add(FundShortName, "Success");
                        }
                        else if (retVal == 1)
                        {
                            lblError.Text = "File Upload Failed";
                        }

                        //break;

                    }
                    else
                    {
                        lblError.Text = "Error Occured in File Process";
                    }
                }
                else
                {
                    strErrorOccured = strErrorOccured + "<br/>Fund Not Found :" + FundShortName;
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
                //if (sqlconn != null)
                //    if (sqlconn.State != System.Data.ConnectionState.Open)
                //        sqlconn.Close();
            }

        }
    }
    public void Proces_Selected_FUND(string TempPath, IOrganizationService service, string MailType)
    {
        clsDB = new DB();

        foreach (System.Web.UI.WebControls.ListItem li in lstFund.Items)
        {
            int retVal = 1;
            string FundShortName = string.Empty;
            string FilePath = string.Empty;
            // string FileName = Path.GetFileName(TempPath);
            // string FinalFilePath = TempPath.Replace(FileName, "");
            string SelectedFund = string.Empty;
            string SelectedFundValue = string.Empty;
            bool bProcess = false;
            string strFileName = string.Empty;
            string SheetName = string.Empty;
            string TypeId = string.Empty;



            if (li.Selected)
            {
                if (MailType.ToLower() == "a1a079a4-d7be-e011-a19b-0019b9e7ee05" || MailType.ToLower() == "81091a9b-2ae9-e011-9141-0019b9e7ee05")//Cap Call Letter - Cap Call Wire
                {
                    strFileName = clsDB.FileName("capitalcall");
                    SheetName = "Capital Call";
                    TypeId = "11";
                }
                else if (MailType.ToLower() == "6d7545da-8164-e111-bd8f-0019b9e7ee05" || MailType.ToLower() == "78612b2b-5add-e011-ad4d-0019b9e7ee05")//Fund Distribution - Fund Distribution  Letter
                {
                    strFileName = clsDB.FileName("funddistribution");
                    SheetName = "Distribution";
                    TypeId = "10";
                }

                SelectedFund = li.Text;
                SelectedFundValue = li.Value;
                //  SelectedFund = lstFund.SelectedItem.Value;
                string Ssi_ShortName = string.Empty;
                DataTable dtFund = (DataTable)ViewState["dtFund"];
                for (int i = 0; i <= dtFund.Rows.Count; i++)
                {
                    // Ssi_ShortName = Convert.ToString(dtFund.Rows[i]["Ssi_ShortName"]);
                    string ssi_name = Convert.ToString(dtFund.Rows[i]["ssi_name"]);
                    if (SelectedFund == ssi_name)
                    {
                        Ssi_ShortName = Convert.ToString(dtFund.Rows[i]["Ssi_ShortName"]);
                        break;
                    }
                }
                try
                {
                    // string FileFundName = pair.Value.ToString();
                    foreach (KeyValuePair<string, string> pair in lstFile)
                    {
                        FilePath = pair.Key.ToString();
                        FundShortName = pair.Value.ToString();
                        if (Ssi_ShortName == FundShortName)
                        {
                            bProcess = true;
                            break;
                        }
                        else
                        {
                            bProcess = false;
                        }
                    }
                    if (bProcess)
                    {
                        bool bCopy = ReadAlPsFile(FilePath, TempPath + strFileName, SheetName);//CleanUp the File from Alps
                        if (bCopy)
                        {

                            if (File.Exists(DTSFilePath + strFileName))
                            {
                                File.Delete(DTSFilePath + strFileName); // Delete File from DTS PAth

                                File.Move(TempPath + strFileName, DTSFilePath + strFileName);//Copy File to DTS PAth
                            }
                            else
                            {
                                File.Move(TempPath + strFileName, DTSFilePath + strFileName);//Copy File to DTS PAth
                            }

                            #region JOBCALL
                            try
                            {

                                SqlConnection Gresham_con = new SqlConnection(con);

                                SqlCommand cmd = new SqlCommand();
                                SqlDataAdapter dagersham = new SqlDataAdapter();

                                DataSet ds_gresham = new DataSet();
                                DataSet ds = new DataSet();

                                Gresham_con = new SqlConnection(con);
                                Gresham_con.Open();
                                string greshamquery = "SP_S_ExecuteJobs @TypeId = " + TypeId;
                                cmd = new SqlCommand();
                                cmd.Connection = Gresham_con;
                                cmd.CommandText = greshamquery;

                                cmd.ExecuteNonQuery();
                                retVal = 0;
                            }
                            catch (Exception exception3)
                            {
                                retVal = 1;
                                strErrorOccured = strErrorOccured + "<br/>Error Occurred in File for Fund :" + SelectedFund;//+ "<br/>" + exception3.Message.ToString();
                            }

                            #endregion

                            if (retVal == 0)
                            {
                                if (MailType.ToLower() == "a1a079a4-d7be-e011-a19b-0019b9e7ee05")
                                {
                                    Create_MailRecordsTemp_CapitalCallLetter(service, SelectedFund, SelectedFundValue);
                                }
                                else if (MailType.ToLower() == "81091a9b-2ae9-e011-9141-0019b9e7ee05")
                                {
                                    Create_MailRecordsTEmp_CapCallWire(service, SelectedFund, SelectedFundValue);
                                }
                                else if (MailType.ToLower() == "78612b2b-5add-e011-ad4d-0019b9e7ee05")
                                {
                                    Create_MailRecordsTemp_FundDistribution(service, SelectedFund, SelectedFundValue);
                                }
                                else if (MailType.ToLower() == "6d7545da-8164-e111-bd8f-0019b9e7ee05")
                                {
                                    Create_MailRecordsTemp_FundDistributionLetter(service, SelectedFund, SelectedFundValue);
                                }
                                lstSuccessFundName.Add(FundShortName, "Success");
                            }
                            else if (retVal == 1)
                            {
                                lblError.Text = "File Upload Failed";
                                strErrorOccured = strErrorOccured + "<br/>File Upload Failed For Fund :" + SelectedFund;
                            }

                            //break;

                        }
                        else
                        {
                            lblError.Text = "Error Occured in File Process";
                        }
                    }
                    else
                    {
                        strErrorOccured = strErrorOccured + "<br/>File Not Found For Fund :" + SelectedFund;
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
                    //if (sqlconn != null)
                    //    if (sqlconn.State != System.Data.ConnectionState.Open)
                    //        sqlconn.Close();
                }
            }
        }
    }
    public int Create_MailRecordsTemp_FundDistributionLetter(IOrganizationService service, string SelectedFund, string SelectedFundValue)
    {
        clsDB = new DB();
        object FundId = "'" + SelectedFundValue + "'";// lstFund.SelectedValue == "" || lstFund.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFund) + "'";
        object LegalEntityId = lstLegalEntity.SelectedValue == "" || lstLegalEntity.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstLegalEntity) + "'";
        object AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
        DataSet loDataset = new DataSet();
        if (chkEmailRecipients.Checked == true)
        {
            loDataset = clsDB.getDataSet("SP_S_FUND_DISTRIBUTION_Letter @FundIdNmbList=" + FundId + ",@AsOfDate=" + AsOfDate + ",@LegalEntityIdNmbList=" + LegalEntityId + ",@IncEmailRecipients=1");
        }
        else
        {
            loDataset = clsDB.getDataSet("SP_S_FUND_DISTRIBUTION_Letter @FundIdNmbList=" + FundId + ",@AsOfDate=" + AsOfDate + ",@LegalEntityIdNmbList=" + LegalEntityId + ",@IncEmailRecipients=0");
        }

        //Response.Write("No. Of Rows" + loDataset.Tables[0].Rows.Count.ToString() + "<br/><br/><br/>");

        for (int i = 0; i < loDataset.Tables[0].Rows.Count; i++)
        {

            //Response.Write(loDataset.Tables[0].Rows.Count.ToString() + "<br/><br/><br/>");

            // objMailRecordsTemp = new ssi_mailrecordstemp();
            Entity objMailRecordsTemp = new Entity("ssi_mailrecordstemp");
            //Mail Type
            //objMailRecordsTemp.ssi_mailtypeid = new Lookup();
            //objMailRecordsTemp.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();//3FB190D9-B2CD-E011-A19B-0019B9E7EE05
            //objMailRecordsTemp.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue); //new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_mailid"]));
            objMailRecordsTemp["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));


            //Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]) != "")
            {
                // objMailRecordsTemp.ssi_name = Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);
                objMailRecordsTemp["ssi_name"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);

            }

            //[Spouse Name]
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]) != "")
            {
                //objMailRecordsTemp.ssi_spousepart_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]);
                objMailRecordsTemp["ssi_spousepart_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]);
            }

            //Ssi_AnzianoID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]) != "")
            {
                //objMailRecordsTemp.ssi_anzianoid = new CrmNumber();
                //objMailRecordsTemp.ssi_anzianoid.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]);
                objMailRecordsTemp["ssi_anzianoid"] = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]);

            }

            //Total Commitment
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Total Commitment"]) != "")
            {
                //objMailRecordsTemp.ssi_totalcommitment_db = new CrmMoney();
                //objMailRecordsTemp.ssi_totalcommitment_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Total Commitment"]));
                objMailRecordsTemp["ssi_totalcommitment_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Total Commitment"]));

            }

            //Capital Distribution
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Capital Distribution"]) != "")
            {
                //objMailRecordsTemp.ssi_capitaldistribution_db = new CrmMoney();
                //objMailRecordsTemp.ssi_capitaldistribution_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Capital Distribution"]));
                objMailRecordsTemp["ssi_capitaldistribution_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Capital Distribution"]));
            }

            //Current Distribution - Percent
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Current Distribution - Percent"]) != "")
            {
                //objMailRecordsTemp.ssi_curdistp_db = new CrmDecimal();
                //objMailRecordsTemp.ssi_curdistp_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Current Distribution - Percent"]));
                objMailRecordsTemp["ssi_curdistp_db"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Current Distribution - Percent"]);

            }

            //Prior Distributions
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Prior Distributions"]) != "")
            {
                //objMailRecordsTemp.ssi_priordistributions_db = new CrmMoney();
                //objMailRecordsTemp.ssi_priordistributions_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Prior Distributions"]));
                objMailRecordsTemp["ssi_priordistributions_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Prior Distributions"]));
            }


            //Prior Distributions - Percent
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Prior Distributions - Percent"]) != "")
            {
                //objMailRecordsTemp.ssi_priordistp_db = new CrmDecimal();
                //objMailRecordsTemp.ssi_priordistp_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Prior Distributions - Percent"]));
                objMailRecordsTemp["ssi_priordistp_db"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Prior Distributions - Percent"]);
            }

            //Distributed to date
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Distributed to date"]) != "")
            {
                //objMailRecordsTemp.ssi_disttodate_db = new CrmMoney();
                //objMailRecordsTemp.ssi_disttodate_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Distributed to date"]));
                objMailRecordsTemp["ssi_disttodate_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Distributed to date"]));
            }

            //Distributed to Date - Percent
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Distributed to Date - Percent"]) != "")
            {
                //objMailRecordsTemp.ssi_distdatep_db = new CrmDecimal();
                //objMailRecordsTemp.ssi_distdatep_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Distributed to Date - Percent"]));
                objMailRecordsTemp["ssi_distdatep_db"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Distributed to Date - Percent"]);
            }

            //Called to date
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Called to date"]) != "")
            {
                //objMailRecordsTemp.ssi_calledtodate_db = new CrmMoney();
                //objMailRecordsTemp.ssi_calledtodate_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Called to date"]));
                objMailRecordsTemp["ssi_calledtodate_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Called to date"]));
            }

            //Remaining Commitment
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Remaining Commitment"]) != "")
            {
                //objMailRecordsTemp.ssi_remainingcommitment_db = new CrmMoney();
                //objMailRecordsTemp.ssi_remainingcommitment_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Remaining Commitment"]));
                objMailRecordsTemp["ssi_remainingcommitment_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Remaining Commitment"]));
            }

            //Percent Called
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Percent Called"]) != "")
            {
                //objMailRecordsTemp.ssi_percentcalled_db = Convert.ToString(loDataset.Tables[0].Rows[i]["Percent Called"]);
                objMailRecordsTemp["ssi_percentcalled_db"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Percent Called"]);
            }

            //Class B fee adjustment
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Class B fee adjustment"]) != "")
            {
                //objMailRecordsTemp.ssi_feeadj_db = new CrmMoney();
                //objMailRecordsTemp.ssi_feeadj_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Class B fee adjustment"]));
                objMailRecordsTemp["ssi_feeadj_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Class B fee adjustment"]));

            }


            //Actual Cash distributions
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Actual Cash distributions"]) != "")
            {
                //objMailRecordsTemp.ssi_actualcashdistributions_db = new CrmMoney();
                //objMailRecordsTemp.ssi_actualcashdistributions_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Actual Cash distributions"]));
                objMailRecordsTemp["ssi_actualcashdistributions_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Actual Cash distributions"]));
            }


            //DDA #
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["DDA"]) != "")
            {
                //objMailRecordsTemp.ssi_outdda_household = Convert.ToString(loDataset.Tables[0].Rows[i]["DDA"]);
                objMailRecordsTemp["ssi_outdda_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["DDA"]);
            }

            ////Legal Entity Name
            //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Name"]) != "")
            //{
            //    objMailRecords.ssi_legalentity = Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Name"]);
            //}

            //TNR ID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]) != "")
            {
                //  objMailRecordsTemp.ssi_tnrid_nv = Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]);
                objMailRecordsTemp["ssi_tnrid_nv"] = Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]);
            }

            //Mailing Contact Type
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]) != "")
            {
                // objMailRecordsTemp.ssi_mailingcontacttype = Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]);
                objMailRecordsTemp["ssi_mailingcontacttype"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]);
            }

            //Custodian Account Number
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account Number"]) != "")
            {
                // objMailRecordsTemp.ssi_accountnumber = Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account Number"]);
                objMailRecordsTemp["ssi_accountnumber"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account Number"]);
            }

            //Account Name1
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Account Name1"]) != "")
            {
                //objMailRecordsTemp.ssi_accountname1 = Convert.ToString(loDataset.Tables[0].Rows[i]["Account Name1"]);
                objMailRecordsTemp["ssi_accountname1"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Account Name1"]);
            }

            //Legal Entity Bank
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Bank"]) != "")
            {
                //objMailRecordsTemp.ssi_legalentitybank = Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Bank"]);
                objMailRecordsTemp["ssi_legalentitybank"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Bank"]);
            }

            //Basic Wire Info /*Distribution tab*/
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]) != "")
            {
                // objMailRecordsTemp.ssi_ssi_basicwireinfo_household = Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]);
                objMailRecordsTemp["ssi_ssi_basicwireinfo_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]);
            }

            //Basic Wire Info /*Capital Call Letter tab*/
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]) != "")
            {
                //objMailRecordsTemp.ssi_ssi_basicwireinfo_household1 = Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]);
                objMailRecordsTemp["ssi_ssi_basicwireinfo_household1"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]);
            }

            //ABA Routing # /*Distribution tab*/
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]) != "")
            {
                //objMailRecordsTemp.ssi_outabarouting_household = Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]);
                objMailRecordsTemp["ssi_outabarouting_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]);
            }

            //ABA Routing # /*Capital Call Letter tab*/
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]) != "")
            {
                //objMailRecordsTemp.ssi_abarouting_household = Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]);
                objMailRecordsTemp["ssi_abarouting_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]);
            }

            //For Further Credit (FFC) Acct #
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Acct #"]) != "")
            {
                //objMailRecordsTemp.ssi_ffcacct_household = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Acct #"]);
                objMailRecordsTemp["ssi_ffcacct_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Acct #"]);
            }

            //For Further Credit (FFC) Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Name"]) != "")
            {
                //objMailRecordsTemp.ssi_ffcname_household = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Name"]);
                objMailRecordsTemp["ssi_ffcname_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Name"]);
            }

            //Response.Write(Convert.ToString(loDataset.Tables[0].Rows[i]["Other Wire Instructions"]));

            //Other Wire Instructions
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Other Wire Instructions"]) != "")
            {
                //objMailRecordsTemp.ssi_otherwireinstr_household = Convert.ToString(loDataset.Tables[0].Rows[i]["Other Wire Instructions"]);
                objMailRecordsTemp["ssi_otherwireinstr_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Other Wire Instructions"]);
            }


            //Fund Name
            if (SelectedFund != "")
            {
                //objMailRecordsTemp.ssi_fundname = lstFund.SelectedItem.Text; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                objMailRecordsTemp["ssi_fundname"] = SelectedFund; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
            }

            //Fund Name (to show on wire instrux)
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name (to show on wire instrux)"]) != "")
            {
                // objMailRecordsTemp.ssi_fundname_fund = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name (to show on wire instrux)"]);
                objMailRecordsTemp["ssi_fundname_fund"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name (to show on wire instrux)"]);
            }

            //Fund Bank
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Bank"]) != "")
            {
                // objMailRecordsTemp.ssi_fundbank = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Bank"]);
                objMailRecordsTemp["ssi_fundbank"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Bank"]);
            }

            //Fund Account Number
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Account Number"]) != "")
            {
                // objMailRecordsTemp.ssi_fundaccountnumber = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Account Number"]);
                objMailRecordsTemp["ssi_fundaccountnumber"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Account Number"]);
            }


            //Distribution Payment Method
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Payment Method"]) != "")
            {
                //objMailRecordsTemp.ssi_distributionpaymentmethod = Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Payment Method"]);
                objMailRecordsTemp["ssi_distributionpaymentmethod"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Payment Method"]);
            }


            //Distribution Method Note
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Method Note"]) != "")
            {
                //objMailRecordsTemp.ssi_distributionmethodnote = Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Method Note"]);
                objMailRecordsTemp["ssi_distributionmethodnote"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Method Note"]);
            }

            //SLOAFlg
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["SLOAFlg"]) != "")
            {
                //objMailRecordsTemp.ssi_sloaflg = new CrmBoolean();
                //objMailRecordsTemp.ssi_sloaflg.Value = Convert.ToBoolean(loDataset.Tables[0].Rows[i]["SLOAFlg"]);
                objMailRecordsTemp["ssi_sloaflg"] = Convert.ToBoolean(Convert.ToString(loDataset.Tables[0].Rows[i]["SLOAFlg"]).ToLower());

            }



            //AsofDate
            if (txtAsofdate.Text != "")
            {
                //objMailRecordsTemp.ssi_asofdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_asofdate.Value = txtAsofdate.Text;
                objMailRecordsTemp["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);

            }

            //MailingID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["MailingID"]) != "")
            {
                //objMailRecordsTemp.ssi_mailingid = new CrmNumber();
                //objMailRecordsTemp.ssi_mailingid.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["MailingID"]);
                objMailRecordsTemp["ssi_mailingid"] = Convert.ToInt32(loDataset.Tables[0].Rows[i]["MailingID"]);

            }

            //Mail
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]) != "")
            {
                //objMailRecordsTemp.ssi_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]);
                objMailRecordsTemp["ssi_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]);
            }

            //Owner First Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]) != "")
            {
                //  objMailRecordsTemp.ssi_ownerfirstname_hh_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]);
                objMailRecordsTemp["ssi_ownerfirstname_hh_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]);
            }

            //Owner Last Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]) != "")
            {
                //objMailRecordsTemp.ssi_ownerlname_hh_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]);
                objMailRecordsTemp["ssi_ownerlname_hh_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]);
            }

            //Contact Household
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]) != "")
            {
                //objMailRecordsTemp.ssi_hholdinst_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]);
                objMailRecordsTemp["ssi_hholdinst_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]);
            }

            //Contact Full Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]) != "")
            {
                // objMailRecordsTemp.ssi_fullname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]);
                objMailRecordsTemp["ssi_fullname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]);
            }


            //Secondary Owner First Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
            {
                //objMailRecordsTemp.ssi_secownerfname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]);
                objMailRecordsTemp["ssi_secownerfname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]);
            }

            //Secondary Owner Last Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
            {
                // objMailRecordsTemp.ssi_secownerlname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                objMailRecordsTemp["ssi_secownerlname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]);
            }


            //Address Line 1
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]) != "")
            {
                //objMailRecordsTemp.ssi_addressline1_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]);
                objMailRecordsTemp["ssi_addressline1_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]);
            }

            //Address Line 2
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]) != "")
            {
                // objMailRecordsTemp.ssi_addressline2_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]);
                objMailRecordsTemp["ssi_addressline2_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]);
            }

            //Address Line 3
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]) != "")
            {
                //objMailRecordsTemp.ssi_addressline3_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]);
                objMailRecordsTemp["ssi_addressline3_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]);
            }

            //City
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["City"]) != "")
            {
                // objMailRecordsTemp.ssi_city_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["City"]);
                objMailRecordsTemp["ssi_city_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["City"]);
            }

            //State or Province
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]) != "")
            {
                //objMailRecordsTemp.ssi_stateprovince_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]);
                objMailRecordsTemp["ssi_stateprovince_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]);
            }

            //ZIP Code
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]) != "")
            {
                // objMailRecordsTemp.ssi_zipcode_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]);
                objMailRecordsTemp["ssi_zipcode_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]);
            }

            //Country or Region
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]) != "")
            {
                //objMailRecordsTemp.ssi_countryregion_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]);
                objMailRecordsTemp["ssi_countryregion_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]);
            }

            //Dear
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]) != "")
            {
                // objMailRecordsTemp.ssi_dear_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]);
                objMailRecordsTemp["ssi_dear_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]);
            }

            //Salutation
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]) != "")
            {
                //   objMailRecordsTemp.ssi_salutation_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]);
                objMailRecordsTemp["ssi_salutation_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]);
            }

            //Mail Preference 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]) != "")
            {
                //objMailRecordsTemp.ssi_mailpreference_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]);
                objMailRecordsTemp["ssi_mailpreference_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]);
            }

            //ssi_mailStatus 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]) != "")
            {
                //objMailRecordsTemp.ssi_mailstatus = new Picklist();
                //objMailRecordsTemp.ssi_mailstatus.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]);
                objMailRecordsTemp["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]));

            }


            //HouseHold lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"]) != "")
            {
                //objMailRecordsTemp.ssi_accountid = new Lookup();
                //objMailRecordsTemp.ssi_accountid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"]));
                objMailRecordsTemp["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"])));

            }

            //Contact lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"]) != "")
            {
                //objMailRecordsTemp.ssi_contactfullnameid = new Lookup();
                //objMailRecordsTemp.ssi_contactfullnameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"]));
                objMailRecordsTemp["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"])));
            }

            //ssi_LegalEntityId lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
            {
                //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                //objMailRecordsTemp.ssi_legalentitynameid = new Lookup();
                //objMailRecordsTemp.ssi_legalentitynameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                objMailRecordsTemp["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"])));
            }


            //Advisor Approval Required
            // objMailRecordsTemp.ssi_advisorapprovalreqd = new CrmBoolean();
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "0" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "False".ToUpper())
            {
                // objMailRecordsTemp.ssi_advisorapprovalreqd.Value = false;
                objMailRecordsTemp["ssi_advisorapprovalreqd"] = false;
            }
            else if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "1" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "True".ToUpper())
            {
                objMailRecordsTemp["ssi_advisorapprovalreqd"] = true;
            }

            //File Name Added on 9 oct 2014
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]) != "")
            {
                //objMailRecordsTemp.ssi_filename = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]);
                objMailRecordsTemp["ssi_filename"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]);
            }
            //added by sasmit 5_3_2017
            //clientportalname 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]) != "")
            {
                //objMailRecordsTemp.ssi_clientportalname = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]);
                objMailRecordsTemp["ssi_clientportalname"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]);
            }
            //added by sasmit 5_3_2017
            //clientreportfolder
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]) != "")
            {
                // objMailRecordsTemp.ssi_clientreportfolder = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]);
                objMailRecordsTemp["ssi_clientreportfolder"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]);
            }
            // CreatedByCustomid Field 
            //Rohit Pawar
            string Userid = GetcurrentUser();

            if (Userid != "")
            {
                //objMailRecordsTemp.ssi_createdbycustomid = new Lookup();
                //objMailRecordsTemp.ssi_createdbycustomid.Value = new Guid(Userid);

                objMailRecordsTemp["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
            }

            #region Logic for Capital Call Approval Process

            // Logic for Capital Call Approval Process

            //if (chkUnify.Checked == true)
            //{
            //    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
            //    //objMailRecordsTemp.ssi_unifiedflg.Value = true;
            //    objMailRecordsTemp["ssi_unifiedflg"] = true;
            //}
            //else if (chkUnify.Checked == false)
            //{
            //    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
            //    //objMailRecordsTemp.ssi_unifiedflg.Value = false;
            //    objMailRecordsTemp["ssi_unifiedflg"] = false;
            //}

            if (MailId != "" && MailId != "0")
            {
                //objMailRecordsTemp.ssi_mailidtemp = new CrmNumber();
                //objMailRecordsTemp.ssi_mailidtemp.Value = Convert.ToInt32(MailId);
                objMailRecordsTemp["ssi_mailidtemp"] = Convert.ToInt32(MailId);

            }


            if (ddlTemplates.SelectedValue != "" && ddlTemplates.SelectedValue != "0")
            {
                //objMailRecordsTemp.ssi_templateid = new Lookup();
                //objMailRecordsTemp.ssi_templateid.Value = new Guid(ddlTemplates.SelectedValue);
                objMailRecordsTemp["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(ddlTemplates.SelectedValue));
            }


            //Wire AsofDate
            if (txtWireAsofDate.Text != "")
            {
                //objMailRecordsTemp.ssi_wireasofdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_wireasofdate.Value = txtWireAsofDate.Text;
                objMailRecordsTemp["ssi_wireasofdate"] = Convert.ToDateTime(txtWireAsofDate.Text);

            }

            //Letter AsofDate
            if (txtLetterDate.Text != "")
            {
                //objMailRecordsTemp.ssi_letterdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_letterdate.Value = txtLetterDate.Text;
                objMailRecordsTemp["ssi_letterdate"] = Convert.ToDateTime(txtLetterDate.Text);
            }


            #endregion


            if (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0" && chkUnify.Checked == false)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
            else if (ddlMailId.SelectedValue != "" && chkUnify.Checked == true)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
            else if (ddlMailId.SelectedValue != "" && chkUnify.Checked == false)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
        }


        #region Update Mailing List

        for (int j = 0; j < loDataset.Tables[1].Rows.Count; j++)
        {
            // objMailingList = new ssi_mailinglist();
            Entity objMailingList = new Entity("ssi_mailinglist");
            if (Convert.ToString(loDataset.Tables[1].Rows[j]["Ssi_MailingListID"]) != "")
            {
                //objMailingList.ssi_mailinglistid = new Key();
                //objMailingList.ssi_mailinglistid.Value = new Guid(Convert.ToString(loDataset.Tables[1].Rows[j]["Ssi_MailingListID"]));
                objMailingList["ssi_mailinglistid"] = new Guid(Convert.ToString(loDataset.Tables[1].Rows[j]["Ssi_MailingListID"]));


            }


            if (txtWireAsofDate.Text != "")
            {
                //objMailingList.ssi_distributiondate = new CrmDateTime();
                //objMailingList.ssi_distributiondate.Value = txtWireAsofDate.Text;
                objMailingList["ssi_distributiondate"] = Convert.ToDateTime(txtWireAsofDate.Text);

            }

            service.Update(objMailingList);
            intResult++;
        }

        #endregion



        if (intResult > 0)
        {
            Success++;
            //lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully";
        }

        if (loDataset.Tables[0].Rows.Count == 0 && loDataset.Tables[1].Rows.Count == 0 && chkUnify.Checked == false && (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0"))
            // lblError.Text = "No Records Found";// commented/changed 27_08_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[0].Rows.Count == 0 && loDataset.Tables[1].Rows.Count == 0 && chkUnify.Checked == false && (ddlMailId.SelectedValue != "" || ddlMailId.SelectedValue != "0"))
            //  lblError.Text = "No Records Found";// commented/changed 27_08_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[0].Rows.Count == 0 && loDataset.Tables[1].Rows.Count == 0 && chkUnify.Checked == true && (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0"))
            //  lblError.Text = "No Records Found";// commented/changed 27_08_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[0].Rows.Count == 0 && loDataset.Tables[1].Rows.Count == 0 && chkUnify.Checked == true && (ddlMailId.SelectedValue != "" || ddlMailId.SelectedValue != "0"))
        {
            UpdateMailRecords(MailId);
            BindMailingId();
            lblError.Text = "No records found to save but records Unified successfully";
        }
        trMonths.Style.Add("display", "none");
        return intResult;
    }
    public int Create_MailRecordsTemp_FundDistribution(IOrganizationService service, string SelectedFund, string SelectedFundValue)
    {
        clsDB = new DB();
        object FundId = "'" + SelectedFundValue + "'"; //lstFund.SelectedValue == "" || lstFund.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFund) + "'";
        object LegalEntityId = lstLegalEntity.SelectedValue == "" || lstLegalEntity.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstLegalEntity) + "'";
        object AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
        DataSet loDataset = new DataSet();
        if (chkEmailRecipients.Checked == true)
        {
            loDataset = clsDB.getDataSet("SP_S_FUND_DISTRIBUTION @FundIdNmbList=" + FundId + ",@AsOfDate=" + AsOfDate + ",@LegalEntityIdNmbList=" + LegalEntityId + ",@IncEmailRecipients=1");
        }
        else
        {
            loDataset = clsDB.getDataSet("SP_S_FUND_DISTRIBUTION @FundIdNmbList=" + FundId + ",@AsOfDate=" + AsOfDate + ",@LegalEntityIdNmbList=" + LegalEntityId + ",@IncEmailRecipients=0");
        }

        for (int i = 0; i < loDataset.Tables[0].Rows.Count; i++)
        {

            //Response.Write(loDataset.Tables[0].Rows.Count.ToString() + "<br/><br/><br/>");

            //objMailRecordsTemp = new ssi_mailrecordstemp();
            Entity objMailRecordsTemp = new Entity("ssi_mailrecordstemp");
            //Mail Type
            //objMailRecordsTemp.ssi_mailtypeid = new Lookup();
            //objMailRecordsTemp.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();//3FB190D9-B2CD-E011-A19B-0019B9E7EE05
            //objMailRecordsTemp.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue); //new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_mailid"]));
            objMailRecordsTemp["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));

            //Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]) != "")
            {
                //objMailRecordsTemp.ssi_name = Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);
                objMailRecordsTemp["ssi_name"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);
            }

            //[Spouse Name]
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]) != "")
            {
                //objMailRecordsTemp.ssi_spousepart_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]);
                objMailRecordsTemp["ssi_spousepart_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]);
            }

            //Ssi_AnzianoID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]) != "")
            {
                //objMailRecordsTemp.ssi_anzianoid = new CrmNumber();
                //objMailRecordsTemp.ssi_anzianoid.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]);
                objMailRecordsTemp["ssi_anzianoid"] = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]);

            }

            //Total Commitment
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Total Commitment"]) != "")
            {
                //objMailRecordsTemp.ssi_totalcommitment_db = new CrmMoney();
                //objMailRecordsTemp.ssi_totalcommitment_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Total Commitment"]));
                objMailRecordsTemp["ssi_totalcommitment_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Total Commitment"]));

            }

            //Capital Distribution
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Capital Distribution"]) != "")
            {
                //objMailRecordsTemp.ssi_capitaldistribution_db = new CrmMoney();
                //objMailRecordsTemp.ssi_capitaldistribution_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Capital Distribution"]));
                objMailRecordsTemp["ssi_capitaldistribution_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Capital Distribution"]));
            }

            //Current Distribution - Percent
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Current Distribution - Percent"]) != "")
            {
                //objMailRecordsTemp.ssi_curdistp_db = new CrmDecimal();
                //objMailRecordsTemp.ssi_curdistp_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Current Distribution - Percent"]));
                objMailRecordsTemp["ssi_curdistp_db"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Current Distribution - Percent"]);

            }

            //Prior Distributions
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Prior Distributions"]) != "")
            {
                //objMailRecordsTemp.ssi_priordistributions_db = new CrmMoney();
                //objMailRecordsTemp.ssi_priordistributions_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Prior Distributions"]));
                objMailRecordsTemp["ssi_priordistributions_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Prior Distributions"]));
            }


            //Prior Distributions - Percent
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Prior Distributions - Percent"]) != "")
            {
                //objMailRecordsTemp.ssi_priordistp_db = new CrmDecimal();
                //objMailRecordsTemp.ssi_priordistp_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Prior Distributions - Percent"]));
                objMailRecordsTemp["ssi_priordistp_db"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Prior Distributions - Percent"]);
            }

            //Distributed to date
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Distributed to date"]) != "")
            {
                //objMailRecordsTemp.ssi_disttodate_db = new CrmMoney();
                //objMailRecordsTemp.ssi_disttodate_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Distributed to date"]));
                objMailRecordsTemp["ssi_disttodate_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Distributed to date"]));

            }

            //Distributed to Date - Percent
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Distributed to Date - Percent"]) != "")
            {
                //objMailRecordsTemp.ssi_distdatep_db = new CrmDecimal();
                //objMailRecordsTemp.ssi_distdatep_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Distributed to Date - Percent"]));
                objMailRecordsTemp["ssi_distdatep_db"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Distributed to Date - Percent"]);

            }

            //Called to date
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Called to date"]) != "")
            {
                //objMailRecordsTemp.ssi_calledtodate_db = new CrmMoney();
                //objMailRecordsTemp.ssi_calledtodate_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Called to date"]));
                objMailRecordsTemp["ssi_calledtodate_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Called to date"]));

            }

            //Remaining Commitment
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Remaining Commitment"]) != "")
            {
                //objMailRecordsTemp.ssi_remainingcommitment_db = new CrmMoney();
                //objMailRecordsTemp.ssi_remainingcommitment_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Remaining Commitment"]));
                objMailRecordsTemp["ssi_remainingcommitment_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Remaining Commitment"]));
            }

            //Percent Called
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Percent Called"]) != "")
            {
                //   objMailRecordsTemp.ssi_percentcalled_db = Convert.ToString(loDataset.Tables[0].Rows[i]["Percent Called"]);
                objMailRecordsTemp["ssi_percentcalled_db"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Percent Called"]);
            }

            //Class B fee adjustment
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Class B fee adjustment"]) != "")
            {
                //objMailRecordsTemp.ssi_feeadj_db = new CrmMoney();
                //objMailRecordsTemp.ssi_feeadj_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Class B fee adjustment"]));
                objMailRecordsTemp["ssi_feeadj_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Class B fee adjustment"]));

            }


            //Actual Cash distributions
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Actual Cash distributions"]) != "")
            {
                //objMailRecordsTemp.ssi_actualcashdistributions_db = new CrmMoney();
                //objMailRecordsTemp.ssi_actualcashdistributions_db.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Actual Cash distributions"]));
                objMailRecordsTemp["ssi_actualcashdistributions_db"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Actual Cash distributions"]));
            }


            //DDA #
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["DDA"]) != "")
            {
                // objMailRecordsTemp.ssi_outdda_household = Convert.ToString(loDataset.Tables[0].Rows[i]["DDA"]);
                objMailRecordsTemp["ssi_outdda_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["DDA"]);
            }

            ////Legal Entity Name
            //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Name"]) != "")
            //{
            //    objMailRecords.ssi_legalentity = Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Name"]);
            //}

            //TNR ID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]) != "")
            {
                //objMailRecordsTemp.ssi_tnrid_nv = Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]);
                objMailRecordsTemp["ssi_tnrid_nv"] = Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]);
            }

            //Mailing Contact Type
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]) != "")
            {
                // objMailRecordsTemp.ssi_mailingcontacttype = Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]);
                objMailRecordsTemp["ssi_mailingcontacttype"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]);
            }

            //Custodian Account Number
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account Number"]) != "")
            {
                // objMailRecordsTemp.ssi_accountnumber = Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account Number"]);
                objMailRecordsTemp["ssi_accountnumber"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account Number"]);
            }

            //Account Name1
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Account Name1"]) != "")
            {
                //objMailRecordsTemp.ssi_accountname1 = Convert.ToString(loDataset.Tables[0].Rows[i]["Account Name1"]);
                objMailRecordsTemp["ssi_accountname1"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Account Name1"]);
            }

            //Legal Entity Bank
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Bank"]) != "")
            {
                //objMailRecordsTemp.ssi_legalentitybank = Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Bank"]);
                objMailRecordsTemp["ssi_legalentitybank"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Bank"]);
            }

            //Basic Wire Info /*Distribution tab*/
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]) != "")
            {
                //objMailRecordsTemp.ssi_ssi_basicwireinfo_household = Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]);
                objMailRecordsTemp["ssi_ssi_basicwireinfo_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]);
            }

            //Basic Wire Info /*Capital Call Letter tab*/
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]) != "")
            {
                // objMailRecordsTemp.ssi_ssi_basicwireinfo_household1 = Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]);
                objMailRecordsTemp["ssi_ssi_basicwireinfo_household1"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]);
            }

            //ABA Routing # /*Distribution tab*/
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]) != "")
            {
                //objMailRecordsTemp.ssi_outabarouting_household = Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]);
                objMailRecordsTemp["ssi_outabarouting_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]);
            }

            //ABA Routing # /*Capital Call Letter tab*/
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]) != "")
            {
                // objMailRecordsTemp.ssi_abarouting_household = Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]);
                objMailRecordsTemp["ssi_abarouting_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]);
            }

            //For Further Credit (FFC) Acct #
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Acct #"]) != "")
            {
                //  objMailRecordsTemp.ssi_ffcacct_household = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Acct #"]);
                objMailRecordsTemp["ssi_ffcacct_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Acct #"]);
            }

            //For Further Credit (FFC) Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Name"]) != "")
            {
                //objMailRecordsTemp.ssi_ffcname_household = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Name"]);
                objMailRecordsTemp["ssi_ffcname_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Name"]);
            }

            //Other Wire Instructions
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Other Wire Instructions"]) != "")
            {
                // objMailRecordsTemp.ssi_otherwireinstr_household = Convert.ToString(loDataset.Tables[0].Rows[i]["Other Wire Instructions"]);
                objMailRecordsTemp["ssi_otherwireinstr_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Other Wire Instructions"]);
            }


            //Fund Name
            if (SelectedFund != "")
            {
                //objMailRecordsTemp.ssi_fundname = lstFund.SelectedItem.Text; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                objMailRecordsTemp["ssi_fundname"] = SelectedFund; //Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
            }

            //Fund Name (to show on wire instrux)
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name (to show on wire instrux)"]) != "")
            {
                // objMailRecordsTemp.ssi_fundname_fund = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name (to show on wire instrux)"]);
                objMailRecordsTemp["ssi_fundname_fund"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name (to show on wire instrux)"]);
            }

            //Fund Bank
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Bank"]) != "")
            {
                //objMailRecordsTemp.ssi_fundbank = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Bank"]);
                objMailRecordsTemp["ssi_fundbank"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Bank"]);
            }

            //Fund Account Number
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Account Number"]) != "")
            {
                //objMailRecordsTemp.ssi_fundaccountnumber = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Account Number"]);
                objMailRecordsTemp["ssi_fundaccountnumber"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Account Number"]);
            }


            //Distribution Payment Method
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Payment Method"]) != "")
            {
                // objMailRecordsTemp.ssi_distributionpaymentmethod = Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Payment Method"]);
                objMailRecordsTemp["ssi_distributionpaymentmethod"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Payment Method"]);
            }


            //Distribution Method Note
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Method Note"]) != "")
            {
                //objMailRecordsTemp.ssi_distributionmethodnote = Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Method Note"]);
                objMailRecordsTemp["ssi_distributionmethodnote"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Distribution Method Note"]);
            }

            //SLOAFlg
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["SLOAFlg"]) != "")
            {
                //objMailRecordsTemp.ssi_sloaflg = new CrmBoolean();
                //objMailRecordsTemp.ssi_sloaflg.Value = Convert.ToBoolean(loDataset.Tables[0].Rows[i]["SLOAFlg"]);
                objMailRecordsTemp["ssi_sloaflg"] = Convert.ToBoolean(Convert.ToString(loDataset.Tables[0].Rows[i]["SLOAFlg"]).ToLower());

            }



            //AsofDate
            if (txtAsofdate.Text != "")
            {
                //objMailRecordsTemp.ssi_asofdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_asofdate.Value = txtAsofdate.Text;
                objMailRecordsTemp["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);

            }

            //MailingID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["MailingID"]) != "")
            {
                //objMailRecordsTemp.ssi_mailingid = new CrmNumber();
                //objMailRecordsTemp.ssi_mailingid.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["MailingID"]);
                objMailRecordsTemp["ssi_mailingid"] = Convert.ToInt32(loDataset.Tables[0].Rows[i]["MailingID"]);

            }

            //Mail
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]) != "")
            {
                // objMailRecordsTemp.ssi_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]);
                objMailRecordsTemp["ssi_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]);
            }

            //Owner First Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]) != "")
            {
                //objMailRecordsTemp.ssi_ownerfirstname_hh_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]);
                objMailRecordsTemp["ssi_ownerfirstname_hh_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]);
            }

            //Owner Last Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]) != "")
            {
                // objMailRecordsTemp.ssi_ownerlname_hh_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]);
                objMailRecordsTemp["ssi_ownerlname_hh_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]);
            }

            //Contact Household
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]) != "")
            {
                //objMailRecordsTemp.ssi_hholdinst_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]);
                objMailRecordsTemp["ssi_hholdinst_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]);
            }

            //Contact Full Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]) != "")
            {
                // objMailRecordsTemp.ssi_fullname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]);
                objMailRecordsTemp["ssi_fullname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]);
            }


            //Secondary Owner First Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
            {
                // objMailRecordsTemp.ssi_secownerfname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]);
                objMailRecordsTemp["ssi_secownerfname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]);
            }

            //Secondary Owner Last Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
            {
                //objMailRecordsTemp.ssi_secownerlname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                objMailRecordsTemp["ssi_secownerlname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]);
            }


            //Address Line 1
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]) != "")
            {
                //  objMailRecordsTemp.ssi_addressline1_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]);
                objMailRecordsTemp["ssi_addressline1_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]);
            }

            //Address Line 2
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]) != "")
            {
                //objMailRecordsTemp.ssi_addressline2_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]);
                objMailRecordsTemp["ssi_addressline2_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]);
            }

            //Address Line 3
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]) != "")
            {
                // objMailRecordsTemp.ssi_addressline3_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]);
                objMailRecordsTemp["ssi_addressline3_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]);
            }

            //City
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["City"]) != "")
            {
                // objMailRecordsTemp.ssi_city_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["City"]);
                objMailRecordsTemp["ssi_city_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["City"]);
            }

            //State or Province
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]) != "")
            {
                //objMailRecordsTemp.ssi_stateprovince_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]);
                objMailRecordsTemp["ssi_stateprovince_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]);
            }

            //ZIP Code
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]) != "")
            {
                //objMailRecordsTemp.ssi_zipcode_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]);
                objMailRecordsTemp["ssi_zipcode_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]);
            }

            //Country or Region
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]) != "")
            {
                // objMailRecordsTemp.ssi_countryregion_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]);
                objMailRecordsTemp["ssi_countryregion_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]);
            }

            //Dear
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]) != "")
            {
                //  objMailRecordsTemp.ssi_dear_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]);
                objMailRecordsTemp["ssi_dear_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]);
            }

            //Salutation
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]) != "")
            {
                //objMailRecordsTemp.ssi_salutation_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]);
                objMailRecordsTemp["ssi_salutation_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]);
            }

            //Mail Preference 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]) != "")
            {
                //objMailRecordsTemp.ssi_mailpreference_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]);
                objMailRecordsTemp["ssi_mailpreference_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]);
            }

            //ssi_mailStatus 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]) != "")
            {
                //objMailRecordsTemp.ssi_mailstatus = new Picklist();
                //objMailRecordsTemp.ssi_mailstatus.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]);
                objMailRecordsTemp["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]));

            }


            //HouseHold lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"]) != "")
            {
                //objMailRecordsTemp.ssi_accountid = new Lookup();
                //objMailRecordsTemp.ssi_accountid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"]));
                objMailRecordsTemp["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"])));
            }

            //Contact lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"]) != "")
            {
                //objMailRecordsTemp.ssi_contactfullnameid = new Lookup();
                //objMailRecordsTemp.ssi_contactfullnameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"]));
                objMailRecordsTemp["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"])));
            }

            //ssi_LegalEntityId lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
            {
                //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                //objMailRecordsTemp.ssi_legalentitynameid = new Lookup();
                //objMailRecordsTemp.ssi_legalentitynameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                objMailRecordsTemp["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"])));
            }


            //Advisor Approval Required
            //objMailRecordsTemp.ssi_advisorapprovalreqd = new CrmBoolean();
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "0" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "False".ToUpper())
            {
                // objMailRecordsTemp.ssi_advisorapprovalreqd.Value = false;
                objMailRecordsTemp["ssi_advisorapprovalreqd"] = false;
            }
            else if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "1" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "True".ToUpper())
            {
                // objMailRecordsTemp.ssi_advisorapprovalreqd.Value = true;
                objMailRecordsTemp["ssi_advisorapprovalreqd"] = true;
            }

            //File Name Added on 9 oct 2014
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]) != "")
            {
                //objMailRecordsTemp.ssi_filename = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]);
                objMailRecordsTemp["ssi_filename"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]);
            }
            //added by sasmit 5_3_2017
            //clientportalname 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]) != "")
            {
                //objMailRecordsTemp.ssi_clientportalname = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]);
                objMailRecordsTemp["ssi_clientportalname"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]);
            }
            //added by sasmit 5_3_2017
            //clientreportfolder
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]) != "")
            {
                // objMailRecordsTemp.ssi_clientreportfolder = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]);
                objMailRecordsTemp["ssi_clientreportfolder"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]);
            }
            // CreatedByCustomid Field 
            //Rohit Pawar
            string Userid = GetcurrentUser();

            if (Userid != "")
            {
                //objMailRecordsTemp.ssi_createdbycustomid = new Lookup();
                //objMailRecordsTemp.ssi_createdbycustomid.Value = new Guid(Userid);

                objMailRecordsTemp["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
            }

            #region Logic for Capital Call Approval Process

            // Logic for Capital Call Approval Process

            //if (chkUnify.Checked == true)
            //{
            //    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
            //    //objMailRecordsTemp.ssi_unifiedflg.Value = true;
            //    objMailRecordsTemp["ssi_unifiedflg"] = true;
            //}
            //else if (chkUnify.Checked == false)
            //{
            //    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
            //    //objMailRecordsTemp.ssi_unifiedflg.Value = false;
            //    objMailRecordsTemp["ssi_unifiedflg"] = false;
            //}

            if (MailId != "" && MailId != "0")
            {
                //objMailRecordsTemp.ssi_mailidtemp = new CrmNumber();
                //objMailRecordsTemp.ssi_mailidtemp.Value = Convert.ToInt32(MailId);
                objMailRecordsTemp["ssi_mailidtemp"] = Convert.ToInt32(MailId);

            }


            if (ddlTemplates.SelectedValue != "" && ddlTemplates.SelectedValue != "0")
            {
                //objMailRecordsTemp.ssi_templateid = new Lookup();
                //objMailRecordsTemp.ssi_templateid.Value = new Guid(ddlTemplates.SelectedValue);
                objMailRecordsTemp["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(ddlTemplates.SelectedValue));
            }

            //Wire AsofDate
            if (txtWireAsofDate.Text != "")
            {
                //objMailRecordsTemp.ssi_wireasofdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_wireasofdate.Value = txtWireAsofDate.Text;
                objMailRecordsTemp["ssi_wireasofdate"] = Convert.ToDateTime(txtWireAsofDate.Text);

            }

            //Letter AsofDate
            if (txtLetterDate.Text != "")
            {
                //objMailRecordsTemp.ssi_letterdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_letterdate.Value = txtLetterDate.Text;
                objMailRecordsTemp["ssi_letterdate"] = Convert.ToDateTime(txtLetterDate.Text);
            }


            #endregion


            if (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0" && chkUnify.Checked == false)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
            else if (ddlMailId.SelectedValue != "" && chkUnify.Checked == true)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
            else if (ddlMailId.SelectedValue != "" && chkUnify.Checked == false)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
        }
        if (intResult > 0)
        {
            Success++;
            // lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully";
        }

        if (loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == false && (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0"))
            //  lblError.Text = "No Records Found";// commented/changed 27_08_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == false && (ddlMailId.SelectedValue != "" || ddlMailId.SelectedValue != "0"))
            // lblError.Text = "No Records Found";// commented/changed 27_08_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == true && (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0"))
            // lblError.Text = "No Records Found";// commented/changed 27_08_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == true && (ddlMailId.SelectedValue != "" || ddlMailId.SelectedValue != "0"))
        {
            UpdateMailRecords(MailId);
            BindMailingId();
            lblError.Text = "No records found to save but records Unified successfully";
        }

        trMonths.Style.Add("display", "none");
        return intResult;
    }
    public int Create_MailRecordsTEmp_CapCallWire(IOrganizationService service, string SelectedFund, string SelectedFundValue)
    {
        clsDB = new DB();
        object FundId = "'" + SelectedFundValue + "'";// lstFund.SelectedValue == "" || lstFund.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFund) + "'";
        object LegalEntityId = lstLegalEntity.SelectedValue == "" || lstLegalEntity.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstLegalEntity) + "'";
        object AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
        DataSet loDataset = new DataSet();
        if (chkEmailRecipients.Checked == true)
        {
            loDataset = clsDB.getDataSet("SP_S_CAPITAL_CALL_WIRE @FundIdNmbList=" + FundId + ",@AsOfDate=" + AsOfDate + ",@LegalEntityIdNmbList=" + LegalEntityId + ",@IncEmailRecipients=1");
        }
        else
        {
            loDataset = clsDB.getDataSet("SP_S_CAPITAL_CALL_WIRE @FundIdNmbList=" + FundId + ",@AsOfDate=" + AsOfDate + ",@LegalEntityIdNmbList=" + LegalEntityId + ",@IncEmailRecipients=0");
        }

        for (int i = 0; i < loDataset.Tables[0].Rows.Count; i++)
        {

            //Response.Write(loDataset.Tables[0].Rows.Count.ToString() + "<br/><br/><br/>");

            // objMailRecordsTemp = new ssi_mailrecordstemp();
            Entity objMailRecordsTemp = new Entity("ssi_mailrecordstemp");

            //Mail Type
            //objMailRecordsTemp.ssi_mailtypeid = new Lookup();
            //objMailRecordsTemp.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();//3FB190D9-B2CD-E011-A19B-0019B9E7EE05
            //objMailRecordsTemp.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue); //new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_mailid"]));
            objMailRecordsTemp["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));

            //Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]) != "")
            {
                //objMailRecordsTemp.ssi_name = Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);
                objMailRecordsTemp["ssi_name"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);
            }

            //[Spouse Name]
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]) != "")
            {
                //objMailRecordsTemp.ssi_spousepart_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]);
                objMailRecordsTemp["ssi_spousepart_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]);
            }

            //Ssi_AnzianoID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]) != "")
            {
                //objMailRecordsTemp.ssi_anzianoid = new CrmNumber();
                //objMailRecordsTemp.ssi_anzianoid.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]);
                objMailRecordsTemp["ssi_anzianoid"] = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]);
            }

            //Ssi_Ttladjcommit_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Ttladjcommit_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_ttladjcommit_ccsf = new CrmMoney();
                //objMailRecordsTemp.ssi_ttladjcommit_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Ttladjcommit_ccsf"]));
                objMailRecordsTemp["ssi_ttladjcommit_ccsf"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_Ttladjcommit_ccsf"]));

            }

            //Ssi_PercentCalled_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PercentCalled_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_percentcalled_ccsf = new CrmDecimal();
                //objMailRecordsTemp.ssi_percentcalled_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PercentCalled_ccsf"]));
                objMailRecordsTemp["ssi_percentcalled_ccsf"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_PercentCalled_ccsf"]);

            }

            //ssi_currentcall_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_CurrentCall_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_currentcall_ccsf = new CrmMoney();
                //objMailRecordsTemp.ssi_currentcall_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_CurrentCall_ccsf"]));
                objMailRecordsTemp["ssi_currentcall_ccsf"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_CurrentCall_ccsf"]));

            }

            //Ssi_PriorCalls_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PriorCalls_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_priorcalls_ccsf = new CrmMoney();
                //objMailRecordsTemp.ssi_priorcalls_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PriorCalls_ccsf"]));
                objMailRecordsTemp["ssi_priorcalls_ccsf"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_PriorCalls_ccsf"]));
            }

            //Ssi_PercentPriorCalls_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PercentPriorCalls_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_percentpriorcalls_ccsf = new CrmDecimal();
                //objMailRecordsTemp.ssi_percentpriorcalls_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PercentPriorCalls_ccsf"]));
                objMailRecordsTemp["ssi_percentpriorcalls_ccsf"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_PercentPriorCalls_ccsf"]);

            }

            //Ssi_Calledtodate_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Calledtodate_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_calledtodate_ccsf = new CrmMoney();
                //objMailRecordsTemp.ssi_calledtodate_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Calledtodate_ccsf"]));
                objMailRecordsTemp["ssi_calledtodate_ccsf"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_Calledtodate_ccsf"]));
            }

            //Ssi_CalledtoDateP_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_CalledtoDateP_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_calledtodatep_ccsf = new CrmDecimal();
                //objMailRecordsTemp.ssi_calledtodatep_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_CalledtoDateP_ccsf"]));
                objMailRecordsTemp["ssi_calledtodatep_ccsf"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_CalledtoDateP_ccsf"]);
            }

            //Ssi_RemainingCommitment_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitment_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_remainingcommitment_ccsf = new CrmMoney();
                //objMailRecordsTemp.ssi_remainingcommitment_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitment_ccsf"]));
                objMailRecordsTemp["ssi_remainingcommitment_ccsf"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitment_ccsf"]));

            }

            //Ssi_RemainingCommitmentP_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitmentP_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_remainingcommitmentp_ccsf = new CrmDecimal();
                //objMailRecordsTemp.ssi_remainingcommitmentp_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitmentP_ccsf"]));
                objMailRecordsTemp["ssi_remainingcommitmentp_ccsf"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitmentP_ccsf"]);
            }

            ////Legal Entity Name
            //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Name"]) != "")
            //{
            //    objMailRecords.ssi_legalentity = Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Name"]);
            //}

            //TNR ID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]) != "")
            {
                //objMailRecordsTemp.ssi_tnrid_nv = Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]);
                objMailRecordsTemp["ssi_tnrid_nv"] = Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]);
            }

            //Mailing Contact Type
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]) != "")
            {
                // objMailRecordsTemp.ssi_mailingcontacttype = Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]);
                objMailRecordsTemp["ssi_mailingcontacttype"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]);
            }

            //Custodian Account Number
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account Number"]) != "")
            {
                // objMailRecordsTemp.ssi_accountnumber = Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account Number"]); objMailRecordsTemp.ssi_accountnumber = Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account Number"]);
                objMailRecordsTemp["ssi_accountnumber"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Custodian Account Number"]);
            }

            //Account Name1
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Account Name1"]) != "")
            {
                // objMailRecordsTemp.ssi_accountname1 = Convert.ToString(loDataset.Tables[0].Rows[i]["Account Name1"]);
                objMailRecordsTemp["ssi_accountname1"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Account Name1"]);
            }

            //Legal Entity Bank
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Bank"]) != "")
            {
                //objMailRecordsTemp.ssi_legalentitybank = Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Bank"]);
                objMailRecordsTemp["ssi_legalentitybank"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Bank"]);
            }

            //Basic Wire Info
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]) != "")
            {
                // objMailRecordsTemp.ssi_ssi_basicwireinfo_household1 = Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]);
                objMailRecordsTemp["ssi_ssi_basicwireinfo_household1"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Basic Wire Info"]);
            }

            //AsofDate
            if (txtAsofdate.Text != "")
            {
                //objMailRecordsTemp.ssi_asofdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_asofdate.Value = txtAsofdate.Text;
                objMailRecordsTemp["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);

            }

            //ABA Routing #
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]) != "")
            {
                // objMailRecordsTemp.ssi_abarouting_household = Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]);
                objMailRecordsTemp["ssi_abarouting_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ABA Routing #"]);
            }

            //For Further Credit (FFC) Acct #
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Acct #"]) != "")
            {
                //objMailRecordsTemp.ssi_ffcacct_household = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Acct #"]);
                objMailRecordsTemp["ssi_ffcacct_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Acct #"]);
            }

            //For Further Credit (FFC) Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Name"]) != "")
            {
                // objMailRecordsTemp.ssi_ffcname_household = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Name"]);
                objMailRecordsTemp["ssi_ffcname_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["For Further Credit (FFC) Name"]);
            }

            //Other Wire Instructions
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Other Wire Instructions"]) != "")
            {
                //objMailRecordsTemp.ssi_otherwireinstr_household = Convert.ToString(loDataset.Tables[0].Rows[i]["Other Wire Instructions"]);
                objMailRecordsTemp["ssi_otherwireinstr_household"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Other Wire Instructions"]);
            }

            //Fund Name
            if (SelectedFund != "")
            {
                // objMailRecordsTemp.ssi_fundname = lstFund.SelectedItem.Text;// Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                objMailRecordsTemp["ssi_fundname"] = SelectedFund;// lstFund.SelectedItem.Text;// Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
            }

            //Fund Name (to show on wire instrux
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name (to show on wire instrux)"]) != "")
            {
                // objMailRecordsTemp.ssi_fundname_fund = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name (to show on wire instrux)"]);
                objMailRecordsTemp["ssi_fundname_fund"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name (to show on wire instrux)"]);
            }

            //Fund Bank
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Bank"]) != "")
            {
                // objMailRecordsTemp.ssi_fundbank = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Bank"]);
                objMailRecordsTemp["ssi_fundbank"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Bank"]);
            }

            //Fund Account Number
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Account Number"]) != "")
            {
                // objMailRecordsTemp.ssi_fundaccountnumber = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Account Number"]);
                objMailRecordsTemp["ssi_fundaccountnumber"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Account Number"]);
            }

            //Capital Call Payment Method
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Capital Call Payment Method"]) != "")
            {
                //  objMailRecordsTemp.ssi_capitalcallpaymentmethod = Convert.ToString(loDataset.Tables[0].Rows[i]["Capital Call Payment Method"]);
                objMailRecordsTemp["ssi_capitalcallpaymentmethod"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Capital Call Payment Method"]);
            }

            //Payment Method Note
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Payment Method Note"]) != "")
            {
                // objMailRecordsTemp.ssi_paymentmethodnote = Convert.ToString(loDataset.Tables[0].Rows[i]["Payment Method Note"]);
                objMailRecordsTemp["ssi_paymentmethodnote"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Payment Method Note"]);
            }

            //SLOAFlg
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["SLOAFlg"]) != "")
            {
                //objMailRecordsTemp.ssi_sloaflg = new CrmBoolean();
                //objMailRecordsTemp.ssi_sloaflg.Value = Convert.ToBoolean(loDataset.Tables[0].Rows[i]["loDataset"]);
                objMailRecordsTemp["ssi_sloaflg"] = Convert.ToBoolean(Convert.ToString(loDataset.Tables[0].Rows[i]["SLOAFlg"]).ToLower());

            }

            //Signer 1
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 1"]) != "")
            {
                // objMailRecordsTemp.ssi_signer1_clientaccount = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 1"]);
                objMailRecordsTemp["ssi_signer1_clientaccount"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 1"]);
            }

            //Signer 2
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 2"]) != "")
            {
                //objMailRecordsTemp.ssi_signer2_clientaccount = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 2"]);
                objMailRecordsTemp["ssi_signer2_clientaccount"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 2"]);
            }

            //Signer 3
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 3"]) != "")
            {
                //objMailRecordsTemp.ssi_signer3_clientaccount = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 3"]);
                objMailRecordsTemp["ssi_signer3_clientaccount"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 3"]);
            }

            //Signer 1 Title
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 1 Title"]) != "")
            {
                // objMailRecordsTemp.ssi_signer1title_clientaccount = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 1 Title"]);
                objMailRecordsTemp["ssi_signer1title_clientaccount"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 1 Title"]);
            }

            //Signer 2 Title
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 2 Title"]) != "")
            {
                // objMailRecordsTemp.ssi_signer2title_clientaccount = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 2 Title"]);
                objMailRecordsTemp["ssi_signer2title_clientaccount"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 2 Title"]);
            }

            //Signer 3 Title
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 3 Title"]) != "")
            {
                //objMailRecordsTemp.ssi_signer3title_clientaccount = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 3 Title"]);
                objMailRecordsTemp["ssi_signer3title_clientaccount"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Signer 3 Title"]);
            }


            //MailingID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["MailingID"]) != "")
            {
                //objMailRecordsTemp.ssi_mailingid = new CrmNumber();
                //objMailRecordsTemp.ssi_mailingid.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["MailingID"]);
                objMailRecordsTemp["ssi_mailingid"] = Convert.ToInt32(loDataset.Tables[0].Rows[i]["MailingID"]);

            }

            //Mail
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]) != "")
            {
                //objMailRecordsTemp.ssi_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]);
                objMailRecordsTemp["ssi_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]);
            }

            //Owner First Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]) != "")
            {
                // objMailRecordsTemp.ssi_ownerfirstname_hh_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]);
                objMailRecordsTemp["ssi_ownerfirstname_hh_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]);
            }

            //Owner Last Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]) != "")
            {
                //objMailRecordsTemp.ssi_ownerlname_hh_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]);
                objMailRecordsTemp["ssi_ownerlname_hh_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]);
            }

            //Contact Household
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]) != "")
            {
                //objMailRecordsTemp.ssi_hholdinst_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]);
                objMailRecordsTemp["ssi_hholdinst_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]);
            }

            //Contact Full Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]) != "")
            {
                //objMailRecordsTemp.ssi_fullname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]);
                objMailRecordsTemp["ssi_fullname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]);
            }

            //Secondary Owner First Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
            {
                //objMailRecordsTemp.ssi_secownerfname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]);
                objMailRecordsTemp["ssi_secownerfname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]);
            }

            //Secondary Owner Last Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
            {
                // objMailRecordsTemp.ssi_secownerlname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                objMailRecordsTemp["ssi_secownerlname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]);
            }


            //Country or Region
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]) != "")
            {
                // objMailRecordsTemp.ssi_countryregion_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]);
                objMailRecordsTemp["ssi_countryregion_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]);
            }

            //Address Line 1
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]) != "")
            {
                //objMailRecordsTemp.ssi_addressline1_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]);
                objMailRecordsTemp["ssi_addressline1_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]);
            }

            //Address Line 2
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]) != "")
            {
                //objMailRecordsTemp.ssi_addressline2_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]);
                objMailRecordsTemp["ssi_addressline2_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]);
            }

            //Address Line 3
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]) != "")
            {
                // objMailRecordsTemp.ssi_addressline3_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]);
                objMailRecordsTemp["ssi_addressline3_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]);
            }

            //City
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["City"]) != "")
            {
                //objMailRecordsTemp.ssi_city_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["City"]);
                objMailRecordsTemp["ssi_city_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["City"]);
            }

            //State or Province
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]) != "")
            {
                // objMailRecordsTemp.ssi_stateprovince_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]);
                objMailRecordsTemp["ssi_stateprovince_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]);
            }

            //ZIP Code
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]) != "")
            {
                // objMailRecordsTemp.ssi_zipcode_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]);
                objMailRecordsTemp["ssi_zipcode_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]);
            }


            //Dear
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]) != "")
            {
                // objMailRecordsTemp.ssi_dear_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]);
                objMailRecordsTemp["ssi_dear_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]);
            }

            //Salutation
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]) != "")
            {
                //objMailRecordsTemp.ssi_salutation_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]);
                objMailRecordsTemp["ssi_salutation_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]);
            }

            //Mail Preference 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]) != "")
            {
                // objMailRecordsTemp.ssi_mailpreference_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]);
                objMailRecordsTemp["ssi_mailpreference_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]);
            }

            //ssi_mailStatus 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]) != "")
            {
                //objMailRecordsTemp.ssi_mailstatus = new Picklist();
                //objMailRecordsTemp.ssi_mailstatus.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]);
                objMailRecordsTemp["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]));

            }

            //HouseHold lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"]) != "")
            {
                //objMailRecordsTemp.ssi_accountid = new Lookup();
                //objMailRecordsTemp.ssi_accountid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"]));
                objMailRecordsTemp["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"])));
            }

            //Contact lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"]) != "")
            {
                //objMailRecordsTemp.ssi_contactfullnameid = new Lookup();
                //objMailRecordsTemp.ssi_contactfullnameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"]));
                objMailRecordsTemp["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"])));
            }

            //ssi_LegalEntityId lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
            {
                //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                //objMailRecordsTemp.ssi_legalentitynameid = new Lookup();
                //objMailRecordsTemp.ssi_legalentitynameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                objMailRecordsTemp["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"])));
            }


            //Advisor Approval Required
            //objMailRecordsTemp.ssi_advisorapprovalreqd = new CrmBoolean();
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "0" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "False".ToUpper())
            {
                // objMailRecordsTemp.ssi_advisorapprovalreqd.Value = false;
                objMailRecordsTemp["ssi_advisorapprovalreqd"] = false;
            }
            else if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "1" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "True".ToUpper())
            {
                //objMailRecordsTemp.ssi_advisorapprovalreqd.Value = true;
                objMailRecordsTemp["ssi_advisorapprovalreqd"] = true;
            }

            //File Name Added on 9 oct 2014
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]) != "")
            {
                // objMailRecordsTemp.ssi_filename = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]);
                objMailRecordsTemp["ssi_filename"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]);
            }
            //added by sasmit 5_3_2017
            //clientportalname 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]) != "")
            {
                // objMailRecordsTemp.ssi_clientportalname = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]);
                objMailRecordsTemp["ssi_clientportalname"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]);
            }
            //added by sasmit 5_3_2017
            //clientreportfolder
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]) != "")
            {
                // objMailRecordsTemp.ssi_clientreportfolder = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]);
                objMailRecordsTemp["ssi_clientreportfolder"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]);
            }
            // CreatedByCustomid Field 
            //Rohit Pawar
            string Userid = GetcurrentUser();

            if (Userid != "")
            {
                //objMailRecordsTemp.ssi_createdbycustomid = new Lookup();
                //objMailRecordsTemp.ssi_createdbycustomid.Value = new Guid(Userid);
                objMailRecordsTemp["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
            }

            #region Logic for Capital Call Approval Process

            // Logic for Capital Call Approval Process

            //if (chkUnify.Checked == true)
            //{
            //    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
            //    //objMailRecordsTemp.ssi_unifiedflg.Value = true;
            //    objMailRecordsTemp["ssi_unifiedflg"] = true;

            //}
            //else if (chkUnify.Checked == false)
            //{
            //    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
            //    //objMailRecordsTemp.ssi_unifiedflg.Value = false;
            //    objMailRecordsTemp["ssi_unifiedflg"] = false;
            //}

            if (MailId != "" && MailId != "0")
            {
                //objMailRecordsTemp.ssi_mailidtemp = new CrmNumber();
                //objMailRecordsTemp.ssi_mailidtemp.Value = Convert.ToInt32(MailId);
                objMailRecordsTemp["ssi_mailidtemp"] = Convert.ToInt32(MailId);

            }


            if (ddlTemplates.SelectedValue != "" && ddlTemplates.SelectedValue != "0")
            {
                //objMailRecordsTemp.ssi_templateid = new Lookup();
                //objMailRecordsTemp.ssi_templateid.Value = new Guid(ddlTemplates.SelectedValue);
                objMailRecordsTemp["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(Convert.ToString(ddlTemplates.SelectedValue)));

            }

            //Wire AsofDate
            if (txtWireAsofDate.Text != "")
            {
                //objMailRecordsTemp.ssi_wireasofdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_wireasofdate.Value = txtWireAsofDate.Text;
                objMailRecordsTemp["ssi_wireasofdate"] = Convert.ToDateTime(txtWireAsofDate.Text);

            }

            //Letter AsofDate
            if (txtLetterDate.Text != "")
            {
                //objMailRecordsTemp.ssi_letterdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_letterdate.Value = txtLetterDate.Text;
                objMailRecordsTemp["ssi_letterdate"] = Convert.ToDateTime(txtLetterDate.Text);
            }


            #endregion


            if (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0" && chkUnify.Checked == false)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
            else if (ddlMailId.SelectedValue != "" && chkUnify.Checked == true)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
            else if (ddlMailId.SelectedValue != "" && chkUnify.Checked == false)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
        }

        if (intResult > 0)
        {
            Success++;
            //lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully";
        }

        if (loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == false && (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0"))
            //lblError.Text = "No Records Found"; // commented/changed 27_08_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == false && (ddlMailId.SelectedValue != "" || ddlMailId.SelectedValue != "0"))
            //lblError.Text = "No Records Found"; // commented/changed 27_08_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == true && (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0"))
            //  lblError.Text = "No Records Found"; // commented/changed 27_08_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == true && (ddlMailId.SelectedValue != "" || ddlMailId.SelectedValue != "0"))
        {
            UpdateMailRecords(MailId);
            BindMailingId();
            lblError.Text = "No records found to save but records Unified successfully";
        }
        return intResult;
    }
    public int Create_MailRecordsTemp_CapitalCallLetter(IOrganizationService service, string SelectedFund, string SelectedFundValue)
    {
        clsDB = new DB();
        DataSet loDataset = new DataSet();
        //object FundId = lstFund.SelectedValue == "" || lstFund.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstFund) + "'";
        object FundId = "'" + SelectedFundValue + "'";// lstFund.SelectedValue == "" || lstFund.SelectedValue == "0" ? "null" : "'" + SelectedFundValue + "'";
        object LegalEntityId = lstLegalEntity.SelectedValue == "" || lstLegalEntity.SelectedValue == "0" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstLegalEntity) + "'";
        object AsOfDate = txtAsofdate.Text == "" ? "null" : "'" + txtAsofdate.Text + "'";
        if (chkEmailRecipients.Checked == true)
        {
            loDataset = clsDB.getDataSet("SP_S_CAPITAL_CALL_LETTER @FundIdNmbList=" + FundId + ",@AsOfDate=" + AsOfDate + ",@LegalEntityIdNmbList=" + LegalEntityId + ",@IncEmailRecipients=1");
        }
        else
        {
            loDataset = clsDB.getDataSet("SP_S_CAPITAL_CALL_LETTER @FundIdNmbList=" + FundId + ",@AsOfDate=" + AsOfDate + ",@LegalEntityIdNmbList=" + LegalEntityId + ",@IncEmailRecipients=0");
        }

        for (int i = 0; i < loDataset.Tables[0].Rows.Count; i++)
        {

            //Response.Write(loDataset.Tables[0].Rows.Count.ToString() + "<br/><br/><br/>");

            // objMailRecordsTemp = new ssi_mailrecordstemp();
            Entity objMailRecordsTemp = new Entity("ssi_mailrecordstemp");
            //Mail Type
            //objMailRecordsTemp.ssi_mailtypeid = new Lookup();
            //objMailRecordsTemp.ssi_mailtypeid.type = EntityName.ssi_mail.ToString();//3FB190D9-B2CD-E011-A19B-0019B9E7EE05
            //objMailRecordsTemp.ssi_mailtypeid.Value = new Guid(ddlMailType.SelectedValue); //new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_mailid"]));
            objMailRecordsTemp["ssi_mailtypeid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_mail", new Guid(Convert.ToString(ddlMailType.SelectedValue)));



            //Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]) != "")
            {
                //objMailRecordsTemp.ssi_name = Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);
                objMailRecordsTemp["ssi_name"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Name"]);
            }


            //[Spouse Name]
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]) != "")
            {
                //objMailRecordsTemp.ssi_spousepart_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]);
                objMailRecordsTemp["ssi_spousepart_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Spouse Name"]);
            }

            //Ssi_AnzianoID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]) != "")
            {
                //objMailRecordsTemp.ssi_anzianoid = new CrmNumber();
                //objMailRecordsTemp.ssi_anzianoid.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]);
                objMailRecordsTemp["ssi_anzianoid"] = Convert.ToInt32(loDataset.Tables[0].Rows[i]["Ssi_AnzianoID"]);

            }

            //Ssi_Ttladjcommit_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Ttladjcommit_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_ttladjcommit_ccsf = new CrmMoney();
                //objMailRecordsTemp.ssi_ttladjcommit_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Ttladjcommit_ccsf"]));
                objMailRecordsTemp["ssi_ttladjcommit_ccsf"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_Ttladjcommit_ccsf"]));

            }

            //Ssi_PercentCalled_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PercentCalled_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_percentcalled_ccsf = new CrmDecimal();
                //objMailRecordsTemp.ssi_percentcalled_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PercentCalled_ccsf"]));
                objMailRecordsTemp["ssi_percentcalled_ccsf"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_PercentCalled_ccsf"]);

            }

            //ssi_currentcall_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_CurrentCall_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_currentcall_ccsf = new CrmMoney();
                //objMailRecordsTemp.ssi_currentcall_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_CurrentCall_ccsf"]));
                objMailRecordsTemp["ssi_currentcall_ccsf"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_CurrentCall_ccsf"]));
            }

            //Ssi_PriorCalls_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PriorCalls_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_priorcalls_ccsf = new CrmMoney();
                //objMailRecordsTemp.ssi_priorcalls_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PriorCalls_ccsf"]));
                objMailRecordsTemp["ssi_priorcalls_ccsf"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_PriorCalls_ccsf"]));
            }

            //Ssi_PercentPriorCalls_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PercentPriorCalls_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_percentpriorcalls_ccsf = new CrmDecimal();
                //objMailRecordsTemp.ssi_percentpriorcalls_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_PercentPriorCalls_ccsf"]));
                objMailRecordsTemp["ssi_percentpriorcalls_ccsf"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_PercentPriorCalls_ccsf"]);

            }

            //Ssi_Calledtodate_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Calledtodate_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_calledtodate_ccsf = new CrmMoney();
                //objMailRecordsTemp.ssi_calledtodate_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_Calledtodate_ccsf"]));
                objMailRecordsTemp["ssi_calledtodate_ccsf"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_Calledtodate_ccsf"]));
            }

            //Ssi_CalledtoDateP_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_CalledtoDateP_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_calledtodatep_ccsf = new CrmDecimal();
                //objMailRecordsTemp.ssi_calledtodatep_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_CalledtoDateP_ccsf"]));
                objMailRecordsTemp["ssi_calledtodatep_ccsf"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_CalledtoDateP_ccsf"]);
            }

            //Ssi_RemainingCommitment_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitment_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_remainingcommitment_ccsf = new CrmMoney();
                //objMailRecordsTemp.ssi_remainingcommitment_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitment_ccsf"]));
                objMailRecordsTemp["ssi_remainingcommitment_ccsf"] = new Microsoft.Xrm.Sdk.Money(Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitment_ccsf"]));

            }

            //Ssi_RemainingCommitmentP_ccsf
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitmentP_ccsf"]) != "")
            {
                //objMailRecordsTemp.ssi_remainingcommitmentp_ccsf = new CrmDecimal();
                //objMailRecordsTemp.ssi_remainingcommitmentp_ccsf.Value = Convert.ToDecimal(Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitmentP_ccsf"]));
                objMailRecordsTemp["ssi_remainingcommitmentp_ccsf"] = Convert.ToDecimal(loDataset.Tables[0].Rows[i]["Ssi_RemainingCommitmentP_ccsf"]);
            }

            ////Legal Entity Name
            //if (Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Name"]) != "")
            //{
            //    objMailRecords.ssi_legalentity = Convert.ToString(loDataset.Tables[0].Rows[i]["Legal Entity Name"]);
            //}
            //Response.Write(Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]));
            //TNR ID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]) != "")
            {
                // objMailRecordsTemp.ssi_tnrid_nv = Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]);
                objMailRecordsTemp["ssi_tnrid_nv"] = Convert.ToString(loDataset.Tables[0].Rows[i]["TNR ID"]);
                //Response.Write(objMailRecords.ssi_tnrid_nv);
            }

            //Mailing Contact Type
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]) != "")
            {
                // objMailRecordsTemp.ssi_mailingcontacttype = Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]);
                objMailRecordsTemp["ssi_mailingcontacttype"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mailing Contact Type"]);
            }

            //Commented
            ////Fund Name
            //if (lstFund.SelectedValue != "")
            //{
            //    // objMailRecordsTemp.ssi_fundname = lstFund.SelectedItem.Text;// Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
            //    objMailRecordsTemp["ssi_fundname"] = lstFund.SelectedItem.Text;// Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
            //}

            //Fund Name
            if (SelectedFund != "")
            {
                // objMailRecordsTemp.ssi_fundname = lstFund.SelectedItem.Text;// Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
                objMailRecordsTemp["ssi_fundname"] = SelectedFund;// Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name"]);
            }

            //Fund Name (to show on wire instrux)
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name to show on wire instrux"]) != "")
            {
                // objMailRecordsTemp.ssi_fundname_fund = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name to show on wire instrux"]);
                objMailRecordsTemp["ssi_fundname_fund"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Fund Name to show on wire instrux"]);
            }

            //Capital Call Payment Method
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Capital Call Payment Method"]) != "")
            {
                //objMailRecordsTemp.ssi_capitalcallpaymentmethod = Convert.ToString(loDataset.Tables[0].Rows[i]["Capital Call Payment Method"]);
                objMailRecordsTemp["ssi_capitalcallpaymentmethod"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Capital Call Payment Method"]);
            }

            //Payment Method Note
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Payment Method Note"]) != "")
            {
                // objMailRecordsTemp.ssi_paymentmethodnote = Convert.ToString(loDataset.Tables[0].Rows[i]["Payment Method Note"]);
                objMailRecordsTemp["ssi_paymentmethodnote"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Payment Method Note"]);
            }

            //AsofDate
            if (txtAsofdate.Text != "")
            {
                //objMailRecordsTemp.ssi_asofdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_asofdate.Value = txtAsofdate.Text;
                objMailRecordsTemp["ssi_asofdate"] = Convert.ToDateTime(txtAsofdate.Text);

            }

            //MailingID
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["MailingID"]) != "")
            {
                //objMailRecordsTemp.ssi_mailingid = new CrmNumber();
                //objMailRecordsTemp.ssi_mailingid.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["MailingID"]);
                objMailRecordsTemp["ssi_mailingid"] = Convert.ToInt32(loDataset.Tables[0].Rows[i]["MailingID"]);

            }

            //Mail
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]) != "")
            {
                //objMailRecordsTemp.ssi_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]);
                objMailRecordsTemp["ssi_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail"]);
            }

            //Owner First Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]) != "")
            {
                // objMailRecordsTemp.ssi_ownerfirstname_hh_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]);
                objMailRecordsTemp["ssi_ownerfirstname_hh_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner First Name"]);
            }

            //Owner Last Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]) != "")
            {
                //objMailRecordsTemp.ssi_ownerlname_hh_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]);
                objMailRecordsTemp["ssi_ownerlname_hh_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Owner Last Name"]);
            }

            //Contact Household
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]) != "")
            {
                // objMailRecordsTemp.ssi_hholdinst_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]);
                objMailRecordsTemp["ssi_hholdinst_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Household"]);
            }

            //Contact Full Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]) != "")
            {
                //objMailRecordsTemp.ssi_fullname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]);
                objMailRecordsTemp["ssi_fullname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Contact Full Name"]);
            }

            //Secondary Owner First Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]) != "")
            {
                //objMailRecordsTemp.ssi_secownerfname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]);
                objMailRecordsTemp["ssi_secownerfname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner First Name"]);
            }

            //Secondary Owner Last Name
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]) != "")
            {
                //objMailRecordsTemp.ssi_secownerlname_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]);
                objMailRecordsTemp["ssi_secownerlname_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Secondary Owner Last Name"]);
            }

            //Address Line 1
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]) != "")
            {
                //objMailRecordsTemp.ssi_addressline1_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]);
                objMailRecordsTemp["ssi_addressline1_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 1"]);
            }

            //Address Line 2
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]) != "")
            {
                //objMailRecordsTemp.ssi_addressline2_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]);
                objMailRecordsTemp["ssi_addressline2_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 2"]);
            }

            //Address Line 3
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]) != "")
            {
                // objMailRecordsTemp.ssi_addressline3_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]);
                objMailRecordsTemp["ssi_addressline3_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Address Line 3"]);
            }

            //City
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["City"]) != "")
            {
                // objMailRecordsTemp.ssi_city_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["City"]);
                objMailRecordsTemp["ssi_city_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["City"]);
            }

            //State or Province
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]) != "")
            {
                // objMailRecordsTemp.ssi_stateprovince_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]);
                objMailRecordsTemp["ssi_stateprovince_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["State or Province"]);
            }

            //ZIP Code
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]) != "")
            {
                //objMailRecordsTemp.ssi_zipcode_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]);
                objMailRecordsTemp["ssi_zipcode_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ZIP Code"]);
            }

            //Country or Region
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]) != "")
            {
                //objMailRecordsTemp.ssi_countryregion_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]);
                objMailRecordsTemp["ssi_countryregion_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Country or Region"]);
            }

            //Dear
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]) != "")
            {
                //objMailRecordsTemp.ssi_dear_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]);
                objMailRecordsTemp["ssi_dear_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Dear"]);
            }

            //Salutation
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]) != "")
            {
                // objMailRecordsTemp.ssi_salutation_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]);
                objMailRecordsTemp["ssi_salutation_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Salutation"]);
            }

            //Mail Preference 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]) != "")
            {
                // objMailRecordsTemp.ssi_mailpreference_mail = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]);
                objMailRecordsTemp["ssi_mailpreference_mail"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Mail Preference"]);
            }

            //ssi_mailStatus 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]) != "")
            {
                //objMailRecordsTemp.ssi_mailstatus = new Picklist();
                //objMailRecordsTemp.ssi_mailstatus.Value = Convert.ToInt32(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]);
                objMailRecordsTemp["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(Convert.ToInt32(loDataset.Tables[0].Rows[i]["ssi_mailStatus"]));

            }

            //HouseHold lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"]) != "")
            {
                //objMailRecordsTemp.ssi_accountid = new Lookup();
                //objMailRecordsTemp.ssi_accountid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"]));
                objMailRecordsTemp["ssi_accountid"] = new Microsoft.Xrm.Sdk.EntityReference("account", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["AccountId"])));
            }

            //Contact lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"]) != "")
            {
                //objMailRecordsTemp.ssi_contactfullnameid = new Lookup();
                //objMailRecordsTemp.ssi_contactfullnameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"]));
                objMailRecordsTemp["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ContactId"])));
            }

            //ssi_LegalEntityId lookup
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]) != "")
            {
                //objMailRecords.ssi_legalentity = new  = new CrmNumber();
                //objMailRecordsTemp.ssi_legalentitynameid = new Lookup();
                //objMailRecordsTemp.ssi_legalentitynameid.Value = new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"]));
                objMailRecordsTemp["ssi_legalentitynameid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_legalentity", new Guid(Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_LegalEntityId"])));
            }


            //Advisor Approval Required
            //objMailRecordsTemp.ssi_advisorapprovalreqd = new CrmBoolean();
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "0" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "False".ToUpper())
            {
                //objMailRecordsTemp.ssi_advisorapprovalreqd.Value = false;
                objMailRecordsTemp["ssi_advisorapprovalreqd"] = false;

            }
            else if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]) == "1" || Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_advisorapprovalreqd"]).ToUpper() == "True".ToUpper())
            {
                //objMailRecordsTemp.ssi_advisorapprovalreqd.Value = true;
                objMailRecordsTemp["ssi_advisorapprovalreqd"] = true;
            }

            //File Name Added on 9 oct 2014
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]) != "")
            {
                //objMailRecordsTemp.ssi_filename = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]);
                objMailRecordsTemp["ssi_filename"] = Convert.ToString(loDataset.Tables[0].Rows[i]["Ssi_FileName"]);
            }

            //added by sasmit 5_3_2017
            //clientportalname 
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]) != "")
            {
                //objMailRecordsTemp.ssi_clientportalname  = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]);
                objMailRecordsTemp["ssi_clientportalname"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientportalname"]);
            }
            //added by sasmit 5_3_2017
            //clientreportfolder
            if (Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]) != "")
            {
                //objMailRecordsTemp.ssi_clientreportfolder = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]);
                objMailRecordsTemp["ssi_clientreportfolder"] = Convert.ToString(loDataset.Tables[0].Rows[i]["ssi_clientreportfolder"]);
            }

            // CreatedByCustomid Field 
            //Rohit Pawar
            string Userid = GetcurrentUser();

            if (Userid != "")
            {
                //objMailRecordsTemp.ssi_createdbycustomid = new Lookup();
                //objMailRecordsTemp.ssi_createdbycustomid.Value = new Guid(Userid);
                objMailRecordsTemp["ssi_createdbycustomid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(Userid));
            }


            #region Logic for Capital Call Approval Process

            // Logic for Capital Call Approval Process

            //if (chkUnify.Checked == true)
            //{
            //    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
            //    //objMailRecordsTemp.ssi_unifiedflg.Value = true;
            //    objMailRecordsTemp["ssi_unifiedflg"] = true;

            //}
            //else if (chkUnify.Checked == false)
            //{
            //    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
            //    //objMailRecordsTemp.ssi_unifiedflg.Value = false;
            //    objMailRecordsTemp["ssi_unifiedflg"] = false;
            //}

            if (MailId != "" && MailId != "0")
            {
                //objMailRecordsTemp.ssi_mailidtemp = new CrmNumber();
                //objMailRecordsTemp.ssi_mailidtemp.Value = Convert.ToInt32(MailId);
                objMailRecordsTemp["ssi_mailidtemp"] = Convert.ToInt32(MailId);

            }


            if (ddlTemplates.SelectedValue != "" && ddlTemplates.SelectedValue != "0")
            {
                //objMailRecordsTemp.ssi_templateid = new Lookup();
                //objMailRecordsTemp.ssi_templateid.Value = new Guid(ddlTemplates.SelectedValue);
                objMailRecordsTemp["ssi_templateid"] = new Microsoft.Xrm.Sdk.EntityReference("ssi_template", new Guid(ddlTemplates.SelectedValue));
            }

            //Wire AsofDate
            if (txtWireAsofDate.Text != "")
            {
                //objMailRecordsTemp.ssi_wireasofdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_wireasofdate.Value = txtWireAsofDate.Text;
                objMailRecordsTemp["ssi_wireasofdate"] = Convert.ToDateTime(txtWireAsofDate.Text);

            }

            //Wire AsofDate
            if (txtLetterDate.Text != "")
            {
                //objMailRecordsTemp.ssi_letterdate = new CrmDateTime();
                //objMailRecordsTemp.ssi_letterdate.Value = txtLetterDate.Text;
                objMailRecordsTemp["ssi_letterdate"] = Convert.ToDateTime(txtLetterDate.Text);
            }


            #endregion


            if (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0" && chkUnify.Checked == false)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
            else if (ddlMailId.SelectedValue != "" && chkUnify.Checked == true)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }
            else if (ddlMailId.SelectedValue != "" && chkUnify.Checked == false)
            {
                service.Create(objMailRecordsTemp);
                intResult++;
            }


            //Response.Write(intResult.ToString());
        }


        #region Update Mailing List

        for (int j = 0; j < loDataset.Tables[1].Rows.Count; j++)
        {
            // objMailingList = new ssi_mailinglist();
            Entity objMailingList = new Entity("ssi_mailinglist");
            if (Convert.ToString(loDataset.Tables[1].Rows[j]["Ssi_MailingListID"]) != "")
            {
                //objMailingList.ssi_mailinglistid = new Key();
                //objMailingList.ssi_mailinglistid.Value = new Guid(Convert.ToString(loDataset.Tables[1].Rows[j]["Ssi_MailingListID"]));
                objMailingList["ssi_mailinglistid"] = new Guid(Convert.ToString(loDataset.Tables[1].Rows[j]["Ssi_MailingListID"]));

            }


            if (txtWireAsofDate.Text != "")
            {
                //objMailingList.ssi_capitalcalldate = new CrmDateTime();
                //objMailingList.ssi_capitalcalldate.Value = txtWireAsofDate.Text;
                objMailingList["ssi_capitalcalldate"] = Convert.ToDateTime(txtWireAsofDate.Text);

            }

            service.Update(objMailingList);
            intResult++;
        }

        #endregion


        if (intResult > 0)
        {
            Success++;
            //lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully";
        }

        if (loDataset.Tables[1].Rows.Count == 0 && loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == false && (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0"))
            // lblError.Text = "No Records Found"; // changed 27_8_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[1].Rows.Count == 0 && loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == false && (ddlMailId.SelectedValue != "" || ddlMailId.SelectedValue != "0"))
            // lblError.Text = "No Records Found"; // changed 27_8_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[1].Rows.Count == 0 && loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == true && (ddlMailId.SelectedValue == "" || ddlMailId.SelectedValue == "0"))
            //  lblError.Text = "No Records Found";// changed 27_8_2019 (Basecamp Request)
            lblError.Text = "No records loaded, please check setup/file.";
        else if (loDataset.Tables[1].Rows.Count == 0 && loDataset.Tables[0].Rows.Count == 0 && chkUnify.Checked == true && (ddlMailId.SelectedValue != "" || ddlMailId.SelectedValue != "0"))
        {
            UpdateMailRecords(MailId);
            BindMailingId();
            lblError.Text = "No records found to save but records Unified successfully";
        }


        return intResult;
    }
    public bool ReadAlPsFile(string SourcePath, string DestinationPath, string sheetname)
    {
        bool bProceed = false;
        bool totalflag = false;
        try
        {
            string License = AppLogic.GetParam(AppLogic.ConfigParam.SpireLicense);
            Spire.License.LicenseProvider.SetLicenseKey(License);
            Spire.License.LicenseProvider.LoadLicense();


            // string lsFileNamforFinalXls = "Alps distrubutionfile.xlsx";

            // string lsFileNamforFinalXls = inputfile;
            // string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);

            string strDirectory = SourcePath;

            Workbook workbook = new Workbook();

            //open an excel file

            workbook.LoadFromFile(strDirectory);

            //get the first worksheet

            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = sheetname;
            // sheet.Name = "Capital Call";

            for (int i = sheet.Pictures.Count - 1; i >= 0; i--)
            {
                sheet.Pictures[i].Remove();
            }


            sheet.DeleteRow(1, 15);


            string freezpane = sheet.IsFreezePanes.ToString();//checking for excel contains frrezpane

            sheet.RemovePanes();//remove freezpane



            int cnt = 0;
            int Rowcnt = 0;

            foreach (CellRange range in sheet.Columns[3])
            {
                var str = range.Text;

                cnt++;
                if (range.Text != null)
                {
                    if (range.Text.ToLower().Contains("total"))
                    {
                        //int rowcount = range.RowCount;
                        Rowcnt = cnt;
                        totalflag = true;
                    }
                }
            }

            //sheet.DeleteRow(103, 104);//delete total line
            //sheet.DeleteRow(105, 109);//delete footnote
            //                          //save the excel file

            if (totalflag)
            {

                sheet.DeleteRow(Rowcnt, Rowcnt + 1);//delete total line
                sheet.DeleteRow(Rowcnt + 2, Rowcnt + 6);//delete footnote
                                                        //save the excel file

            }
            sheet.DeleteColumn(1);//delete # column that contains id  

            // workbook.SaveToFile("sample.xlsx", ExcelVersion.Version2016);
            workbook.SaveToFile(DestinationPath, ExcelVersion.Version2016);

            /* destination folder */
            // System.Diagnostics.Process.Start(workbook.FileName);
            bProceed = true;
        }
        catch (Exception ex)
        {
            bProceed = false;
        }
        return bProceed;

    }
    public int ProcessZip(string TempPath, string dateTime, string MailType)
    {
        int count = 0;
        try
        {
            List<string> folderlist = new List<string>();
            if (FileUpload1.HasFile == true)
            {

                if (System.IO.Path.GetExtension(FileUpload1.FileName) == ".zip")
                {
                    //  int count = 0;


                    string ZipFileNamewithExtension = FileUpload1.FileName;
                    string ZipFileNamewithoutExtension = Path.GetFileNameWithoutExtension(FileUpload1.FileName);
                    string BackupFileName = ZipFileNamewithoutExtension + dateTime + ".zip";
                    string ZipBackupPath = string.Empty;// Server.MapPath("") + @"\ExcelTemplate\" + BackupFileName;


                    if (MailType.ToLower() == "a1a079a4-d7be-e011-a19b-0019b9e7ee05" || MailType.ToLower() == "81091a9b-2ae9-e011-9141-0019b9e7ee05")//Cap Call Letter - Cap Call Wire
                    {
                        ZipBackupPath = AppLogic.GetParam(AppLogic.ConfigParam.CapCallBackupPath);

                    }
                    else if (MailType.ToLower() == "6d7545da-8164-e111-bd8f-0019b9e7ee05" || MailType.ToLower() == "78612b2b-5add-e011-ad4d-0019b9e7ee05")//Fund Distribution Letter - Fund Distribution  
                    {
                        ZipBackupPath = AppLogic.GetParam(AppLogic.ConfigParam.DistributionBackupPath);
                    }


                    ZipBackupPath = ZipBackupPath + BackupFileName;

                    string TempFilePath = TempPath + ZipFileNamewithExtension;

                    if (!Directory.Exists(TempPath))
                    {
                        Directory.CreateDirectory(TempPath);
                    }

                    //Backup Zip File                           

                    FileUpload1.PostedFile.SaveAs(ZipBackupPath);// Copy File To Backup Folder.
                    FileUpload1.PostedFile.SaveAs(TempPath + ZipFileNamewithExtension);//Copy File To Temp Folder For Processing

                    //  ZipFile.ExtractToDirectory(ZipBackupPath, TempPath);
                    ZipFile.ExtractToDirectory(ZipBackupPath, TempPath);
                    string DirectoryName1 = Path.GetFileNameWithoutExtension(TempPath + ZipFileNamewithExtension);

                    string[] Folderindirectory = Directory.GetDirectories(TempPath);
                    foreach (string subdir in Folderindirectory)
                    {

                        string[] filesindirectory = Directory.GetFiles(subdir);
                        foreach (string FileinFolder in filesindirectory)
                        {
                            string Fileextension = Path.GetExtension(FileinFolder);//Extension of the File Inside The Folder of The uploaded Zip.
                            if (Fileextension == ".xlsx")
                            {
                                string DirectoryName = Path.GetFileNameWithoutExtension(subdir); // Name of The Fund - FolderName inside the Zip.
                                File.Copy(FileinFolder, TempPath + DirectoryName + Fileextension, true); // Copy files in the TempFolder
                                lstFile.Add(TempPath + DirectoryName + Fileextension, DirectoryName);
                                count++;
                            }
                        }

                    }
                }
                else
                {
                    lblError.Text = "Please Upload Zip File";
                }
            }
        }
        catch (Exception ex)
        {
            lblError.Text = "ERROR : " + ex.Message.ToString();
            bProceed = false;

        }
        return count;
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
        da.SelectCommand.CommandTimeout = 300;
        DataSet ds = new DataSet();
        da.Fill(ds);
        da.Dispose();
        CloseConnection();
        return (ds);
    }
    private string GetcurrentUser()
    {
        string UserID = string.Empty;
        string strName = string.Empty;
        //  System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
        // string strName = Request.LogonUserIdentity.Name;// p.Identity.Name;
        //Response.Write("p.Identity.Name:" + strName + "<br/><br/>");
        //strName = HttpContext.Current.User.Identity.Name.ToString();
        //Response.Write("HttpContext.Current.User.Identity.Name:" + strName + "<br/><br/>");
        //strName = Request.ServerVariables["AUTH_USER"]; //Finding with name


        if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
        {
            strName = "corp\\gbhagia";
        }
        else
        {
            IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
            strName = claimsIdentity.Name;

        }


        //Response.Write("AUTH_USER:" + strName + "<br/><br/>");
        //////////
        //"select top 1 internalemailaddress,systemuserid from systemuser where domainname= 'Signature\\" + strName + "'";
        string sqlstr = "select top 1 internalemailaddress,systemuserid from systemuser where domainname= '" + strName + "'";
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


    private void UpdateMailRecords(string MailId)
    {
        int intResult = 0;
        //string test = ddlMailType.SelectedValue;
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://Crm01/";
        ////string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
        ////string orgName = "Webdev";
        //CrmService service = null;
        IOrganizationService service = null;

        string test = ddlMailType.SelectedValue;

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


        try
        {
            //Response.Write(service.Url);
            //service.PreAuthenticate = true;
            //service.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        catch (NullReferenceException ne)
        {
            //Response.Write(ne.StackTrace + "<br/>" + ne.Message);
        }





        try
        {

            #region Update Mailing records to Unify
            clsDB = new DB();
            string strSql = "SP_S_MAIL_UNIFY @MailingId = " + Convert.ToInt32(MailId);
            DataSet UpdateTempRecords = clsDB.getDataSet(strSql);

            // ssi_mailrecordstemp objMailRecordsTemp = null;

            if (UpdateTempRecords.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < UpdateTempRecords.Tables[0].Rows.Count; i++)
                {
                    // objMailRecordsTemp = new ssi_mailrecordstemp();
                    Entity objMailRecordsTemp = new Entity("ssi_mailrecordstemp");
                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_mailrecordstempId"]) != "")
                    {
                        //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
                        //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_mailrecordstempId"]));
                        objMailRecordsTemp["ssi_mailrecordstempid"] = new Guid(Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_mailrecordstempId"]));

                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_contactfullname"]) != "")
                    {
                        //objMailRecordsTemp.ssi_contactfullname = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_contactfullname"]);
                        objMailRecordsTemp["ssi_contactfullname"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_contactfullname"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_contactfullname = "";
                        objMailRecordsTemp["ssi_contactfullname"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_contactfullnameid"]) != "")
                    {
                        //objMailRecordsTemp.ssi_contactfullnameid = new Lookup();
                        //objMailRecordsTemp.ssi_contactfullnameid.Value = new Guid(Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_contactfullnameid"]));
                        objMailRecordsTemp["ssi_contactfullnameid"] = new Microsoft.Xrm.Sdk.EntityReference("contact", new Guid(Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_contactfullnameid"])));
                    }


                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_dear_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_dear_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_dear_mail"]);
                        objMailRecordsTemp["ssi_dear_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_dear_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_dear_mail = "";
                        objMailRecordsTemp["ssi_dear_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_salutation_mail"]) != "")
                    {
                        //  objMailRecordsTemp.ssi_salutation_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_salutation_mail"]);
                        objMailRecordsTemp["ssi_salutation_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_salutation_mail"]);
                    }
                    else
                    {
                        //   objMailRecordsTemp.ssi_salutation_mail = "";
                        objMailRecordsTemp["ssi_salutation_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline1_mail"]) != "")
                    {
                        // objMailRecordsTemp.ssi_addressline1_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline1_mail"]);
                        objMailRecordsTemp["ssi_addressline1_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline1_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_addressline1_mail = "";
                        objMailRecordsTemp["ssi_addressline1_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline2_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_addressline2_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline2_mail"]);
                        objMailRecordsTemp["ssi_addressline2_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline2_mail"]);
                    }
                    else
                    {
                        // objMailRecordsTemp.ssi_addressline2_mail = "";
                        objMailRecordsTemp["ssi_addressline2_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline3_mail"]) != "")
                    {
                        // objMailRecordsTemp.ssi_addressline3_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline3_mail"]);
                        objMailRecordsTemp["ssi_addressline3_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_addressline3_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_addressline3_mail = "";
                        objMailRecordsTemp["ssi_addressline3_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_city_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_city_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_city_mail"]);
                        objMailRecordsTemp["ssi_city_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_city_mail"]);
                    }
                    else
                    {
                        //  objMailRecordsTemp.ssi_city_mail = "";
                        objMailRecordsTemp["ssi_city_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_stateprovince_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_stateprovince_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_stateprovince_mail"]);
                        objMailRecordsTemp["ssi_stateprovince_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_stateprovince_mail"]);
                    }
                    else
                    {
                        // objMailRecordsTemp.ssi_stateprovince_mail = "";
                        objMailRecordsTemp["ssi_stateprovince_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_zipcode_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_zipcode_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_zipcode_mail"]);
                        objMailRecordsTemp["ssi_zipcode_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_zipcode_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_zipcode_mail = "";
                        objMailRecordsTemp["ssi_zipcode_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_countryregion_mail"]) != "")
                    {
                        //objMailRecordsTemp.ssi_countryregion_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_countryregion_mail"]);
                        objMailRecordsTemp["ssi_countryregion_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_countryregion_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_countryregion_mail = "";
                        objMailRecordsTemp["ssi_countryregion_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_fullname_mail"]) != "")
                    {
                        // objMailRecordsTemp.ssi_fullname_mail = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_fullname_mail"]);
                        objMailRecordsTemp["ssi_fullname_mail"] = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["ssi_fullname_mail"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_fullname_mail = "";
                        objMailRecordsTemp["ssi_fullname_mail"] = "";
                    }

                    //File Name Added on 9 oct 2014 //commented becuase not required in unify logic
                    //if (Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_Filename"]) != "")
                    //{
                    //    objMailRecordsTemp.ssi_filename = Convert.ToString(UpdateTempRecords.Tables[0].Rows[i]["Ssi_Filename"]);
                    //}

                    //objMailRecordsTemp.ssi_updateunifyflg = new CrmBoolean();
                    //objMailRecordsTemp.ssi_updateunifyflg.Value = true;
                    objMailRecordsTemp["ssi_updateunifyflg"] = true;

                    service.Update(objMailRecordsTemp);
                    intResult++;
                }
            }



            #endregion

            #region Update Unify Flg

            if (UpdateTempRecords.Tables[1].Rows.Count > 0)
            {
                for (int i = 0; i < UpdateTempRecords.Tables[1].Rows.Count; i++)
                {
                    // objMailRecordsTemp = new ssi_mailrecordstemp();
                    Entity objMailRecordsTemp = new Entity("ssi_mailrecordstemp");
                    if (ddlMailId.SelectedValue != "")
                    {
                        if (Convert.ToString(UpdateTempRecords.Tables[1].Rows[i]["Ssi_mailrecordstempId"]) != "")
                        {
                            //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
                            //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(UpdateTempRecords.Tables[1].Rows[i]["Ssi_mailrecordstempId"]));
                            objMailRecordsTemp["ssi_mailrecordstempid"] = new Guid(Convert.ToString(UpdateTempRecords.Tables[1].Rows[i]["Ssi_mailrecordstempId"]));

                        }

                        //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
                        //objMailRecordsTemp.ssi_unifiedflg.Value = true;
                        objMailRecordsTemp["ssi_unifiedflg"] = true;

                        service.Update(objMailRecordsTemp);
                        intResult++;
                    }
                }
            }

            #endregion

            #region Update Salutation to Unify

            if (UpdateTempRecords.Tables[2].Rows.Count > 0)
            {
                for (int i = 0; i < UpdateTempRecords.Tables[2].Rows.Count; i++)
                {
                    //objMailRecordsTemp = new ssi_mailrecordstemp();
                    Entity objMailRecordsTemp = new Entity("ssi_mailrecordstemp");
                    if (Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Ssi_mailrecordstempId"]) != "")
                    {
                        //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
                        //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Ssi_mailrecordstempId"]));
                        objMailRecordsTemp["ssi_mailrecordstempid"] = new Guid(Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Ssi_mailrecordstempId"]));

                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Sas_JointSalutation"]) != "")
                    {
                        //objMailRecordsTemp.ssi_salutation_mail = Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Sas_JointSalutation"]);
                        objMailRecordsTemp["ssi_salutation_mail"] = Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["Sas_JointSalutation"]);
                    }
                    else
                    {
                        //objMailRecordsTemp.ssi_salutation_mail = "";
                        objMailRecordsTemp["ssi_salutation_mail"] = "";
                    }

                    if (Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["sas_dear2"]) != "")
                    {
                        // objMailRecordsTemp.ssi_dear_mail = Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["sas_dear2"]);
                        objMailRecordsTemp["ssi_dear_mail"] = Convert.ToString(UpdateTempRecords.Tables[2].Rows[i]["sas_dear2"]);
                    }
                    else
                    {
                        // objMailRecordsTemp.ssi_dear_mail = "";
                        objMailRecordsTemp["ssi_dear_mail"] = "";
                    }

                    //objMailRecordsTemp.ssi_unifiedflg = new CrmBoolean();
                    //objMailRecordsTemp.ssi_unifiedflg.Value = true;
                    objMailRecordsTemp["ssi_unifiedflg"] = true;


                    service.Update(objMailRecordsTemp);
                    intResult++;
                }
            }



            #endregion


            if (intResult > 0)
            {
                if (ddlMailType.SelectedItem.Text != "")
                {
                    lblError.Text = ddlMailType.SelectedItem.Text + " records saved successfully and records Unified successfully";
                    //lblError.Text = ddlMailType.SelectedItem.Text + " records Unified successfully";
                }
                else
                {
                    lblError.Text = "Records Unified successfully";
                }
            }
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
    protected void ddlMailId_SelectedIndexChanged(object sender, EventArgs e)
    {
        // added 5_27_2019 - Zip File Upload fro CapitalCall and Fund Distribution
        lblmailid.Text = "";
        lblmailid.Visible = false;

        if (ddlMailId.SelectedValue != "" && ddlMailId.SelectedValue != "0")
        {
            clsDB = new DB();
            string sql = "SP_S_TemplateTemp_List @MailID=" + ddlMailId.SelectedValue;
            DataSet table = clsDB.getDataSet(sql);

            if (Convert.ToString(table.Tables[0].Rows[0]["ssi_templateid"]) != "")
            {
                ddlTemplates.SelectedValue = Convert.ToString(table.Tables[0].Rows[0]["ssi_templateid"]);
                ddlTemplates.Enabled = false;
            }


            if (Convert.ToString(table.Tables[0].Rows[0]["ssi_templateid"]) != "" && ddlMailId.SelectedValue != "" && ddlMailId.SelectedValue != "0")
            {
                string strsql = "SP_S_TemplateLettDate_List @MailID=" + ddlMailId.SelectedValue + ",@Ssi_Templateid='" + Convert.ToString(table.Tables[0].Rows[0]["ssi_templateid"]) + "'";
                DataSet ds = clsDB.getDataSet(strsql);

                if (Convert.ToString(ds.Tables[0].Rows[0]["Ssi_WireAsOfDate"]) != "" || Convert.ToString(ds.Tables[0].Rows[0]["Ssi_LetterDate"]) != "")
                {
                    txtWireAsofDate.Text = Convert.ToString(ds.Tables[0].Rows[0]["Ssi_WireAsOfDate"]);
                    txtWireAsofDate.Enabled = false;
                    img1.Style.Add("display", "none");

                    txtLetterDate.Text = Convert.ToString(ds.Tables[0].Rows[0]["Ssi_LetterDate"]);
                    txtLetterDate.Enabled = false;
                    img2.Style.Add("display", "none");
                }
                else
                {
                    txtWireAsofDate.Enabled = false;
                    txtLetterDate.Enabled = false;
                    img1.Style.Add("display", "none");
                    img2.Style.Add("display", "none");
                }
            }

        }
        else
        {
            ddlTemplates.Enabled = true;
            txtWireAsofDate.Enabled = true;
            txtLetterDate.Enabled = true;
            img1.Style.Add("display", "inline");
            img2.Style.Add("display", "inline");
        }
    }
    protected void ddlMailType_SelectedIndexChanged(object sender, EventArgs e)
    {
        lbtnExceptionReport.Visible = false;
        ddlTemplates.Enabled = true;
        txtWireAsofDate.Enabled = true;
        txtLetterDate.Enabled = true;
        img1.Style.Add("display", "table-row");
        img2.Style.Add("display", "");
        ddlTemplates.SelectedValue = "0";
        ddlMailId.Visible = true;
        lblmailid.Visible = false;
        trAutoDebitDate.Style.Add("display", "none");
        Label3.Text = "Position As Of Date:";
        trFund.Style.Add("display", "");

        ddlMailId.SelectedValue = "0";

        if (ddlMailType.SelectedValue == "a1a079a4-d7be-e011-a19b-0019b9e7ee05" || ddlMailType.SelectedValue == "6d7545da-8164-e111-bd8f-0019b9e7ee05" || ddlMailType.SelectedValue == "78612b2b-5add-e011-ad4d-0019b9e7ee05" || ddlMailType.SelectedValue == "81091a9b-2ae9-e011-9141-0019b9e7ee05")//Capital Call Letter
        {
            trUnify.Style.Add("display", "table-row");
            trBrowsefiles.Style.Add("display", "table-row");
            trMonths.Style.Add("display", "table-row");//Commented after mail type = billing added 
            trMonths.Style.Add("display", "none");
            trMailID.Style.Add("display", "table-row");
            trReportTemplate.Style.Add("display", "table-row");
            trWireAsof.Style.Add("display", "table-row");
            trLetter.Style.Add("display", "table-row");
            trLegalentity.Style.Add("display", "table-row");
            lstFund.SelectionMode = ListSelectionMode.Multiple; // added 5_27_2019 - Zip File Upload fro CapitalCall and Fund Distribution

        }
        else if (ddlMailType.SelectedValue == "3cbaf86d-5edd-e011-ad4d-0019b9e7ee05") //Fund Mailing (Signature Required)
        {
            trUnify.Style.Add("display", "table-row");
            trBrowsefiles.Style.Add("display", "none");
            trMonths.Style.Add("display", "none");
            trMailID.Style.Add("display", "table-row");
            trReportTemplate.Style.Add("display", "table-row");
            trWireAsof.Style.Add("display", "none");
            trLetter.Style.Add("display", "table-row");
            trLegalentity.Style.Add("display", "table-row");
            lstFund.ClearSelection();// added 5_27_2019 - Zip File Upload fro CapitalCall and Fund Distribution
            lstFund.SelectionMode = ListSelectionMode.Single; // added 5_27_2019 - Zip File Upload fro CapitalCall and Fund Distribution
        }
        else if (ddlMailType.SelectedValue == "3fb190d9-b2cd-e011-a19b-0019b9e7ee05") //Billing 
        {
            trFund.Style.Add("display", "none");
            trUnify.Style.Add("display", "none");
            trBrowsefiles.Style.Add("display", "table-row");
            trMonths.Style.Add("display", "table-row");
            // trMailID.Style.Add("display", "none");
            trMailID.Style.Add("display", "table-row");
            trReportTemplate.Style.Add("display", "none");
            trWireAsof.Style.Add("display", "none");
            trLetter.Style.Add("display", "table-row");
            trLegalentity.Style.Add("display", "none");
            Label3.Text = "AUM As Of Date:";
            trBrowsefiles.Style.Add("display", "none");
            trAutoDebitDate.Style.Add("display", "");
            lstFund.ClearSelection();// added 5_27_2019 - Zip File Upload fro CapitalCall and Fund Distribution
            lstFund.SelectionMode = ListSelectionMode.Single; // added 5_27_2019 - Zip File Upload fro CapitalCall and Fund Distribution
        }
        else
        {
            trUnify.Style.Add("display", "none");
            trBrowsefiles.Style.Add("display", "none");
            trMonths.Style.Add("display", "none");
            trMailID.Style.Add("display", "none");
            trReportTemplate.Style.Add("display", "none");
            trWireAsof.Style.Add("display", "none");
            trLegalentity.Style.Add("display", "none");
            //trLetter.Style.Add("display", "none");
            lstFund.ClearSelection();// added 5_27_2019 - Zip File Upload fro CapitalCall and Fund Distribution
            lstFund.SelectionMode = ListSelectionMode.Single; // added 5_27_2019 - Zip File Upload fro CapitalCall and Fund Distribution
        }
    }





    protected void lbtnExceptionReport_Click(object sender, EventArgs e)
    {
        string SampleFilePath = ViewState["ExcetionReportPath"].ToString();
        Download_File(SampleFilePath, "ExceptionReport.xlsx");
    }
}
