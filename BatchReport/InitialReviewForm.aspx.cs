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
using System.Net.Mail;

using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using System.Threading;
using Microsoft.IdentityModel.Claims;

public partial class _InitialReviewForm : System.Web.UI.Page
{
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
    bool chk;
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName, AdvisorFlag, BatchStatusID;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            BindGridView();
            BindDropDown();
        }
        if (ddlAction.SelectedValue != "5")
        {
            tblBrowse.Style.Add("display", "none");
        }
        else if (ddlAction.SelectedValue == "5")
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

        GridView1.Columns[9].Visible = true;
        GridView1.Columns[10].Visible = true;
        GridView1.Columns[11].Visible = true;
        GridView1.Columns[12].Visible = true;
        GridView1.Columns[13].Visible = true;
        GridView1.Columns[14].Visible = true;
        GridView1.Columns[15].Visible = true;
        GridView1.Columns[16].Visible = true;

        GridView1.Columns[18].Visible = true;
        GridView1.Columns[19].Visible = true;
        GridView1.Columns[20].Visible = true;

        // GridView1.Columns[15].Visible = true;

        GridView1.DataSource = loDataset;
        GridView1.DataBind();

        GridView1.Columns[9].Visible = false;
        GridView1.Columns[10].Visible = false;
        GridView1.Columns[11].Visible = false;
        GridView1.Columns[12].Visible = false;
        GridView1.Columns[13].Visible = false;
        GridView1.Columns[14].Visible = false;
        GridView1.Columns[15].Visible = false;
        GridView1.Columns[16].Visible = false;

        GridView1.Columns[18].Visible = false;
        GridView1.Columns[19].Visible = false;
        GridView1.Columns[20].Visible = false;

        // GridView1.Columns[15].Visible = false;


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


    public void checkforbillinguser()
    {
        string strName = string.Empty;
        //Changed Windows to - ADFS Claims Login 8_9_2019
        if (HttpContext.Current.Request.Url.Host.ToLower() == "localhost")
        {
            strName = "corp\\gbhagia";
        }
        else
        {
            //  IClaimsIdentity claimsIdentity = Thread.CurrentPrincipal.Identity as IClaimsIdentity;
            //   strName = claimsIdentity.Name;

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
        BindAsOfDate();
        BindCreatedBy();
    }


    public void BindAsOfDate()
    {
        //ddl.Items.Clear();
        sqlstr = "SP_S_InitialReview_Asofdate";

        clsGM.getListForBindDDL(ddlAsOfDate, sqlstr, "ssi_AsOfDate", "ssi_AsOfDate");

        ddlAsOfDate.Items.Insert(0, "All");
        ddlAsOfDate.Items[0].Value = "0";
        ddlAsOfDate.SelectedIndex = 0;
    }

    public void BindCreatedBy()
    {
        //ddl.Items.Clear();
        sqlstr = "SP_S_InitialReview_CreatedBy";

        clsGM.getListForBindDDL(ddlCreatedBy, sqlstr, "ssi_createdbycustomidname", "ssi_createdbycustomid");

        ddlCreatedBy.Items.Insert(0, "All");
        ddlCreatedBy.Items[0].Value = "0";
        ddlCreatedBy.SelectedIndex = 0;
    }


    public void CallMailqueform()
    {
        string BatchType = string.Empty;
        string BatchIdListTxt = string.Empty;
        foreach (GridViewRow row in GridView1.Rows)
        {
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
            string batchid = row.Cells[9].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
            BatchType = "4";
            if (chkSelectNC.Checked)
            {
                if (BatchIdListTxt == "")
                    BatchIdListTxt = batchid;
                else
                    BatchIdListTxt = BatchIdListTxt + "," + batchid;
            }
        }

        Session["BatchIdList"] = BatchIdListTxt;

        string Billingflg = "true";
        string csname2 = "ClientScript";
        System.Text.StringBuilder cstext2 = new System.Text.StringBuilder();
        cstext2.Append("<script type=\"text/javascript\"> ");
        cstext2.Append("window.open('MailQueue.aspx?btypeid=" + BatchType + "&bflag=" + Billingflg + "') </");//?bidlist=" + BatchIdListTxt + "'
        cstext2.Append("script>");
        // RegisterClientScriptBlock(csname2, cstext2.ToString());

        ClientScript.RegisterClientScriptBlock(GetType(), csname2, cstext2.ToString());
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
        sqlstr = "SP_S_IR_Recipient";
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
        //object RecipientId = ddlRecipient.SelectedValue == "0" || ddlRecipient.SelectedValue == "" ? "null" : "'" + ddlRecipient.SelectedValue + "'";
        object BatchType = ddlBatchtype.SelectedValue == "0" || ddlBatchtype.SelectedValue == "" ? "null" : ddlBatchtype.SelectedValue;

        lstBox.Items.Clear();
        //",@RecipientId=" + RecipientId +
        sqlstr = "SP_S_IR_HouseHoldName @AdvisorId=" + AdvisorId + ",@AssociateId=" + AssociatedId + ",@BatchTypeId=" + BatchType;
        clsGM.getListForBindListBox(lstBox, sqlstr, "HouseholdName", "ssi_householdid");

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

    public string BindGrid()
    {
        String lsSQL = "";//

        object AsOfDate = ddlAsOfDate.SelectedValue == "0" || ddlAsOfDate.SelectedValue == "" ? "null" : "'" + ddlAsOfDate.SelectedValue + "'";
        object BatchType = ddlBatchtype.SelectedValue == "0" || ddlBatchtype.SelectedValue == "" ? "null" : ddlBatchtype.SelectedValue;
        object Advisor = ddlAdvisor.SelectedValue == "0" || ddlAdvisor.SelectedValue == "" ? "null" : "'" + ddlAdvisor.SelectedValue + "'";
        object Associate = ddlAssociate.SelectedValue == "0" || ddlAssociate.SelectedValue == "" ? "null" : "'" + ddlAssociate.SelectedValue + "'";
        object HouseHold = lstHouseHold.SelectedValue == "0" || lstHouseHold.SelectedValue == "" ? "null" : "'" + clsGM.GetMultipleSelectedItemsFromListBox(lstHouseHold) + "'";
        //object BatchOwner = ddlBatchOwner.SelectedValue == "0" || ddlBatchOwner.SelectedValue == "" ? "null" : "'" + ddlBatchOwner.SelectedValue + "'";
        //object BatchStatus = ddlBatchstatus.SelectedValue == "0" || ddlBatchstatus.SelectedValue == "" ? "null" : ddlBatchstatus.SelectedValue;
        object MailStatus = ddlMailStatus.SelectedValue == "0" || ddlMailStatus.SelectedValue == "" || ddlMailStatus.SelectedValue == "9999" ? "null" : ddlMailStatus.SelectedValue;
        object Recipient = ddlRecipient.SelectedValue == "0" || ddlRecipient.SelectedValue == "" ? "null" : "'" + ddlRecipient.SelectedValue + "'";
        object CreatedBy = ddlCreatedBy.SelectedValue == "0" || ddlCreatedBy.SelectedValue == "" ? "null" : "'" + ddlCreatedBy.SelectedValue + "'";

        lsSQL = "SP_S_INITIAL_REVIEW @AsofDate=" + AsOfDate + ",@BatchTypeId=" + BatchType + ",@AdvisorId=" + Advisor + ",@AssociateId=" + Associate + ",@HouseHoldIdNmbList=" + HouseHold + ",@MailStatusId=" + MailStatus + ",@ReceipentId=" + Recipient + ",@createdbycustomid=" + CreatedBy;

        return lsSQL;

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
    #region Commented OLDCODE CRM2016 Upgrade
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
    protected void btnSubmit_Click(object sender, EventArgs e)
    {
        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://crm01/";
        ////string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
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
        int count = 0;
        bool bproceed = true;
        bool bpathproceed = true;
        bool bBatchStatus = true;
        bool bBillingFlag = false;
        bool bcreatedflg = false;

        try
        {
            // service = GetCrmService(crmServerUrl, orgName);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        // catch (System.Web.Services.Protocols.SoapException exc)
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

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;


        if (ddlAction.SelectedValue == "5")
        {
            //int count = 0;
            //bool bproceed = true;
            //bool bpathproceed = true;
            //bool bBatchStatus = true;
            //ssi_batch objBatch = null;

            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string ssi_batchid = row.Cells[9].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                //string BatchStatusID = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                string BatchFilePath = row.Cells[12].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string BatchFileName = row.Cells[11].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
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

            //if (!bBatchStatus)
            //{
            //    lblError.Text = "Can not merge pdf once batch status is 'OPS Approved'";
            //    lblError.Visible = true;
            //    return;
            //}


            foreach (GridViewRow row in GridView1.Rows)
            {
                CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
                string ssi_batchid = row.Cells[9].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                //string BatchStatusID = row.Cells[19].Text.Trim().Replace("BatchStatusID", "").Replace("&nbsp;", "");
                string BatchFilePath = row.Cells[12].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
                string BatchFileName = row.Cells[11].Text.Trim().Replace("ssi_batchdisplayfilename", "").Replace("&nbsp;", "");
                string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + BatchFileName);

                if (chkSelectNC.Checked == true)
                {
                    if (FileUpload1.HasFile == true)
                    {
                        if (count == 1)// && BatchStatusID != "8"
                        {

                            string str = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + BatchFileName); //FileName
                            FileUpload1.PostedFile.SaveAs(Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + FileUpload1.FileName);
                            string filename = Path.GetFileName(FileUpload1.FileName);

                            string strClientPath = Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + filename;
                            // File.Copy(BatchFilePath, strDirectory, true);
                            string str2 = BatchFilePath;

                            string[] str3 = new string[2];
                            str3[0] = str2;
                            str3[1] = strClientPath;
                            PDFMerge pdfMerge = new PDFMerge();
                            pdfMerge.MergeFiles(str, str3);

                            File.Copy(str, BatchFilePath, true);

                            if (File.Exists(strClientPath))
                            {
                                File.Delete(strClientPath);
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
        if (ddlAction.SelectedValue == "4") //Generate single PDF Report.
        {
            string status = GenerateMergeTypeConsolidatedPDF();
            //Response.Write("Status:" + status);
            if (status != "")
            {
                lblError.Text = "Report generated successfully";
                lblError.Visible = true;

                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();
                sb.Append("\n<script type=text/javascript>\n");
                sb.Append("\nwindow.open('ViewReport.aspx?" + status + "', 'mywindow');");
                sb.Append("</script>");
                ClientScript.RegisterClientScriptBlock(tp, "Script", sb.ToString());
            }
            else
            {
                lblError.Text = "Please select only batch related mail type";
            }

            return;
        }



        //1	Approve
        //2	Reject

        //ssi_batch objBatch = null;
        //ssi_mailrecords objMailRecords = null;
        //ssi_mailrecordstemp objMailRecordsTemp = null;
        //ssi_wireexecution ObjWireExe = null;


        foreach (GridViewRow row in GridView1.Rows)
        {
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");
            DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");

            string Batchid = row.Cells[9].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
            string BatchName = row.Cells[3].Text.Trim().Replace("Batch Name", "").Replace("&nbsp;", "");
            string secondaryownerid = row.Cells[13].Text.Trim().Replace("ssi_secondaryownerid", "").Replace("&nbsp;", "");
            string MailrecordsId = row.Cells[10].Text.Trim().Replace("ssi_mailrecordsid", "").Replace("&nbsp;", "");

            string BillingInvoiceid = row.Cells[20].Text.Trim().Replace("ssi_BillingInvoiceId", "").Replace("&nbsp;", ""); ;
            //string MailrecordsTempId = row.Cells[14].Text.Trim().Replace("ssi_mailrecordstempid", "").Replace("&nbsp;", "");
            string BatchTypeId = row.Cells[16].Text.Trim().Replace("ssi_mailrecordstempid", "").Replace("&nbsp;", "");

            string AdvisorApprovalFlg = row.Cells[18].Text.Trim().Replace("&nbsp;", "");//AdvisorApprovalFlg
            string AssociateApprovalFlg = row.Cells[19].Text.Trim().Replace("&nbsp;", "");//AssociateApprovalFlg

            string MailingStatus = row.Cells[5].Text.Trim().Replace("&nbsp;", "");//MailingStatus

            //objBatch = new ssi_batch();
            //objMailRecords = new ssi_mailrecords();
            //objMailRecordsTemp = new ssi_mailrecordstemp();
            Entity objBatch = new Entity("ssi_batch");
            Entity objMailRecords = new Entity("ssi_mailrecords");
            Entity objMailRecordsTemp = new Entity("ssi_mailrecordstemp");
            try
            {
                if (chkSelectNC.Checked == true)
                {
                    if (ddlAction.SelectedValue == "1")//Approve
                    {

                        if (ddlBatchtype.SelectedItem.Text.ToUpper() == "BILLING")
                        {
                            if (MailingStatus.ToUpper() == "CREATED")
                            {
                                //lblError.Text = "The selected batches have already been Created and Saved to SharePoint.";
                                //return;
                                bcreatedflg = true;

                            }

                            else
                            {

                                #region Update batch

                                if (Batchid != "")
                                {
                                    //objBatch.ssi_batchid = new Key();
                                    //objBatch.ssi_batchid.Value = new Guid(Batchid);
                                    objBatch["ssi_batchid"] = new Guid(Batchid);

                                    //objBatch.ssi_reporttrackerstatus = new Picklist();
                                    //objBatch.ssi_reporttrackerstatus.Value = 6;//Pend Approval

                                    //if(ddlBatchtype.SelectedItem.Text.ToUpper()!="BILLING" && MailingStatus.ToUpper()!="CREATED")
                                    //{ 

                                    if (AdvisorApprovalFlg.ToUpper() == "FALSE" && AssociateApprovalFlg.ToUpper() == "FALSE")
                                    {
                                        //if (ddlBatchtype.SelectedItem.Text.ToUpper() == "BILLING")
                                        //{
                                        //    if (MailingStatus.ToUpper() == "CREATED")
                                        //    {
                                        //        bcreatedflg = true;
                                        //    }
                                        //}


                                        //else
                                        //{
                                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(8);//8: OPS APPROVED
                                        bBillingFlag = true;
                                        //}
                                    }
                                    else
                                    {
                                        objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(6);
                                    }

                                    if (secondaryownerid != "")
                                    {
                                        //SecurityPrincipal assignee = new SecurityPrincipal();
                                        //assignee.PrincipalId = new Guid(secondaryownerid);////Advisor

                                        //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();
                                        //targetAssign.EntityId = new Guid(Batchid);
                                        //targetAssign.EntityName = EntityName.ssi_batch.ToString();

                                        //AssignRequest assign = new AssignRequest();
                                        //assign.Assignee = assignee;
                                        //assign.Target = targetAssign;

                                        //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);


                                        AssignRequest assignRequest = new AssignRequest
                                        {
                                            Assignee = new EntityReference("systemuser",
                                             new Guid(secondaryownerid)),
                                            Target = new EntityReference("ssi_batch",
                                             new Guid(Batchid))
                                        };



                                        service.Execute(assignRequest);

                                    }

                                    //if (!bcreatedflg)
                                    service.Update(objBatch);
                                    selectedCount++;
                                }
                                #endregion

                                #region Update MailRecords

                                if (MailrecordsId != "")
                                {
                                    //objMailRecords.ssi_mailrecordsid = new Key();
                                    //objMailRecords.ssi_mailrecordsid.Value = new Guid(MailrecordsId);
                                    objMailRecords["ssi_mailrecordsid"] = new Guid(MailrecordsId);

                                    //objMailRecords.ssi_mailstatus = new Picklist();
                                    //objMailRecords.ssi_mailstatus.Value = 5;//Pend Approval

                                    if (AdvisorApprovalFlg.ToUpper() == "FALSE" && AssociateApprovalFlg.ToUpper() == "FALSE")
                                    {
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(7);//7-Approved
                                    }
                                    else
                                    {
                                        objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(5);
                                    }

                                    //objMailRecords.ssi_ir_status = new Picklist();
                                    //objMailRecords.ssi_ir_status.Value = 2;//Approved
                                    objMailRecords["ssi_ir_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(2);

                                    service.Update(objMailRecords);
                                    selectedCount++;
                                }

                                #endregion

                                #region Update Mail Records Temp

                                if (Batchid != "")
                                {
                                    string sql = "SP_S_MailRecordsTempID_Batch @BatchID='" + Batchid + "'";
                                    DataSet BatchDataset = clsDB.getDataSet(sql);

                                    for (int i = 0; i < BatchDataset.Tables[0].Rows.Count; i++)
                                    {
                                        if (Convert.ToString(BatchDataset.Tables[0].Rows[i]["ssi_mailrecordstempid"]) != "")
                                        {
                                            //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
                                            //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(BatchDataset.Tables[0].Rows[i]["ssi_mailrecordstempid"]));
                                            objMailRecordsTemp["ssi_mailrecordstempid"] = new Guid(Convert.ToString(BatchDataset.Tables[0].Rows[i]["ssi_mailrecordstempid"]));

                                            //objMailRecordsTemp.ssi_batchstatus = new Picklist();
                                            //objMailRecordsTemp.ssi_batchstatus.Value = 2; //Approve

                                            objMailRecordsTemp["ssi_batchstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(2);

                                            service.Update(objMailRecordsTemp);
                                            selectedCount++;
                                        }
                                    }
                                }

                                #endregion

                                #region Wire Execution
                                //if (Batchid != "")
                                //{

                                //    if (BatchTypeId == "1") // Capital Call 
                                //    {
                                //        #region Capital Call

                                //        string strsql = "SP_S_CapitalCall_WireExecution @BatchIdList='" + Batchid + "'";
                                //        DataSet WireExeDataset = clsDB.getDataSet(strsql);

                                //        for (int i = 0; i < WireExeDataset.Tables[0].Rows.Count; i++)
                                //        {
                                //            ObjWireExe = new ssi_wireexecution();

                                //            //ssi_name
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["Name"]) != "")
                                //            {
                                //                ObjWireExe.ssi_name = Convert.ToString(WireExeDataset.Tables[0].Rows[i]["Name"]);
                                //            }


                                //            //ssi_typeid
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_typeid"]) != "")
                                //            {
                                //                ObjWireExe.ssi_type = new Picklist();
                                //                ObjWireExe.ssi_type.Value = Convert.ToInt32(WireExeDataset.Tables[0].Rows[i]["ssi_typeid"]);
                                //            }


                                //            //ssi_legalentitynameid
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_legalentityid"]) != "")
                                //            {
                                //                ObjWireExe.ssi_legalentityid = new Lookup();
                                //                ObjWireExe.ssi_legalentityid.Value = new Guid(Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_legalentityid"]));
                                //            }

                                //            //ssi_Householdid
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_Householdid"]) != "")
                                //            {
                                //                ObjWireExe.ssi_householdid = new Lookup();
                                //                ObjWireExe.ssi_householdid.Value = new Guid(Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_Householdid"]));
                                //            }

                                //            //ssi_totaladjustedcommitment
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_totaladjustedcommitment"]) != "")
                                //            {
                                //                ObjWireExe.ssi_totaladjustedcommitment = new CrmMoney();
                                //                ObjWireExe.ssi_totaladjustedcommitment.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_totaladjustedcommitment"]);
                                //            }


                                //            //ssi_capitalcalltotalpercentcalled
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcalltotalpercentcalled"]) != "")
                                //            {
                                //                ObjWireExe.ssi_capitalcallpercentcalled = new CrmDecimal();
                                //                ObjWireExe.ssi_capitalcallpercentcalled.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcalltotalpercentcalled"]);
                                //            }

                                //            //ssi_amount
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_amount"]) != "")
                                //            {
                                //                ObjWireExe.ssi_amount = new CrmMoney();
                                //                ObjWireExe.ssi_amount.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_amount"]);
                                //            }

                                //            //ssi_capitalcallpriorcalls
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpriorcalls"]) != "")
                                //            {
                                //                ObjWireExe.ssi_capitalcallpriorcalls = new CrmMoney();
                                //                ObjWireExe.ssi_capitalcallpriorcalls.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpriorcalls"]);
                                //            }


                                //            //ssi_capitalcallpercentpriorcalls
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentpriorcalls"]) != "")
                                //            {
                                //                ObjWireExe.ssi_capitalcallpercentpriorcalls = new CrmDecimal();
                                //                ObjWireExe.ssi_capitalcallpercentpriorcalls.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentpriorcalls"]);
                                //            }


                                //            //ssi_capitalcallpercentpriorcalls
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallcalledtodate"]) != "")
                                //            {
                                //                ObjWireExe.ssi_capitalcallcalledtodate = new CrmMoney();
                                //                ObjWireExe.ssi_capitalcallcalledtodate.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallcalledtodate"]);
                                //            }

                                //            //ssi_capitalcallpercentpriorcalls
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentcalledtodate"]) != "")
                                //            {
                                //                ObjWireExe.ssi_capitalcallpercentcalledtodate = new CrmDecimal();
                                //                ObjWireExe.ssi_capitalcallpercentcalledtodate.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentcalledtodate"]);
                                //            }

                                //            //ssi_commitmentremainingcommitment
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitment"]) != "")
                                //            {
                                //                ObjWireExe.ssi_commitmentremainingcommitment = new CrmMoney();
                                //                ObjWireExe.ssi_commitmentremainingcommitment.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitment"]);
                                //            }


                                //            //ssi_commitmentremainingcommitmentpercent
                                //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitmentpercent"]) != "")
                                //            {
                                //                ObjWireExe.ssi_commitmentremainingcommitmentpercent = new CrmDecimal();
                                //                ObjWireExe.ssi_commitmentremainingcommitmentpercent.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitmentpercent"]);
                                //            }

                                //            service.Create(ObjWireExe);
                                //            selectedCount++;

                                //        }

                                //        #endregion
                                //    }
                                //    else if (BatchTypeId == "2")//Distribution
                                //    {
                                //        #region Distribution

                                //        string strsql = "SP_S_Distribution_WireExecution  @BatchIdList='" + Batchid + "'";
                                //        DataSet WireExeDistributionDataset = clsDB.getDataSet(strsql);

                                //        for (int j = 0; j < WireExeDistributionDataset.Tables[0].Rows.Count; j++)
                                //        {

                                //            ObjWireExe = new ssi_wireexecution();
                                //            //ssi_name
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["Name"]) != "")
                                //            {
                                //                ObjWireExe.ssi_name = Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["Name"]);
                                //            }


                                //            //Type
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_type"]) != "")
                                //            {
                                //                ObjWireExe.ssi_type = new Picklist();
                                //                ObjWireExe.ssi_type.Value = Convert.ToInt32(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_type"]);
                                //            }

                                //            //ssi_LegalEntityid
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_LegalEntityid"]) != "")
                                //            {
                                //                ObjWireExe.ssi_legalentityid = new Lookup();
                                //                ObjWireExe.ssi_legalentityid.Value = new Guid(Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_LegalEntityid"]));
                                //            }

                                //            //ssi_Householdid
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_Householdid"]) != "")
                                //            {
                                //                ObjWireExe.ssi_householdid = new Lookup();
                                //                ObjWireExe.ssi_householdid.Value = new Guid(Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_Householdid"]));
                                //            }


                                //            //ssi_totalcommitment
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_totalcommitment"]) != "")
                                //            {
                                //                ObjWireExe.ssi_totalcommitment = new CrmMoney();
                                //                ObjWireExe.ssi_totalcommitment.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_totalcommitment"]);
                                //            }


                                //            //ssi_amount
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_amount"]) != "")
                                //            {
                                //                ObjWireExe.ssi_amount = new CrmMoney();
                                //                ObjWireExe.ssi_amount.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_amount"]);
                                //            }


                                //            //ssi_distributionpercent
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpercent"]) != "")
                                //            {
                                //                ObjWireExe.ssi_distributionpercent = new CrmDecimal();
                                //                ObjWireExe.ssi_distributionpercent.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpercent"]);
                                //            }


                                //            //ssi_distributionpriordistributions
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributions"]) != "")
                                //            {
                                //                ObjWireExe.ssi_distributionpriordistributions = new CrmMoney();
                                //                ObjWireExe.ssi_distributionpriordistributions.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributions"]);
                                //            }

                                //            //ssi_distributionpriordistributionspercent
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributionspercent"]) != "")
                                //            {
                                //                ObjWireExe.ssi_distributionpriordistributionspercent = new CrmDecimal();
                                //                ObjWireExe.ssi_distributionpriordistributionspercent.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributionspercent"]);
                                //            }

                                //            //ssi_distributionstodate
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionstodate"]) != "")
                                //            {
                                //                ObjWireExe.ssi_distributionstodate = new CrmMoney();
                                //                ObjWireExe.ssi_distributionstodate.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionstodate"]);
                                //            }

                                //            //ssi_DistributionstoDatePct
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_DistributionstoDatePct"]) != "")
                                //            {
                                //                ObjWireExe.ssi_distributionstodatepct = new CrmDecimal();
                                //                ObjWireExe.ssi_distributionstodatepct.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_DistributionstoDatePct"]);
                                //            }


                                //            //ssi_capitalcallcalledtodate
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallcalledtodate"]) != "")
                                //            {
                                //                ObjWireExe.ssi_capitalcallcalledtodate = new CrmMoney();
                                //                ObjWireExe.ssi_capitalcallcalledtodate.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallcalledtodate"]);
                                //            }


                                //            //ssi_commitmentremainingcommitment
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_commitmentremainingcommitment"]) != "")
                                //            {
                                //                ObjWireExe.ssi_commitmentremainingcommitment = new CrmMoney();
                                //                ObjWireExe.ssi_commitmentremainingcommitment.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_commitmentremainingcommitment"]);
                                //            }


                                //            //ssi_capitalcallpercentcalledtodate
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallpercentcalledtodate"]) != "")
                                //            {
                                //                ObjWireExe.ssi_capitalcallpercentcalledtodate = new CrmDecimal();
                                //                ObjWireExe.ssi_capitalcallpercentcalledtodate.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallpercentcalledtodate"]);
                                //            }


                                //            //ssi_distributionsclassbfeeadjustment
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsclassbfeeadjustment"]) != "")
                                //            {
                                //                ObjWireExe.ssi_distributionsclassbfeeadjustement = new CrmMoney();
                                //                ObjWireExe.ssi_distributionsclassbfeeadjustement.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsclassbfeeadjustment"]);
                                //            }



                                //            //ssi_distributionsactualcashdistributions 
                                //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsactualcashdistributions"]) != "")
                                //            {
                                //                ObjWireExe.ssi_distributionsactualcashdistributions = new CrmMoney();
                                //                ObjWireExe.ssi_distributionsactualcashdistributions.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsactualcashdistributions"]);
                                //            }

                                //            service.Create(ObjWireExe);
                                //            selectedCount++;

                                //        }



                                //        #endregion
                                //    }
                                //}

                                #endregion
                            }

                        }

                        else
                        {

                            #region Update batch

                            if (Batchid != "")
                            {
                                //objBatch.ssi_batchid = new Key();
                                //objBatch.ssi_batchid.Value = new Guid(Batchid);
                                objBatch["ssi_batchid"] = new Guid(Batchid);

                                //objBatch.ssi_reporttrackerstatus = new Picklist();
                                //objBatch.ssi_reporttrackerstatus.Value = 6;//Pend Approval

                                //if(ddlBatchtype.SelectedItem.Text.ToUpper()!="BILLING" && MailingStatus.ToUpper()!="CREATED")
                                //{ 

                                if (AdvisorApprovalFlg.ToUpper() == "FALSE" && AssociateApprovalFlg.ToUpper() == "FALSE")
                                {
                                    //if (ddlBatchtype.SelectedItem.Text.ToUpper() == "BILLING")
                                    //{
                                    //    if (MailingStatus.ToUpper() == "CREATED")
                                    //    {
                                    //        bcreatedflg = true;
                                    //    }
                                    //}


                                    //else
                                    //{
                                    objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(8);//8: OPS APPROVED
                                    bBillingFlag = true;
                                    //}
                                }
                                else
                                {
                                    objBatch["ssi_reporttrackerstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(6);
                                }

                                if (secondaryownerid != "")
                                {
                                    //SecurityPrincipal assignee = new SecurityPrincipal();
                                    //assignee.PrincipalId = new Guid(secondaryownerid);////Advisor

                                    //TargetOwnedDynamic targetAssign = new TargetOwnedDynamic();
                                    //targetAssign.EntityId = new Guid(Batchid);
                                    //targetAssign.EntityName = EntityName.ssi_batch.ToString();

                                    //AssignRequest assign = new AssignRequest();
                                    //assign.Assignee = assignee;
                                    //assign.Target = targetAssign;

                                    //AssignResponse assignResponse = (AssignResponse)service.Execute(assign);


                                    AssignRequest assignRequest = new AssignRequest
                                    {
                                        Assignee = new EntityReference("systemuser",
                                         new Guid(secondaryownerid)),
                                        Target = new EntityReference("ssi_batch",
                                         new Guid(Batchid))
                                    };



                                    service.Execute(assignRequest);

                                }

                                //if (!bcreatedflg)
                                service.Update(objBatch);
                                selectedCount++;
                            }
                            #endregion

                            #region Update MailRecords

                            if (MailrecordsId != "")
                            {
                                //objMailRecords.ssi_mailrecordsid = new Key();
                                //objMailRecords.ssi_mailrecordsid.Value = new Guid(MailrecordsId);
                                objMailRecords["ssi_mailrecordsid"] = new Guid(MailrecordsId);

                                //objMailRecords.ssi_mailstatus = new Picklist();
                                //objMailRecords.ssi_mailstatus.Value = 5;//Pend Approval

                                if (AdvisorApprovalFlg.ToUpper() == "FALSE" && AssociateApprovalFlg.ToUpper() == "FALSE")
                                {
                                    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(7);//7-Approved
                                }
                                else
                                {
                                    objMailRecords["ssi_mailstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(5);
                                }

                                //objMailRecords.ssi_ir_status = new Picklist();
                                //objMailRecords.ssi_ir_status.Value = 2;//Approved
                                objMailRecords["ssi_ir_status"] = new Microsoft.Xrm.Sdk.OptionSetValue(2);

                                service.Update(objMailRecords);
                                selectedCount++;
                            }

                            #endregion

                            #region Update Mail Records Temp

                            if (Batchid != "")
                            {
                                string sql = "SP_S_MailRecordsTempID_Batch @BatchID='" + Batchid + "'";
                                DataSet BatchDataset = clsDB.getDataSet(sql);

                                for (int i = 0; i < BatchDataset.Tables[0].Rows.Count; i++)
                                {
                                    if (Convert.ToString(BatchDataset.Tables[0].Rows[i]["ssi_mailrecordstempid"]) != "")
                                    {
                                        //objMailRecordsTemp.ssi_mailrecordstempid = new Key();
                                        //objMailRecordsTemp.ssi_mailrecordstempid.Value = new Guid(Convert.ToString(BatchDataset.Tables[0].Rows[i]["ssi_mailrecordstempid"]));
                                        objMailRecordsTemp["ssi_mailrecordstempid"] = new Guid(Convert.ToString(BatchDataset.Tables[0].Rows[i]["ssi_mailrecordstempid"]));

                                        //objMailRecordsTemp.ssi_batchstatus = new Picklist();
                                        //objMailRecordsTemp.ssi_batchstatus.Value = 2; //Approve

                                        objMailRecordsTemp["ssi_batchstatus"] = new Microsoft.Xrm.Sdk.OptionSetValue(2);

                                        service.Update(objMailRecordsTemp);
                                        selectedCount++;
                                    }
                                }
                            }

                            #endregion

                            #region Wire Execution
                            //if (Batchid != "")
                            //{

                            //    if (BatchTypeId == "1") // Capital Call 
                            //    {
                            //        #region Capital Call

                            //        string strsql = "SP_S_CapitalCall_WireExecution @BatchIdList='" + Batchid + "'";
                            //        DataSet WireExeDataset = clsDB.getDataSet(strsql);

                            //        for (int i = 0; i < WireExeDataset.Tables[0].Rows.Count; i++)
                            //        {
                            //            ObjWireExe = new ssi_wireexecution();

                            //            //ssi_name
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["Name"]) != "")
                            //            {
                            //                ObjWireExe.ssi_name = Convert.ToString(WireExeDataset.Tables[0].Rows[i]["Name"]);
                            //            }


                            //            //ssi_typeid
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_typeid"]) != "")
                            //            {
                            //                ObjWireExe.ssi_type = new Picklist();
                            //                ObjWireExe.ssi_type.Value = Convert.ToInt32(WireExeDataset.Tables[0].Rows[i]["ssi_typeid"]);
                            //            }


                            //            //ssi_legalentitynameid
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_legalentityid"]) != "")
                            //            {
                            //                ObjWireExe.ssi_legalentityid = new Lookup();
                            //                ObjWireExe.ssi_legalentityid.Value = new Guid(Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_legalentityid"]));
                            //            }

                            //            //ssi_Householdid
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_Householdid"]) != "")
                            //            {
                            //                ObjWireExe.ssi_householdid = new Lookup();
                            //                ObjWireExe.ssi_householdid.Value = new Guid(Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_Householdid"]));
                            //            }

                            //            //ssi_totaladjustedcommitment
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_totaladjustedcommitment"]) != "")
                            //            {
                            //                ObjWireExe.ssi_totaladjustedcommitment = new CrmMoney();
                            //                ObjWireExe.ssi_totaladjustedcommitment.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_totaladjustedcommitment"]);
                            //            }


                            //            //ssi_capitalcalltotalpercentcalled
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcalltotalpercentcalled"]) != "")
                            //            {
                            //                ObjWireExe.ssi_capitalcallpercentcalled = new CrmDecimal();
                            //                ObjWireExe.ssi_capitalcallpercentcalled.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcalltotalpercentcalled"]);
                            //            }

                            //            //ssi_amount
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_amount"]) != "")
                            //            {
                            //                ObjWireExe.ssi_amount = new CrmMoney();
                            //                ObjWireExe.ssi_amount.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_amount"]);
                            //            }

                            //            //ssi_capitalcallpriorcalls
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpriorcalls"]) != "")
                            //            {
                            //                ObjWireExe.ssi_capitalcallpriorcalls = new CrmMoney();
                            //                ObjWireExe.ssi_capitalcallpriorcalls.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpriorcalls"]);
                            //            }


                            //            //ssi_capitalcallpercentpriorcalls
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentpriorcalls"]) != "")
                            //            {
                            //                ObjWireExe.ssi_capitalcallpercentpriorcalls = new CrmDecimal();
                            //                ObjWireExe.ssi_capitalcallpercentpriorcalls.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentpriorcalls"]);
                            //            }


                            //            //ssi_capitalcallpercentpriorcalls
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallcalledtodate"]) != "")
                            //            {
                            //                ObjWireExe.ssi_capitalcallcalledtodate = new CrmMoney();
                            //                ObjWireExe.ssi_capitalcallcalledtodate.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallcalledtodate"]);
                            //            }

                            //            //ssi_capitalcallpercentpriorcalls
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentcalledtodate"]) != "")
                            //            {
                            //                ObjWireExe.ssi_capitalcallpercentcalledtodate = new CrmDecimal();
                            //                ObjWireExe.ssi_capitalcallpercentcalledtodate.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_capitalcallpercentcalledtodate"]);
                            //            }

                            //            //ssi_commitmentremainingcommitment
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitment"]) != "")
                            //            {
                            //                ObjWireExe.ssi_commitmentremainingcommitment = new CrmMoney();
                            //                ObjWireExe.ssi_commitmentremainingcommitment.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitment"]);
                            //            }


                            //            //ssi_commitmentremainingcommitmentpercent
                            //            if (Convert.ToString(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitmentpercent"]) != "")
                            //            {
                            //                ObjWireExe.ssi_commitmentremainingcommitmentpercent = new CrmDecimal();
                            //                ObjWireExe.ssi_commitmentremainingcommitmentpercent.Value = Convert.ToDecimal(WireExeDataset.Tables[0].Rows[i]["ssi_commitmentremainingcommitmentpercent"]);
                            //            }

                            //            service.Create(ObjWireExe);
                            //            selectedCount++;

                            //        }

                            //        #endregion
                            //    }
                            //    else if (BatchTypeId == "2")//Distribution
                            //    {
                            //        #region Distribution

                            //        string strsql = "SP_S_Distribution_WireExecution  @BatchIdList='" + Batchid + "'";
                            //        DataSet WireExeDistributionDataset = clsDB.getDataSet(strsql);

                            //        for (int j = 0; j < WireExeDistributionDataset.Tables[0].Rows.Count; j++)
                            //        {

                            //            ObjWireExe = new ssi_wireexecution();
                            //            //ssi_name
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["Name"]) != "")
                            //            {
                            //                ObjWireExe.ssi_name = Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["Name"]);
                            //            }


                            //            //Type
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_type"]) != "")
                            //            {
                            //                ObjWireExe.ssi_type = new Picklist();
                            //                ObjWireExe.ssi_type.Value = Convert.ToInt32(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_type"]);
                            //            }

                            //            //ssi_LegalEntityid
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_LegalEntityid"]) != "")
                            //            {
                            //                ObjWireExe.ssi_legalentityid = new Lookup();
                            //                ObjWireExe.ssi_legalentityid.Value = new Guid(Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_LegalEntityid"]));
                            //            }

                            //            //ssi_Householdid
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_Householdid"]) != "")
                            //            {
                            //                ObjWireExe.ssi_householdid = new Lookup();
                            //                ObjWireExe.ssi_householdid.Value = new Guid(Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_Householdid"]));
                            //            }


                            //            //ssi_totalcommitment
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_totalcommitment"]) != "")
                            //            {
                            //                ObjWireExe.ssi_totalcommitment = new CrmMoney();
                            //                ObjWireExe.ssi_totalcommitment.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_totalcommitment"]);
                            //            }


                            //            //ssi_amount
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_amount"]) != "")
                            //            {
                            //                ObjWireExe.ssi_amount = new CrmMoney();
                            //                ObjWireExe.ssi_amount.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_amount"]);
                            //            }


                            //            //ssi_distributionpercent
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpercent"]) != "")
                            //            {
                            //                ObjWireExe.ssi_distributionpercent = new CrmDecimal();
                            //                ObjWireExe.ssi_distributionpercent.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpercent"]);
                            //            }


                            //            //ssi_distributionpriordistributions
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributions"]) != "")
                            //            {
                            //                ObjWireExe.ssi_distributionpriordistributions = new CrmMoney();
                            //                ObjWireExe.ssi_distributionpriordistributions.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributions"]);
                            //            }

                            //            //ssi_distributionpriordistributionspercent
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributionspercent"]) != "")
                            //            {
                            //                ObjWireExe.ssi_distributionpriordistributionspercent = new CrmDecimal();
                            //                ObjWireExe.ssi_distributionpriordistributionspercent.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionpriordistributionspercent"]);
                            //            }

                            //            //ssi_distributionstodate
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionstodate"]) != "")
                            //            {
                            //                ObjWireExe.ssi_distributionstodate = new CrmMoney();
                            //                ObjWireExe.ssi_distributionstodate.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionstodate"]);
                            //            }

                            //            //ssi_DistributionstoDatePct
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_DistributionstoDatePct"]) != "")
                            //            {
                            //                ObjWireExe.ssi_distributionstodatepct = new CrmDecimal();
                            //                ObjWireExe.ssi_distributionstodatepct.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_DistributionstoDatePct"]);
                            //            }


                            //            //ssi_capitalcallcalledtodate
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallcalledtodate"]) != "")
                            //            {
                            //                ObjWireExe.ssi_capitalcallcalledtodate = new CrmMoney();
                            //                ObjWireExe.ssi_capitalcallcalledtodate.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallcalledtodate"]);
                            //            }


                            //            //ssi_commitmentremainingcommitment
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_commitmentremainingcommitment"]) != "")
                            //            {
                            //                ObjWireExe.ssi_commitmentremainingcommitment = new CrmMoney();
                            //                ObjWireExe.ssi_commitmentremainingcommitment.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_commitmentremainingcommitment"]);
                            //            }


                            //            //ssi_capitalcallpercentcalledtodate
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallpercentcalledtodate"]) != "")
                            //            {
                            //                ObjWireExe.ssi_capitalcallpercentcalledtodate = new CrmDecimal();
                            //                ObjWireExe.ssi_capitalcallpercentcalledtodate.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_capitalcallpercentcalledtodate"]);
                            //            }


                            //            //ssi_distributionsclassbfeeadjustment
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsclassbfeeadjustment"]) != "")
                            //            {
                            //                ObjWireExe.ssi_distributionsclassbfeeadjustement = new CrmMoney();
                            //                ObjWireExe.ssi_distributionsclassbfeeadjustement.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsclassbfeeadjustment"]);
                            //            }



                            //            //ssi_distributionsactualcashdistributions 
                            //            if (Convert.ToString(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsactualcashdistributions"]) != "")
                            //            {
                            //                ObjWireExe.ssi_distributionsactualcashdistributions = new CrmMoney();
                            //                ObjWireExe.ssi_distributionsactualcashdistributions.Value = Convert.ToDecimal(WireExeDistributionDataset.Tables[0].Rows[j]["ssi_distributionsactualcashdistributions"]);
                            //            }

                            //            service.Create(ObjWireExe);
                            //            selectedCount++;

                            //        }



                            //        #endregion
                            //    }
                            //}

                            #endregion
                        }
                    }
                    else if (ddlAction.SelectedValue == "2")//Reject
                    {

                        if (MailrecordsId != "")
                        {
                            //objMailRecords.ssi_mailrecordsid = new Key();
                            //objMailRecords.ssi_mailrecordsid.Value = new Guid(MailrecordsId);
                            objMailRecords["ssi_mailrecordsid"] = new Guid(MailrecordsId);

                            //objMailRecords.ssi_initialreviewer_reject = new CrmBoolean();
                            //objMailRecords.ssi_initialreviewer_reject.Value = true;
                            objMailRecords["ssi_initialreviewer_reject"] = true;

                            //objMailRecords.ssi_review_reject = new CrmBoolean();
                            //objMailRecords.ssi_review_reject.Value = false;
                            objMailRecords["ssi_review_reject"] = false;

                            //objMailRecords.ssi_deleterecord_flg = new CrmBoolean();
                            //objMailRecords.ssi_deleterecord_flg.Value = true;
                            objMailRecords["ssi_deleterecord_flg"] = true;

                            string UserId = GetcurrentUser();

                            if (UserId != "")
                            {
                                //objMailRecords.ssi_rejectedbyuserid = new Lookup();
                                //objMailRecords.ssi_rejectedbyuserid.Value = new Guid(UserId);
                                objMailRecords["ssi_rejectedbyuserid"] = new Microsoft.Xrm.Sdk.EntityReference("systemuser", new Guid(UserId));

                            }
                            //lblError.Text = UserId;
                            //Response.Write(UserId);

                            service.Update(objMailRecords);
                            selectedCount++;
                        }

                        #region Update completed flag on billing invoice

                        //  string strbillingComplteFlag = "SP_S_MailRecordsTempID_List @MailIDList=" + ViewState["MailId"].ToString();// +",@LegalEntityNameID='" + LegalEntityId + "',@ContactFullnameID='" + ContactId + "'";



                        Entity objBillingInvoice = new Entity("ssi_billinginvoice");
                        if (chkSelectNC.Checked)
                        {
                            if (BillingInvoiceid != "")
                            {
                                objBillingInvoice["ssi_billinginvoiceid"] = new Guid(Convert.ToString(BillingInvoiceid));

                                objBillingInvoice["ssi_completed"] = false;

                                objBillingInvoice["ssi_invoicedate"] = null;

                                service.Update(objBillingInvoice);


                                SendEmail(BatchName); /* send email when billing completed flag uncheck added on 04/24/2020*/
                            }
                        }

                        //}


                        #endregion



                    }
                    else if (ddlAction.SelectedValue == "3")
                    {
                        if (Batchid != "")
                        {
                            string sqlMailRecords = "SP_D_MailRecord_Batch @BatchIdList='" + Batchid + "'";
                            //string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords);
                            // string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords, EntityName.ssi_mailrecords, service);
                            string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords, "ssi_mailrecords", service);

                            string sqlBatch = "SP_D_Batch @BatchIdList='" + Batchid + "'";
                            //string DelBatch = clsDB.DeleteRecord(sqlBatch);
                            // string DelBatch = clsDB.DeleteRecord(sqlBatch, EntityName.ssi_batch, service);
                            string DelBatch = clsDB.DeleteRecord(sqlBatch, "ssi_batch", service);
                        }

                        #region Update completed flag on billing invoice

                        //  string strbillingComplteFlag = "SP_S_MailRecordsTempID_List @MailIDList=" + ViewState["MailId"].ToString();// +",@LegalEntityNameID='" + LegalEntityId + "',@ContactFullnameID='" + ContactId + "'";

                        Entity objBillingInvoice = new Entity("ssi_billinginvoice");

                        if (BillingInvoiceid != "")
                        {
                            objBillingInvoice["ssi_billinginvoiceid"] = new Guid(Convert.ToString(BillingInvoiceid));

                            objBillingInvoice["ssi_completed"] = false;


                            objBillingInvoice["ssi_invoicedate"] = null;

                            service.Update(objBillingInvoice);
                        }

                        //}

                        #endregion


                    }
                }
            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                bProceed = false;
                strDescription = "Error occured, Error Detail: " + exc.Detail.InnerText;
                lblError.Text = strDescription;
            }
            catch (Exception exc)
            {
                bProceed = false;
                strDescription = "Error occured, Error Detail: " + exc.Message;
                lblError.Text = strDescription;
            }
        }


        if (bBillingFlag)
            CallMailqueform();



        if (ddlAction.SelectedValue == "1")
        {
            if (selectedCount > 0)
            {
                BindGridView();
                lblError.Visible = true;
                if (bcreatedflg)
                    lblError.Text = "The selected batches have already been Created and Saved to SharePoint.";
                else
                    lblError.Text = "Approved Successfully";

            }

            if (bcreatedflg)
                lblError.Text = "The selected batches have already been Created and Saved to SharePoint.";
        }
        else if (ddlAction.SelectedValue == "2")
        {
            if (selectedCount > 0)
            {
                BindGridView();
                System.Threading.Thread.Sleep(20000);
                DeleteBatchAndMailRecords();
                lblError.Text = "Batch and Mail Records rejected.";
                lblError.Visible = true;
            }
        }
        else if (ddlAction.SelectedValue == "3")
        {
            BindGridView();

            lblError.Text = "Batch and Mail Records removed successfully.";
            lblError.Visible = true;
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
    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {

            string FileName = e.Row.Cells[10].Text.Trim().Replace("ssi_batchfilename", "").Replace("&nbsp;", "");
            ImageButton imgApprovedFile = (ImageButton)e.Row.FindControl("imgApprovedFile");

            string ChkSelect = e.Row.Cells[0].Text.Trim().Replace("chkSelectNC", "").Replace("&nbsp;", "");
            if (ChkSelect != "")
            {
                chk = Convert.ToBoolean(ChkSelect);
            }

            CheckBox chkSelectNC = (CheckBox)e.Row.FindControl("chkSelectNC");

            if (chk == true)
            {
                chkSelectNC.Checked = true;
            }
            else
            {
                chkSelectNC.Checked = false;
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


    private void GenerateReport()
    {
        try
        {
            ////lblError.Text = "";
            //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://crm01/";
            ////string crmServerURL = "http://server:5555/";

            //string orgName = "GreshamPartners";
            string currentuser = null;
            //string orgName = "Webdev";
            // CrmService service = null;
            IOrganizationService service = null;
            Boolean checkrunreport = false;
            String DestinationPath = string.Empty;
            string ConsolidatePdfFileName = string.Empty;
            string ReportOpFolder = string.Empty;
            string ApprovedReports = AppLogic.GetParam(AppLogic.ConfigParam.ApprovedReports);// "\\\\fs01\\opsreports$\\Approved Reports\\"; //"\\\\Fs01\\shared$\\OPS REPORTS\\Approved Reports\\";
            try
            {
                // service = GetCrmService(crmServerUrl, orgName);
                service = clsGM.GetCrmService();
                strDescription = "Crm Service starts successfully";
            }
            // catch (System.Web.Services.Protocols.SoapException exc)
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

            //UserName_YYYYMMDD_Timewhere 

            //ViewState["ParentFolder"] = CurrentDateTime.Replace(":", "-").Replace("/", "-"); // orig

            ViewState["ParentFolder"] = strUserName + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond;

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
                ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.BatchReports);// "\\\\Fs01\\shared$\\BATCH REPORTS\\";

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                ReportOpFolder = @"C:\Reports\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            }
            else
            {
                ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.OpsReports);//"\\\\fs01\\opsreports$";//"\\\\Fs01\\shared$\\OPS REPORTS\\";// +Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

                if (ddlAction.SelectedValue == "2" || ddlAction.SelectedValue == "3")
                    ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.BatchReports);//"\\\\Fs01\\shared$\\BATCH REPORTS\\";
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
                string FileSec = FileDateTime.Second.ToString().Length < 2 ? "0" + FileDateTime.Second.ToString() : FileDateTime.Second.ToString();

                string CurrentTimeStamp = FileYear + "_" + FileMonth + "_" + FileDay + "_" + FileHour + "_" + FileMinute + "_" + FileSec;

                if (chkBox.Checked)
                {
                    checkrunreport = true;
                    String BatchIdListTxt = Convert.ToString(GridView1.Rows[j].Cells[10].Text);
                    dtBatch = GetDataTable(BatchIdListTxt);

                    //String TempName =  GridView1.Rows[j].Cells[6].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();

                    //String HHName = GridView1.Rows[j].Cells[6].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                    //string ssi_batchid = GridView1.Rows[j].Cells[10].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
                    String HHName = GridView1.Rows[j].Cells[16].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();

                    //string TempName = HttpContext.Current.User.Identity.Name.ToString() + "_" + 

                    sourcefilecount = dtBatch.Rows.Count + 1;
                    SourceFileArray = new string[sourcefilecount];

                    for (int i = 0; i < dtBatch.Rows.Count; i++)
                    {
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

                        //if (i == 0)
                        //{
                        //    SourceFileArray[i] = lsCoversheet.Replace(".xls", ".pdf");
                        //    if (CombinedFileName == true)
                        //        SourceFileArray[i + 1] = lsExcleSavePath.Replace(".xls", ".pdf");
                        //}
                        if (i == 0)
                        {
                            SourceFileArray[i] = lsCoversheet.Replace(".xls", ".pdf");
                            if (CombinedFileName == true)
                                SourceFileArray[i + 1] = lsExcleSavePath.Replace(".xls", ".pdf");
                            else if (CombinedFileName == false)
                            {
                                lblError.Text = "No Record Found";
                                lblError.Visible = true;
                                return;
                            }
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

                    //string DisplayFileName = HHName + "_" + strYear + "-" + strMonth + strDay + ".pdf";

                    string DisplayFileName = HHName + " " + strYear + "-" + strMonth + strDay + ".pdf";

                    DisplayFileName = GeneralMethods.RemoveSpecialCharacters(DisplayFileName);
                    DisplayFileName = DisplayFileName.Replace(" Family", "").Replace(",", "");

                    if (!File.Exists(ReportOpFolder + "\\" + ConsolidatePdfFileName))
                        File.Copy(ReportOpFolder + "\\" + ContactFolderName + "\\Coversheet.pdf", ReportOpFolder + "\\" + ConsolidatePdfFileName);

                    DestinationPath = ReportOpFolder + "\\" + GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);

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
                            // objBatch.ssi_batchfilename = DestinationPath;
                            objBatch["ssi_batchfilename"] = DestinationPath;
                        }

                        if (BatchIdListTxt != "")
                        {
                            service.Update(objBatch);
                        }

                        File.Copy(DestinationPath, ApprovedReports + ConsolidatePdfFileName);

                        #endregion
                    }

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

                    strReportFiles = strReportFiles + "<br/>" + "<a href=file:" + AppLogic.GetParam(AppLogic.ConfigParam.OutPutReports) + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1).Replace(" ", "%20") + ">" + DestinationPath.Substring(DestinationPath.LastIndexOf("\\") + 1) + " </a>";
                }
            }

            ////////////////////////////////////

            if (ddlAction.SelectedValue != "1")//Approved
            {
                if (NoOfBatches == 1)
                {
                    string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ConsolidatePdfFileName);

                    File.Copy(DestinationPath, strDirectory, true);
                    File.Delete(DestinationPath);
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
        loChunk = new Chunk("The values shown for the current period and the prior period are subject to the availability of information. In particular, certain non-marketable investments such as commercial real estate and private equity holdings do not provide frequent valuations. In these and other cases, we have either carried the investments at cost or used the general partner's most recent valuation estimates adjusted for subsequent investments or distributions. \"Prior Period Net Worth\" includes the most recent manager provided updated balances, some of which may remain estimated values.", setFontsAll(8, 0, 1, new iTextSharp.text.Color(150, 150, 150)));
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
        Random rand = new Random();
        rand.Next();

        String lsFileNamforFinalXls = Convert.ToString(rand.Next()) + System.DateTime.Now.ToString("MMddyyhhmmss") + ".xls";
        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\coversheet.xls");
        string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);

        string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls.Replace("xls", "xml"));

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
        string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);//"Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";


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
            lsSQL = "SP_R_Advent_Report_Allocation @AllocationGroupNameTxt='" + fsAllocationGroup + "', ";
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

        string str9 = GeneralMethods.RemoveSpecialCharacters(GridView1.Rows[rowIndex].Cells[11].Text);

        if (GridView1.Rows[rowIndex].Cells[12].Text != "")
        {
            string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + str9);

            //Response.Write(strDirectory);

            //File.Copy(AppLogic.GetParam(AppLogic.ConfigParam.OpsReports) + str9, strDirectory, true);

            File.Copy(GridView1.Rows[rowIndex].Cells[12].Text, strDirectory, true);
            //Directory.Delete(ReportOpFolder, true);

            try
            {
                //Response.Write("<script>");
                //string lsFileNamforFinal = "./ExcelTemplate/" + GridView1.Rows[rowIndex].Cells[18].Text;
                //Response.Write("window.open('ViewReport.aspx?" + GridView1.Rows[rowIndex].Cells[18].Text + "', 'mywindow')");
                //Response.Write("</script>");Capital Call - Roger L. Howe Irrev. Trust #1 (Howe GST)-Lerner D 2012-1006.pdf
                Session["id"] = str9;//GridView1.Rows[rowIndex].Cells[9].Text;

                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                Type tp = this.GetType();
                sb.Append("\n<script type=text/javascript>\n");
                sb.Append("\nwindow.open('ViewReport.aspx?" + str9 + "', 'mywindow');");
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
        BindGridView();
    }


    protected void ddlAsOfDate_SelectedIndexChanged(object sender, EventArgs e)
    {
        ClearControls();
        //BindHouseHold(lstHouseHold);

        BindGridView();
    }


    private void DeleteBatchAndMailRecords()
    {

        //string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);//"http://crm01/";
        ////string crmServerURL = "http://server:5555/";
        //string orgName = "GreshamPartners";
        ////string orgName = "Webdev";
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
            // service = GetCrmService(crmServerUrl, orgName);
            service = clsGM.GetCrmService();
            strDescription = "Crm Service starts successfully";
        }
        // catch (System.Web.Services.Protocols.SoapException exc)
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

        //service.PreAuthenticate = true;
        //service.Credentials = System.Net.CredentialCache.DefaultCredentials;

        //1	Approve
        //2	Reject

        //ssi_batch objBatch = null;
        //ssi_mailrecords objMailRecords = null;

        foreach (GridViewRow row in GridView1.Rows)
        {
            CheckBox chkSelectNC = (CheckBox)row.FindControl("chkSelectNC");

            DropDownList ddlHoldReport = (DropDownList)row.FindControl("ddlHoldReport");

            string Batchid = row.Cells[9].Text.Trim().Replace("ssi_batchid", "").Replace("&nbsp;", "");
            string MailRecordsDelete = row.Cells[15].Text.Trim().Replace("ssi_mailrecords_del", "").Replace("&nbsp;", "");
            string MailrecordsId = row.Cells[10].Text.Trim().Replace("ssi_mailrecordsid", "").Replace("&nbsp;", "");

            //objBatch = new ssi_batch();
            //objMailRecords = new ssi_mailrecords();
            Entity objBatch = new Entity("ssi_batch");
            Entity objMailRecords = new Entity("ssi_mailrecords");

            try
            {
                if (ddlAction.SelectedValue == "2")//Approve
                {
                    if (MailRecordsDelete.ToUpper() == "TRUE")
                    {
                        if (Batchid != "")
                        {
                            System.Threading.Thread.Sleep(20000);

                            string sqlMailRecords = "SP_D_MailRecord_Batch @BatchIdList='" + Batchid + "'";
                            //string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords);
                            //string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords,EntityName.ssi_mailrecords,service);
                            string DelMailRecords = clsDB.DeleteRecord(sqlMailRecords, "ssi_mailrecords", service);

                            string sqlBatch = "SP_D_Batch @BatchIdList='" + Batchid + "'";
                            //string DelBatch = clsDB.DeleteRecord(sqlBatch);
                            // string DelBatch = clsDB.DeleteRecord(sqlBatch,EntityName.ssi_batch,service);
                            string DelBatch = clsDB.DeleteRecord(sqlBatch, "ssi_batch", service);
                        }
                    }
                }
            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                bProceed = false;
                strDescription = "Error occured, Error Detail: " + exc.Detail.InnerText;
                lblError.Text = strDescription;
            }
            catch (Exception exc)
            {
                bProceed = false;
                strDescription = "Error occured, Error Detail: " + exc.Message;
                lblError.Text = strDescription;
            }
        }


        if (ddlAction.SelectedValue == "2")//Approve
        {
            BindGridView();
        }
    }



    protected void ddlCreatedBy_SelectedIndexChanged(object sender, EventArgs e)
    {
        BindGridView();
    }



    public string GenerateMergeTypeConsolidatedPDF()
    {
        try
        {
            string[] SourceFileName = new string[0];
            string BatchTypeId = string.Empty;

            int NoOfBatches = 1;
            int checkBoxChecked = 0;
            //this loop will get the number of files to merge according to conditions.
            for (int j = 0; j < GridView1.Rows.Count; j++)
            {
                CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");
                //  string ReportPath = (string)GridView1.Rows[j].Cells[10].Text.Replace("&nbsp;", "");
                string ReportPath = (string)GridView1.Rows[j].Cells[11].Text.Replace("&nbsp;", "");
                if (chkBox.Checked == true && ReportPath.Replace("&nbsp;", "") != "")
                {
                    NoOfBatches = NoOfBatches + 1;
                }
            }

            string FileName = string.Empty;
            int NoofFiles = 0;
            NoofFiles = NoOfBatches * 2;

            SourceFileName = new string[NoofFiles];

            int FileNo = 0;
            //this loop will get the paths of files to merge according to conditions.
            for (int j = 0; j < GridView1.Rows.Count; j++)
            {
                CheckBox chkBox = (CheckBox)GridView1.Rows[j].FindControl("chkSelectNC");
                string ReportPath1 = (string)GridView1.Rows[j].Cells[11].Text;
                if (chkBox.Checked == true && ReportPath1.Replace("&nbsp;", "") != "")
                {
                    string ReportName = (string)GridView1.Rows[j].Cells[11].Text.Replace("&nbsp;", "");
                    string ReportPath = (string)GridView1.Rows[j].Cells[12].Text.Replace("&nbsp;", "");

                    SourceFileName[FileNo] = ReportPath.Replace("&nbsp;", "");
                    FileNo++;
                    if (chkBox.Checked == true && ReportPath1.Replace("&nbsp;", "") != "")
                    {
                        SourceFileName[FileNo] = Server.MapPath("") + "/ExcelTemplate/Template/EndReport.pdf";
                        FileNo++;
                    }
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
            string CombinedFileName = AppLogic.GetParam(AppLogic.ConfigParam.CombinedPdfs) + ConsolidatedPDFFileName + ".pdf";

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                DestinationFileName = "D:\\GRESHAM\\AdventReport\\ExcelTemplate\\BATCH REPORTS\\" + ConsolidatedPDFFileName + ".pdf"; //Server.MapPath("") + "\\ExcelTemplate\\" + ConsolidatedPDFFileName + ".pdf";
            }
            else
                // DestinationFileName = "\\\\GRPAO1-VWFS01\\opsreports$\\Combined PDFs\\" + ConsolidatedPDFFileName + ".pdf";
                DestinationFileName = CombinedFileName;


            string FilePath = string.Empty;
            string FilePath1 = string.Empty;
            //string strTemplate = SourceFileName.Split('|');

            for (int l = 0; l < SourceFileName.Length; l++)
            {
                if (FilePath != "")
                {
                    if (SourceFileName[l] != "" && SourceFileName[l] != "null")
                    {
                        FilePath = FilePath + "|" + SourceFileName[l];
                    }
                }
                else
                {
                    if (SourceFileName[l] != "" && SourceFileName[l] != "null")
                    {
                        FilePath = "|" + SourceFileName[l];
                    }
                }
            }

            FilePath = FilePath.Substring(1, FilePath.Length - 1);
            string[] strPath = FilePath.Split('|');
            for (int m = 0; m < strPath.Length; m++)
            {
                if (FilePath1 != "")
                {
                    if (strPath[m] != "" && strPath[m] != "null")
                    {
                        FilePath1 = FilePath1 + "|" + strPath[m];
                    }
                }
                else
                {
                    if (strPath[m] != "" && strPath[m] != "null")
                    {
                        FilePath1 = "|" + strPath[m];
                    }
                }
            }

            FilePath1 = FilePath1.Substring(1, FilePath1.Length - 1);
            string[] strFiles = FilePath1.Split('|');
            PDFMerge PDF = new PDFMerge();
            //Response.Write("FilePath:" + FilePath + "<br/><br/>");
            PDF.MergeFiles(DestinationFileName, strFiles);
            string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ConsolidatedPDFFileName + ".pdf");

            File.Copy(DestinationFileName, strDirectory, true);
            //Response.Write("Copied" + "<br/><br/>");
            //Response.Write(DestinationFileName);
            return ConsolidatedPDFFileName + ".pdf";
        }
        catch (Exception ex)
        {
            // Response.Write(lblError.Text = ex.ToString());
            return "";
        }
    }
}