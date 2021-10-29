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
//using CrmSdk;
using System.IO;
using System.Linq;
using System.Data.Common;
using Spire.Xls;
using System.Drawing;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Collections.Generic;
using GemBox.Document;
using GemBox.Document.Tables;
using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Xls;



public partial class BatchReportTEST : System.Web.UI.Page
{
    public StreamWriter sw = null;
    string strDescription = string.Empty;
    bool bProceed = true;
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    public int numIndexPageCount = 1;  //Index page count -- if count of batch records is > 22 then it will come on next page 
    public int numIndexPageSize = 20; // Size of index page 
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    String fsReportingName = "";


    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName, lsGAorTIAHeader;
    protected void Page_Load(object sender, EventArgs e)
    {
        bool isAdvancedMode = false;
        if (!IsPostBack)
        {
            Session.Abandon();
            //  Session["CurPageInBatch"] 
            //   Session["BatchDic"] = "";
            //string strUserName = HttpContext.Current.User.Identity.Name.ToString();
            //Response.Write("UserName: "+ strUserName);


            // to find windows user 
            // System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
            //string strName = p.Identity.Name;
            //Response.Write("<br/>p.Identity.Name:" + strName);
            //strName = HttpContext.Current.User.Identity.Name.ToString();
            //Response.Write("<br/>HttpContext.Current.User.Identity.Name:" + strName);
            //strName = Request.ServerVariables["AUTH_USER"]; //Finding with name
            //Response.Write("<br/>AUTH_USER:" + strName);
            ////////
            // string value = string.Empty;


            /***********Added 12/19/2016 NightlyLoad***********************/
            isAdvancedMode = (Request.QueryString["id"] ?? String.Empty).Equals("Scheduled");

            /***********Added 12/19/2016 NightlyLoad***********************/

            fillHousehold();
            DataTable dtBatch = GetBatchList("0", "0");
            gvList.Columns[0].Visible = true;
            gvList.Columns[5].Visible = true;
            gvList.Columns[6].Visible = true;
            gvList.Columns[7].Visible = true;
            gvList.DataSource = dtBatch;
            gvList.DataBind();
            gvList.Columns[0].Visible = false;
            gvList.Columns[5].Visible = false;
            gvList.Columns[6].Visible = false;
            gvList.Columns[7].Visible = false;
            btnGenerateReport.Visible = true;


            /***********Added 12/19/2016 NightlyLoad***********************/
            if (isAdvancedMode)
            {
                try
                {

                    btnGenerateReport.Visible = false;// TESTING 
                    btnGenerateReport_Click(btnGenerateReport, new EventArgs());


                }
                catch (Exception ex)
                {
                    sw.WriteLine("Error" + ex.Message + DateTime.Now);
                }

            }
            /***********Added 12/19/2016 NightlyLoad***********************/
        }
    }
    public void DeleteImageFiles()
    {
        string DeletePath = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\";
        int count = 0; 
        string FileNAme = string.Empty;
        //string[] files = Directory.GetFiles(DeletePath);
        try
        {
            sw.WriteLine("Delete started in directory = "  + DeletePath + " " + DateTime.Now);
            DirectoryInfo di = new DirectoryInfo(DeletePath);
            //int count = 0;
            foreach (FileInfo fi in di.GetFiles())
            {
                 FileNAme = fi.FullName.ToString();

                DateTime Current = DateTime.Today.Date;
                DateTime FileCreated = fi.CreationTime.Date;

                int result = DateTime.Compare(FileCreated, Current);
                if (result == 0)
                {
                    fi.Delete();
                    count++;
                }
            }

            sw.WriteLine("Delete Done, No. of files ="+count + " " +  DateTime.Now);
        }
        catch (Exception ex)
        {
            sw.WriteLine("Error Deleting File" +FileNAme+ " ERROR :" + ex.Message.ToString()+ " " + DateTime.Now);
        }
    }
    public void fillHousehold()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("sp_s_Get_HouseHoldName");
        ddlHouseHold.Items.Clear();
        ddlHouseHold.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHouseHold.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }

    protected void btnGenerateReport_Click(object sender, EventArgs e)
    {
        bool Nodata = false;
        try
        {

            lblError.Text = "";
            Session.Remove("CurPageInBatch");
            string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl);// "http://Crm01/";
            //string crmServerURL = "http://server:5555/";

            string orgName = "GreshamPartners";
            string currentuser = null;
            //string orgName = "Webdev";
           // CrmService service = null;
            Boolean checkrunreport = false;
            String DestinationPath = string.Empty;
            string ConsolidatePdfFileName = string.Empty;
            string ReportOpFolder = string.Empty;
            clsCombinedReports objCombinedReports = new clsCombinedReports();

            /***********Added 12/19/2016 NightlyLoad***********************/
            string LogFileName = string.Empty;
            LogFileName = "Log-" + DateTime.Now + ".txt";
            //string FolderName = "Logs";
            LogFileName = LogFileName.Replace(":", "-");
            LogFileName = LogFileName.Replace("/", "-");
            sw = new StreamWriter(Server.MapPath("") + @"\Logs" + "/" + LogFileName);
            /***********Added 12/19/2016 NightlyLoad***********************/



            try
            {
                //service = GetCrmService(crmServerUrl, orgName);
                //WhoAmIRequest userRequest = new WhoAmIRequest();
                // Execute the request.
                //WhoAmIResponse user = (WhoAmIResponse)service.Execute(userRequest);
                //currentuser = user.UserId.ToString();
            }
            catch (System.Web.Services.Protocols.SoapException exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Detail.InnerText;
                // Response.Write(strDescription);
                lblError.Text = strDescription;
            }
            catch (Exception exc)
            {
                bProceed = false;
                strDescription = "Crm Service failed to start, Error Detail: " + exc.Message;
                //Response.Write(strDescription);
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

            string strUserName = HttpContext.Current.User.Identity.Name.ToString();

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


            ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.BatchReports) + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                ReportOpFolder = @"C:\Reports\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            }
            else
                ReportOpFolder = AppLogic.GetParam(AppLogic.ConfigParam.BatchReports) + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            /********* Added 12/19/2016 NightlyLOad***********/
            bool isAdvancedMode = (Request.QueryString["id"] ?? String.Empty).Equals("Scheduled");
            if (isAdvancedMode)
            {
               // ReportOpFolder = @"\\GRPAO1-VWFS01\Shared$\BATCH REPORTS\Daily Report Dump\";
                ReportOpFolder = @"\\GRPAO1-VWFS01\Shared$\BATCH REPORTS\Daily_Report_Dump_PROD\";
            }
            /********* Added 12/19/2016 NightlyLOad***********/

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
            for (int j = 0; j < gvList.Rows.Count; j++)
            {
                CheckBox chkBox = (CheckBox)gvList.Rows[j].FindControl("chkbSelectBatch");

                if (chkBox.Checked)
                {
                    NoOfBatches++;
                }
            }

            for (int j = 0; j < gvList.Rows.Count; j++)
            {
                
                try
                {
                    CheckBox chkBox = (CheckBox)gvList.Rows[j].FindControl("chkbSelectBatch");

                    if (chkBox.Checked)
                    {
                        /**********Added 12/19/2016 NIGHTLY LOAD*******************/
                        sw.WriteLine("--------------------------------");
                        sw.WriteLine("started for" + " " + Convert.ToString(gvList.Rows[j].Cells[2].Text + DateTime.Now));
                        sw.WriteLine("--------------------------------");
                        /**********Added 12/19/2016 NIGHTLY LOAD*******************/

                        numIndexPageCount = 1;  //Index page count -- if count of batch records is > 22 then it will come on next page 
                        numIndexPageSize = 20; // Size of index page 
                        checkrunreport = true;
                        String BatchIdListTxt = Convert.ToString(gvList.Rows[j].Cells[0].Text);
                        dtBatch = GetDataTable(BatchIdListTxt, "");

                        //String TempName =  gvList.Rows[j].Cells[6].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();

                        //String HHName = gvList.Rows[j].Cells[6].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();

                        String HHName = "";

                        //string TempName = HttpContext.Current.User.Identity.Name.ToString() + "_" + 

                        double total = (double)dtBatch.Rows.Count / numIndexPageSize;
                        int liTotalPage = Convert.ToInt32(Math.Ceiling(total));
                        numIndexPageCount = numIndexPageCount + liTotalPage;

                        sourcefilecount = dtBatch.Rows.Count + (numIndexPageCount + 1);
                        SourceFileArray = new string[sourcefilecount];
                       // Session["Count"] = dtBatch.Rows.Count;

                        /***********Added 12/19/2016 NightlyLoad ( check for No data if else clause added)***********************/
                        if (dtBatch.Rows.Count > 0)
                        {
                           
                            //Nodata = true;
                            for (int i = 0; i < dtBatch.Rows.Count; i++)
                            {
                                
                                if (Convert.ToString(dtBatch.Rows[i]["ssi_spvfilename"]) != "")
                                {
                                    HHName = Convert.ToString(dtBatch.Rows[i]["ssi_spvfilename"]);
                                    HHName = HHName.Replace("/", "");
                                }
                                else
                                {
                                    HHName = gvList.Rows[j].Cells[7].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString().Replace("&#39;", "'").ToString();
                                    HHName = HHName.Replace("/", "");
                                }



                                ContactFolderName = gvList.Rows[j].Cells[5].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString().Replace("&#39;", "'").ToString();
                                //ContactFolderName = Convert.ToString(dtBatch.Rows[i]["Ssi_ContactIdName"]).Replace(",", "");
                                bool isExist = System.IO.Directory.Exists(ReportOpFolder + "\\" + ContactFolderName);

                                if (!isExist)
                                {
                                    //  Response.Write("Folder: " + ReportOpFolder + "\\" + ContactFolderName);
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
                                if (chkSuppressManagerDetail.Checked)
                                    fsVersion = "No";

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

                                if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["HouseHoldReportTitle"])))
                                    lsFinalTitleAfterChange = Convert.ToString(dtBatch.Rows[i]["HouseHoldReportTitle"]);

                                if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["AllocationGroupReportTitle"])))
                                    lsFinalTitleAfterChange = Convert.ToString(dtBatch.Rows[i]["AllocationGroupReportTitle"]);

                                String fsFooterTxt = String.Empty;
                                if (!String.IsNullOrEmpty(Convert.ToString(dtBatch.Rows[i]["GreshamFooterTxt"])))
                                    fsFooterTxt = Convert.ToString(dtBatch.Rows[i]["GreshamFooterTxt"]);


                                // Added By Rohit for Direct Manager Report

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

                                    if (fsGreshReportIdName != "Asset Distribution" && fsGreshReportIdName != "Asset Distribution Comparison")
                                    {
                                        CombinedFileName = generateCombinedPDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath.Replace(".xls", ".pdf"), fsFooterTxt, fsGreshReportIdName, LegalEntity, FundID, CommitmentReportHeader, fsGAorTIAflag, fsReportRollupGroupIdName, fsHHreportparametersId);
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
                                        SetValuesToVariable(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg);
                                        // generatesExcelsheets(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg);
                                        generatePDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg);
                                        CombinedFileName = true;
                                    }

                                    loCoversheetCheck = new FileInfo(lsCoversheet);
                                    if (!loCoversheetCheck.Exists)
                                    {
                                        generateCoversheetPDF(fsAsofDate, lsCoversheet, fsAllocationGroup, fsHouseholdName, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange, fsCoverSheetPageTitle, fsGAorTIAflag, fsDiscretionaryFlg);
                                        generatesCoverExcel(fsAsofDate, fsHouseholdName, fsAllocationGroup, lsCoversheet, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange, fsCoverSheetPageTitle);
                                    }
                                }
                                else
                                {
                                    CombinedFileName = true;
                                }

                                /* Array fill with the PATH + Fullname of PDF*/

                                if (i == 0)
                                {
                                    SourceFileArray[i] = lsCoversheet.Replace(".xls", ".pdf");
                                    for (int PageCnt = 1; PageCnt < numIndexPageCount; PageCnt++)
                                    {
                                        SourceFileArray[i + PageCnt] = (Server.MapPath("") + @"\ExcelTemplate\Blank.pdf");
                                    }
                                    if (CombinedFileName == true)
                                        SourceFileArray[i + (numIndexPageCount)] = lsExcleSavePath.Replace(".xls", ".pdf");
                                }
                                else
                                {
                                    if (CombinedFileName == true)
                                        SourceFileArray[i + (numIndexPageCount)] = lsExcleSavePath.Replace(".xls", ".pdf");

                                }



                                /* Array fill with the PATH + Fullname of PDF*/

                                #region Region to update USER currently not used
                                //  code to update updatedate in batch ety of crm
                                //ssi_batch objBatch = new ssi_batch();
                                //objBatch.ssi_batchid = new Key();

                                //objBatch.ssi_batchid.Value = new Guid(dtBatch.Rows[i]["ssi_batchid"].ToString());

                                //objBatch.ssi_updatedate = new CrmDateTime();
                                //objBatch.ssi_updatedate.Value = DateTime.Now.ToString();

                                //objBatch.ssi_updateuserid = new Lookup();
                                //objBatch.ssi_updateuserid.type = EntityName.systemuser.ToString();
                                //objBatch.ssi_updateuserid.Value = new Guid(currentuser);

                                //service.Update(objBatch);
                                //  Response.Write("<br>Batch ID" + objBatch.ssi_batchid.Value);
                                // Response.Write("<br>Current User" + currentuser);  
                                #endregion
                            }

                            // Consolidate File Logic ORIGINAL
                            //File.Copy(ReportOpFolder + " " + TempName + "\\" + ContactFolderName + "\\Coversheet.pdf", ReportOpFolder + " " + TempName + "\\" + ContactFolderName + "\\" + "ConsolidatedReport.pdf");
                            //String DestinationPath = ReportOpFolder + " " + TempName + "\\" + ContactFolderName + "\\" + "ConsolidatedReport.pdf";

                            // Consolidate File Logic NEW
                            DateTime dtAsOfDate = Convert.ToDateTime(ViewState["AsOfDate"]);

                            strYear = dtAsOfDate.Year.ToString().Length < 2 ? "0" + dtAsOfDate.Year.ToString() : dtAsOfDate.Year.ToString();
                            strMonth = dtAsOfDate.Month.ToString().Length < 2 ? "0" + dtAsOfDate.Month.ToString() : dtAsOfDate.Month.ToString();
                            strDay = dtAsOfDate.Day.ToString().Length < 2 ? "0" + dtAsOfDate.Day.ToString() : dtAsOfDate.Day.ToString();

                            ConsolidatePdfFileName = HHName + "_" + strYear + "-" + strMonth + strDay + ".pdf";

                            ConsolidatePdfFileName = GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);

                            if (!File.Exists(ReportOpFolder + "\\" + ConsolidatePdfFileName))
                                File.Copy(ReportOpFolder + "\\" + ContactFolderName + "\\Coversheet.pdf", ReportOpFolder + "\\" + ConsolidatePdfFileName);

                            DestinationPath = ReportOpFolder + "\\" + GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);


                            //string pathn = @"C:\Reports\tttttt.pdf";
                            if (ContactFolderName.Contains("MTGBK")) //generate without coversheet
                            {
                               
                                string[] target = new string[sourcefilecount - (numIndexPageCount)];
                                Array.Copy(SourceFileArray, (numIndexPageCount), target, 0, sourcefilecount - (numIndexPageCount));
                                PDF.MergeFiles(DestinationPath, target);

                                /***********Added 12/19/2016 NightlyLoad***********************/
                                Nodata = true;// checking
                                /***********Added 12/19/2016 NightlyLoad***********************/
                            }
                            else //generate with coversheet
                            {
                              
                                PDF.MergeFiles(DestinationPath, SourceFileArray);
                                //System.Threading.Thread.Sleep(15000);
                                string DestinationPath1 = objCombinedReports.addPageIndex(DestinationPath, dtBatch);

                                //string strCoverLetterPath = getCoverLetter(BatchIdListTxt, "1");
                                //if (strCoverLetterPath != "")
                                //{
                                //    string[] DestiFiles = new string[2];
                                //    DestiFiles[0] = strCoverLetterPath;
                                //    DestiFiles[1] = DestinationPath1;
                                //    PDF.MergeFiles(pathn, DestiFiles);
                                //}
                                //else
                                //{

                                File.Copy(DestinationPath1, DestinationPath, true);

                                /***********Added 12/19/2016 NightlyLoad***********************/
                                Nodata = true;// checking
                                /***********Added 12/19/2016 NightlyLoad***********************/
                                //}
                            }

                            //Response.Write(Convert.ToString(Session["CurPageInBatch"]));
                            //Dictionary<string, int> dicNumFilesCount = (Dictionary<string, int>)Session["BatchDic"];

                            //foreach (KeyValuePair<string, int> pair in dicNumFilesCount)
                            //{
                            //    Response.Write(pair.Key.ToString() + " : " + pair.Value.ToString() + "<br/>");
                            //}                   


                            //Session.Remove("BatchDic");
                            Session.Remove("CurPageInBatch");
                            Directory.Delete(ReportOpFolder + "\\" + ContactFolderName, true);
                            sw.WriteLine("completed" + " " + Convert.ToString(gvList.Rows[j].Cells[2].Text) + DateTime.Now);
                        }
                        else
                        {
                            /***********Added 12/19/2016 NightlyLoad***********************/
                            Session.Remove("CurPageInBatch");
                            checkrunreport = false;
                            //sw.WriteLine("No DAta Found" + " " + Convert.ToString(gvList.Rows[j].Cells[2].Text) + DateTime.Now);
                            string Family = Convert.ToString(Session["FamilyName"]);
                            sw.WriteLine("No DAta Found" + " " + Family);
                            Session.Remove("FamilyName");
                            lblError.Text = "No Data Found";
                            /***********Added 12/19/2016 NightlyLoad***********************/

                        }
                    }
                    
                }
                /**********Added 12/19/2016 NIGHTLY LOAD*******************/
                catch (Exception ex)
                {
                    Nodata = false;
                    sw.WriteLine("------------------------------------------------------------------");
                    sw.WriteLine(ex.ToString() + Convert.ToString(gvList.Rows[j].Cells[2].Text));
                    sw.WriteLine("------------------------------------------------------------------");
                    Session.Remove("CurPageInBatch");
                }
                /**********Added 12/19/2016 NIGHTLY LOAD*******************/
            }

            ////////////////////////////////////
            
                if (NoOfBatches == 1)
                {
                    if (Nodata)
                    {
                        string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + ConsolidatePdfFileName);

                        File.Copy(DestinationPath, strDirectory, true);
                        Directory.Delete(ReportOpFolder, true);

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
            
            ////////////////////////////////////

            if (checkrunreport)
            {
                sw.WriteLine("---------------------------------------"+Nodata);
                lblError.Text = "Reports generated successfully";
                ClearCheckBoxSelection();
            }
            else if (!Nodata)
            {
                sw.WriteLine("---------------------------------------" + Nodata);
                lblError.Text = "No Data Found.";
            }
            else
            {
                sw.WriteLine("---------------------------------------" + Nodata);
                lblError.Text = "Please Select a batch to run report.";
            }
            Nodata = false;
        }
        catch (Exception ex)
        {
            Nodata = false;
            lblError.Text = "Error Generating Report " + ex.ToString();
        }
        finally
        {
            DeleteImageFiles();
            sw.Close();
        }
    }

    private void ClearCheckBoxSelection()
    {
        //Loop through all the rows in gridview
        foreach (GridViewRow gvrow in gvList.Rows)
        {
            //Finiding checkbox control in gridview for particular row
            CheckBox chkbSelectBatch = (CheckBox)gvrow.FindControl("chkbSelectBatch");
            //Condition to check checkbox selected or not
            if (chkbSelectBatch.Checked)
            {
                chkbSelectBatch.Checked = false;
            }
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

    private string getCoverLetter(string BatchIdListTxt, string CoverletterFlg)
    {
        try
        {
            DataTable dt = GetDataTable(BatchIdListTxt, "1");

            if (dt.Rows.Count > 0)
            {
                string strTemplateFilePath = Convert.ToString(dt.Rows[0]["ssi_TemplateFilePath"]);
                if (strTemplateFilePath != "")
                {
                    //FOR -- TESTING 
                    if (Request.Url.AbsoluteUri.Contains("localhost"))
                        strTemplateFilePath = @"C:\Reports\Commentaries.pdf";

                    return strTemplateFilePath;
                }
                else
                    return "";
            }
            else
                return "";

        }
        catch (Exception ex)
        {
            return "";
        }


    }
    public DateTime NearestQuarterEnd(DateTime date)
    {
        IEnumerable<DateTime> candidates =//Enumarateor used To show candiadtes by LastDate
            QuartersInYear(date.Year).Union(QuartersInYear(date.Year - 1));
        return candidates.Where(d => d < date.Date).OrderBy(d => d).Last();
    }

    IEnumerable<DateTime> QuartersInYear(int year)
    {
        return new List<DateTime>() {
        new DateTime(year, 3, 31),
        new DateTime(year, 6, 30),
        new DateTime(year, 9, 30),
        new DateTime(year, 12, 31),
    };
    }

    private DataTable GetDataTable(String BatchIdListTxt, string CoverletterFlg)
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
            object PriorDate = txtPriorDate.Text == "" ? "null" : "'" + txtPriorDate.Text + "'";
            //object EndDate = txtEndDate.Text == "" ? "null" : "'" + txtEndDate.Text + "'"; //12/19/2016
            object Coverletter = CoverletterFlg == "" ? "null" : "'1'";

            /********* Added 12/19/2016 NightlyLOad***********/
            object EndDate = null;
            bool isAdvancedMode = (Request.QueryString["id"] ?? String.Empty).Equals("Scheduled");
            if (isAdvancedMode)
            {
                DateTime dtAsOfDate = DateTime.Now;
                DateTime lastDay = new DateTime(dtAsOfDate.Year, dtAsOfDate.Month, 1); //1st Day of Current Month
                lastDay = lastDay.AddDays(-1);  //last date of previous month

                DateTime date = DateTime.Now;
                DateTime quarterEnd = NearestQuarterEnd(date);
                txtEndDate.Text = quarterEnd.ToShortDateString();
                EndDate = "'" + txtEndDate.Text + "'";

                //object EndDate = txtEndDate.Text == "" ? "null" : "'" + txtEndDate.Text + "'";
            }
            else
            {
                EndDate = txtEndDate.Text == "" ? "null" : "'" + txtEndDate.Text + "'";
            }
            /********* Added 12/19/2016 NightlyLOad***********/


            //object NoComparison = chkNoComparison.Checked == false ? 0 : 1;
            greshamquery = "SP_S_BATCH @BatchIdListTxt='" + BatchIdListTxt + "',@PriorDT=" + PriorDate + ",@EndDT=" + EndDate + ",@NoComparisonLineFlg=" + Convert.ToBoolean(chkNoComparison.Checked);
            //greshamquery = "SP_S_BATCH @BatchIdListTxt='" + BatchIdListTxt + "',@PriorDT=" + PriorDate + ",@EndDT=" + EndDate + ",@NoComparisonLineFlg=" + Convert.ToBoolean(chkNoComparison.Checked) + ",@CoverletterFlg= " + Coverletter + "";

                sw.WriteLine("Query" + greshamquery );

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

    private DataTable GetBatchList(string HouseholdID, string BatchType)
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
            HouseholdID = HouseholdID == "0" ? "null" : "'" + HouseholdID + "'";
            BatchType = BatchType == "0" ? "null" : "'" + BatchType + "'";
            //greshamquery = "sp_s_batch_list @HouseHoldId =" + HouseholdID + ",@BatchType=" + BatchType;

            greshamquery = "SP_S_BATCH_LIST_CONSOLIDETED_NIGHTLY @HouseHoldId =" + HouseholdID + ",@BatchType=" + BatchType;
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


    /// <summary>
    /// Set up the CRM Service.
    /// </summary>
    /// <param name="organizationName">My Organization</param>
    /// <returns>CrmService configured with AD Authentication</returns>
  //  public static CrmService GetCrmService(string crmServerUrl, string organizationName)
  //  {
  //      // Get the CRM Users appointments
  //      // Setup the Authentication Token
  //      CrmAuthenticationToken token = new CrmAuthenticationToken();
  //      token.AuthenticationType = 0; // Use Active Directory authentication.
  //      token.OrganizationName = organizationName;
  //      string username = WindowsIdentity.GetCurrent().Name;

  //      if (username == "CORP\\gbhagia")
  //     {
    //        // Use the global user ID of the system user that is to be impersonated.
   //         token.CallerId = new Guid("EE8E3A77-59E2-DD11-831F-001D09665E8F");//deb
   //         //token.CallerId = new Guid("C42C7E05-8303-DE11-A38C-001D09665E8F");//gary                
   //     }
   //     token.CallerId = new Guid("EE8E3A77-59E2-DD11-831F-001D09665E8F");//deb
   //     CrmService service = new CrmService();

   //     if (crmServerUrl != null &&
   //         crmServerUrl.Length > 0)
    //    {
     //       UriBuilder builder = new UriBuilder(crmServerUrl);
   //         builder.Path = "//MSCRMServices//2007//CrmService.asmx";
  //          service.Url = builder.Uri.ToString();
   //     }

   //     service.CrmAuthenticationTokenValue = token;
    //    service.Credentials = System.Net.CredentialCache.DefaultCredentials;

    //    return service;
  //  }

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

        if (chkNoComparison.Checked)
            lsSQL = lsSQL + ",@ComparisonFlg = 1";

        //  Response.Write("<br><br><br>" + lsSQL + "<br><br><br>");
        return lsSQL;
    }

    public void generatesExcelsheets(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFinalReportTitle, String lsFooterTxt, String fsGAorTIAflag, String fsDiscretionaryFlg)
    {
        //  String lsSQL = "SP_R_Adventure_Report @UUID = '" + System.Guid.NewGuid().ToString() + "'";

        String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, fsGAorTIAflag, fsDiscretionaryFlg);

        DB clsDB = new DB();
        DataSet lodataset;
        lodataset = null;
        lodataset = clsDB.getDataSet(lsSQL);

        DataSet loInsertblankRow = lodataset.Copy();
        lodataset.Tables[0].Clear();
        lodataset.Clear();
        lodataset = null;
        lodataset = loInsertblankRow.Clone();
        int liBlankCounter = 1;
        for (int liBlankRow = 0; liBlankRow < loInsertblankRow.Tables[0].Rows.Count; liBlankRow++)
        {
            if (liBlankRow != 0 && (loInsertblankRow.Tables[0].Rows[liBlankRow]["_Ssi_BoldFlg"].ToString() == "True" || loInsertblankRow.Tables[0].Rows[liBlankRow]["_Ssi_SuperBoldFlg"].ToString() == "True"))
            {
                String lsdsd = loInsertblankRow.Tables[0].Rows[liBlankRow][0].ToString();
                if (!lsdsd.Contains("NET CHANGE %"))
                {

                    //if ((!String.IsNullOrEmpty(fsSPriorDate) || !String.IsNullOrEmpty(fsAllocationGroup)))
                    //{

                    DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                    newCustomersRow[0] = "test";
                    lodataset.Tables[0].Rows.Add(newCustomersRow);
                    liBlankCounter = liBlankCounter + 1;
                    // }
                    //else if (fsAllignment != "Horizontal")
                    //{
                    //    DataRow newCustomersRow = lodataset.Tables[0].NewRow();
                    //    newCustomersRow[0] = "test";
                    //    lodataset.Tables[0].Rows.Add(newCustomersRow);
                    //    liBlankCounter = liBlankCounter + 1;
                    //}
                }
            }
            lodataset.Tables[0].ImportRow(loInsertblankRow.Tables[0].Rows[liBlankRow]);
        }
        lodataset.AcceptChanges();
        DataSet loInsertdataset = lodataset.Copy();
        int liTtrow = 0;
        for (int liNewdataset = 0; liNewdataset < lodataset.Tables[0].Columns.Count; liNewdataset++)
        {
            if (!lodataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") && !lodataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
            {
                liTtrow = liTtrow + 1;
            }

        }
        for (int liNewdataset = lodataset.Tables[0].Columns.Count - 1; liNewdataset > -1; liNewdataset--)
        {

            if (lodataset.Tables[0].Columns[liNewdataset].ColumnName.Contains("_") || lodataset.Tables[0].Columns[liNewdataset].ColumnName.Trim().Equals("1"))
            {
                loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[liNewdataset]);
            }

        }
        loInsertdataset.AcceptChanges();
        // loInsertdataset.Tables[0].Columns.Remove(loInsertdataset.Tables[0].Columns[1]);
        // loInsertdataset.AcceptChanges();
        String lsFileNamforFinalXls = System.DateTime.Now.ToString("MMddyyhhmmss") + ".xls";
        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\Book_" + liTtrow + ".xls");
        string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls);
        string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\" + lsFileNamforFinalXls.Replace("xls", "xml"));
        // Response.Write(strDirectory);
        string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strDirectory + ";Extended Properties=\"Excel 8.0;HDR=YES;\"";
        DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");



        FileInfo loFile = new FileInfo(strDirectory1);
        loFile.CopyTo(strDirectory, true);





        String lsFirstColumn = "Insert into [Sheet1$] (";
        String lsFiled = "";
        String lsFieldvalue = "";
        for (int liColumns = 0; liColumns < loInsertdataset.Tables[0].Columns.Count; liColumns++)
        {

            lsFieldvalue += "'" + loInsertdataset.Tables[0].Columns[liColumns].ColumnName.Replace("'", "''") + "'";
            lsFiled += "id" + (liColumns + 1);
            if (liColumns < loInsertdataset.Tables[0].Columns.Count - 1)
            {
                lsFieldvalue = lsFieldvalue + ",";
                lsFiled = lsFiled + ",";
            }

        }
        lsFirstColumn = lsFirstColumn + lsFiled + ")" + " Values (" + lsFieldvalue + ")";



        using (DbConnection connection = factory.CreateConnection())
        {
            connection.ConnectionString = connectionString;

            using (DbCommand command = connection.CreateCommand())
            {
                try
                {
                    command.CommandText = lsFirstColumn;

                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
                catch
                {
                    Response.Write(lsFirstColumn);
                }
            }
        }

        for (int liCounter = 0; liCounter < loInsertdataset.Tables[0].Rows.Count; liCounter++)
        {

            lsFirstColumn = "Insert into [Sheet1$] (";

            lsFieldvalue = "";
            for (int liColumns = 0; liColumns < loInsertdataset.Tables[0].Columns.Count; liColumns++)
            {
                lsFieldvalue += "'" + loInsertdataset.Tables[0].Rows[liCounter][liColumns].ToString().Replace("'", "''") + "'";
                if (liColumns < loInsertdataset.Tables[0].Columns.Count - 1)
                {
                    lsFieldvalue = lsFieldvalue + ",";
                }
            }
            lsFirstColumn = lsFirstColumn + lsFiled + ")" + " Values (" + lsFieldvalue + ")";
            using (DbConnection connection = factory.CreateConnection())
            {
                connection.ConnectionString = connectionString;

                using (DbCommand command = connection.CreateCommand())
                {
                    //if (liCounter == 0 || liCounter == 2)
                    //{
                    //    connection.Open();
                    //    command.CommandText = "INSERT INTO [Sheet1$] (id1) VALUES('')";
                    //    command.ExecuteNonQuery();
                    //    connection.Close();
                    //}
                    try
                    {
                        command.CommandText = lsFirstColumn;
                        //  Response.Write(lsFirstColumn);
                        connection.Open();
                        command.ExecuteNonQuery();
                        connection.Close();
                    }
                    catch
                    {
                        Response.Write(lsFirstColumn);
                        Response.End();
                    }
                }
            }
        }

        /*---------NEW CODE FOR FOOTER IN EXCEL FILE-------------*/
        if (!String.IsNullOrEmpty(lsFooterTxt))
        {
            String lsFooterRow = "Insert into [Sheet1$] (id1) Values ('" + lsFooterTxt + "')";
            using (DbConnection connection = factory.CreateConnection())
            {
                connection.ConnectionString = connectionString;

                using (DbCommand command = connection.CreateCommand())
                {
                    try
                    {
                        command.CommandText = lsFooterRow;
                        //  Response.Write(lsFirstColumn);
                        connection.Open();
                        command.ExecuteNonQuery();
                        connection.Close();
                    }
                    catch
                    {
                        Response.Write(lsFooterRow);
                        Response.End();
                    }
                }
            }
        }

        /*---------END OF NEW CODE FOR FOOTER IN EXCEL FILE------*/

        #region StyleUsing Spire.xls
        Workbook workbook = new Workbook();
        workbook.LoadFromFile(strDirectory);

        //Gets first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        // Worksheet sheetCover = workbook.Worksheets[0];

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

        sheet.Range["A2"].Text = lsfamilyName;
        sheet.Range["A4"].Text = Convert.ToDateTime(fsAsofDate).ToString("MMMM dd, yyyy") + "";
        sheet.Range["A2"].VerticalAlignment = VerticalAlignType.Center;
        if (fsAllignment != "Horizontal")
            sheet.Range["A3"].Text = "ASSET DISTRIBUTION COMPARISON";
        sheet.Range["A3"].VerticalAlignment = VerticalAlignType.Center;
        sheet.Range["A4"].VerticalAlignment = VerticalAlignType.Center;

        //Set for Pdf
        lsDistributionName = sheet.Range["A3"].Text;
        lsFamiliesName = lsfamilyName;
        lsDateName = sheet.Range["A4"].Text;
        // sheetCover.Range["A21"].Text = lsfamilyName;
        // sheetCover.Range["A23"].Text = Convert.ToDateTime(fsAsofDate).ToString("MMMM dd, yyyy") + "";
        // sheetCover.Range[1, 1, 500, 1].ColumnWidth = 23.1;
        // sheetCover.Range["A21"].RowHeight = 37;
        sheet.Range["A2"].VerticalAlignment = VerticalAlignType.Center;
        sheet.GridLinesVisible = false;
        for (int liRemoveheader = 1; liRemoveheader < 23; liRemoveheader++)
        {
            sheet.Range[1, liRemoveheader].Text = "";
        }

        for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
        {
            int lisrc = liCounter + 7;


            for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
            {
                if (liColumns != 1 && liColumns != loInsertdataset.Tables[0].Columns.Count && !String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text))
                {
                    try
                    {
                        if (!sheet.Range[lisrc, liColumns].Text.Contains("E"))
                            sheet.Range[lisrc, liColumns].Text = Convert.ToString(Math.Round(Convert.ToDecimal(sheet.Range[lisrc, liColumns].Text), 2));
                        else
                        {
                            sheet.Range[lisrc, liColumns].Text = Convert.ToString(Math.Round(Convert.ToDecimal(Convert.ToDouble(sheet.Range[lisrc, liColumns].Text))));
                        }
                        // sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(sheet.Range[lisrc, liColumns].Text);
                        //   sheet.Range[lisrc, liColumns].NumberFormat = "_(* #,##0_);_(* \\(#,##0\\);_(* &quot;-&quot;??_);_(@_)";

                    }
                    catch
                    {
                        Response.Write(sheet.Range[lisrc, liColumns].Text);
                    }
                }
                //Header Setting           
                if (liCounter == 0)
                {
                    sheet.Range[6, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                    //28/02/2011
                    //sheet.Range[6, liColumns].Style.Font.Size = 9;
                    sheet.Range[6, liColumns].Style.Font.Size = 7;
                    sheet.Range[6, liColumns].RowHeight = 12;
                    sheet.Range[6, liColumns].VerticalAlignment = VerticalAlignType.Center;

                    sheet.Range[6, liColumns].Style.Font.IsBold = true;

                    sheet.Range[6, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Right;



                }

                sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 255);
                sheet.Range[lisrc, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                //28/02/2011
                //sheet.Range[lisrc, liColumns].Style.Font.Size = 8;
                sheet.Range[lisrc, liColumns].Style.Font.Size = 7;
                sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
                //sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.FromArgb(216, 216, 216);
                sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.FromArgb(216, 216, 216);

                if (liColumns != 1)
                    sheet.Range[lisrc, liColumns].Style.HorizontalAlignment = HorizontalAlignType.Right;
                sheet.Range[lisrc, liColumns].VerticalAlignment = VerticalAlignType.Center;


            }
            if (lodataset.Tables[0].Rows[liCounter]["_Ssi_BoldFlg"].ToString() == "True")
            {
                sheet.Range[lisrc, 1].Style.Font.IsBold = true;
                sheet.Range[lisrc - 1, 1].Text = " ";

                for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {
                    sheet.Range[lisrc - 1, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;
                    sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(216, 216, 216);
                    sheet.Range[lisrc, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                    //sheet.Range[lisrc, liColumns].Style.Font.Size = 9;
                    sheet.Range[lisrc, liColumns].Style.Font.Size = 8;
                    sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;

                }
            }
            if (lodataset.Tables[0].Rows[liCounter]["_Ssi_UnderlineFlg"].ToString() != "True" && lodataset.Tables[0].Rows[liCounter]["_Ssi_SuperBoldFlg"].ToString() != "True")
            {
                if (!String.IsNullOrEmpty(Convert.ToString(lodataset.Tables[0].Rows[liCounter][1])))
                {
                    String abc = "          " + lodataset.Tables[0].Rows[liCounter][1].ToString();
                    sheet.Range[lisrc, 1].Text = abc;
                }
            }
            if (lodataset.Tables[0].Rows[liCounter]["_Ssi_UnderlineFlg"].ToString() == "True")
            {
                for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {
                    String abc = "          " + "          " + lodataset.Tables[0].Rows[liCounter][0].ToString();
                    sheet.Range[lisrc, 1].Text = abc;
                    sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;
                    sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
                    sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeTop].Color = System.Drawing.Color.FromArgb(0, 0, 0);
                }

            }
            if (lodataset.Tables[0].Rows[liCounter]["_Ssi_SuperBoldFlg"].ToString() == "True")
            {
                for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {
                    //ExcelColors backKnownColor1 = (ExcelColors)(49);
                    //  sheet.Range[lisrc, liColumns].Style.Interior.FillPattern = ExcelPatternType.Gradient;
                    // sheet.Range[lisrc, liColumns].Style.Interior.Gradient.BackKnownColor = backKnownColor1;
                    // sheet.Range[lisrc, liColumns].Style.Interior.Gradient.ForeKnownColor = backKnownColor1;
                    //sheet.Range[lisrc, liColumns].Style.Interior.Gradient.GradientStyle = GradientStyleType.Vertical;
                    //  sheet.Range[lisrc, liColumns].Style.Interior.Gradient.GradientVariant = GradientVariantsType.ShadingVariants4; 
                    sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(51, 204, 204);
                    sheet.Range[lisrc, liColumns].Style.Font.FontName = "Frutiger 55 Roman";
                    if (liColumns == 1)
                    {
                        //sheet.Range[lisrc, liColumns].Style.Font.Size = 9;
                        sheet.Range[lisrc, liColumns].Style.Font.Size = 8;
                    }
                    else
                    {
                        //sheet.Range[lisrc, liColumns].Style.Font.Size = 8;
                        sheet.Range[lisrc, liColumns].Style.Font.Size = 7;
                    }


                    sheet.Range[lisrc, liColumns].Style.Font.IsBold = true;
                    sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;

                    sheet.Range[lisrc - 1, 1].Text = "";

                    sheet.Range[lisrc, liColumns].VerticalAlignment = VerticalAlignType.Center;
                    sheet.Range[lisrc - 1, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;
                }
            }
            if (lodataset.Tables[0].Rows[liCounter]["_Ssi_TabFlg"].ToString() == "True" && lodataset.Tables[0].Rows[liCounter]["_Ssi_UnderlineFlg"].ToString() != "True")
            {

                String abc = "          " + "          " + lodataset.Tables[0].Rows[liCounter][1].ToString();
                sheet.Range[lisrc, 1].Text = abc;



            }
            if (lodataset.Tables[0].Rows[liCounter]["_ssi_greylineflg"].ToString() == "True")
            {
                for (int liColumns = 1; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
                {
                    //sheet.Range[lisrc, liColumns].Style.Font.Color = System.Drawing.Color.FromArgb(165, 165, 165);
                    sheet.Range[lisrc, liColumns].Style.Font.Color = System.Drawing.Color.FromArgb(99, 99, 99);
                }
            }
            for (int liColumns = 2; liColumns <= loInsertdataset.Tables[0].Columns.Count; liColumns++)
            {


                //  Response.Write("<br>String :"+sheet.Range[lisrc, liColumns].Text+" " + " Colums: " +liColumns+ "  "+loInsertdataset.Tables[0].Columns.Count);
                if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && liColumns != loInsertdataset.Tables[0].Columns.Count)
                {


                    if (sheet.Range[lisrc, 1].Text == "NET CHANGE %")
                    {

                        String lsa = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(sheet.Range[lisrc, liColumns].Text));
                        sheet.Range[lisrc, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 255);
                        sheet.Range[lisrc - 1, liColumns].Style.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 255);
                        sheet.Range[lisrc, 1].Style.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 255);
                        sheet.Range[lisrc - 1, 1].Style.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 255);

                        sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Style.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 255);
                        sheet.Range[lisrc - 1, loInsertdataset.Tables[0].Columns.Count].Style.Interior.Color = System.Drawing.Color.FromArgb(255, 255, 255);

                        sheet.Range[lisrc, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;
                        sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;
                        sheet.Range[lisrc, 1].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.None;
                        sheet.Range[lisrc - 1, liColumns].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
                        //sheet.Range[lisrc - 1, liColumns].Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.FromArgb(216, 216, 216);

                        sheet.Range[lisrc - 1, liColumns].Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.FromArgb(216, 216, 216);


                        sheet.Range[lisrc - 1, 1].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
                        //sheet.Range[lisrc - 1, 1].Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.FromArgb(216, 216, 216);
                        sheet.Range[lisrc - 1, 1].Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.FromArgb(216, 216, 216);
                        sheet.Range[lisrc - 1, 1].Text = "NET CHANGE";

                        //28/02/2011

                        sheet.Range[lisrc - 1, liColumns].Style.Font.Size = 7;
                        sheet.Range[lisrc, liColumns].Style.Font.Size = 7;

                        sheet.Range[lisrc - 1, 1].Style.Font.Size = 8;
                        sheet.Range[lisrc, 1].Style.Font.Size = 8;

                        sheet.Range[lisrc - 1, liColumns].Style.Font.IsBold = true;
                        sheet.Range[lisrc, liColumns].Style.Font.IsBold = true;

                        /*end*/


                        if (lsa.Contains(")"))
                        {
                            sheet.Range[lisrc, liColumns].Text = lsa.Replace(")", "%)");
                        }
                        else
                        {
                            sheet.Range[lisrc, liColumns].Text = lsa + "%";
                        }
                    }
                    else
                    {
                        sheet.Range[lisrc, liColumns].Text = String.Format("{0:#,###0;(#,###0)}", Convert.ToDecimal(sheet.Range[lisrc, liColumns].Text));
                    }

                }
                if (liColumns == loInsertdataset.Tables[0].Columns.Count && !String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text))
                {
                    try { sheet.Range[lisrc, liColumns].Text = String.Format("{0:#,###0.0;(#,###0.0)}", Convert.ToDecimal(sheet.Range[lisrc, liColumns].Text)); }
                    catch { }

                }
            }
        }

        sheet.Range[6, 1, 500, 1].ColumnWidth = 35;
        for (int liCounter = 0; liCounter < lodataset.Tables[0].Rows.Count; liCounter++)
        {
            int lisrc = liCounter + 7;
            int liColumnHigeshWidth = 0;
            for (int liColumns = 2; liColumns < loInsertdataset.Tables[0].Columns.Count; liColumns++)
            {
                try
                {
                    liColumnHigeshWidth = 0;

                    liColumnHigeshWidth = sheet.Range[6, liColumns].Text.Length;
                    if (liColumnHigeshWidth < 9)
                        liColumnHigeshWidth = 9;
                    sheet.Range[6, liColumns, 500, liColumns].ColumnWidth = liColumnHigeshWidth;

                    if (sheet.Range[6, liColumns].Text.Contains(" Market Value"))
                    {
                        sheet.Range[6, liColumns].Text = sheet.Range[6, liColumns].Text.Replace(" Market Value", "   Market Value");
                        sheet.Range[6, liColumns].Style.WrapText = true;
                        sheet.Range[6, liColumns, 500, liColumns].ColumnWidth = 12;
                        sheet.Range[6, liColumns].RowHeight = 24;
                    }
                }

                catch { }
                try
                {
                    if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && !sheet.Range[lisrc, liColumns].Text.Contains("%"))
                    {
                        if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                            sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                        sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(sheet.Range[lisrc, liColumns].Text);
                        sheet.Range[lisrc, liColumns].NumberFormat = "#,##0_);\\(#,##0\\)";
                    }
                    if (!String.IsNullOrEmpty(sheet.Range[lisrc, liColumns].Text) && sheet.Range[lisrc, liColumns].Text.Contains("%"))
                    {
                        sheet.Range[lisrc, liColumns].Text = sheet.Range[lisrc, liColumns].Text.Replace("%", "");
                        if (sheet.Range[lisrc, liColumns].Text.Contains("("))
                            sheet.Range[lisrc, liColumns].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, liColumns].Text.Replace("(", "").Replace(")", ""))).ToString();
                        sheet.Range[lisrc, liColumns].NumberValue = Convert.ToDouble(Convert.ToDouble(sheet.Range[lisrc, liColumns].Text) / 100);
                        sheet.Range[lisrc, liColumns].NumberFormat = "#,##0.0%_);\\(#,##0.0%\\)";
                    }
                    if (!String.IsNullOrEmpty(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text))
                    {
                        if (sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text.Contains("("))
                            sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text = Convert.ToDouble((-1) * Convert.ToDouble(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text.Replace("(", "").Replace(")", ""))).ToString();
                        sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberValue = Convert.ToDouble(sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].Text);
                        //sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberFormat = "#,##0.0_);\\(#,##0.0\\)";
                        if (fsAllignment == "Horizontal")
                            sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberFormat = "#,##0.0_);\\(#,##0.0\\)";

                        else

                            // Response.Write("ll");
                            sheet.Range[lisrc, loInsertdataset.Tables[0].Columns.Count].NumberFormat = "#,##0_);\\(#,##0\\)";

                    }
                }
                catch
                {
                    Response.Write("<br>Error: " + lisrc + "  " + liColumns + " " + sheet.Range[lisrc, liColumns].Text);
                }
            }
        }

        if (!String.IsNullOrEmpty(lsFooterTxt))
        {
            int lsCellCount = sheet.Columns[0].CellsCount;
            sheet.Range[lsCellCount, 1].Style.Font.FontName = "Frutiger 55 Roman";
            //sheet.Range[lsCellCount, 1].Style.Font.Size = 9;
            sheet.Range[lsCellCount, 1].Style.Font.Size = 7;
        }

        //Save workbook to disk
        // workbook.Save();
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
                                    {
                                        //lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&C&\"Frutiger 55 Roman,Regular\"&8Page &P of &N&R&\"Frutiger 55 Roman,Italic\"&8&KD8D8D8&D,&T";
                                        lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&C&\"Frutiger 55 Roman,Regular\"&7Page &P of &N&R&\"Frutiger 55 Roman,Italic\"&7&KD8D8D8&D,&T";
                                    }
                                    else
                                    {
                                        //lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&R&\"Frutiger 55 Roman,Italic\"&8&KD8D8D8&D,&T";
                                        lxPagingSetup.ChildNodes[0].Attributes[1].InnerText = "&R&\"Frutiger 55 Roman,Italic\"&7&KD8D8D8&D,&T";
                                    }

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
                                        //lxNodessss.Attributes["ss:Color"].InnerText = "#F2F2F2";
                                        lxNodessss.Attributes["ss:Color"].InnerText = "#D8D8D8";
                                    }
                                }

                            }
                        }

                        foreach (XmlNode lxNodess in lxNodes.ChildNodes)
                        {
                            if (lxNodess.Name == "ss:Font")
                            {

                                if (lxNodess.Attributes["ss:Color"].InnerText == "#808080")
                                {
                                    //lxNodessss.Attributes["ss:Color"].InnerText = "#F2F2F2";
                                    lxNodess.Attributes["ss:Color"].InnerText = "#A5A5A5";
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

        #region delete spire.xls Region


        #endregion
        // Response.Write("<br><br>Final Location " + fsFinalLocation);


    }

    public void setTopWidthBlack(Cell foCell)
    {
        foCell.BorderColor = iTextSharp.text.Color.BLACK;
        foCell.Border = iTextSharp.text.Rectangle.TOP_BORDER;
        foCell.BorderWidth = 0.1F;
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

    public iTextSharp.text.Table addFooter(int liTotalPages, int liCurrentPage, int liLastPageData, Boolean footerflg, String FooterTxt)
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

    public void generatePDF(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFooterTxt, String fsGAorTIAflag, String fsDiscretionaryFlg)
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
                    // document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty));//Commented -- FooterLogic
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
                document.Add(addFooter(liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt));
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

    public bool generateCombinedPDF(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate,
        String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment,
        String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFooterTxt, String ReportName,
        String LegalEntityId, String FundId, String CommitmentReportHeader, String GAorTIAflag, String ReportRollupGroupIdName, String fsHHreportparametersId)
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

    public void generateCoversheetPDF(String lsDateString, String fsFinalLocation, String fsAllocationGroup, String fsHouseholdName, String fsContactId, DataTable foTable, String fsKeyContactID, String fsHouseHoldTitle, String fsContactFullname, String fsDisplayContactName, String lsFinalReportTitle, String lsCoverSheetPageTitle, String fsGAorTIAflag, String fsDiscretionaryFlg)
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

        UpperspaceCount = GetEmptyRowSpace(foTable, UpperspaceCount);
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

        RptTitleCount = GetEmptyRowSpace(foTable, RptTitleCount);
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
        //                    loChunk = new Chunk("GA " + Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]) + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
        //                else
        //                    loChunk = new Chunk("GA " + Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]) + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
        //            }
        //            else
        //            {
        //                if (fsDiscretionaryFlg.ToUpper() == "TRUE")
        //                    loChunk = new Chunk(Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]) + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
        //                else
        //                    loChunk = new Chunk(Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]) + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));//setFontsAll(10, 0, 1));
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

    private int GetEmptyRowSpace(DataTable foTable, int RptTitleCount)
    {

        String lsFinalTitleAfterChange1 = String.Empty;
        for (int j = 0; j < foTable.Rows.Count; j++)
        {
            if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["HouseHoldReportTitle"])))
                lsFinalTitleAfterChange1 = Convert.ToString(foTable.Rows[j]["HouseHoldReportTitle"]);

            if (!String.IsNullOrEmpty(Convert.ToString(foTable.Rows[j]["AllocationGroupReportTitle"])))
                lsFinalTitleAfterChange1 = Convert.ToString(foTable.Rows[j]["AllocationGroupReportTitle"]);
            string FullRptName = Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]) + ": " + lsFinalTitleAfterChange1;
            if (FullRptName.Length > 74 && foTable.Rows.Count != 1)
                RptTitleCount--;
        }
        return RptTitleCount;
    }

    public void generatesCoverExcel(String fsAsofDate, String fsHouseholdName, String fsAllocationGroup, String fsFinalLocation, String fsContactID, DataTable foTable, String fsKeyContactID, String fsHouseHoldTitle, String fsContactFullname, String fsDisplayContactName, String lsFinalReportTitle, String lsCoverSheetPageTitle)
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

    protected void ddlHouseHold_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        chkNoComparison.Checked = false;
        txtPriorDate.Enabled = true;
        img1.Disabled = false;
        img1.Visible = true;

        DataTable dtBatch = GetBatchList(ddlHouseHold.SelectedValue, ddlBatchType.SelectedValue);

        gvList.Columns[0].Visible = true;
        gvList.Columns[5].Visible = true;
        gvList.Columns[6].Visible = true;
        gvList.Columns[7].Visible = true;
        gvList.DataSource = dtBatch;
        gvList.DataBind();
        gvList.Columns[0].Visible = false;
        gvList.Columns[5].Visible = false;
        gvList.Columns[6].Visible = false;
        gvList.Columns[7].Visible = false;

        if (dtBatch.Rows.Count > 0)
            btnGenerateReport.Visible = true;
        else
            btnGenerateReport.Visible = false;

    }

    protected void ddlBatchType_SelectedIndexChanged(object sender, EventArgs e)
    {
        lblError.Text = "";
        DataTable dtBatch = GetBatchList(ddlHouseHold.SelectedValue, ddlBatchType.SelectedValue);

        gvList.Columns[0].Visible = true;
        gvList.Columns[5].Visible = true;
        gvList.Columns[6].Visible = true;
        gvList.Columns[7].Visible = true;
        gvList.DataSource = dtBatch;
        gvList.DataBind();
        gvList.Columns[0].Visible = false;
        gvList.Columns[5].Visible = false;
        gvList.Columns[6].Visible = false;
        gvList.Columns[7].Visible = false;

        if (dtBatch.Rows.Count > 0)
            btnGenerateReport.Visible = true;
        else
            btnGenerateReport.Visible = false;
    }

    protected void gvList_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            CheckBox chk = (CheckBox)e.Row.FindControl("chkbSelectBatch");
            chk.Attributes.Add("onclick", "ClearLabel();");

            /***********Added 12/19/2016 NightlyLoad***********************/
            bool isAdvancedMode = (Request.QueryString["id"] ?? String.Empty).Equals("Scheduled");
            if (isAdvancedMode)
            {
                chk.Checked = true;
                //sw.WriteLine("checked" + DateTime.Now);
            }
            /***********Added 12/19/2016 NightlyLoad***********************/
        }
    }

    protected void chkNoComparison_CheckedChanged(object sender, EventArgs e)
    {
        if (chkNoComparison.Checked)
        {
            txtPriorDate.Enabled = false;
            txtPriorDate.Text = null;
            img1.Disabled = true;
            img1.Visible = false;
        }
        else
        {
            txtPriorDate.Enabled = true;
            img1.Disabled = false;
            img1.Visible = true;
        }
    }


}


//DATE:   15March2011
//BY :    Ankit
//DESC:   Added no comparison line check box 

