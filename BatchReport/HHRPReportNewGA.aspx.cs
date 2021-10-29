
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
using System.Security.Principal;
using System.Data.SqlClient;
//using CrmSdk;
using System.IO;
using System.Data.Common;
using Spire.Xls;
using System.Drawing;
using System.Xml;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Web.UI.DataVisualization.Charting;
using System.Linq;
using GemBox.Spreadsheet;
using GemBox.Document;
using Microsoft.SharePoint.Client;
using System.Security;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Text;
using System.Threading;
using Microsoft.IdentityModel.Claims;
public partial class BatchReport_HHRPReportNewGA : System.Web.UI.Page
{
    ClientContext context;
    public StreamWriter sw = null;
    string strDescription = string.Empty;
    bool bProceed = true;
    public int liPageSize = 29;//30 -- CHANGE THIS VALUE IN THE GENERATEPDF METHOD WHEN CHANGED HERE.
    //public int liPageSize = 27;
    public string lsStringName = "frutigerce-roman";
    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName, lsGAorTIAHeader;

    string ColorTIA1 = "#558ED5"; //Blue
    string ColorNetInvestedCap = "#77933C"; //Green
    string ColorTIA2 = "#558ED5"; //Blue
    string ColorInflationAdjInvCap = "#E46C0A";//Orange
    double extendedrangePerc = 1.05;
    double max1 = 0.0;
    double min1 = 0.0;
    int SelectedRptCnt = 0;
    clsCombinedReports objCombinedReports = new clsCombinedReports();
    String fsReportingName = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            //string strUserName = HttpContext.Current.User.Identity.Name.ToString();
            //Response.Write("UserName: "+ strUserName);

            // to find windows user 
            //System.Security.Principal.WindowsPrincipal p = System.Threading.Thread.CurrentPrincipal as System.Security.Principal.WindowsPrincipal;
            //string strName = p.Identity.Name;
            //Response.Write("<br/>p.Identity.Name:" + strName);
            //strName = HttpContext.Current.User.Identity.Name.ToString();
            //Response.Write("<br/>HttpContext.Current.User.Identity.Name:" + strName);
            //strName = Request.ServerVariables["AUTH_USER"]; //Finding with name
            //Response.Write("<br/>AUTH_USER:" + strName);
            ////////

            fillHouseholdType();
            fillHousehold();

        }


    }

    public void fillHousehold()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        DB clsDB = new DB();

        object HouseHoldType = ddlHouseHoldType.SelectedValue == "0" || ddlHouseHoldType.SelectedValue == "" ? "null" : "'" + ddlHouseHoldType.SelectedValue + "'";
        //DataSet loDataset = clsDB.getDataSet("sp_s_Get_HouseHoldName");
        DataSet loDataset = clsDB.getDataSet("sp_s_Get_HouseHoldName @RelationShipStatusID = " + HouseHoldType);

        ddlHouseHold.Items.Clear();
        ddlHouseHold.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHouseHold.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }

    }
    public void fillHouseholdType()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("SP_S_HH_Relationship_Status @ReportFlg = 1");
        ddlHouseHoldType.Items.Clear();
        ddlHouseHoldType.Items.Add(new System.Web.UI.WebControls.ListItem("ALL", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHouseHoldType.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][1])));
        }

    }
    public void BindHHReportType()
    {
        //ddlHousehold.Items.Add(new ListItem("fdf","dfsdf"));
        DB clsDB = new DB();
        object HHUUID = Convert.ToString(ddlHouseHold.SelectedValue == "0" ? "null" : "'" + ddlHouseHold.SelectedValue + "'");
        DataSet loDataset = clsDB.getDataSet("SP_S_GET_HH_GreshamReportType @HHUUID = " + HHUUID + " ");
        ddlReportType.Items.Clear();
        ddlReportType.Items.Add(new System.Web.UI.WebControls.ListItem("All", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlReportType.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][1])));
        }
    }

    public void BindHHRP()
    {
        DB clsDB = new DB();
        string Str = "";
        if (ddlReportType.SelectedValue == "0")
            Str = "SP_S_GET_HH_Parameters @HHUUID='" + ddlHouseHold.SelectedValue + "',@ReportTypeUUID =  null";
        else
            Str = "SP_S_GET_HH_Parameters @HHUUID='" + ddlHouseHold.SelectedValue + "',@ReportTypeUUID = '" + ddlReportType.SelectedValue + "'";
        DataSet loDataset = clsDB.getDataSet(Str);
        ddlHHRP.Items.Clear();
        ddlHHRP.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHHRP.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][1])));
        }

        if (ddlHHRP.SelectedValue != "0")
        {
            lblError.Text = "";
        }

    }

    protected void btnGenerateReport_Click(object sender, EventArgs e)
    {

        string ReportOpFolder = string.Empty;
        string ContactFolderName = string.Empty;
        string ParentFolder = string.Empty;
        string TempFolderPath = string.Empty;
        try
        {
            lblError.Text = "";
            string crmServerUrl = AppLogic.GetParam(AppLogic.ConfigParam.CRMServerurl); //"http://Crm01/";
            //string crmServerURL = "http://server:5555/";

            string orgName = "GreshamPartners";
            string currentuser = null;
            //string orgName = "Webdev";
            //  CrmService service = null;
            Boolean checkrunreport = false;
            String DestinationPath = string.Empty;
            string ConsolidatePdfFileName = string.Empty;

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


            //Response.Write(strUserName);

            strUserName = strUserName.Substring(strUserName.IndexOf("\\") + 1);

            //UserName_YYYYMMDD_Timewhere 
            //ViewState["ParentFolder"] = CurrentDateTime.Replace(":", "-").Replace("/", "-"); // orig

            ParentFolder = strUserName + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond + strMilliSecond;

            //string ReportOpFolder = "\\\\Fs01\\_ops_C_I_R_group\\Quarterly_Reports\\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            ReportOpFolder = Request.MapPath("ExcelTemplate\\TempFolder\\");  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            if (Request.Url.AbsoluteUri.Contains("localhost"))
            {
                ReportOpFolder = Request.MapPath("ExcelTemplate\\TempFolder\\");  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            }
            else
                ReportOpFolder = Request.MapPath("ExcelTemplate\\TempFolder\\");  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();


            //ReportOpFolder = "\\\\Fs01\\shared$\\BATCH REPORTS\\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();

            //if (Request.Url.AbsoluteUri.Contains("localhost"))
            //{
            //    ReportOpFolder = @"C:\Reports\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();
            //}
            //else
            //    ReportOpFolder = "\\\\Fs01\\shared$\\BATCH REPORTS\\" + Convert.ToString(ViewState["ParentFolder"]);  // ConfigurationManager.AppSettings.Keys["ReportOutputFolder"].ToString();



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

            for (int j = 0; j < 1; j++)
            {
                //CheckBox chkBox = (CheckBox)gvList.Rows[j].FindControl("chkbSelectBatch");

                //if (chkBox.Checked)
                //{
                DB clsDB = new DB();
                checkrunreport = true;
                String HHRPIDListTxt = Convert.ToString(ddlHHRP.SelectedValue);
                dtBatch = GetDataTable(HHRPIDListTxt);

                if (dtBatch == null)
                {
                    //lblError.Text = "Report can not be generated, Please try again.";
                    return;
                }



                String HHName = "";  //gvList.Rows[j].Cells[7].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                HHName = HHName.Replace("/", "");
                sourcefilecount = dtBatch.Rows.Count + 1;
                SourceFileArray = new string[sourcefilecount];
                Random rnd = new Random();
                string strRndNumber = Convert.ToString(rnd.Next(99999));
                for (int i = 0; i < dtBatch.Rows.Count; i++)
                {
                    if (Convert.ToString(dtBatch.Rows[i]["ssi_spvfilename"]) != "")
                    {
                        HHName = Convert.ToString(dtBatch.Rows[i]["ssi_spvfilename"]);
                        HHName = HHName.Replace("/", "");
                    }
                    else
                    {
                        HHName = ddlHHRP.SelectedItem.Text;
                        HHName = HHName.Replace("/", "");
                    }


                    ContactFolderName = ddlHHRP.SelectedItem.Text; //gvList.Rows[j].Cells[5].Text.Replace("/", "").Replace(":", "").Replace("*", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "").ToString();
                    ContactFolderName = ContactFolderName.Replace("/", "");

                    ////added 07-jun-2018 (sasmit - randomnumber for contact folder)
                    //ContactFolderName = ContactFolderName + "_" + strRndNumber;

                    //ContactFolderName = Convert.ToString(dtBatch.Rows[i]["Ssi_ContactIdName"]).Replace(",", "");
                    bool isExist = System.IO.Directory.Exists(ReportOpFolder + "\\" + ParentFolder);
                    TempFolderPath = ReportOpFolder + "\\" + ParentFolder;
                    if (!isExist)
                    {
                        //  Response.Write("Folder: " + ReportOpFolder + "\\" + ContactFolderName);
                        System.IO.Directory.CreateDirectory(ReportOpFolder + "\\" + ParentFolder);
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
                    String fsSummaryFlag = Convert.ToString(dtBatch.Rows[i]["Ssi_SummaryDetail"]);
                    String fsAllignment = Convert.ToString(dtBatch.Rows[i]["Ssi_Alignment"]);
                    String fsDisplayContactName = Convert.ToString(dtBatch.Rows[i]["ContactName"]);
                    String fsContactId = Convert.ToString(dtBatch.Rows[i]["ssi_ContactID"]);
                    String fsKeyContactID = Convert.ToString(dtBatch.Rows[i]["ssi_keycontactId"]);
                    String fsHousholdReportTitle = Convert.ToString(dtBatch.Rows[i]["ssi_householdreporttitle"]);
                    String fsGreshReportIdName = Convert.ToString(dtBatch.Rows[i]["ssi_GreshamReportIdName"]);
                    String fsGAorTIAflag = Convert.ToString(dtBatch.Rows[i]["ssi_gaortia"]);
                    String lsFinalTitleAfterChange = String.Empty;
                    String fsReportRollupGroupIdName = Convert.ToString(dtBatch.Rows[i]["Ssi_ReportRollupGroupIdName"]).Replace("'", "''");
                    String fsDiscretionaryFlg = Convert.ToString(dtBatch.Rows[i]["Discretionary Flag"]);
                    fsReportingName = Convert.ToString(dtBatch.Rows[i]["Ssi_ReportingName"]);

                    //added 2_1_2019 Non marketable (DYNAMO)
                    String fsReportRollupGroupId = Convert.ToString(dtBatch.Rows[i]["Ssi_ReportRollupGroupId"]);
                    String fsrHouseholdId = Convert.ToString(dtBatch.Rows[i]["Ssi_HouseholdId"]);
                    String fsFundIRR = Convert.ToString(dtBatch.Rows[i]["ssi_FundIRR"]);
                    String fsGreshamReportId = Convert.ToString(dtBatch.Rows[i]["ssi_GreshamReportId"]);
                    //added 5_20_2019 -- LegalEntity -- Title
                    String fsLegalEntityTitle = Convert.ToString(dtBatch.Rows[i]["Ssi_LegalEntityIdName"]);


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



                    //string abc = "See notes on Total Investments Assets and Asset Distribution for important information 1111111";
                    //string xyz = "See notes on Total Investments Assets and Asset Distribution for important information 1111111";
                    //string wer = abc + "\n" + xyz;

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

                    // Convert To Asset Dist Comparison logic
                    if (chkConvertToAssetDistComp.Checked)
                    {
                        fsSummaryFlag = "Summary Column";
                        fsGreshReportIdName = "Asset Distribution Comparison";
                        fsAllignment = "Vertical";
                    }

                    //overrid value of Underlying Manager Detail if Suppress manager detail is checked
                    if (chkSuppressManagerDetail.Checked)
                        fsVersion = "No";
                    /* END OF CHANGE*/

                    string strGUID = Guid.NewGuid().ToString();
                    // strGUID = strGUID.Substring(0, 5);
                    //String lsExcleSavePath = ReportOpFolder + "\\" + ContactFolderName + "\\" + fsHouseholdName.Replace(",", "") + "_" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + strGUID + ".xls";
                    //String lsExcleSavePath = ReportOpFolder + "\\" + ContactFolderName + "\\" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + lsFinalTitleAfterChange.Replace(",", "").Replace("/", "").Replace("\\", "") + "_" + Convert.ToDateTime(fsAsofDate).ToString("yyyyMMdd") + "_" + strGUID + ".xls";
                    String lsExcleSavePath = ReportOpFolder + "\\" + ParentFolder + "\\" + Convert.ToString(dtBatch.Rows[i]["Ssi_OrderNumber"]) + "_" + strGUID + ".xls";

                    String lsCoversheet = ReportOpFolder + "\\" + ParentFolder + "\\Coversheet.xls";
                    //String fsHouseHoldReportTitle = "";

                    // Generate report on excel and pdf

                    bool bContinueBatch = true;

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
                            //SharepointPath = SharepointPath.Replace("https:/greshampartners.sharepoint.com/clientserv/", "");
                            //SharepointPath = SharepointPath.Replace("http:/greshampartners.sharepoint.com/clientserv/", "");
                            SharepointPath = SharepointPath.Replace(AppLogic.GetParam(AppLogic.ConfigParam.clientservURL) + "/", "");
                            SharepointPath = SharepointPath.Replace(AppLogic.GetParam(AppLogic.ConfigParam.httpclientservURL) + "/ ", "");


                            SharepointPath = SharepointPath.Replace("%20", " ");
                            SharepointPath = SharepointPath.Replace(FileName, "");

                            string LocalPath = ReportOpFolder + "\\" + ParentFolder + "\\";

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
                        bContinueBatch = false;
                        //int numofPage = objCombinedReports.GetPageCountFromPDF(strTemplateFilePath);
                        //int CurPage = Convert.ToInt32(Convert.ToString(Session["CurPageInBatch"])) + 1;
                        //if (numofPage > 0)
                        //{
                        //    numofPage--;
                        //    dtBatch.Rows[i]["numPageNo"] = CurPage;
                        //    Session["CurPageInBatch"] = numofPage + CurPage;
                        //    bContinueBatch = false;
                        //}
                        //else
                        //    dtBatch.Rows[i]["numPageNo"] = 0;

                    }



                    bool CombinedFileName = false;
                    if (bContinueBatch)
                    {



                        //generatesExcelsheets(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt, fsGAorTIAflag);
                        //generatePDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, fsFooterTxt);

                        if (fsGreshReportIdName != "Asset Distribution" && fsGreshReportIdName != "Asset Distribution Comparison")
                        {
                            CombinedFileName = generateCombinedPDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath.Replace(".xls", ".pdf"), fsFooterTxt, fsGreshReportIdName, LegalEntity, FundID, CommitmentReportHeader, fsGAorTIAflag, fsReportRollupGroupIdName, fsReportRollupGroupId, fsrHouseholdId, fsFundIRR, HHRPIDListTxt, fsGreshamReportId, fsLegalEntityTitle, TempFolderPath, ssi_FooterLocation, ClientFooterTxt, Ssi_GreshamClientFooter);
                        }
                        else
                        {
                            SetValuesToVariable(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg);
                            // generatesExcelsheets(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, lsFinalTitleAfterChange, fsFooterTxt, fsGAorTIAflag);
                            bool bSuccess = generatePDF(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, lsExcleSavePath, fsFooterTxt, fsGAorTIAflag, fsDiscretionaryFlg, ssi_FooterLocation, ClientFooterTxt, Ssi_GreshamClientFooter);

                            if (!bSuccess)
                                return;
                            CombinedFileName = true;
                        }

                    }
                    else
                    {
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
                        generateCoversheetPDF(fsAsofDate, lsCoversheet, fsAllocationGroup, fsHouseholdName, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange, fsGAorTIAflag, fsDiscretionaryFlg, TempFolderPath);
                        generatesCoverExcel(fsAsofDate, fsHouseholdName, fsAllocationGroup, lsCoversheet, fsContactId, dtBatch, fsKeyContactID, fsHousholdReportTitle, fsContactFullname, fsDisplayContactName, lsFinalTitleAfterChange, fsGAorTIAflag, TempFolderPath);
                    }


                    /* Array fill with the PATH + Fullname of PDF*/

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

                //string ConsolidatePdfFileName = ContactFolderName + "_" + strYear + strMonth + strDay + ".pdf";
                ConsolidatePdfFileName = HHName + "_" + strYear + "-" + strMonth + strDay + ".pdf";

                ConsolidatePdfFileName = GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);
                string WatermarkCopy = "W_" + ConsolidatePdfFileName;
                if (!System.IO.File.Exists(ReportOpFolder + "\\" + ParentFolder + "\\" + ConsolidatePdfFileName))
                {
                    System.IO.File.Copy(ReportOpFolder + "\\" + ParentFolder + "\\Coversheet.pdf", ReportOpFolder + "\\" + ParentFolder + "\\" + ConsolidatePdfFileName);
                }

                DestinationPath = ReportOpFolder + "\\" + GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);
                //string  WatermarkDestinationPath = ReportOpFolder + "\\" + GeneralMethods.RemoveSpecialCharacters(WatermarkCopy);
                //   string WatermarkDestinationPath = HttpContext.Current.Server.MapPath("") + @"\\ExcelTemplate\\TempFolder\\" + GeneralMethods.RemoveSpecialCharacters(WatermarkCopy);
                string WatermarkDestinationPath = TempFolderPath + "\\" + GeneralMethods.RemoveSpecialCharacters(WatermarkCopy);
                //DestinationPath = Server.MapPath("") + @"\ExcelTemplate\" + ConsolidatePdfFileName;
                // Response.Write("LENGTH= " +WatermarkDestinationPath.Length.ToString());

                if (ContactFolderName.Contains("MTGBK")) //generate without coversheet
                {
                    string[] target = new string[sourcefilecount - 1];
                    Array.Copy(SourceFileArray, 1, target, 0, sourcefilecount - 1);
                    PDF.MergeFiles(DestinationPath, target);
                    #region Watermark
                    string Flg = string.Empty;
                    DataSet ds1 = new DataSet();
                    String HHRPIDListTxt1 = Convert.ToString(ddlHHRP.SelectedValue);
                    string lsSQL1 = "[SP_S_SETWATERMARK] @HHRPUUID='" + HHRPIDListTxt1 + "'";
                    ds1 = clsDB.getDataSet(lsSQL1);
                    if (ds1.Tables[0].Rows.Count > 0)
                    {
                        for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                        {

                            Flg = Convert.ToString(ds1.Tables[0].Rows[i][1]);

                        }
                        if (Flg == "1")
                        {
                            System.IO.File.Copy(DestinationPath, WatermarkDestinationPath, true);
                            objCombinedReports.WatermarkPdf(WatermarkDestinationPath, DestinationPath, "");
                            System.IO.File.Delete(WatermarkDestinationPath);
                        }
                    }
                    #endregion

                }
                else //generate with coversheet
                {
                    PDF.MergeFiles(DestinationPath, SourceFileArray);
                    #region Watermark
                    string Flg = string.Empty;
                    DataSet ds1 = new DataSet();
                    String HHRPIDListTxt1 = Convert.ToString(ddlHHRP.SelectedValue);
                    string lsSQL1 = "[SP_S_SETWATERMARK] @HHRPUUID='" + HHRPIDListTxt1 + "'";
                    ds1 = clsDB.getDataSet(lsSQL1);
                    if (ds1.Tables[0].Rows.Count > 0)
                    {
                        for (int i = 0; i < ds1.Tables[0].Rows.Count; i++)
                        {
                            Flg = Convert.ToString(ds1.Tables[0].Rows[i][1]);

                        }
                        if (Flg == "1")
                        {
                            System.IO.File.Copy(DestinationPath, WatermarkDestinationPath, true);
                            objCombinedReports.WatermarkPdf(WatermarkDestinationPath, DestinationPath, "");
                            System.IO.File.Delete(WatermarkDestinationPath);
                        }
                    }
                    #endregion
                }

                //  Directory.Delete(ReportOpFolder + "\\" + ParentFolder, true);

                //}

            }


            ////////////////////////////////////

            if (1 == 1)
            {

                //ConsolidatePdfFileName = ConsolidatePdfFileName.Replace("&", "and").Replace("'","");


                ConsolidatePdfFileName = GeneralMethods.RemoveSpecialCharacters(ConsolidatePdfFileName);


                string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + ConsolidatePdfFileName);

                //string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + ConsolidatePdfFileName.Replace(",", "").Replace("(", "").Replace(")", "").Replace("'", ""));

                //  System.IO.File.Copy(DestinationPath, strDirectory, true);
                //Directory.Delete(ReportOpFolder, true);
                string lsFileNamforFinal;
                lsFileNamforFinal = "./ExcelTemplate/TempFolder/" + ConsolidatePdfFileName;
                //lsFileNamforFinal = "./ExcelTemplate/" + ConsolidatePdfFileName.Replace(",", "").Replace("(", "").Replace(")", "").Replace("'", "");
                string newWindow = string.Empty;
                try
                {
                    ////loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
                    Response.Write("<script>");
                    lsFileNamforFinal = "./ExcelTemplate/TempFolder/" + ConsolidatePdfFileName;
                    Response.Write("window.open('ViewReport.aspx?" + ConsolidatePdfFileName + "', 'mywindow')");
                    Response.Write("</script>");
                }
                catch (Exception exc)
                {
                    Response.Write(exc.Message);
                }

                /*
            string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\" + ConsolidatePdfFileName.Replace(",", "").Replace("(", "").Replace(")", "").Replace("'", ""));

             File.Copy(DestinationPath, strDirectory, true);

             try
             {
                 //loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

                 Response.Write("<script>");
                 string lsFileNamforFinal = "./ExcelTemplate/" + ConsolidatePdfFileName.Replace(",", "").Replace("(", "").Replace(")", "").Replace("'", "");
                 Response.Write("window.open('" + lsFileNamforFinal + "', 'mywindow')");
                 Response.Write("</script>");

             }
             catch (Exception exc)
             {
                 Response.Write(exc.Message);
             }
               
             */
            }
            ////////////////////////////////////

            ////////////////////////////////////////////////////////////////////////////////////
            #region Commented
            //DirectoryInfo dir = new DirectoryInfo(DestinationPath);
            //string[] SourceFiles;
            //SourceFiles = new string[dir.GetFiles().Length];
            //int fileNo = 0;

            ////foreach (string strFile in Directory.GetFiles(DestinationPath, "*.pdf"))
            ////{
            ////    SourceFiles[fileNo] = file.FullName;// lsExcleSavePath.Replace(".xls", ".pdf");
            ////    fileNo++;
            ////}

            //foreach (FileInfo file in dir.GetFiles())
            //{
            //    SourceFiles[fileNo] = file.FullName;// lsExcleSavePath.Replace(".xls", ".pdf");
            //    fileNo++;
            //}

            //DestinationPath = ReportOpFolder + "\\" + "ConsolidatePDFNEW.pdf";

            //PDF.MergeFiles(DestinationPath, SourceFiles);

            #endregion
            ////////////////////////////////////////////////////////////////////////////////////

            //dtBatch.Clear();
            //dtBatch = null;
            //dtBatch = GetBatchList(ddlHouseHold.SelectedValue, ddlBatchType.SelectedValue);

            //gvList.Columns[0].Visible = true;
            //gvList.Columns[5].Visible = true;
            //gvList.Columns[6].Visible = true;
            //gvList.DataSource = dtBatch;
            //gvList.DataBind();
            //gvList.Columns[0].Visible = false;
            //gvList.Columns[5].Visible = false;
            //gvList.Columns[6].Visible = false;

            if (checkrunreport)
                lblError.Text = "Reports generated successfully";
            else
                lblError.Text = "Please Select a batch to run report.";
        }
        catch (Exception ex)
        {
            //if (Directory.Exists(ReportOpFolder + "\\" + ParentFolder))
            //{
            //    Directory.Delete(ReportOpFolder + "\\" + ParentFolder, true);
            //}
            lblError.Text = "Error Generating Report " + ex.ToString();
        }
        finally
        {
            if (Directory.Exists(ReportOpFolder + "\\" + ParentFolder))
            {
                Directory.Delete(ReportOpFolder + "\\" + ParentFolder, true);
            }
        }
    }
    public string sharepointFile(string FileName, string path, string finalPath)
    {
        string Value = null;

        string siteUrl = AppLogic.GetParam(AppLogic.ConfigParam.clientservURL);
      //  string siteUrl = "https://greshampartners.sharepoint.com/clientserv";
        context = new ClientContext(siteUrl);
        SecureString passWord = new SecureString();

        //foreach (var c in "w!ldWind36") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("sp_workflow@greshampartners.com", passWord);
        //foreach (var c in "51ngl3malt") passWord.AppendChar(c);
        //context.Credentials = new SharePointOnlineCredentials("gbhagia@greshampartners.com", passWord);

        string user = AppLogic.GetParam(AppLogic.ConfigParam.SPUserEmailID).ToString();
        string Pass = AppLogic.GetParam(AppLogic.ConfigParam.SPUserPassword).ToString();
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
        String LegalEntityId, String FundId, String CommitmentReportHeader, String GAorTIAflag, String ReportRollupGroupIdName, String fsReportRollupGroupId, String fsHouseholdId, String fsFundIRR, String HHParameterTxt, String fsGreshamReportId, String fsLegalEntityTitle, String TempFolderPath, String FooterLocation, String ClientFooterTxt, String Ssi_GreshamClientFooter)
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
        objCombinedReports.ReportSource = "AdventInd";
        objCombinedReports.PriorDate = fsSPriorDate;

        //added 2_1_2019 - Non Marketable(DYNAMO)
        objCombinedReports.ReportRollupGroupId = fsReportRollupGroupId;
        objCombinedReports.HouseholdId = fsHouseholdId;
        objCombinedReports.FundIRR = fsFundIRR;
        objCombinedReports.HHParameterTxt = HHParameterTxt;
        objCombinedReports.ReportingID = fsGreshamReportId;

        objCombinedReports.Footerlocation = FooterLocation;

        objCombinedReports.ClientFooterTxt = ClientFooterTxt;

        objCombinedReports.Ssi_GreshamClientFooter = Ssi_GreshamClientFooter;



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
        objCombinedReports.ReportName = ReportName;
        if (fsReportingName != "")
            objCombinedReports.ReportingName = fsReportingName;

        if (ReportName == "Client Goals" || ReportName == "Absolute Returns" || ReportName == "Capital Protection" || ReportName == "Short Term Performance")
        {
            string Gresham_String = AppLogic.GetParam(AppLogic.ConfigParam.DBConnectionstring);// "Password=slater6;Persist Security Info=True;User ID=mpiuser;Initial Catalog=GreshamPartners_MSCRM;Data Source=sql01";


            SqlConnection Gresham_con = new SqlConnection(Gresham_String);
            String HHRPIDListTxt = Convert.ToString(ddlHHRP.SelectedValue);
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
            DB clsDB = new DB();

            //if (ReportName == "Client Goals")
            //{
            //    DataSet ds = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt1LineChart));

            //    string _Chart1 = getLineChartReport1(ds);
            //    objCombinedReports.Chart1 = _Chart1;
            //}
            //if (ReportName == "Absolute Returns")
            //{
            //    DataSet ds = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt3LineChart));
            //    string _Chart1 = getLineChartReport3(ds);
            //    objCombinedReports.Chart1 = _Chart1;

            //    DataSet ds1 = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt3BarChart));
            //    string _Chart2 = getBarChartReport3(ds1);
            //    objCombinedReports.Chart2 = _Chart2;
            //}
            //if (ReportName == "Capital Protection")
            //{
            //    DataSet ds = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt4ShapeChart));
            //    string _Chart1 = getShapeChartReport4("4", ds);
            //    objCombinedReports.Chart1 = _Chart1;

            //    DataSet ds1 = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt4ColumnChart));
            //    string _Chart2 = getColumnChartReport4(fsHouseholdName, ds1);
            //    objCombinedReports.Chart2 = _Chart2;
            //}
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


    private DataTable GetDataTable(String HHRPIdListTxt)
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
            object EndDate = txtEndDate.Text == "" ? "null" : "'" + txtEndDate.Text + "'";

            //object NoComparison = chkNoComparison.Checked == false ? 0 : 1;
            greshamquery = "[SP_S_BATCH_HH_PARAMETER] @HHParameterListTxt='" + HHRPIdListTxt + "',@PriorDT=" + PriorDate + ",@EndDT=" + EndDate + ",@NoComparisonLineFlg=" + Convert.ToBoolean(chkNoComparison.Checked);

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
            Response.Write("sp_s_batch sp fails error desc:" + exc.Message);
            //LogMessage(sw, service, strDescription, 62, "Anziano Position");
        }

        if (ds_gresham.Tables.Count > 0)
            return ds_gresham.Tables[0];
        else
            return null;

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
            greshamquery = "sp_s_batch_list_CONSOLIDETED @HouseHoldId =" + HouseholdID + ",@BatchType=" + BatchType;
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
    //public static CrmService GetCrmService(string crmServerUrl, string organizationName)
    //{
    //    // Get the CRM Users appointments
    //    // Setup the Authentication Token
    //    CrmAuthenticationToken token = new CrmAuthenticationToken();
    //    token.AuthenticationType = 0; // Use Active Directory authentication.
    //    token.OrganizationName = organizationName;
    //    string username = WindowsIdentity.GetCurrent().Name;

    //    if (username == "CORP\\gbhagia")
    //    {
    //        // Use the global user ID of the system user that is to be impersonated.
    //        token.CallerId = new Guid("EE8E3A77-59E2-DD11-831F-001D09665E8F");//deb
    //        //token.CallerId = new Guid("C42C7E05-8303-DE11-A38C-001D09665E8F");//gary                
    //    }
    //    token.CallerId = new Guid("EE8E3A77-59E2-DD11-831F-001D09665E8F");//deb
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

    public string getFinalSp(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsGAorTIAflag, String fsDiscretionaryFlg)
    {
        String lsSQL = "";
        //if (chkNewGA.Checked)
        //{
        if (!String.IsNullOrEmpty(fsAllocationGroup))
        {
            lsSQL = "SP_R_Advent_Report_Allocation_NEW_GA @AllocationGroupNameTxt='" + fsAllocationGroup + "', ";
        }
        else
        {
            lsSQL = "SP_R_Advent_Report_Other_NEW_GA";
        }
        //}
        //else
        //{
        //    if (!String.IsNullOrEmpty(fsAllocationGroup))
        //    {
        //        lsSQL = "SP_R_Advent_Report_Allocation @AllocationGroupNameTxt='" + fsAllocationGroup + "', ";
        //    }
        //    else
        //    {
        //        lsSQL = "SP_R_Advent_Report_Other";
        //    }
        //}

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

        if (chkNoComparison.Checked && chkConvertToAssetDistComp.Checked)
        {
            lsSQL = lsSQL + ",@ComparisonFlg = 1"; //Override
        }
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
        String lsFileNamforFinalXls = System.DateTime.Now.ToString("MMddyyHHmmss") + ".xls";
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
                //                foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216); //change by abhi
                foCell.BorderColorBottom = new iTextSharp.text.Color(191, 191, 191);
            }
        }
        catch { }
    }

    public void setGreyBorder(Cell foCell)
    {

        foCell.BorderWidthBottom = 0.1F;
        //foCell.BorderColorBottom = new iTextSharp.text.Color(242, 242, 242);
        // foCell.BorderColorBottom = new iTextSharp.text.Color(216, 216, 216);
        foCell.BorderColorBottom = new iTextSharp.text.Color(191, 191, 191);  //change by abhi
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
        loCell.Leading = 13F;
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



                    for (int liCounter = 0; liCounter < liPageSize - 2 - liLastPageData - EndOfReportPageCnt - 1 ; liCounter++)
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
    public bool generatePDF(String fsAllocationGroup, String fsHouseholdName, String fsAsofDate, String fsSPriorDate, String fsLookthrogh, String fsContactFullname, String fsVersion, String fsSummaryFlag, String fsAllignment, String fsReportGroupflag, String fsReportgroupflag2, String fsFinalLocation, String lsFooterTxt, String fsGAorTIAflag, String fsDiscretionaryFlg, String FooterLocation, String ClientFooterTxt, String Ssi_GreshamClientFooter)
    {
        bool bSuccess = true;
        
     //   liPageSize = 28;//commented page size 28 to 26 as discuss with sir.
        liPageSize = 28;
        DataSet lodataset; DB clsDB = new DB();
        lodataset = null;

        String lsSQL = getFinalSp(fsAllocationGroup, fsHouseholdName, fsAsofDate, fsSPriorDate, fsLookthrogh, fsContactFullname, fsVersion, fsSummaryFlag, fsAllignment, fsReportGroupflag, fsReportgroupflag2, fsGAorTIAflag, fsDiscretionaryFlg);
        // Response.Write(lsSQL);
        lodataset = clsDB.getDataSet(lsSQL);


        if (lodataset.Tables[0].Rows.Count < 1)
        {
            lblError.Text = "No Records found";
            return bSuccess;
        }

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
        String ls = fsFinalLocation.Replace(".xls", ".pdf"); // Server.MapPath("") + "/" + System.DateTime.Now.ToString("MMddyyHHmmss") + ".pdf";
        iTextSharp.text.pdf.PdfWriter.GetInstance(document, new FileStream(ls, FileMode.Create));
        document.Open();

        if (loInsertdataset.Tables[0].Columns.Count > 12)
        {
            lblError.Text = "Number of Report Rollup Groups are greater then 9.<br/> It will not fit in the report so please select another household.";
            return false;
        }

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
            //liTotalPage = liTotalPage + 1;
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
                    document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, liPageSize, false, String.Empty, FooterLocation, ClientFooterTxt, Ssi_GreshamClientFooter));
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
                        loCell.BackgroundColor = new iTextSharp.text.Color(191, 191, 191);//(216, 216, 216);
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
                document.Add(addFooter(lsDateTime, liTotalPage, liCurrentPage, loInsertdataset.Tables[0].Rows.Count % liPageSize, true, lsFooterTxt, FooterLocation, ClientFooterTxt, Ssi_GreshamClientFooter));
            }
        }

        document.Close();

        //FileInfo loFile = new FileInfo(ls);
        //try
        //{
        //    loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));
        //}
        //catch { }

        return bSuccess;
    }

    public void generateCoversheetPDF(String lsDateString, String fsFinalLocation, String fsAllocationGroup, String fsHouseholdName, String fsContactId, DataTable foTable, String fsKeyContactID, String fsHouseHoldTitle, String fsContactFullname, String fsDisplayContactName, String lsFinalReportTitle, String fsGAorTIAflag, String fsDiscretionaryFlg, String TempFolderPath)
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
        //String ls = Server.MapPath("") + "/a" + System.DateTime.Now.ToString("MMddyyHHmmss") + System.Guid.NewGuid().ToString() + ".pdf";
        String ls = TempFolderPath + "\\" + Guid.NewGuid().ToString() + ".pdf";
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

                    fsGAorTIAflag = Convert.ToString(foTable.Rows[j]["ssi_gaortia"]);
                    fsDiscretionaryFlg = Convert.ToString(foTable.Rows[j]["Discretionary Flag"]);

                    if (fsGAorTIAflag == "GA")
                    {
                        if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                            loChunk = new Chunk("GA " + Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
                        else
                            loChunk = new Chunk("GA " + Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
                    }
                    else
                    {
                        if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                            loChunk = new Chunk(Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
                        else
                            loChunk = new Chunk(Convert.ToString(foTable.Rows[j]["ssi_greshamreportidname"]).Replace("v2.1", "") + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));//setFontsAll(10, 0, 1));
                    }

                    if (chkConvertToAssetDistComp.Checked)
                    {
                        if (fsGAorTIAflag == "GA")
                        {
                            if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                                loChunk = new Chunk("GA Asset Distribution Comparison" + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
                            else
                                loChunk = new Chunk("GA Asset Distribution Comparison" + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));
                        }
                        else
                        {
                            if (fsDiscretionaryFlg.ToUpper() == "TRUE")
                                loChunk = new Chunk("Asset Distribution Comparison" + " - Discretionary: " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));//setFontsAll(10, 0, 1));
                            else
                                loChunk = new Chunk("Asset Distribution Comparison" + ": " + lsFinalTitleAfterChange, setFontsAll(10, 0, 0));//setFontsAll(10, 0, 1));
                        }
                    }

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


    public void generatesCoverExcel(String fsAsofDate, String fsHouseholdName, String fsAllocationGroup, String fsFinalLocation, String fsContactID, DataTable foTable, String fsKeyContactID, String fsHouseHoldTitle, String fsContactFullname, String fsDisplayContactName, String lsFinalReportTitle, String fsGAorTIAflag, String TempFolderPath)
    {

        //String lsFileNamforFinalXls = System.DateTime.Now.ToString("MMddyyHHmmss") + ".xls";
        String lsFileNamforFinalXls = Guid.NewGuid().ToString() + ".xls";
        string strDirectory1 = (Server.MapPath("") + @"\ExcelTemplate\coversheet.xls");
        //string strDirectory = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls);        
        //string strDirectory2 = (Server.MapPath("") + @"\ExcelTemplate\TempFolder\" + lsFileNamforFinalXls.Replace("xls", "xml"));
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
        chkConvertToAssetDistComp.Checked = false;
        chkSuppressManagerDetail.Checked = false;
        txtPriorDate.Enabled = true;
        img1.Disabled = false;
        img1.Visible = true;

        if (ddlHouseHold.SelectedValue != "0")
        {
            BindHHReportType();
            BindHHRP();
        }
        else
        {
            ddlReportType.Items.Clear();
            ddlReportType.Items.Add(new System.Web.UI.WebControls.ListItem("All", "0"));

            ddlHHRP.Items.Clear();
            ddlHHRP.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        }


        //DataTable dtBatch = GetBatchList(ddlHouseHold.SelectedValue, ddlBatchType.SelectedValue);

        //gvList.Columns[0].Visible = true;
        //gvList.Columns[5].Visible = true;
        //gvList.Columns[6].Visible = true;
        //gvList.Columns[7].Visible = true;
        //gvList.DataSource = dtBatch;
        //gvList.DataBind();
        //gvList.Columns[0].Visible = false;
        //gvList.Columns[5].Visible = false;
        //gvList.Columns[6].Visible = false;
        //gvList.Columns[7].Visible = false;

        //if (dtBatch.Rows.Count > 0)
        //    btnGenerateReport.Visible = true;
        //else
        //    btnGenerateReport.Visible = false;

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
            lblError.Text = "";
            // calenderAnchor.on = true;
        }
        else
        {
            txtPriorDate.Enabled = true;
            img1.Disabled = false;
            img1.Visible = true;
            lblError.Text = "";
            //calenderAnchor.Disabled = false;
        }

    }
    protected void btnGetReport_Click(object sender, EventArgs e)
    {

    }

    protected void ddlReportType_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (ddlHouseHold.SelectedValue != "0")
            BindHHRP();
        else
        {
            ddlHHRP.Items.Clear();
            ddlHHRP.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        }
    }


    private string getLineChartReport1(DataSet ds)
    {

        System.Random rand = new System.Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmssfff") + rand.Next().ToString();


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + strGUID + ".xls";
        String fsFinalLocation1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\P_OP_" + strGUID + ".xls";

        DB clsDB = new DB();
        // string TIAGrp = ddlTIAGrp.SelectedValue.Replace("'", "''") == "" ? "null" : "'" + ddlTIAGrp.SelectedValue.Replace("'", "''") + "'";
        //  DataSet ds = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt1LineChart));

        string Dt1;
        string sDay;
        string sMonth;
        string sYear;
        double max1 = 0.0;
        double max2 = 0.0;
        double max4 = 0.0;
        double maxx1 = 0.0;
        double maxx2 = 0.0;
        DataTable dtatble = ds.Tables[0];

        LineChart1.DataSource = dtatble;
        LineChart1.DataBind();

        LineChart2.DataSource = dtatble;
        LineChart2.DataBind();

        // Set series chart type
        LineChart1.Series[0].ChartType = SeriesChartType.Line;
        LineChart1.Series[1].ChartType = SeriesChartType.Line;

        LineChart2.Series[0].ChartType = SeriesChartType.Line;
        LineChart2.Series[1].ChartType = SeriesChartType.Line;

        DateTime minDate = new DateTime();
        if (ds.Tables[0].Rows.Count > 0)
        {
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                Double s1val = Convert.ToDouble(ds.Tables[0].Rows[i]["value"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                Double s2val = Convert.ToDouble(ds.Tables[0].Rows[i]["NetInvestments"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                Double s4val = Convert.ToDouble(ds.Tables[0].Rows[i]["Infl. Adj. Net InvestMent"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                Dt1 = ds.Tables[0].Rows[i]["Date"].ToString();
                sDay = DateTime.Parse(Dt1).ToString("dd");
                sMonth = DateTime.Parse(Dt1).ToString("MM");
                sYear = DateTime.Parse(Dt1).ToString("yyyy");

                DateTime dtDate = Convert.ToDateTime(Dt1);

                if (Convert.ToString(minDate) != "" && dtDate.Month == 12)
                    minDate = dtDate;

                LineChart1.Series["Series1"].Points.AddXY(dtDate, s1val);
                LineChart1.Series["Series2"].Points.AddXY(dtDate, s2val);

                LineChart2.Series["Series1"].Points.AddXY(dtDate, s1val);
                LineChart2.Series["Series2"].Points.AddXY(dtDate, s4val);

                //To set max point on chart
                if (s1val > max1)
                    max1 = s1val;

                if (s2val > max2)
                    max2 = s2val;

                if (s4val > max4)
                    max4 = s4val;

            }

            //Set Max value of Y-Axis -- Left Line chart 
            if (max1 > max2)
                maxx1 = RoundToMax(max1);

            if (max2 > max1)
                maxx1 = RoundToMax(max2);

            if (maxx1 != 0.0)
                LineChart1.ChartAreas[0].AxisY.Maximum = maxx1;

            //Set Max value of Y-Axis -- Right Line chart 
            if (max1 > max4)
                maxx2 = RoundToMax(max1);

            if (max4 > max1)
                maxx2 = RoundToMax(max4);

            if (maxx2 != 0.0)
                LineChart2.ChartAreas[0].AxisY.Maximum = maxx2;

            if (maxx1 > 5000000 && maxx1 < 60000000)
                LineChart1.ChartAreas["ChartArea1"].AxisY.Interval = 5000000;

            if (maxx2 > 5000000 && maxx2 < 60000000)
                LineChart2.ChartAreas["ChartArea1"].AxisY.Interval = 5000000;

            Double S1LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["value"].ToString());
            Double S2LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["NetInvestments"].ToString());
            Double S3LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Infl. Adj. Net InvestMent"].ToString());

            LineChart1.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);
            LineChart1.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);

            LineChart1.Series[0].BorderWidth = 10;
            LineChart1.Series[1].BorderWidth = 10;


            LineChart1.Series[0].Color = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChart1.Series[1].Color = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);

            if (ds.Tables[0].Rows.Count < 12) //1years --MONTHLY
            {
                LineChart1.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChart1.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;

            }
            else if (ds.Tables[0].Rows.Count >= 12 && ds.Tables[0].Rows.Count < 36) //Quaterly
            {
                LineChart1.ChartAreas["ChartArea1"].AxisX.Interval = 3;
                LineChart1.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
            }
            else
            {
                // LineChart1.ChartAreas["ChartArea1"].AxisX.Minimum = 10000;
                LineChart1.ChartAreas["ChartArea1"].AxisX.Interval = 1;

                LineChart1.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Years;
                LineChart1.ChartAreas["ChartArea1"].AxisX.IntervalOffset = -1;
                LineChart1.ChartAreas["ChartArea1"].AxisX.IntervalOffsetType = DateTimeIntervalType.Days;
                // LineChart1.ChartAreas["ChartArea1"].AxisX.IsMarginVisible = true;
            }



            LineChart1.Series[0].Name = "Total Investment Assets (TIA)";
            LineChart1.Series[1].Name = "Net Invested Capital";


            //LineChart1.Series[0].BorderColor = S;
            LineChart1.Series[0].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChart1.Series[1].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);


            LineChart1.Series[0].Points[ds.Tables[0].Rows.Count - 1].Label = S1LastValue.ToString("C0");
            LineChart1.Series[1].Points[ds.Tables[0].Rows.Count - 1].Label = S2LastValue.ToString("C0");

            LineChart1.Series[0].SmartLabelStyle.Enabled = true;
            LineChart1.Series[0].SmartLabelStyle.Enabled = true;

            if (S1LastValue > S2LastValue)
            {
                LineChart1.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                LineChart1.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            }
            else
            {
                LineChart1.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                LineChart1.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            }

            LineChart1.ChartAreas[0].AxisX.IsMarginVisible = false;
            LineChart1.ChartAreas[0].AxisX.IsStartedFromZero = true;
            LineChart2.ChartAreas[0].AxisX.IsMarginVisible = false;
            LineChart2.ChartAreas[0].AxisX.IsStartedFromZero = true;

            //Remove Extra for label to display 
            LineChart1.Series[0].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChart1.Series[0].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChart1.Series[0].SmartLabelStyle.CalloutLineWidth = 0;

            LineChart1.Series[1].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChart1.Series[1].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChart1.Series[1].SmartLabelStyle.CalloutLineWidth = 0;



            LineChart1.Series[0].IsVisibleInLegend = true;
            LineChart1.Series[1].IsVisibleInLegend = true;


            LineChart1.ChartAreas[0].AxisX.LabelStyle.Format = "MMM-yy";
            LineChart1.ChartAreas[0].AxisX.LabelStyle.Angle = -90;



            LineChart2.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);
            LineChart2.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);

            LineChart2.Series[0].BorderWidth = 10;
            LineChart2.Series[1].BorderWidth = 10;


            LineChart2.Series[0].Color = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChart2.Series[1].Color = System.Drawing.ColorTranslator.FromHtml(ColorInflationAdjInvCap);


            LineChart2.Series[0].Name = "Total Investment Assets (TIA)";
            LineChart2.Series[1].Name = "Inflation Adj. Net Invested Capital";


            //LineChart2.Series[0].BorderColor = S;
            LineChart2.Series[0].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChart2.Series[1].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorInflationAdjInvCap);


            LineChart2.Series[0].Points[ds.Tables[0].Rows.Count - 1].Label = S1LastValue.ToString("C0");
            LineChart2.Series[1].Points[ds.Tables[0].Rows.Count - 1].Label = S3LastValue.ToString("C0");

            if (S1LastValue > S3LastValue)
            {
                LineChart2.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                LineChart2.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;

            }
            else
            {
                LineChart2.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                LineChart2.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            }

            LineChart2.Series[0].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChart2.Series[0].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChart2.Series[0].SmartLabelStyle.CalloutLineWidth = 0;

            LineChart2.Series[1].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChart2.Series[1].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChart2.Series[1].SmartLabelStyle.CalloutLineWidth = 0;



            if (ds.Tables[0].Rows.Count < 12) //1years --MONTHLY
            {
                LineChart2.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChart2.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
            }
            else if (ds.Tables[0].Rows.Count >= 12 && ds.Tables[0].Rows.Count < 36) //Quaterly
            {
                LineChart2.ChartAreas["ChartArea1"].AxisX.Interval = 3;
                LineChart2.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
            }
            else
            {
                LineChart2.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChart2.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Years;
                LineChart2.ChartAreas["ChartArea1"].AxisX.IntervalOffset = -1;
                LineChart2.ChartAreas["ChartArea1"].AxisX.IntervalOffsetType = DateTimeIntervalType.Days;
            }


            LineChart2.Series[0].IsVisibleInLegend = true;
            LineChart2.Series[1].IsVisibleInLegend = true;


            LineChart2.ChartAreas[0].AxisX.LabelStyle.Format = "MMM-yy";
            LineChart2.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
        }



        //LineChart1.ChartAreas[0].Position.X = 5;
        //LineChart1.ChartAreas[0].Position.Y = 8;
        //LineChart1.ChartAreas[0].Position.Height = 82;
        //LineChart1.ChartAreas[0].Position.Width = 97;

        //LineChart2.ChartAreas[0].Position.X = 5;
        //LineChart2.ChartAreas[0].Position.Y = 8;
        //LineChart2.ChartAreas[0].Position.Height = 82;
        //LineChart2.ChartAreas[0].Position.Width = 97;

        System.Random rnd = new System.Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + RNum + ".bmp";
        string filename1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\A_OP_" + RNum + ".bmp";

        // filename = Server.MapPath("~") + @"\\TempImages\\ChartImage-" + RNum + ".bmp";

        Bitmap bm = new Bitmap(1300, 800);
        Bitmap bm1 = new Bitmap(1300, 800);

        bm.SetResolution(300, 300);
        bm1.SetResolution(300, 300);

        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);
        System.Drawing.Graphics gGraphics1 = System.Drawing.Graphics.FromImage(bm1);


        LineChart1.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 1300, 800));
        LineChart2.Paint(gGraphics1, new System.Drawing.Rectangle(0, 0, 1300, 800));

        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);
        bm1.Save(filename1, System.Drawing.Imaging.ImageFormat.Bmp);


        //  Chart1.SaveImage(filename, ChartImageFormat.Bmp);


        foreach (var series in LineChart2.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }




        return filename;



    }

    private double RoundToMax(double maxfromdb)
    {
        try
        {
            maxfromdb = maxfromdb + (int)(maxfromdb * 0.12);
            int length = Convert.ToInt64(maxfromdb).ToString().Length;
            // length++;
            int num = 0;
            if (length == 6)
                num = 50000;
            if (length == 7)
                num = 500000;
            if (length == 8)
                num = 5000000;
            if (length == 9)
                num = 50000000;
            if (length == 10)
                num = 500000000;
            double i = Math.Ceiling(maxfromdb / (Double)num) * num;

            return i;
        }
        catch (Exception ex)
        {
            return 0.0;
        }
    }

    private string getLineChartReport3(DataSet ds)
    {
        clsCombinedReports objCombinedReports = new clsCombinedReports();
        System.Random rand = new System.Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmssfff") + rand.Next().ToString();


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\2OP_" + strGUID + ".xls";
        String fsFinalLocation1 = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\P_2OP_" + strGUID + ".xls";

        DB clsDB = new DB();

        //  DataSet ds = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt3LineChart));

        string Dt1;
        string sDay;
        string sMonth;
        string sYear;

        //chart1
        //  TimeSeries s1 = new TimeSeries("Gresham Advised Assets");
        //  TimeSeries s2 = new TimeSeries("Net Invested Capital");
        // TimeSeries s3 = new TimeSeries("Inflation Adj. Net Invested Capital");

        DataTable dtatble = ds.Tables[0];

        LineChart.DataSource = dtatble;
        LineChart.DataBind();

        // Set series chart type
        LineChart.Series[0].ChartType = SeriesChartType.Line;
        LineChart.Series[1].ChartType = SeriesChartType.Line;
        LineChart.Series[2].ChartType = SeriesChartType.Line;
        double max1 = 0.0, max2 = 0.0, max3 = 0.0;
        double maxx1 = 0.0;


        if (ds.Tables[0].Rows.Count > 0)
        {
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                Double s1val = Convert.ToDouble(ds.Tables[0].Rows[i]["Gresham Advised Assets"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                Double s2val = Convert.ToDouble(ds.Tables[0].Rows[i]["Net Invested Capital"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                Double s3val = Convert.ToDouble(ds.Tables[0].Rows[i]["Infl. Adj. Net InvestMent"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                Dt1 = ds.Tables[0].Rows[i]["Year"].ToString();
                sDay = DateTime.Parse(Dt1).ToString("dd");
                sMonth = DateTime.Parse(Dt1).ToString("MM");
                sYear = DateTime.Parse(Dt1).ToString("yyyy");

                DateTime dtDate = Convert.ToDateTime(Dt1);



                //chart 1
                // s1.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s1val);
                //  s2.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s2val);
                //  s3.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s3val);
                string d = Convert.ToInt16(sMonth) + "/" + Convert.ToInt16(sDay) + "/" + Convert.ToInt32(sYear);

                LineChart.Series["Series1"].Points.AddXY(dtDate, s1val);
                LineChart.Series["Series2"].Points.AddXY(dtDate, s2val);
                LineChart.Series["Series3"].Points.AddXY(dtDate, s3val);

                //To set max point on chart
                if (s1val > max1)
                    max1 = s1val;

                if (s2val > max2)
                    max2 = s2val;

                if (s3val > max3)
                    max3 = s3val;
            }


            //Set Max value of Y-Axis -- Left Line chart 
            if (max1 > max2 && max1 > max3)
                maxx1 = RoundToMax(max1);

            if (max2 > max1 && max2 > max3)
                maxx1 = RoundToMax(max2);

            if (max3 > max1 && max3 > max2)
                maxx1 = RoundToMax(max3);

            if (maxx1 != 0.0)
                LineChart.ChartAreas[0].AxisY.Maximum = maxx1;

            if (maxx1 > 5000000 && maxx1 < 50000000)
                LineChart.ChartAreas["ChartArea1"].AxisY.Interval = 5000000;

            Double S1LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Gresham Advised Assets"].ToString());
            Double S2LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Net Invested Capital"].ToString());
            Double S3LastValue = Convert.ToDouble(ds.Tables[0].Rows[ds.Tables[0].Rows.Count - 1]["Infl. Adj. Net InvestMent"].ToString());

            LineChart.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);
            LineChart.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Regular);

            LineChart.Series[0].BorderWidth = 10;
            LineChart.Series[1].BorderWidth = 10;
            LineChart.Series[2].BorderWidth = 10;

            LineChart.Series[0].Color = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChart.Series[1].Color = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);
            LineChart.Series[2].Color = System.Drawing.ColorTranslator.FromHtml(ColorInflationAdjInvCap);

            if (ds.Tables[0].Rows.Count < 12) //1years --MONTHLY
            {
                LineChart.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChart.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
            }
            else if (ds.Tables[0].Rows.Count >= 12 && ds.Tables[0].Rows.Count < 36) //Quaterly
            {
                LineChart.ChartAreas["ChartArea1"].AxisX.Interval = 3;
                LineChart.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Months;
                //LineChart.ChartAreas["ChartArea1"].AxisX.IntervalOffset = -1; //To start with december
                //LineChart.ChartAreas["ChartArea1"].AxisX.IntervalOffsetType = DateTimeIntervalType.Months;
            }
            else
            {
                LineChart.ChartAreas["ChartArea1"].AxisX.Interval = 1;
                LineChart.ChartAreas["ChartArea1"].AxisX.IntervalType = DateTimeIntervalType.Years;
                LineChart.ChartAreas["ChartArea1"].AxisX.IntervalOffset = -1; //To start with december
                LineChart.ChartAreas["ChartArea1"].AxisX.IntervalOffsetType = DateTimeIntervalType.Days;
            }



            LineChart.ChartAreas[0].AxisX.IsMarginVisible = false;
            LineChart.ChartAreas[0].AxisX.IsStartedFromZero = true;

            LineChart.Series[0].Name = "Gresham Advised Assets";
            LineChart.Series[1].Name = "Net Invested Capital";
            LineChart.Series[2].Name = "Inflation Adj. Net Invested Capital";

            //LineChart.Series[0].BorderColor = S;
            LineChart.Series[0].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorTIA1);
            LineChart.Series[1].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorNetInvestedCap);
            LineChart.Series[2].BorderColor = System.Drawing.ColorTranslator.FromHtml(ColorInflationAdjInvCap);

            LineChart.Series[0].Points[ds.Tables[0].Rows.Count - 1].Label = S1LastValue.ToString("C0");
            LineChart.Series[1].Points[ds.Tables[0].Rows.Count - 1].Label = S2LastValue.ToString("C0");
            LineChart.Series[2].Points[ds.Tables[0].Rows.Count - 1].Label = S3LastValue.ToString("C0");

            int S1 = 0, S2 = 0, S3 = 0;
            double MaxPoint = 0.0, MinPoint = 0.0;
            double[] values = { S1LastValue, S2LastValue, S3LastValue };
            double maxval = values.Max();
            double minval = values.Min();

            if (S1LastValue == maxval)
            {
                LineChart.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                S1 = 1;
            }
            else if (S2LastValue == maxval)
            {
                LineChart.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                S2 = 1;
            }
            else
            {
                LineChart.Series[2].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                S3 = 1;
            }

            if (S1LastValue == minval)
            {
                LineChart.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                S1 = 1;
            }
            else if (S2LastValue == minval)
            {
                LineChart.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                S2 = 1;
            }
            else
            {
                LineChart.Series[2].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
                S3 = 1;
            }

            if (S1 == 0)
            {
                double diff1 = maxval - S1LastValue;
                double diff2 = S1LastValue - minval;

                if (diff1 > diff2)
                    LineChart.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                else
                    LineChart.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            }

            if (S2 == 0)
            {
                double diff1 = maxval - S2LastValue;
                double diff2 = S2LastValue - minval;

                if (diff1 > diff2)
                    LineChart.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                else
                    LineChart.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            }

            if (S3 == 0)
            {
                double diff1 = maxval - S3LastValue;
                double diff2 = S3LastValue - minval;

                if (diff1 > diff2)
                    LineChart.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
                else
                    LineChart.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.BottomLeft;
            }

            LineChart.Series[2].SmartLabelStyle.IsMarkerOverlappingAllowed = true;
            LineChart.Series[1].SmartLabelStyle.IsMarkerOverlappingAllowed = true;
            LineChart.Series[0].SmartLabelStyle.IsMarkerOverlappingAllowed = true;


            //To Set Last DataPoint (on Lines) Position 
            //Maximum Value will come on above line
            //Minimum Value will come on below line
            //Middle Value will come accordingly Max and Min value (difference)
            //if (S1LastValue > S2LastValue && S1LastValue > S3LastValue)
            //{
            //    LineChart.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            //    S1 = 1; //To indicate s1 is max
            //    MaxPoint = S1LastValue;
            //}
            //if (S2LastValue > S1LastValue && S2LastValue > S3LastValue)
            //{
            //    LineChart.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            //    S2 = 1; //To indicate s1 is max
            //    MaxPoint = S2LastValue;
            //}

            //if (S3LastValue > S2LastValue && S3LastValue > S1LastValue)
            //{

            //    LineChart.Series[2].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            //    S3 = 1; //To indicate s1 is max
            //    MaxPoint = S3LastValue;
            //}

            //if (S1LastValue < S2LastValue && S1LastValue < S3LastValue)
            //{
            //    LineChart.Series[0].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            //    S1 = 1; //To indicate s1 is max
            //    MaxPoint = S1LastValue;
            //}
            //if (S2LastValue < S1LastValue && S2LastValue < S3LastValue)
            //{
            //    LineChart.Series[1].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            //    S2 = 1; //To indicate s1 is max
            //    MaxPoint = S2LastValue;
            //}

            //if (S3LastValue < S2LastValue && S3LastValue < S1LastValue)
            //{

            //    LineChart.Series[2].SmartLabelStyle.MovingDirection = LabelAlignmentStyles.TopLeft;
            //    S3 = 1; //To indicate s1 is max
            //    MaxPoint = S3LastValue;
            //}


            LineChart.Series[2].SmartLabelStyle.IsOverlappedHidden = false;
            LineChart.Series[1].SmartLabelStyle.IsOverlappedHidden = false;
            LineChart.Series[0].SmartLabelStyle.IsOverlappedHidden = false;



            //Remove Extra for label to display 
            LineChart.Series[0].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChart.Series[0].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChart.Series[0].SmartLabelStyle.CalloutLineWidth = 0;

            LineChart.Series[1].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChart.Series[1].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChart.Series[1].SmartLabelStyle.CalloutLineWidth = 0;

            LineChart.Series[2].SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None;
            LineChart.Series[2].SmartLabelStyle.CalloutLineColor = System.Drawing.Color.White;
            LineChart.Series[2].SmartLabelStyle.CalloutLineWidth = 0;

            LineChart.Series[0].SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No;
            LineChart.Series[1].SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No;
            LineChart.Series[2].SmartLabelStyle.AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.No;


            LineChart.Series[0].IsVisibleInLegend = true;
            LineChart.Series[1].IsVisibleInLegend = true;
            LineChart.Series[2].IsVisibleInLegend = true;

            LineChart.ChartAreas[0].AxisX.LabelStyle.Format = "MMM-yy";
            LineChart.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
        }



        System.Random rnd = new System.Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\A_" + RNum + ".bmp";

        // filename = Server.MapPath("~") + @"\\TempImages\\ChartImage-" + RNum + ".bmp";

        Bitmap bm = new Bitmap(2600, 600);

        bm.SetResolution(300, 300);

        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);

        LineChart.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 2600, 600));

        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);

        //  Chart1.SaveImage(filename, ChartImageFormat.Bmp);


        foreach (var series in LineChart.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }


        return filename;
    }

    private string getBarChartReport3(DataSet ds)
    {
        clsCombinedReports objCombinedReports = new clsCombinedReports();
        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_11122.xls";

        // JFreeChart chart = ChartFactory.createBarChart(
        DB clsDB = new DB();



        // DataSet ds = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt3BarChart));

        string Dt1;
        string sDay;
        string sMonth;
        string sYear;
        DataTable dtatble = ds.Tables[0];

        Chart1.DataSource = dtatble;
        Chart1.DataBind();

        // Set series chart type
        Chart1.Series[0].ChartType = SeriesChartType.Column;

        if (ds.Tables[0].Rows.Count > 0)
        {
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string strReturn = Convert.ToString(ds.Tables[0].Rows[i]["Return"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["Return"]);
                Double s1val = Convert.ToDouble(strReturn, System.Globalization.CultureInfo.InvariantCulture);

                Dt1 = ds.Tables[0].Rows[i]["Dates"].ToString();
                sDay = "01";
                sMonth = "01";
                sYear = ds.Tables[0].Rows[i]["Dates"].ToString();

                //chart 1
                //   s1.add(new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear)), s1val);
                // Day day = new Day(Convert.ToInt16(sDay), Convert.ToInt16(sMonth), Convert.ToInt32(sYear));
                string d = Convert.ToInt16(sDay) + "/" + Convert.ToInt16(sMonth) + "/" + Convert.ToInt32(sYear);
                Chart1.Series[0].Points.AddXY(sYear, s1val * 100);
            }
        }

        Chart1.Series[0].IsValueShownAsLabel = true;
        Chart1.Series[0].LabelFormat = "{0.0}%";
        Chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = true; //disabled inner gridlines
        Chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false; //disabled inner gridlines
        Chart1.ChartAreas[0].AxisY.MajorGrid.LineWidth = 2;

        Chart1.ChartAreas[0].AxisY.MinorGrid.Enabled = false; //disabled inner gridlines
        Chart1.ChartAreas[0].AxisX.MinorGrid.Enabled = false; //disabled inner gridlines

        Chart1.ChartAreas[0].AxisX.IsStartedFromZero = true;
        Chart1.ChartAreas[0].AxisX.IsMarginVisible = true;

        //Chart1.ChartAreas[0].AxisX.IsReversed = true;
        //clsDB.getConfiguration();

        System.Random rnd = new System.Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\A_" + RNum + ".bmp";

        // filename = Server.MapPath("~") + @"\\TempImages\\ChartImage-" + RNum + ".bmp";

        Bitmap bm = new Bitmap(2600, 600);

        bm.SetResolution(300, 300);

        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);

        Chart1.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 2600, 600));

        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);

        //  Chart1.SaveImage(filename, ChartImageFormat.Bmp);


        foreach (var series in Chart1.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }


        return filename;
    }


    private string getShapeChartReport4(string rptno, DataSet ds)
    {
        clsCombinedReports objCombinedReports = new clsCombinedReports();
        System.Random rand = new System.Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmssfff") + rand.Next().ToString();


        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\2OP_" + strGUID + ".xls";

        double Xmax = 0.0;
        double Ymax = 0.0;
        double axismax = 0.0;
        DB clsDB = new DB();


        //  DataSet ds = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt4ShapeChart));

        DataTable dt = GetFormatedDatatable(ds);

        // Populate series data with random data
        System.Random random = new System.Random();
        //for (int pointIndex = 0; pointIndex < 10; pointIndex++)
        //{
        //    ShapeChartRpt4.Series["Series1"].Points.AddY(random.Next(5, 60));
        //}

        Double s1X = 0.0; Double s1Y = 0.0; Double s2X = 0.0; Double s2Y = 0.0; Double s3X = 0.0; Double s3Y = 0.0;
        Double s4X = 0.0; Double s4Y = 0.0;

        if (Convert.ToString(dt.Rows[0]["X"]) != "")
            s1X = Convert.ToDouble(dt.Rows[0]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

        if (Convert.ToString(dt.Rows[0]["Y"]) != "")
            s1Y = Convert.ToDouble(dt.Rows[0]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

        if (Convert.ToString(dt.Rows[1]["X"]) != "")
            s2X = Convert.ToDouble(dt.Rows[1]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

        if (Convert.ToString(dt.Rows[1]["Y"]) != "")
            s2Y = Convert.ToDouble(dt.Rows[1]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

        if (Convert.ToString(dt.Rows[2]["X"]) != "")
            s3X = Convert.ToDouble(dt.Rows[2]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

        if (Convert.ToString(dt.Rows[2]["Y"]) != "")
            s3Y = Convert.ToDouble(dt.Rows[2]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

        if (Convert.ToString(dt.Rows[3]["X"]) != "")
            s4X = Convert.ToDouble(dt.Rows[3]["X"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

        if (Convert.ToString(dt.Rows[3]["Y"]) != "")
            s4Y = Convert.ToDouble(dt.Rows[3]["Y"].ToString(), System.Globalization.CultureInfo.InvariantCulture);

        double[] values = { s1X, s1Y, s2X, s2Y, s3X, s3Y, s4X, s4Y };
        double maxval = values.Max();
        maxval = Math.Ceiling(maxval);

        if (maxval % 2 != 0)
            maxval++;



        ShapeChartRpt4.Series["Series1"].Points.AddXY(s1X, s1Y);
        ShapeChartRpt4.Series["Series2"].Points.AddXY(s2X, s2Y);
        ShapeChartRpt4.Series["Series3"].Points.AddXY(s3X, s3Y);
        ShapeChartRpt4.Series["Series4"].Points.AddXY(s4X, s4Y);

        ShapeChartRpt4.Series["Series1"].Points[0].Label = dt.Rows[0]["Name"].ToString();
        ShapeChartRpt4.Series["Series2"].Points[0].Label = dt.Rows[1]["Name"].ToString();
        ShapeChartRpt4.Series["Series3"].Points[0].Label = dt.Rows[2]["Name"].ToString();
        ShapeChartRpt4.Series["Series4"].Points[0].Label = dt.Rows[3]["Name"].ToString();

        // Set point chart type
        ShapeChartRpt4.Series["Series1"].ChartType = SeriesChartType.Point;
        ShapeChartRpt4.Series["Series2"].ChartType = SeriesChartType.Point;
        ShapeChartRpt4.Series["Series3"].ChartType = SeriesChartType.Point;
        ShapeChartRpt4.Series["Series4"].ChartType = SeriesChartType.Point;


        // Enable data points labels
        // ShapeChartRpt4.Series["Series1"].IsValueShownAsLabel = true;
        //  ShapeChartRpt4.Series["Series1"]["LabelStyle"] = "Center";

        // Set marker size
        ShapeChartRpt4.Series["Series1"].MarkerSize = 10;
        ShapeChartRpt4.Series["Series2"].MarkerSize = 10;
        ShapeChartRpt4.Series["Series3"].MarkerSize = 10;
        ShapeChartRpt4.Series["Series4"].MarkerSize = 10;

        ShapeChartRpt4.ChartAreas[0].Position.X = 0;
        ShapeChartRpt4.ChartAreas[0].Position.Y = 2;
        ShapeChartRpt4.ChartAreas[0].Position.Height = 95;
        ShapeChartRpt4.ChartAreas[0].Position.Width = 95;

        ShapeChartRpt4.ChartAreas[0].AxisX.Maximum = maxval;
        ShapeChartRpt4.ChartAreas[0].AxisY.Maximum = maxval;
        // ShapeChartRpt4.ChartAreas[0].InnerPlotPosition.Height = 80;
        // ShapeChartRpt4.ChartAreas[0].InnerPlotPosition.Width = 80;

        // Set marker shape
        ShapeChartRpt4.Series["Series1"].MarkerStyle = MarkerStyle.Square; //Total GAA
        ShapeChartRpt4.Series["Series2"].MarkerStyle = MarkerStyle.Circle; //Marketable GAA
        ShapeChartRpt4.Series["Series3"].MarkerStyle = MarkerStyle.Diamond; //Strategic Benchmark
        ShapeChartRpt4.Series["Series4"].MarkerStyle = MarkerStyle.Triangle; //MSCI

        // Set marker color 
        ShapeChartRpt4.Series["Series1"].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#548ACF");
        ShapeChartRpt4.Series["Series2"].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#8064A2");
        ShapeChartRpt4.Series["Series3"].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#B7DEE8");
        ShapeChartRpt4.Series["Series4"].MarkerColor = System.Drawing.ColorTranslator.FromHtml("#17375E");

        // Set marker border -  Strategic Benchmark
        ShapeChartRpt4.Series["Series3"].MarkerBorderWidth = 1;
        ShapeChartRpt4.Series["Series3"].MarkerBorderColor = System.Drawing.ColorTranslator.FromHtml("#215968");

        ShapeChartRpt4.ChartAreas["ChartArea1"].AxisX.Interval = 2;
        ShapeChartRpt4.ChartAreas["ChartArea1"].AxisY.Interval = 2;
        ShapeChartRpt4.ChartAreas["ChartArea1"].AxisX.Minimum = 0.0;

        ShapeChartRpt4.ChartAreas["ChartArea1"].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8, FontStyle.Regular);
        ShapeChartRpt4.ChartAreas["ChartArea1"].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8, FontStyle.Regular);

        ShapeChartRpt4.Titles[0].Font = new System.Drawing.Font("Frutiger65", 9, FontStyle.Bold);
        ShapeChartRpt4.Titles[0].Docking = Docking.Top;
        ShapeChartRpt4.Titles[0].DockingOffset = -2;

        // Enable 3D
        // ShapeChartRpt4.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = true;

        System.Random rnd = new System.Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\OP_" + RNum + ".bmp";

        Bitmap bm = new Bitmap(1280, 1000);
        bm.SetResolution(300, 300);
        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);
        ShapeChartRpt4.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 1280, 1000));
        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);

        foreach (var series in ShapeChartRpt4.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }

        return filename;


    }

    private string getColumnChartReport4(string Fname, DataSet ds)
    {
        clsCombinedReports objCombinedReports = new clsCombinedReports();
        System.Random rand = new System.Random();
        string strGUID = System.DateTime.Now.ToString("MMddyyHHmmssfff") + rand.Next().ToString();

        String fsFinalLocation = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\CC_" + strGUID + ".xls";

        // JFreeChart chart = ChartFactory.createBarChart(
        DB clsDB = new DB();

        //DataSet ds = clsDB.getDataSet(objCombinedReports.getFinalSp(clsCombinedReports.ReportType.Rpt4ColumnChart));

        string Dt1;
        string sDay;
        string sMonth;
        string sYear;

        DataTable dtatble = ds.Tables[0];

        ColumnChartRpt4.DataSource = dtatble;
        ColumnChartRpt4.DataBind();

        // Set series chart type
        ColumnChartRpt4.Series[0].ChartType = SeriesChartType.Column;
        ColumnChartRpt4.Series[1].ChartType = SeriesChartType.RangeColumn;
        if (ds.Tables[0].Rows.Count > 0)
        {
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                string strVal1 = Convert.ToString(ds.Tables[0].Rows[i]["Honore"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["Honore"]);
                string strVal2 = Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[i]["MSCI AC World Index"]);
                Double s1val = Convert.ToDouble(strVal1, System.Globalization.CultureInfo.InvariantCulture);
                Double s2val = Convert.ToDouble(strVal2, System.Globalization.CultureInfo.InvariantCulture);

                string strVal3 = Convert.ToString(ds.Tables[0].Rows[0]["Honor Max Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["Honor Max Drawdown"]);
                string strVal4 = Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index Drawdown"]);
                Double s3val = Convert.ToDouble(strVal3, System.Globalization.CultureInfo.InvariantCulture);
                Double s4val = Convert.ToDouble(strVal4, System.Globalization.CultureInfo.InvariantCulture);

                //decimal ss = 12.5m;
                //ss = Math.Round(ss, 0, MidpointRounding.AwayFromZero);
                s1val = Math.Round(s1val * 100, 0, MidpointRounding.AwayFromZero);
                s2val = Math.Round(s2val * 100, 0, MidpointRounding.AwayFromZero);

                s3val = Math.Round(s3val * 100, 0, MidpointRounding.AwayFromZero);
                s4val = Math.Round(s4val * 100, 0, MidpointRounding.AwayFromZero);

                string sDate = ds.Tables[0].Rows[i]["Month"].ToString();

                //bardataset.setValue(s1val, "1", sDate);
                //bardataset.setValue(s2val, "2", sDate);


                ColumnChartRpt4.Series[0].Points.AddXY(sDate, s1val);
                ColumnChartRpt4.Series[1].Points.AddXY(sDate, s2val);
                ColumnChartRpt4.Series[2].Points.AddXY(sDate, s2val);

                //To Set axis max-min value
                if (s1val <= min1)
                    min1 = s1val;
                if (s2val <= min1)
                    min1 = s2val;
                if (s3val <= min1)
                    min1 = s3val;
                if (s4val <= min1)
                    min1 = s4val;

                if (s1val >= max1)
                    max1 = s1val;
                if (s2val >= max1)
                    max1 = s2val;
                if (s3val >= max1)
                    max1 = s1val;
                if (s4val >= max1)
                    max1 = s4val;

            }

            string _strVal3 = Convert.ToString(ds.Tables[0].Rows[0]["Honor Max Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["Honor Max Drawdown"]);
            string _strVal4 = Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index Drawdown"]) == "" ? "0" : Convert.ToString(ds.Tables[0].Rows[0]["MSCI AC World Index Drawdown"]);
            Double _s3val = Convert.ToDouble(_strVal3, System.Globalization.CultureInfo.InvariantCulture);
            Double _s4val = Convert.ToDouble(_strVal4, System.Globalization.CultureInfo.InvariantCulture);

            _s3val = Math.Round(_s3val * 100, 0, MidpointRounding.AwayFromZero);
            _s4val = Math.Round(_s4val * 100, 0, MidpointRounding.AwayFromZero);

            string _sDate = ds.Tables[0].Rows[0]["MinDate"].ToString();

            //   bardataset1.setValue(s3val, "Marks", ".");
            //  bardataset1.setValue(s4val, "Marks1", ".");

            ColumnChartRpt4.Series[0].Points.AddXY(" ", _s3val);
            ColumnChartRpt4.Series[1].Points.AddXY(" ", _s4val);
            ColumnChartRpt4.Series[2].Points.AddXY(" ", _s4val);

        }


        ColumnChartRpt4.Series[0].IsValueShownAsLabel = true;
        ColumnChartRpt4.Series[0].LabelFormat = "{0}%";
        ColumnChartRpt4.Series[0].Font = new System.Drawing.Font("Frutiger55", 6F, System.Drawing.FontStyle.Regular);
        ColumnChartRpt4.Series[2].Font = new System.Drawing.Font("Frutiger55", 6F, System.Drawing.FontStyle.Regular);

        ColumnChartRpt4.Series[1].IsValueShownAsLabel = false;


        ColumnChartRpt4.Series[2].IsValueShownAsLabel = true;
        ColumnChartRpt4.Series[2].LabelFormat = "{0}%";


        ColumnChartRpt4.ChartAreas[0].AxisY.MajorGrid.Enabled = false; //disabled inner gridlines
        ColumnChartRpt4.ChartAreas[0].AxisX.MajorGrid.Enabled = false; //disabled inner gridlines
        // ColumnChartRpt4.ChartAreas[0].AxisY.MajorGrid.LineWidth = 2;

        ColumnChartRpt4.ChartAreas[0].AxisY.MinorGrid.Enabled = false; //disabled inner gridlines
        ColumnChartRpt4.ChartAreas[0].AxisX.MinorGrid.Enabled = false; //disabled inner gridlines

        ColumnChartRpt4.ChartAreas[0].AxisY2.MajorGrid.Enabled = false; //disabled inner gridlines
        ColumnChartRpt4.ChartAreas[0].AxisX2.MajorGrid.Enabled = false; //disabled inner gridlines
        // ColumnChartRpt4.ChartAreas[0].AxisY.MajorGrid.LineWidth = 2;

        ColumnChartRpt4.ChartAreas[0].AxisY2.MinorGrid.Enabled = false; //disabled inner gridlines
        ColumnChartRpt4.ChartAreas[0].AxisX2.MinorGrid.Enabled = false; //disabled inner gridlines

        ColumnChartRpt4.ChartAreas[0].AxisX.Enabled = AxisEnabled.False;

        ColumnChartRpt4.ChartAreas[0].AxisX.IsStartedFromZero = true;
        ColumnChartRpt4.ChartAreas[0].AxisX2.IsStartedFromZero = true;
        ColumnChartRpt4.ChartAreas[0].AxisX2.Title = "Return %";
        ColumnChartRpt4.ChartAreas[0].AxisX2.TitleFont = new System.Drawing.Font("Frutiger55", 7F, System.Drawing.FontStyle.Bold);
        // ColumnChartRpt4.ChartAreas[0].AxisX2.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8F, System.Drawing.FontStyle.Bold);


        ColumnChartRpt4.ChartAreas[0].AxisY2.LabelStyle.Format = "{N0}%";
        ColumnChartRpt4.ChartAreas[0].AxisY.LabelStyle.Format = "{N0}%";

        ColumnChartRpt4.ChartAreas[0].AxisX.IsMarginVisible = false;
        ColumnChartRpt4.ChartAreas[0].AxisX2.IsMarginVisible = false;

        //ColumnChartRpt4.ChartAreas[0].AxisY.IsMarginVisible = false;
        //ColumnChartRpt4.ChartAreas[0].AxisY2.IsMarginVisible = false;

        // ColumnChartRpt4.Series[0]["PointWidth"] = "0.5";
        //// ColumnChartRpt4.Series[1]["PointWidth"] = "1";
        // ColumnChartRpt4.Series[1]["PointWidth"] = "0.5";
        // ColumnChartRpt4.Series[2]["PointWidth"] = "0.5";

        ColumnChartRpt4.Series[0]["PixelPointWidth"] = "190";
        ColumnChartRpt4.Series[1]["PixelPointWidth"] = "25";

        //  ColumnChartRpt4.Series[1]["PixelPointWidth"].PadRight(0);
        ColumnChartRpt4.Series[2]["PixelPointWidth"] = "190";

        ColumnChartRpt4.Series[0].Name = Fname;
        ColumnChartRpt4.Series[2].Name = Convert.ToString(ds.Tables[0].Rows[0]["BenchMarkName"]);
        ColumnChartRpt4.Series[1].IsVisibleInLegend = false;

        //  ColumnChartRpt4.Series[0].BorderWidth = 5;
        // ColumnChartRpt4.Series[0].BorderColor = System.Drawing.Color.Aqua;


        ColumnChartRpt4.Series[0].XAxisType = System.Web.UI.DataVisualization.Charting.AxisType.Primary;
        ColumnChartRpt4.Series[0].YAxisType = System.Web.UI.DataVisualization.Charting.AxisType.Primary;

        //vertical line above x axis 
        VerticalLineAnnotation annotation = new VerticalLineAnnotation();
        annotation.AnchorDataPoint = ColumnChartRpt4.Series[0].Points[3];
        annotation.AxisX = ColumnChartRpt4.ChartAreas[0].AxisX;
        annotation.AxisY = ColumnChartRpt4.ChartAreas[0].AxisY;
        annotation.AnchorY = -0;
        annotation.AnchorOffsetX = -1;
        annotation.X = 3.5;
        annotation.Height = -8;
        annotation.LineWidth = 2;
        annotation.StartCap = LineAnchorCapStyle.None;
        annotation.EndCap = LineAnchorCapStyle.None;
        annotation.LineDashStyle = ChartDashStyle.Dash;
        annotation.IsInfinitive = false;

        //vertical line below x axis 
        VerticalLineAnnotation annotation1 = new VerticalLineAnnotation();
        annotation1.AnchorDataPoint = ColumnChartRpt4.Series[0].Points[3];
        annotation1.AxisX = ColumnChartRpt4.ChartAreas[0].AxisX;
        annotation1.AxisY = ColumnChartRpt4.ChartAreas[0].AxisY;
        annotation1.AnchorY = -0;
        annotation1.AnchorOffsetX = -1;
        annotation1.X = 3.5;
        annotation1.Height = 65;
        annotation1.LineWidth = 2;
        annotation1.StartCap = LineAnchorCapStyle.None;
        annotation1.EndCap = LineAnchorCapStyle.None;
        annotation1.LineDashStyle = ChartDashStyle.Dash;
        annotation1.IsInfinitive = false;


        System.Web.UI.DataVisualization.Charting.TextAnnotation annotation2 =
          new System.Web.UI.DataVisualization.Charting.TextAnnotation();
        annotation2.Text = "Worst Market Period " + Convert.ToString(ds.Tables[0].Rows[0]["MinDate"]);
        annotation2.X = 71;
        annotation2.Y = 4;
        annotation2.IsMultiline = true;
        annotation2.Width = 30;
        annotation2.Height = 5;
        annotation2.Font = new System.Drawing.Font("Frutiger55", 7, System.Drawing.FontStyle.Bold);
        annotation2.ForeColor = System.Drawing.Color.Black;
        ColumnChartRpt4.Annotations.Add(annotation2);

        ColumnChartRpt4.Annotations.Add(annotation);
        ColumnChartRpt4.Annotations.Add(annotation1);

        //  ColumnChartRpt4.Series[0].YAxisType = AxisType.Secondary;

        //   ColumnChartRpt4.Series[1].XAxisType = AxisType.Primary;
        //  ColumnChartRpt4.Series[1].YAxisType = AxisType.Primary;
        ColumnChartRpt4.Series[1].XAxisType = System.Web.UI.DataVisualization.Charting.AxisType.Secondary;
        ColumnChartRpt4.Series[1].YAxisType = System.Web.UI.DataVisualization.Charting.AxisType.Secondary;


        ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 6, FontStyle.Regular);
        ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY2.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 6, FontStyle.Regular);
        ColumnChartRpt4.ChartAreas["ChartArea1"].AxisX.LabelStyle.Font = new System.Drawing.Font("Frutiger55", 8, FontStyle.Bold);



        if (min1 <= 0.0 && min1 >= -10.00)
        {
            ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY.Interval = 5;
            ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY2.Interval = 5;
            min1 = min1 - 5.0;
        }
        else if (max1 > 0.0 && max1 <= 10.00)
        {
            ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY.Interval = 5;
            ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY2.Interval = 5;
        }
        else
        {
            ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY.Interval = 10;
            ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY2.Interval = 10;

        }

        if (min1 < -10.0)
            min1 = min1 - 10.0;

        if (max1 <= 0.0)
        {
            max1 = 0.0;
            ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY.IsStartedFromZero = true;
            ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY2.IsStartedFromZero = true;
        }


        //   ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY.Minimum = min1;
        //   ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY2.Minimum = min1;

        ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY.Maximum = max1;
        ColumnChartRpt4.ChartAreas["ChartArea1"].AxisY2.Maximum = max1;

        ColumnChartRpt4.ChartAreas[0].Position.X = 5;
        ColumnChartRpt4.ChartAreas[0].Position.Y = 0;
        ColumnChartRpt4.ChartAreas[0].Position.Height = 93;
        ColumnChartRpt4.ChartAreas[0].Position.Width = 100;

        ColumnChartRpt4.ChartAreas[0].AxisX2.LabelStyle.Angle = 0;
        ColumnChartRpt4.ChartAreas[0].AxisX2.IsLabelAutoFit = false;

        ColumnChartRpt4.Legends["Legend1"].Position.Auto = false;
        ColumnChartRpt4.Legends["Legend1"].Position = new ElementPosition(2, 92, 100, 8);

        ColumnChartRpt4.Titles[0].Font = new System.Drawing.Font("Frutiger65", 9, FontStyle.Bold);
        ColumnChartRpt4.Titles[0].Docking = Docking.Top;
        ColumnChartRpt4.Titles[0].DockingOffset = -2;

        //ColumnChartRpt4.Series[1].YAxisType = AxisType.Secondary;

        //   ColumnChartRpt4.ChartAreas[0].AxisY.IsReversed = true;
        //    ColumnChartRpt4.ChartAreas[0].AxisY2.IsReversed = true;
        // ColumnChartRpt4.ChartAreas[0].AxisY.IsReversed = true;
        // ColumnChartRpt4.ChartAreas[0].AxisY2.IsReversed = true;
        // ColumnChartRpt4.ChartAreas[0].AxisX.IsReversed = true;
        // ColumnChartRpt4.ChartAreas[0].AxisX2.IsReversed = true;
        //clsDB.getConfiguration();

        System.Random rnd = new System.Random();
        string RNum = Convert.ToString(rnd.Next(999999999));

        string filename = HttpContext.Current.Server.MapPath("") + @"\ExcelTemplate\pdfOutput\A_" + RNum + ".bmp";

        // filename = Server.MapPath("~") + @"\\TempImages\\ChartImage-" + RNum + ".bmp";

        Bitmap bm = new Bitmap(1280, 1000);

        bm.SetResolution(300, 300);

        System.Drawing.Graphics gGraphics = System.Drawing.Graphics.FromImage(bm);

        ColumnChartRpt4.Paint(gGraphics, new System.Drawing.Rectangle(0, 0, 1280, 1000));

        bm.Save(filename, System.Drawing.Imaging.ImageFormat.Bmp);

        //  ColumnChartRpt4.SaveImage(filename, ChartImageFormat.Bmp);


        foreach (var series in ColumnChartRpt4.Series) //clear all points to reuse chart for multiple records
        {
            series.Points.Clear();
        }


        return filename;
    }

    public DataTable GetFormatedDatatable(DataSet ds)
    {
        DataTable dt = new DataTable();
        dt.Clear();

        dt.Columns.Add("Name");
        dt.Columns.Add("X");
        dt.Columns.Add("Y");

        DataRow dr1 = dt.NewRow();
        dr1["Name"] = "Total GAA";
        dr1["X"] = Convert.ToString(ds.Tables[0].Rows[1]["Return"]);
        dr1["Y"] = Convert.ToString(ds.Tables[0].Rows[0]["Return"]);

        DataRow dr2 = dt.NewRow();
        dr2["Name"] = "Marketable GAA";
        dr2["X"] = Convert.ToString(ds.Tables[0].Rows[3]["Return"]);
        dr2["Y"] = Convert.ToString(ds.Tables[0].Rows[2]["Return"]);

        DataRow dr3 = dt.NewRow();
        dr3["Name"] = "Weighted Benchmark";
        dr3["X"] = Convert.ToString(ds.Tables[0].Rows[5]["Return"]);
        dr3["Y"] = Convert.ToString(ds.Tables[0].Rows[4]["Return"]);

        DataRow dr4 = dt.NewRow();
        dr4["Name"] = ds.Tables[0].Rows[8]["ReturnName"].ToString();
        dr4["X"] = Convert.ToString(ds.Tables[0].Rows[7]["Return"]);
        dr4["Y"] = Convert.ToString(ds.Tables[0].Rows[6]["Return"]);

        dt.Rows.Add(dr1);
        dt.Rows.Add(dr2);
        dt.Rows.Add(dr3);
        dt.Rows.Add(dr4);

        return dt;

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



    protected void ddlHouseHoldType_SelectedIndexChanged(object sender, EventArgs e)
    {
        fillHousehold();
        lblError.Text = "";
        chkNoComparison.Checked = false;
        chkConvertToAssetDistComp.Checked = false;
        chkSuppressManagerDetail.Checked = false;
        txtPriorDate.Enabled = true;
        img1.Disabled = false;
        img1.Visible = true;

        if (ddlHouseHold.SelectedValue != "0")
        {
            BindHHReportType();
            BindHHRP();
        }
        else
        {
            ddlReportType.Items.Clear();
            ddlReportType.Items.Add(new System.Web.UI.WebControls.ListItem("All", "0"));

            ddlHHRP.Items.Clear();
            ddlHHRP.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", "0"));
        }


    }
}

