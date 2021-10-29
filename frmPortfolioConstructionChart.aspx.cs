using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System;
using System.Data;
using System.IO;
using Microsoft.IdentityModel.Claims;
using System.Threading;

public partial class frmPortfolioConstructionChart : System.Web.UI.Page
{
    public int liPageSize = 29;
    public string Text = "";
    public string lsTotalNumberofColumns, lsDistributionName, lsFamiliesName, lsDateName, FooterText, AsOfDate, GreshamAdvisedFlag;


    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            fillHousehold();
            if (Request.QueryString.Count > 0)
            {
                if (Request.QueryString["type"] == "new")
                {
                    trRollupgrp.Style.Add("display", "");
                    trGAFlag.Style.Add("display", "");
                    trHouseHoldRptTitle.Style.Add("display", "");
                }
            }
            else
            {
                lblNote.Text = "";
                trRollupgrp.Style.Add("display", "none");
                trGAFlag.Style.Add("display", "none");
                trHouseHoldRptTitle.Style.Add("display", "none");
            }
        }
    }


    public void fillHousehold()
    {
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("sp_s_Get_HouseHoldName");
        ddlHousehold.Items.Clear();
        ddlHousehold.Items.Add(new System.Web.UI.WebControls.ListItem("Please Select", ""));
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlHousehold.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][1]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
        }
    }
    public void fillHouseholdTitle()
    {
        DB clsDB = new DB();
        drpHouseHoldReportTitle.Items.Clear();
        if (!String.IsNullOrEmpty(ddlHousehold.SelectedValue))
        {
            drpHouseHoldReportTitle.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
            //    Response.Write(" SP_S_HouseHoldTitle @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'","''") + "'");
            DataSet loDataset = clsDB.getDataSet(" SP_S_HouseHoldTitle @Flag=1,@HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", "''") + "'");
            for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
            {
                drpHouseHoldReportTitle.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter][0]), Convert.ToString(loDataset.Tables[0].Rows[liCounter][0])));
            }
        }

    }

    protected void ddlHousehold_SelectedIndexChanged(object sender, EventArgs e)
    {
        FillGroup();
        FillAllocationGroup();
        FillReportRollUpGroup();
        fillHouseholdTitle();
    }

    protected void ddlAllocationGroup_SelectedIndexChanged(object sender, EventArgs e)
    {
        //FillReportRollUpGroup();
        //fillGroupAllocationTitle();
    }

    public void FillGroup()
    {
        DB clsDB = new DB();
        ddlGroup.Items.Clear();
        ddlGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        DataSet loDataset = clsDB.getDataSet("SP_S_HOUSEHOLD_GROUPNAME  @HouseHoldName ='" + ddlHousehold.SelectedItem.Text.Replace("'", " ") + "'");
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter]["sas_name"]), Convert.ToString(loDataset.Tables[0].Rows[liCounter]["sas_name"])));
        }

    }

    public void FillAllocationGroup()
    {
        DB clsDB = new DB();
        ddlAllocationGroup.Items.Clear();
        ddlAllocationGroup.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        DataSet loDataset = clsDB.getDataSet("SP_S_Advent_Allocation_Group  @Householdname ='" + ddlHousehold.SelectedItem.Text.Replace("'", " ") + "'");
        for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
        {
            ddlAllocationGroup.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter]["AllocationGroupName"]), Convert.ToString(loDataset.Tables[0].Rows[liCounter]["AllocationGroupName"])));
        }

    }
    public void FillReportRollUpGroup()
    {
        string HHUID = ddlHousehold.SelectedValue == "" ? "" : "'" + ddlHousehold.SelectedValue + "'";
        DB clsDB = new DB();
        ddlReportRollupgrp.Items.Clear();
        ddlReportRollupgrp.Items.Add(new System.Web.UI.WebControls.ListItem("Select", "0"));
        if (HHUID != "")
        {
            DataSet loDataset = clsDB.getDataSet("SP_S_GROUPNAME  @HHUUID =" + HHUID + "");
            for (int liCounter = 0; liCounter < loDataset.Tables[0].Rows.Count; liCounter++)
            {
                ddlReportRollupgrp.Items.Add(new System.Web.UI.WebControls.ListItem(Convert.ToString(loDataset.Tables[0].Rows[liCounter]["GroupName"]), Convert.ToString(loDataset.Tables[0].Rows[liCounter]["sas_reportrollupgroupid"])));
            }
        }
    }

    protected void Button1_Click(object sender, EventArgs e)
    {

    }

    public void GetPDF()
    {
        string ReportOpFolder = string.Empty;
        string ContactFolderName = string.Empty;
        string ParentFolder = string.Empty;
        string TempFolderPath = string.Empty;
        try
        { 
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

        strUserName = strUserName.Substring(strUserName.IndexOf("\\") + 1);

        ParentFolder = strUserName + "_" + strYear + strMonth + strDay + "_" + strHour + strMinute + strSecond + strMilliSecond;

        ReportOpFolder = Request.MapPath("ExcelTemplate\\TempFolder\\");

        bool isExist = System.IO.Directory.Exists(ReportOpFolder + "\\" + ParentFolder);
        TempFolderPath = ReportOpFolder + ParentFolder;
        if (!isExist)
        {
            //  Response.Write("Folder: " + ReportOpFolder + "\\" + ContactFolderName);
            System.IO.Directory.CreateDirectory(ReportOpFolder + "\\" + ParentFolder);
        }


        string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");

        // String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";
        String fsFinalLocation = ReportOpFolder + strGUID + ".xls";

        clsCombinedReports objCombinedReports = new clsCombinedReports();

        string lsfamilyName = "";
        if (ddlAllocationGroup.SelectedValue != "0")
        {
            lsfamilyName = GetAllocationGrpTitle();
        }
        else if (ddlReportRollupgrp.SelectedValue != "0" && drpHouseHoldReportTitle.SelectedValue == "0")
        {
            lsfamilyName = ddlReportRollupgrp.SelectedItem.Text;
        }
        else if (ddlReportRollupgrp.SelectedValue != "0" && drpHouseHoldReportTitle.SelectedValue != "0")
        {
            lsfamilyName = drpHouseHoldReportTitle.SelectedItem.Text;
        }
        else
        {
            if (drpHouseHoldReportTitle.SelectedValue != "0")
                lsfamilyName = drpHouseHoldReportTitle.SelectedItem.Text;
            else
                lsfamilyName = ddlHousehold.SelectedItem.Text;
        }

        objCombinedReports.HouseHoldValue = "";
        objCombinedReports.HouseHoldText = ddlHousehold.SelectedValue == "0" ? "" : ddlHousehold.SelectedItem.Text.Replace("'", "''");
        objCombinedReports.AllocationGroupValue = "";
        objCombinedReports.AllocationGroupText = ddlAllocationGroup.SelectedValue == "0" ? "" : ddlAllocationGroup.SelectedItem.Text.Replace("'", "''");
        objCombinedReports.AsOfDate = txtAsofdate.Text.Trim() == "" ? "null" : txtAsofdate.Text.Trim();
        // objCombinedReports.lsFamiliesName = lsfamilyName;
        objCombinedReports.CommitmentReportHeader = lsfamilyName;

        objCombinedReports.FooterText = "";
        objCombinedReports.ReportRollUpGroupValue = ddlReportRollupgrp.SelectedValue == "0" ? "" : ddlReportRollupgrp.SelectedValue;
        objCombinedReports.GreshamAdvisedFlag = ddlGreshamAdvisedFlg.SelectedItem.Text;
        objCombinedReports.TempFolderPath = TempFolderPath;
        if (Request.QueryString.Count > 0)
        {
            if (Request.QueryString["type"] == "new")
                objCombinedReports.PortFolioConChartRptVer = "";
        }
        else
            objCombinedReports.PortFolioConChartRptVer = "old";

        string filepdfname = objCombinedReports.MergeReports(fsFinalLocation, "Portfolio Construction Chart v2.1");

        FileInfo loFile = new FileInfo(filepdfname);

        loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

        //Response.Write("<script>");
        //string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
        //Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
        //Response.Write("</script>");
        Response.Write("<script>");
        Response.Write("window.open('ViewReport.aspx?" + strGUID + ".pdf" + "', 'mywindow')");
        Response.Write("</script>");
    }
        catch (Exception ex)
        {
        }
        finally
        {
            if (Directory.Exists(ReportOpFolder + "\\" + ParentFolder))
            {
                Directory.Delete(ReportOpFolder + "\\" + ParentFolder, true);
            }
        }
    }

    private string GetAllocationGrpTitle()
    {
        string val = string.Empty;
        DB clsDB = new DB();
        DataSet loDataset = clsDB.getDataSet("SP_S_Advent_Allocation_Group  @Householdname ='" + ddlHousehold.SelectedItem.Text.Replace("'", " ") + "'");
        DataTable dt = loDataset.Tables[0];
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            string allocationgrpName = Convert.ToString(dt.Rows[i]["AllocationGroupName"]);
            if (allocationgrpName == ddlAllocationGroup.SelectedValue)
            {
                val = Convert.ToString(dt.Rows[i]["AllocationGroupTitle"]);
            }
        }
        return val;
    }
    protected void btnSubmit_Click(object sender, System.EventArgs e)
    {
        if (HttpContext.Current.Request.Url.AbsoluteUri.Contains("localhost"))
        {
            GetPDF();
            /*
            string strGUID = System.DateTime.Now.ToString("MMddyyhhmmss");
            String fsFinalLocation = Server.MapPath("") + @"\ExcelTemplate\pdfOutput\" + strGUID + ".xls";
            AsOfDate = txtAsofdate.Text.Trim() == "" ? "null" : txtAsofdate.Text.Trim();
            GreshamAdvisedFlag = ddlGreshamAdvisedFlg.SelectedItem.Text;
            lsFamiliesName = ddlHousehold.SelectedItem.Text;
            FooterText = "";

            DataSet newdataset;
            DB clsDB = new DB();
            newdataset = null;
            String lsFooterTxt = String.Empty;
            String lsSQL = getFinalSp();

            newdataset = clsDB.getDataSet(lsSQL);

            string filepdfname = generatePDFFinal_New(newdataset);

            FileInfo loFile = new FileInfo(filepdfname);

            loFile.MoveTo(fsFinalLocation.Replace(".xls", ".pdf"));

            Response.Write("<script>");
            string lsFileNamforFinalXls = "./ExcelTemplate/pdfOutput/" + strGUID + ".pdf";
            Response.Write("window.open('" + lsFileNamforFinalXls + "', 'mywindow')");
            Response.Write("</script>");
            //generatePDFFinal();
            */
        }
        else
            GetPDF();

    }
}